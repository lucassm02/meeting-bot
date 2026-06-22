import { StorageProvider, UploadOptions } from './storage-provider';
import config from '../../config';
import { promises as fs, mkdirSync, writeFileSync, unlinkSync } from 'fs';
import path from 'path';

export class LocalStorageProvider implements StorageProvider {
  readonly name = 'local' as const;

  private getBaseDir(): string {
    const baseDir = config.localStorage.baseDir;
    if (!baseDir || !baseDir.trim()) {
      throw new Error('Local storage configuration is not set. Missing: LOCAL_STORAGE_DIR');
    }
    return path.resolve(baseDir);
  }

  /**
   * Resolve a storage key to an absolute path inside baseDir and guarantee it
   * stays contained. The key embeds request-controlled values (userId, namePrefix),
   * so this prevents path traversal (e.g. userId="../../etc/...") from writing
   * outside baseDir. See spec 001 risk R4.
   */
  private resolveSafePath(baseDir: string, key: string): string {
    const dest = path.resolve(baseDir, key);
    const root = baseDir.endsWith(path.sep) ? baseDir : baseDir + path.sep;
    if (dest !== baseDir && !dest.startsWith(root)) {
      throw new Error(`Refusing to write outside local storage directory (key resolves outside baseDir): ${key}`);
    }
    return dest;
  }

  validateConfig(): void {
    const baseDir = this.getBaseDir();
    // Verify we can create and write inside baseDir, failing fast with a clear error.
    // (mkdir + write/remove a probe file synchronously is cheap and surfaces
    // permission / missing-mount problems before a recording is attempted.)
    const probe = path.join(baseDir, `.write-test-${process.pid}`);
    try {
      mkdirSync(baseDir, { recursive: true });
      writeFileSync(probe, '');
      unlinkSync(probe);
    } catch (err) {
      throw new Error(`Local storage directory is not writable: ${baseDir} (${(err as Error).message})`);
    }
  }

  async uploadFile(options: UploadOptions): Promise<boolean> {
    const baseDir = this.getBaseDir();
    const dest = this.resolveSafePath(baseDir, options.key);

    try {
      options.logger.info(`Starting local storage write for ${options.key}`);
      await fs.mkdir(path.dirname(dest), { recursive: true });

      // Copy (not move): the caller deletes the temp file after upload and may retry,
      // so the source must remain intact. See spec 001 risk R2.
      await fs.copyFile(options.filePath, dest);

      // Write a sidecar metadata file for parity with object-storage metadata.
      if (options.metadata && Object.keys(options.metadata).length > 0) {
        const metaPath = `${dest}.meta.json`;
        await fs.writeFile(metaPath, JSON.stringify(options.metadata, null, 2));
      }

      options.logger.info(`Local storage write complete for ${dest}`);
      return true;
    } catch (err) {
      options.logger.error(`Local storage write failed for ${options.key}`, err as Error);
      return false;
    }
  }

  async getSignedUrl(key: string): Promise<string> {
    const baseDir = this.getBaseDir();
    const dest = this.resolveSafePath(baseDir, key);
    const publicBaseUrl = config.localStorage.publicBaseUrl;
    if (publicBaseUrl) {
      return `${publicBaseUrl.replace(/\/$/, '')}/${encodeURI(key)}`;
    }
    return `file://${dest}`;
  }

  async exists(key: string): Promise<boolean> {
    const baseDir = this.getBaseDir();
    const dest = this.resolveSafePath(baseDir, key);
    try {
      await fs.access(dest);
      return true;
    } catch {
      return false;
    }
  }

  async delete(key: string): Promise<void> {
    const baseDir = this.getBaseDir();
    const dest = this.resolveSafePath(baseDir, key);
    await fs.rm(dest, { force: true });
    await fs.rm(`${dest}.meta.json`, { force: true });
  }

  async list(prefix: string): Promise<string[]> {
    const baseDir = this.getBaseDir();
    const searchDir = this.resolveSafePath(baseDir, prefix);
    const results: string[] = [];

    const walk = async (dir: string): Promise<void> => {
      let entries: import('fs').Dirent[];
      try {
        entries = await fs.readdir(dir, { withFileTypes: true });
      } catch {
        return;
      }
      for (const entry of entries) {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) {
          await walk(full);
        } else if (!entry.name.endsWith('.meta.json')) {
          results.push(path.relative(baseDir, full));
        }
      }
    };

    await walk(searchDir);
    return results;
  }
}
