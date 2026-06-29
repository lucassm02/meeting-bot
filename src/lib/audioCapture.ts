// One-shot capture of a short system-audio clip, used to grab the
// Teams audio CAPTCHA. Records the PulseAudio monitor (same device the main
// ffmpeg recorder uses) for a fixed duration and returns it as base64 mp3.
import { spawn } from 'child_process';
import { mkdtemp, readFile, rm } from 'fs/promises';
import { tmpdir } from 'os';
import { join } from 'path';
import { Logger } from 'winston';

export interface CapturedAudio {
  base64: string;
  mimeType: string;
}

/**
 * Captures `seconds` of audio from the PulseAudio monitor into a temporary mp3
 * and returns it base64-encoded. Resolves null if ffmpeg fails or yields no audio.
 */
export async function captureSystemAudioClip(seconds: number, logger: Logger): Promise<CapturedAudio | null> {
  const dir = await mkdtemp(join(tmpdir(), 'captcha-audio-'));
  const outputPath = join(dir, 'captcha.mp3');

  const ffmpegEnv = {
    ...process.env,
    XDG_RUNTIME_DIR: process.env.XDG_RUNTIME_DIR || '/run/user/1001',
    DISPLAY: process.env.DISPLAY || ':99',
  };

  const args = [
    '-y',
    '-loglevel', 'error',
    '-f', 'pulse',
    '-ac', '2',
    '-ar', '44100',
    '-i', 'virtual_output.monitor',
    '-t', String(Math.max(1, Math.floor(seconds))),
    '-ac', '1',
    '-ar', '16000',
    '-c:a', 'libmp3lame',
    '-b:a', '64k',
    outputPath,
  ];

  try {
    const ok = await new Promise<boolean>((resolve) => {
      const proc = spawn('ffmpeg', args, { env: ffmpegEnv });
      let stderr = '';
      proc.stderr.on('data', (chunk) => { stderr += chunk.toString(); });
      proc.on('error', (err) => {
        logger.warn('ffmpeg failed to start for CAPTCHA audio capture', { error: err.message });
        resolve(false);
      });
      proc.on('close', (code) => {
        if (code !== 0) {
          logger.warn('ffmpeg exited non-zero during CAPTCHA audio capture', { code, stderr: stderr.slice(0, 500) });
        }
        resolve(code === 0);
      });
    });

    if (!ok) return null;

    const buffer = await readFile(outputPath);
    if (buffer.length === 0) {
      logger.warn('CAPTCHA audio capture produced an empty file');
      return null;
    }
    return { base64: buffer.toString('base64'), mimeType: 'audio/mpeg' };
  } catch (err) {
    logger.warn('CAPTCHA audio capture failed', { error: err instanceof Error ? err.message : String(err) });
    return null;
  } finally {
    await rm(dir, { recursive: true, force: true }).catch(() => undefined);
  }
}
