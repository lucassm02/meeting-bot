import { StorageProvider } from './storage-provider';
import { S3StorageProvider } from './s3-storage-provider';
import { AzureBlobStorageProvider } from './azure-blob-storage-provider';
import { LocalStorageProvider } from './local-storage-provider';
import config from '../../config';

export function getStorageProvider(): StorageProvider {
  // 'local' is selected via the UPLOADER_TYPE layer; s3/azure via STORAGE_PROVIDER.
  if (config.uploaderType === 'local') {
    return new LocalStorageProvider();
  }
  if (config.storageProvider === 'azure') {
    return new AzureBlobStorageProvider();
  }
  return new S3StorageProvider();
}
