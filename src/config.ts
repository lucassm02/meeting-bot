import dotenv from 'dotenv';
import { UploaderType } from './types';
dotenv.config();

const ENVIRONMENTS = [
  'production',
  'staging',
  'development',
  'cli',
  'test',
] as const;

export type Environment = (typeof ENVIRONMENTS)[number];
export const NODE_ENV: Environment = ENVIRONMENTS.includes(
  process.env.NODE_ENV as Environment
)
  ? (process.env.NODE_ENV as Environment)
  : 'staging';

console.log('NODE_ENV', process.env.NODE_ENV);

const requiredSettings = [
  'GCP_DEFAULT_REGION',
  'GCP_MISC_BUCKET',
];
const missingSettings = requiredSettings.filter((s) => !process.env[s]);
if (missingSettings.length > 0) {
  missingSettings.forEach((ms) =>
    console.error(`ENV settings ${ms} is missing.`)
  );
}

const constructRedisUri = () => {
  const host = process.env.REDIS_HOST || 'redis';
  const port = process.env.REDIS_PORT || 6379;
  const username = process.env.REDIS_USERNAME;
  const password = process.env.REDIS_PASSWORD;

  if (username && password) {
    return `redis://${username}:${password}@${host}:${port}`;
  } else if (password) {
    return `redis://:${password}@${host}:${port}`;
  } else {
    return `redis://${host}:${port}`;
  }
};

const normalizeFileExtension = (extension?: string) => {
  if (!extension) return '.webm';
  return extension.startsWith('.') ? extension : `.${extension}`;
};

const parseOptionalNumber = (value?: string) => {
  if (typeof value === 'undefined' || value.trim() === '') return undefined;
  return Number(value);
};

export default {
  port: process.env.PORT || 3000,
  db: {
    host: process.env.DB_HOST || 'localhost',
    user: process,
  },
  authBaseUrlV2: process.env.AUTH_BASE_URL_V2 ?? 'http://localhost:8081/v2',
  botStatusPatchEnabled: process.env.BOT_STATUS_PATCH_ENABLED !== 'false',
  // Unset MAX_RECORDING_DURATION_MINUTES to use default upper limit on duration
  maxRecordingDuration: process.env.MAX_RECORDING_DURATION_MINUTES ?
    Number(process.env.MAX_RECORDING_DURATION_MINUTES) :
    180, // There's an upper limit on meeting duration 3 hours
  chromeExecutablePath: process.env.CHROME_PATH || '/usr/bin/google-chrome', // We use Google Chrome with Playwright for recording
  googleChromeCdpUrl: process.env.GOOGLE_CHROME_CDP_URL,
  googleChromeUserDataDir: process.env.GOOGLE_CHROME_USER_DATA_DIR,
  googleChromeStorageStatePath: process.env.GOOGLE_CHROME_STORAGE_STATE_PATH,
  googleAnonymousJoinRequestAttempts: process.env.GOOGLE_ANONYMOUS_JOIN_REQUEST_ATTEMPTS ?
    Number(process.env.GOOGLE_ANONYMOUS_JOIN_REQUEST_ATTEMPTS) :
    10,
  inactivityLimit: process.env.MEETING_INACTIVITY_MINUTES ? Number(process.env.MEETING_INACTIVITY_MINUTES) : 1,
  activateInactivityDetectionAfter: process.env.INACTIVITY_DETECTION_START_DELAY_MINUTES ? Number(process.env.INACTIVITY_DETECTION_START_DELAY_MINUTES) :  1,
  // Once `inactivityLimit` is reached, the bot no longer ends the recording right away —
  // it notifies the orchestrator (idle-started) and keeps recording. Only if silence
  // continues for this many additional minutes does it fall back to ending on its own,
  // exactly like before this feature existed (safety net for an unattended empty call).
  idleFallbackExtraMinutes: process.env.MEETING_IDLE_FALLBACK_EXTRA_MINUTES ? Number(process.env.MEETING_IDLE_FALLBACK_EXTRA_MINUTES) : 10,
  loneParticipantExitDelaySeconds: process.env.LONE_PARTICIPANT_EXIT_DELAY_SECONDS ? Number(process.env.LONE_PARTICIPANT_EXIT_DELAY_SECONDS) : 10,
  serviceKey: process.env.SCREENAPP_BACKEND_SERVICE_API_KEY,
  joinWaitTime: process.env.JOIN_WAIT_TIME_MINUTES ? Number(process.env.JOIN_WAIT_TIME_MINUTES) : 10,
  // Number of retries for transient errors (not applied to WaitingAtLobbyRetryError)
  retryCount: process.env.RETRY_COUNT ? Number(process.env.RETRY_COUNT) : 2,
  teamsPrewarmEnabled: process.env.TEAMS_PREWARM_ENABLED === 'true',
  teamsAudioStabilizationMs: process.env.TEAMS_AUDIO_STABILIZATION_MS ? Number(process.env.TEAMS_AUDIO_STABILIZATION_MS) : 1000,
  // Automatic resolution of the Teams anonymous-join CAPTCHA (Microsoft HIP image-text challenge).
  // Enabled by default — set TEAMS_CAPTCHA_SOLVER_ENABLED=false to keep the previous "detect and abort" behaviour.
  teamsCaptchaSolverEnabled: process.env.TEAMS_CAPTCHA_SOLVER_ENABLED !== 'false',
  // API key of a 2Captcha-compatible image solving service (2Captcha, CapMonster, etc.)
  teamsCaptchaSolverApiKey: process.env.TEAMS_CAPTCHA_SOLVER_API_KEY,
  // Base URL of the 2Captcha-compatible endpoint (override for self-hosted/compatible providers)
  teamsCaptchaSolverBaseUrl: process.env.TEAMS_CAPTCHA_SOLVER_BASE_URL || 'https://2captcha.com',
  // How many solve attempts before failing with TeamsCaptchaError (each retry refreshes the challenge)
  teamsCaptchaMaxRetries: process.env.TEAMS_CAPTCHA_MAX_RETRIES ? Number(process.env.TEAMS_CAPTCHA_MAX_RETRIES) : 3,
  // Upper bound for a single solve (submit + polling the provider for the answer)
  teamsCaptchaSolverTimeoutMs: process.env.TEAMS_CAPTCHA_SOLVER_TIMEOUT_MS ? Number(process.env.TEAMS_CAPTCHA_SOLVER_TIMEOUT_MS) : 120000,
  // Solver provider: 'openai' (delegates to the orchestrator, which owns the
  // OpenAI key) or '2captcha' (external image-solving service). Defaults to
  // 'openai' (higher accuracy, no external paid service).
  teamsCaptchaSolverProvider: process.env.TEAMS_CAPTCHA_SOLVER_PROVIDER || 'openai',
  // Base URL of the orchestrator's internal CAPTCHA endpoints. Derived from the
  // callback URL the orchestrator already injects (NOTIFY_WEBHOOK_URL minus the
  // /bot-callback suffix) unless explicitly overridden. Used by the 'openai' provider.
  orchestratorInternalUrl: process.env.ORCHESTRATOR_INTERNAL_URL
    || (process.env.NOTIFY_WEBHOOK_URL ? process.env.NOTIFY_WEBHOOK_URL.replace(/\/bot-callback\/?$/, '') : undefined),
  // When the image CAPTCHA is exhausted, fall back to the audio challenge.
  // Lower success rate (adversarial audio); off by default.
  teamsCaptchaAudioFallbackEnabled: process.env.TEAMS_CAPTCHA_AUDIO_FALLBACK_ENABLED === 'true',
  // Seconds of the audio challenge to capture from the PulseAudio monitor.
  teamsCaptchaAudioCaptureSeconds: process.env.TEAMS_CAPTCHA_AUDIO_CAPTURE_SECONDS ? Number(process.env.TEAMS_CAPTCHA_AUDIO_CAPTURE_SECONDS) : 8,
  miscStorageBucket: process.env.GCP_MISC_BUCKET,
  miscStorageFolder: process.env.GCP_MISC_BUCKET_FOLDER ? process.env.GCP_MISC_BUCKET_FOLDER : 'meeting-bot',
  region: process.env.GCP_DEFAULT_REGION,
  accessKey: process.env.GCP_ACCESS_KEY_ID ?? '',
  accessSecret: process.env.GCP_SECRET_ACCESS_KEY ?? '',
  redisQueueName: process.env.REDIS_QUEUE_NAME ?? 'jobs:meetbot:list',
  redisProcessingQueueName: process.env.REDIS_PROCESSING_QUEUE_NAME ?? 'jobs:meetbot:processing',
  redisUri: constructRedisUri(),
  // Notification: Webhook (disabled by default)
  notifyWebhookEnabled: process.env.NOTIFY_WEBHOOK_ENABLED === 'true',
  notifyWebhookUrl: process.env.NOTIFY_WEBHOOK_URL,
  // Optional secret to sign payloads (HMAC-SHA256). If set, signature will be sent in X-Webhook-Signature header
  notifyWebhookSecret: process.env.NOTIFY_WEBHOOK_SECRET,
  // Notification: Redis. Explicitly enabled via NOTIFY_REDIS_ENABLED, and enabled
  // automatically for Redis-worker mode so completed jobs are written to result list.
  notifyRedisEnabled: process.env.NOTIFY_REDIS_ENABLED === 'true' || process.env.REDIS_CONSUMER_ENABLED === 'true',
  // If not provided, uses redisUri with specified database selection
  notifyRedisUri: process.env.NOTIFY_REDIS_URI, // optional override
  notifyRedisDb: parseOptionalNumber(process.env.NOTIFY_REDIS_DB),
  notifyRedisList: process.env.NOTIFY_REDIS_LIST ?? 'jobs:meetbot:recordings',
  notifyRedisFailureList: process.env.NOTIFY_REDIS_FAILURE_LIST ?? 'jobs:meetbot:failures',
  uploaderFileExtension: normalizeFileExtension(process.env.UPLOADER_FILE_EXTENSION),
  isRedisEnabled: process.env.REDIS_CONSUMER_ENABLED === 'true',
  s3CompatibleStorage: {
    endpoint: process.env.S3_ENDPOINT,
    region: process.env.S3_REGION,
    accessKeyId: process.env.S3_ACCESS_KEY_ID,
    secretAccessKey: process.env.S3_SECRET_ACCESS_KEY,
    bucket: process.env.S3_BUCKET_NAME,
    forcePathStyle: process.env.S3_USE_MINIO_COMPATIBILITY === 'true',
  },
  // Object storage provider selection: 's3' (default) or 'azure'
  storageProvider: (process.env.STORAGE_PROVIDER === 'azure' ? 'azure' : 's3') as 's3' | 'azure',
  azureBlobStorage: {
    // Either provide full connection string OR account + key/SAS OR managed identity
    connectionString: process.env.AZURE_STORAGE_CONNECTION_STRING,
    accountName: process.env.AZURE_STORAGE_ACCOUNT,
    accountKey: process.env.AZURE_STORAGE_ACCOUNT_KEY, // optional when using connection string
    sasToken: process.env.AZURE_STORAGE_SAS_TOKEN, // starts with ?sv=...
    useManagedIdentity: process.env.AZURE_USE_MANAGED_IDENTITY === 'true',
    container: process.env.AZURE_STORAGE_CONTAINER,
    blobPrefix: process.env.AZURE_BLOB_PREFIX || '',
    signedUrlTtlSeconds: process.env.AZURE_SIGNED_URL_TTL_SECONDS ? Number(process.env.AZURE_SIGNED_URL_TTL_SECONDS) : 3600,
    uploadConcurrency: process.env.AZURE_UPLOAD_CONCURRENCY ? Number(process.env.AZURE_UPLOAD_CONCURRENCY) : 4,
  },
  // Persistent working directory for in-progress recordings. When set,
  // temp files + their .meta.json sidecars live here — point it at a mounted volume
  // so a container recreation doesn't lose the recording and the startup recovery
  // can finish orphaned uploads. Must be separate from the final recordings dir.
  // Falls back to dist/_tempvideo (ephemeral) to preserve the previous behaviour.
  recordingWorkDir: process.env.RECORDING_WORK_DIR,
  // Local filesystem storage (when UPLOADER_TYPE=local)
  localStorage: {
    // Root directory where recordings are written. REQUIRED when UPLOADER_TYPE=local
    // (no default — validateConfig() throws if missing).
    baseDir: process.env.LOCAL_STORAGE_DIR,
    // Optional public base URL to build the notification blobUrl (e.g. nginx serving the volume).
    // When absent, blobUrl is a file:// absolute path.
    publicBaseUrl: process.env.LOCAL_STORAGE_PUBLIC_BASE_URL || undefined,
  },
  uploaderType: process.env.UPLOADER_TYPE ? (process.env.UPLOADER_TYPE as UploaderType) : 's3' as UploaderType,
};
