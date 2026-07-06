import axios from 'axios';
import crypto from 'crypto';
import { Logger } from 'winston';
import config from '../config';
import { KnownError } from '../error';
import { getErrorType } from '../util/logger';

export interface RecordingCompletedPayload {
  recordingId: string;
  meetingLink?: string;
  status: 'completed' | string;
  blobUrl?: string; // generic storage url (S3, Azure blob, etc.)
  timestamp: string; // ISO string
  metadata?: Record<string, any>;
}

export interface MeetingFailedPayload {
  recordingId: string;
  meetingLink?: string;
  status: 'failed';
  timestamp: string;
  error: {
    type: string;
    message: string;
    name?: string;
    retryable?: boolean;
    maxRetries?: number;
  };
  metadata: {
    userId: string;
    teamId: string;
    botId?: string;
    eventId?: string;
    provider?: string;
    meetingName?: string;
    timezone?: string;
  };
}

export interface MeetingFailureContext {
  url: string;
  name?: string;
  teamId: string;
  timezone?: string;
  userId: string;
  botId?: string;
  eventId?: string;
  provider?: string;
}

function signPayload(body: string, secret?: string): string | undefined {
  if (!secret) return undefined;
  return crypto.createHmac('sha256', secret).update(body).digest('hex');
}

async function sendWebhook(payload: RecordingCompletedPayload | MeetingFailedPayload, logger: Logger) {
  if (!config.notifyWebhookEnabled) return;
  if (!config.notifyWebhookUrl) {
    logger.warn('Webhook enabled but NOTIFY_WEBHOOK_URL is not set. Skipping.');
    return;
  }

  const body = JSON.stringify(payload);
  const signature = signPayload(body, config.notifyWebhookSecret);

  try {
    await axios.post(config.notifyWebhookUrl, body, {
      headers: {
        'Content-Type': 'application/json',
        ...(signature ? { 'X-Webhook-Signature': signature } : {}),
      },
      timeout: 10000,
    });
    logger.info(`${payload.status === 'failed' ? 'Meeting failed' : 'Recording completed'} webhook delivered.`);
  } catch (err) {
    logger.error('Failed to deliver meeting webhook', err as any);
  }
}

export async function notifyRecordingCompleted(payload: RecordingCompletedPayload, logger: Logger) {
  await sendWebhook(payload, logger);
}

export function createMeetingFailedPayload(context: MeetingFailureContext, error: unknown): MeetingFailedPayload {
  const entityId = context.botId ?? context.eventId ?? context.userId;
  const errorType = getErrorType(error);
  const message = error instanceof Error ? error.message : String(error ?? 'Unknown error');

  return {
    recordingId: entityId,
    meetingLink: context.url,
    status: 'failed',
    timestamp: new Date().toISOString(),
    error: {
      type: errorType,
      message,
      ...(error instanceof Error ? { name: error.name } : {}),
      ...(error instanceof KnownError ? {
        retryable: error.retryable,
        maxRetries: error.maxRetries,
      } : {}),
    },
    metadata: {
      userId: context.userId,
      teamId: context.teamId,
      botId: context.botId,
      eventId: context.eventId,
      provider: context.provider,
      meetingName: context.name,
      timezone: context.timezone,
    },
  };
}

export async function notifyMeetingFailed(payload: MeetingFailedPayload, logger: Logger) {
  await sendWebhook(payload, logger);
}
