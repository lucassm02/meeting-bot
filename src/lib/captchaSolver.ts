// Pluggable solver for the Microsoft Teams anonymous-join CAPTCHA.
//
// The Teams challenge is a Microsoft HIP (Human Interaction Proof): a distorted
// *text-image* CAPTCHA, not Google reCAPTCHA. Text-image is reliably solvable by
// commodity solving services that expose the well-known 2Captcha HTTP API
// (2Captcha itself, CapMonster, and many compatible providers), so a single
// implementation with a configurable base URL covers all of them.
//
// Future providers, such as audio-via-Whisper, can be
// added behind the CaptchaSolver interface and selected in getCaptchaSolver().
import axios from 'axios';
import { Logger } from 'winston';
import config from '../config';

export interface CaptchaSolver {
  /**
   * Solve a text-image CAPTCHA. Receives the PNG of the challenge image as
   * base64 (no data: prefix) and resolves to the recognised characters.
   * Throws on provider error or timeout.
   */
  solveImage(imageBase64: string): Promise<string>;

  /**
   * Solve the audio variant of the challenge. Receives the recorded clip as
   * base64 plus its mime-type and resolves to the recognised characters.
   * Optional: providers that only handle images omit it, and the bot skips the
   * audio fallback when it is absent.
   */
  solveAudio?(audioBase64: string, mimeType: string): Promise<string>;
}

/**
 * Solver for any service implementing the 2Captcha HTTP API
 * (in.php to submit, res.php to poll).
 */
class TwoCaptchaCompatibleSolver implements CaptchaSolver {
  constructor(
    private readonly apiKey: string,
    private readonly baseUrl: string,
    private readonly timeoutMs: number,
    private readonly logger: Logger,
  ) {}

  async solveImage(imageBase64: string): Promise<string> {
    const captchaId = await this.submit(imageBase64);
    this.logger.info('CAPTCHA submitted to solver provider', { captchaId });
    return this.poll(captchaId);
  }

  private async submit(imageBase64: string): Promise<string> {
    const body = new URLSearchParams({
      key: this.apiKey,
      method: 'base64',
      body: imageBase64,
      json: '1',
    });

    const { data } = await axios.post(`${this.baseUrl}/in.php`, body.toString(), {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      timeout: 30000,
    });

    if (data?.status !== 1) {
      throw new Error(`CAPTCHA solver rejected submission: ${data?.request ?? 'unknown error'}`);
    }
    return String(data.request);
  }

  private async poll(captchaId: string): Promise<string> {
    const deadline = Date.now() + this.timeoutMs;
    // Providers ask callers to wait a few seconds before the first poll.
    while (Date.now() < deadline) {
      await new Promise((resolve) => setTimeout(resolve, 5000));

      const { data } = await axios.get(`${this.baseUrl}/res.php`, {
        params: { key: this.apiKey, action: 'get', id: captchaId, json: 1 },
        timeout: 30000,
      });

      if (data?.status === 1) {
        return String(data.request).trim();
      }
      if (data?.request !== 'CAPCHA_NOT_READY') {
        // Anything other than the "still working" marker is a hard error.
        throw new Error(`CAPTCHA solver error while polling: ${data?.request ?? 'unknown error'}`);
      }
    }
    throw new Error(`CAPTCHA solve timed out after ${this.timeoutMs}ms`);
  }
}

/**
 * Solver backed by the orchestrator's internal CAPTCHA endpoints. The OpenAI key
 * lives only in the orchestrator, so the bot ships the challenge (image/audio)
 * and receives the characters back. Supports both modalities:
 *  - image: GPT-4o vision reads the HIP text-image (primary path).
 *  - audio: Whisper transcribes the audio challenge (fallback).
 */
class OpenAiOrchestratorSolver implements CaptchaSolver {
  constructor(
    private readonly baseUrl: string,
    private readonly timeoutMs: number,
    private readonly logger: Logger,
  ) {}

  async solveImage(imageBase64: string): Promise<string> {
    return this.post('/internal/captcha/image', { imageBase64 });
  }

  async solveAudio(audioBase64: string, mimeType: string): Promise<string> {
    return this.post('/internal/captcha/audio', { audioBase64, mimeType });
  }

  private async post(path: string, payload: Record<string, string>): Promise<string> {
    const { data } = await axios.post(`${this.baseUrl}${path}`, payload, {
      headers: { 'Content-Type': 'application/json' },
      timeout: this.timeoutMs,
      // Large base64 payloads (image/audio) must not be capped by axios defaults.
      maxBodyLength: Infinity,
      maxContentLength: Infinity,
    });
    const text = typeof data?.text === 'string' ? data.text.trim() : '';
    if (!text) {
      throw new Error(`orchestrator CAPTCHA endpoint returned no text (${path})`);
    }
    return text;
  }
}

/**
 * Returns a configured solver, or null when the feature is disabled / not
 * configured. A null result means the caller should fall back to the previous
 * "detect and abort" behaviour.
 */
export function getCaptchaSolver(logger: Logger): CaptchaSolver | null {
  if (!config.teamsCaptchaSolverEnabled) {
    return null;
  }

  if (config.teamsCaptchaSolverProvider === 'openai') {
    if (!config.orchestratorInternalUrl) {
      logger.warn('TEAMS_CAPTCHA_SOLVER_PROVIDER=openai but the orchestrator URL is unknown (set ORCHESTRATOR_INTERNAL_URL or NOTIFY_WEBHOOK_URL); CAPTCHA cannot be solved');
      return null;
    }
    return new OpenAiOrchestratorSolver(
      config.orchestratorInternalUrl,
      config.teamsCaptchaSolverTimeoutMs,
      logger,
    );
  }

  if (!config.teamsCaptchaSolverApiKey) {
    logger.warn('TEAMS_CAPTCHA_SOLVER_ENABLED is true but TEAMS_CAPTCHA_SOLVER_API_KEY is missing; CAPTCHA cannot be solved');
    return null;
  }
  return new TwoCaptchaCompatibleSolver(
    config.teamsCaptchaSolverApiKey,
    config.teamsCaptchaSolverBaseUrl,
    config.teamsCaptchaSolverTimeoutMs,
    logger,
  );
}
