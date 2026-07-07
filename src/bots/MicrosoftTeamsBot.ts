import { Locator, Page } from 'playwright';
import { JoinParams } from './AbstractMeetBot';
import { BotStatus } from '../types';
import config from '../config';
import { RecordingUploadFailedError, TeamsCaptchaError, WaitingAtLobbyRetryError } from '../error';
import { handleWaitingAtLobbyError, MeetBotBase } from './MeetBotBase';
import { v4 } from 'uuid';
import { patchBotStatus } from '../services/botService';
import { IUploader } from '../middleware/disk-uploader';
import { Logger } from 'winston';
import { retryActionWithWait } from '../util/resilience';
import { uploadDebugImage } from '../services/bugService';
import createBrowserContext from '../lib/chromium';
import { getCaptchaSolver } from '../lib/captchaSolver';
import { captureSystemAudioClip } from '../lib/audioCapture';
import { browserLogCaptureCallback } from '../util/logger';
import { MICROSOFT_REQUEST_DENIED } from '../constants';
import { FFmpegRecorder } from '../lib/ffmpegRecorder';
import { clearActiveRecording, registerActiveRecording } from '../lib/activeRecording';
import { notifyMeetingIdle } from '../services/notificationService';
import * as path from 'path';
import * as fs from 'fs';
import { exec } from 'child_process';
import { promisify } from 'util';

const execAsync = promisify(exec);

export class MicrosoftTeamsBot extends MeetBotBase {
  private _logger: Logger;
  private _correlationId: string;
  constructor(logger: Logger, correlationId: string) {
    super();
    this.slightlySecretId = v4();
    this._logger = logger;
    this._correlationId = correlationId;
  }
  async join({ url, name, bearerToken, teamId, timezone, userId, eventId, botId, uploader }: JoinParams): Promise<void> {
    const _state: BotStatus[] = ['processing'];

    const handleUpload = async () => {
      this._logger.info('Begin recording upload to server', { userId, teamId });
      const uploadResult = await uploader.uploadRecordingToRemoteStorage();
      this._logger.info('Recording upload result', { uploadResult, userId, teamId });
      return uploadResult;
    };

    try {
      const pushState = (st: BotStatus) => _state.push(st);
      await this.joinMeeting({ url, name, bearerToken, teamId, timezone, userId, eventId, botId, pushState, uploader });

      // Finish the upload from the temp video
      const uploadResult = await handleUpload();

      if (_state.includes('finished') && !uploadResult) {
        _state.splice(_state.indexOf('finished'), 1, 'failed');
        this._logger.error('Recording completed but upload failed; raising non-retryable failure so JobStore does not rejoin the ended meeting', { botId, userId, teamId });
        throw new RecordingUploadFailedError('Microsoft Teams recording completed but upload failed');
      } else if (uploadResult) {
        this._logger.info('Recording and upload completed successfully', { botId, userId, teamId });
      }

      await patchBotStatus({ botId, eventId, provider: 'microsoft', status: _state, token: bearerToken }, this._logger);
    } catch(error) {
      // Log the actual error that occurred
      this._logger.error('Error in Microsoft Teams bot join process', {
        error: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
        botId,
        userId,
        teamId,
        currentState: _state
      });

      if (!_state.includes('finished') && !_state.includes('failed'))
        _state.push('failed');

      // Try to update bot status (may fail if API is unreachable, but that's OK)
      await patchBotStatus({ botId, eventId, provider: 'microsoft', status: _state, token: bearerToken }, this._logger);

      if (error instanceof WaitingAtLobbyRetryError)
        await handleWaitingAtLobbyError({ token: bearerToken, botId, eventId, provider: 'microsoft', error }, this._logger);

      throw error;
    } finally {
      // Guarantee chrome subprocess tree is reaped regardless of exit path.
      // No-op if a deeper code path already closed the browser.
      try {
        const browser = this.page?.context().browser();
        if (browser?.isConnected()) {
          await browser.close();
          this._logger.info('Browser closed in join finally');
        } else if (this.page?.context()) {
          await this.page.context().close();
          this._logger.info('Persistent browser context closed in join finally');
        }
      } catch (cleanupErr) {
        this._logger.warn('Browser cleanup in join finally failed (non-fatal)', { error: cleanupErr });
      }
    }
  }

  private async joinMeeting({ url, name, teamId, userId, eventId, botId, pushState, uploader }: JoinParams & { pushState(state: BotStatus): void }): Promise<void> {
    const joinButtonSelectors = [
      // Locale-independent: the launcher page links directly to light-meetings
      'a[href*="light-meetings"]',
      // English aria-labels
      'button[aria-label="Join meeting from this browser"]',
      'button[aria-label="Continue on this browser"]',
      'button[aria-label="Join on this browser"]',
      // English text
      'button:has-text("Continue on this browser")',
      'button:has-text("Continue in this browser")',
      'button:has-text("Join from browser")',
      // Portuguese (pt-BR)
      'button:has-text("Continuar neste navegador")',
      'button:has-text("Continuar no navegador")',
      // Spanish
      'button:has-text("Continuar en este explorador")',
      'button:has-text("Continuar en el explorador")',
      // French
      'button:has-text("Continuer dans ce navigateur")',
      // German
      'button:has-text("In diesem Browser fortfahren")',
      'button:has-text("Im Browser fortfahren")',
      // Italian
      'button:has-text("Continua nel browser")',
    ];

    const clickFirstVisibleSelector = async (page: Page, selectors: string[], timeoutMs: number, logPrefix: string): Promise<boolean> => {
      const startedAt = Date.now();
      for (const selector of selectors) {
        this._logger.info(`${logPrefix}: checking selector`, { selector });
      }

      while ((Date.now() - startedAt) < timeoutMs) {
        for (const selector of selectors) {
          const button = page.locator(selector).first();
          if (await button.isVisible({ timeout: 500 }).catch(() => false)) {
            this._logger.info(`${logPrefix}: found button`, { selector });
            await button.click({ force: true });
            return true;
          }
        }
        await page.waitForTimeout(500);
      }

      return false;
    };

    if (config.teamsPrewarmEnabled) {
      // First run: Navigate to pre-join screen to trigger Chrome dialogs, then close.
      // Disabled by default because the fake-device flags normally avoid those dialogs.
      this._logger.info('Pre-warming: Opening browser to trigger first-run dialogs...');
      let warmupPage: Page | undefined;
      try {
        warmupPage = await createBrowserContext(url, this._correlationId, 'microsoft');
        this._logger.info('Pre-warming: Navigating to Teams meeting...');
        await warmupPage.goto(url, { waitUntil: 'domcontentloaded' });
        await clickFirstVisibleSelector(warmupPage, joinButtonSelectors, 8000, 'Pre-warming');
        await warmupPage.locator('input[data-tid="prejoin-display-name-input"]').waitFor({ state: 'visible', timeout: 8000 }).catch(() => undefined);
        this._logger.info('Pre-warming complete - dialogs triggered');
      } catch (error) {
        this._logger.warn('Pre-warming failed (non-fatal):', error);
      } finally {
        // Guarantee the warmup chrome tree is reaped even if the block above threw
        // mid-flight. join()'s outer finally only covers this.page, not warmupPage.
        try {
          const browser = warmupPage?.context().browser();
          if (browser?.isConnected()) {
            this._logger.info('Pre-warming: Closing warmup browser...');
            await browser.close();
          }
        } catch (cleanupErr) {
          this._logger.warn('Pre-warming: warmup browser cleanup failed (non-fatal)', { error: cleanupErr });
        }
      }
    } else {
      this._logger.info('Teams pre-warming disabled; launching the meeting browser directly');
    }

    // Second run: Actual meeting join
    this._logger.info('Launching browser for actual meeting...');

    this.page = await createBrowserContext(url, this._correlationId, 'microsoft');

    this._logger.info('Navigating to Microsoft Teams Meeting URL...');
    await this.page.goto(url, { waitUntil: 'domcontentloaded' });

    // Try to find and click "Join from browser" button
    this._logger.info('Waiting for Join meeting from browser button...');
    const buttonClicked = await clickFirstVisibleSelector(this.page, joinButtonSelectors, 60000, 'Join from browser');

    if (!buttonClicked) {
      this._logger.warn('Join from browser button not found after 60s; capturing debug screenshot and proceeding...');
      await uploadDebugImage(await this.page.screenshot({ type: 'png', fullPage: true }), 'launcher-button-not-found', userId, this._logger, botId);
    }

    this._logger.info('Waiting for pre-join screen to load...');

    // Fill the pre-join name field (optional; some flows don't ask for a name).
    await this.fillDisplayName(name);

    // Toggle off camera and mute microphone before joining.
    await this.toggleDevices();

    this._logger.info('Clicking the join button...');
    await retryActionWithWait(
      'Clicking the join button',
      async () => { await this.clickJoinNow(); },
      this._logger,
      3,
      15000,
      async () => {
        await uploadDebugImage(await this.page.screenshot({ type: 'png', fullPage: true }), 'join-button-click', userId, this._logger, botId);
      }
    );

    // Teams shows a Microsoft HIP (image-text) CAPTCHA for anonymous
    // guests on tenants whose policy requires verification. Detect it and try to
    // solve via the configured solver before giving up.
    await this.page.waitForTimeout(2000);
    if (await this.isCaptchaVisible()) {
      this._logger.warn('CAPTCHA detected on Teams join (tenant anonymous-join policy)', { botId, userId });
      const solved = await this.solveCaptchaWithRetries(name, userId, botId);
      if (!solved) {
        await uploadDebugImage(await this.page.screenshot({ type: 'png', fullPage: true }), 'captcha-detected', userId, this._logger, botId);
        throw new TeamsCaptchaError(`Microsoft Teams CAPTCHA could not be solved after ${config.teamsCaptchaMaxRetries} attempt(s). Tenant policy requires guest verification; consider an authenticated session.`);
      }
      this._logger.info('CAPTCHA solved successfully; continuing to join', { botId, userId });
    }

    // Do this to ensure meeting bot has joined the meeting
    try {
      const wanderingTime = config.joinWaitTime * 60 * 1000; // Give some time to be let in
      const callButton = this.page.getByRole('button', { name: /Leave|Sair|Salir|Quitter|Verlassen/i });
      await callButton.waitFor({ timeout: wanderingTime });
      this._logger.info('Bot is entering the meeting...');
    } catch (error) {
      const bodyText = await this.page.evaluate(() => document.body.innerText);

      const userDenied = (bodyText || '')?.includes(MICROSOFT_REQUEST_DENIED);

      this._logger.error('Cant finish wait at the lobby check', { userDenied, waitingAtLobbySuccess: false, bodyText });

      this._logger.error('Closing the browser on error...', error);
      await this.page.context().browser()?.close();

      // Don't retry lobby errors - if user doesn't admit bot, retrying won't help
      throw new WaitingAtLobbyRetryError('Microsoft Teams Meeting bot could not enter the meeting...', bodyText ?? '', false, 0);
    }

    pushState('joined');

    const dismissDeviceChecksAndNotifications = async () => {
      const closeSelectors = ['button[aria-label=Close]:visible', 'button[title="Close"]:visible'];
      const startedAt = Date.now();
      let closeButtonsClicked = 0;
      let emptyPasses = 0;

      while ((Date.now() - startedAt) < 3000) {
        let clickedOnPass = false;
        for (const selector of closeSelectors) {
          const visibleButtons = await this.page.locator(selector).all();
          for (const btn of visibleButtons) {
            try {
              await btn.click({ timeout: 1000 });
              closeButtonsClicked++;
              clickedOnPass = true;
            } catch (err) {
              this._logger.warn('Close button click failed, possibly already dismissed', { error: err });
            }
          }
        }

        if (!clickedOnPass) {
          emptyPasses++;
          if (emptyPasses >= 2) {
            break;
          }
          await this.page.waitForTimeout(250);
        } else {
          emptyPasses = 0;
        }
      }

      this._logger.info('Finished dismissing device checks and notifications', { closeButtonsClicked });
    };
    await dismissDeviceChecksAndNotifications();

    // Wait for mic to be fully muted and any initial beeps to stop
    if (config.teamsAudioStabilizationMs > 0) {
      this._logger.info('Waiting briefly for audio to stabilize before recording...', { ms: config.teamsAudioStabilizationMs });
      await this.page.waitForTimeout(config.teamsAudioStabilizationMs);
    }

    // Recording the meeting page with ffmpeg
    this._logger.info('Begin recording with ffmpeg...');
    await this.recordMeetingPageWithFFmpeg({ teamId, userId, eventId, botId, uploader });

    pushState('finished');
  }

  // Turns off camera and mutes the mic on the pre-join screen. Best-effort and
  // idempotent (only clicks a toggle that is in the "on" state). Shared by the
  // initial join and retry-screen recovery — the rejoin returns to a fresh
  // pre-join screen with the camera back on, so it must run there too.
  private async toggleDevices(): Promise<void> {
    try {
      this._logger.info('Attempting to turn off camera and mute microphone...');

      // Turn off camera
      try {
        const cameraSelectors = [
          'input[data-tid="toggle-video"][checked]',
          'input[type="checkbox"][title*="Turn camera off" i]',
          'input[role="switch"][data-tid="toggle-video"]',
          'button[aria-label*="Turn camera off" i]',
          'button[aria-label*="Camera off" i]',
        ];

        for (const selector of cameraSelectors) {
          const cameraButton = this.page.locator(selector).first();
          const isVisible = await cameraButton.isVisible({ timeout: 1000 }).catch(() => false);
          if (isVisible) {
            const label = await cameraButton.getAttribute('aria-label');
            this._logger.info(`Clicking camera toggle: ${label}`);
            await cameraButton.click();
            await this.page.waitForTimeout(250);
            break;
          }
        }
      } catch (err) {
        this._logger.info('Could not toggle camera', err instanceof Error ? err.message : String(err));
      }

      // Mute microphone
      try {
        const micSelectors = [
          'input[data-tid="toggle-mute"]:not([checked])',
          'input[type="checkbox"][title*="Mute mic" i]',
          'input[role="switch"][data-tid="toggle-mute"]',
          'button[aria-label*="Mute microphone" i]',
          'button[aria-label*="Mute mic" i]',
        ];

        for (const selector of micSelectors) {
          const micButton = this.page.locator(selector).first();
          const isVisible = await micButton.isVisible({ timeout: 1000 }).catch(() => false);
          if (isVisible) {
            const label = await micButton.getAttribute('aria-label');
            this._logger.info(`Clicking microphone toggle: ${label}`);
            await micButton.click();
            await this.page.waitForTimeout(250);
            break;
          }
        }
      } catch (err) {
        this._logger.info('Could not toggle microphone', err instanceof Error ? err.message : String(err));
      }

      this._logger.info('Finished toggling camera and microphone');
    } catch (error) {
      this._logger.warn('Error toggling devices', error instanceof Error ? error.message : String(error));
    }
  }

  // Fills the pre-join display-name field. Best-effort: skips silently if the
  // field never appears. Shared by the initial join and retry-screen recovery.
  private async fillDisplayName(name: string | undefined, timeoutMs = 45000): Promise<void> {
    try {
      this._logger.info('Looking for name input field...');
      // Old Teams web uses data-tid; the new light-meetings UI uses placeholders.
      const nameInputSelectors = [
        'input[data-tid="prejoin-display-name-input"]',
        'input[placeholder*="nome" i]',
        'input[placeholder*="name" i]',
        'input[placeholder*="nombre" i]',
        'input[placeholder*="nom" i]',
        'input[aria-label*="nome" i]',
        'input[aria-label*="name" i]',
      ].join(', ');
      const nameInput = this.page.locator(nameInputSelectors).first();
      await nameInput.waitFor({ state: 'visible', timeout: timeoutMs });
      this._logger.info('Found name input field, filling with bot name...');
      await nameInput.fill(name ? name : 'ScreenApp Notetaker');
    } catch (err) {
      this._logger.info('Name input field not found, skipping...', err instanceof Error ? err.message : String(err));
    }
  }

  // Clicks the Teams "Join now" button across locale variants. Throws if none
  // is visible. Shared by the initial join and retry-screen recovery.
  private async clickJoinNow(): Promise<void> {
    const possibleTexts = [
      'Join now', 'Join', 'Ask to join', 'Join meeting',
      // Portuguese (pt-BR) — light-meetings UI
      'Ingressar agora', 'Ingressar', 'Solicitar para ingressar',
      // Spanish
      'Unirse ahora', 'Unirse',
      // French
      'Rejoindre maintenant', 'Rejoindre',
      // German
      'Jetzt teilnehmen', 'Teilnehmen',
    ];

    for (const text of possibleTexts) {
      try {
        const button = this.page.getByRole('button', { name: new RegExp(text, 'i') });
        if (await button.isVisible({ timeout: 1000 }).catch(() => false)) {
          await button.click();
          this._logger.info(`Successfully clicked "${text}" button`);
          return;
        }
      } catch (err) {
        this._logger.info(`Unable to click "${text}" button, trying next...`);
      }
    }

    throw new Error('Unable to find any join button variant');
  }

  // Root of the Microsoft HIP verification dialog. Anchored on the ARIA
  // relationship the Fluent Dialog exposes (role="dialog" describedby the
  // instructions element), which is stable across image/audio modes and
  // locale-independent. Fluent `fui-*` class hashes must never be selectors.
  private captchaDialog(): Locator {
    return this.page
      .locator('[role="dialog"][aria-describedby="captcha-dialog-instructions"], [role="dialog"]:has(#captcha-dialog-instructions)')
      .first();
  }

  // The answer field. Same element in both modes — targeted by its stable
  // data-tid, not by placeholder (which changes: "...caracteres que você vê"
  // for image vs "...o que você está ouvindo" for audio, and per locale).
  private captchaInput(): Locator {
    return this.captchaDialog().locator('[data-tid="HIP-Captcha-Input"]').first();
  }

  // The verification dialog is identified by locale- and mode-independent
  // anchors: the dialog's ARIA relationship, the instructions id, and the
  // input's data-tid. No text/placeholder matching, no Fluent class hashes.
  private async isCaptchaVisible(): Promise<boolean> {
    const captchaLocator = this.page.locator([
      '[role="dialog"][aria-describedby="captcha-dialog-instructions"]',
      '#captcha-dialog-instructions',
      '[data-tid="HIP-Captcha-Input"]',
      '[data-tid="HIP-Captcha-Audio-Play"]',
    ].join(', ')).first();
    return captchaLocator.isVisible({ timeout: 1000 }).catch(() => false);
  }

  // The "Não foi possível verificar você" screen Teams shows when the submitted
  // answer is WRONG. The captcha dialog is gone at this point, so this must be
  // checked before concluding a solve succeeded — otherwise a wrong answer reads
  // as success. Recovery is via the rejoin button (see recoverFromRetryScreen).
  private async isCaptchaRetryScreenVisible(): Promise<boolean> {
    return this.page.locator('[data-tid="calling-retry-screen"]').first()
      .isVisible({ timeout: 1000 }).catch(() => false);
  }

  // Classifies the page state after submitting an answer:
  //  - 'present':  the captcha dialog is still up (validate in place → refresh).
  //  - 'rejected': the "could not verify you" screen appeared (wrong answer).
  //  - 'solved':   neither is showing → the challenge was accepted.
  // Order matters: dialog first, then the retry screen, else solved.
  private async classifyCaptchaOutcome(): Promise<'present' | 'rejected' | 'solved'> {
    if (await this.isCaptchaVisible()) return 'present';
    if (await this.isCaptchaRetryScreenVisible()) return 'rejected';
    return 'solved';
  }

  // Recovers from the retry screen by clicking "Reingressar na chamada", which
  // returns to the pre-join screen; then re-fills the name and clicks join so
  // the challenge is presented again for another attempt.
  private async recoverFromRetryScreen(name: string | undefined, userId: string, botId?: string): Promise<boolean> {
    const rejoin = this.page.locator('[data-tid="calling-retry-rejoinbutton"]').first();
    if (!(await rejoin.isVisible({ timeout: 2000 }).catch(() => false))) {
      return false;
    }
    await rejoin.click().catch(() => undefined);

    // The rejoin lands on a fresh pre-join screen: re-fill the name and turn the
    // camera/mic back off (they reset to "on") before clicking join again.
    await this.fillDisplayName(name, 15000);
    await this.toggleDevices();
    try {
      await this.clickJoinNow();
    } catch (err) {
      this._logger.warn('Could not click join after rejoin', { error: err instanceof Error ? err.message : String(err), botId, userId });
      return false;
    }

    // Wait for the challenge to come back before the next solve attempt.
    await this.page.waitForTimeout(2000);
    return this.isCaptchaVisible();
  }

  // One-shot diagnostic: dumps every data-tid inside the captcha dialog so any
  // future change to the widget's stable anchors is visible in production logs.
  private async logCaptchaAnchors(botId?: string, userId?: string): Promise<void> {
    try {
      const tids = await this.captchaDialog().evaluate((el) =>
        Array.from(el.querySelectorAll('[data-tid]'))
          .map((n) => n.getAttribute('data-tid'))
          .filter((v): v is string => Boolean(v)),
      );
      this._logger.info('CAPTCHA dialog data-tid anchors', { tids, botId, userId });
    } catch (err) {
      this._logger.warn('Could not read CAPTCHA dialog anchors', { error: err instanceof Error ? err.message : String(err), botId, userId });
    }
  }

  // Solve the HIP CAPTCHA. Between attempts it either refreshes the challenge
  // (dialog still up) or rejoins the meeting (wrong answer bounced us to the
  // retry screen). Returns true once the challenge is accepted, false when
  // attempts are exhausted.
  private async solveCaptchaWithRetries(name: string | undefined, userId: string, botId?: string): Promise<boolean> {
    const solver = getCaptchaSolver(this._logger);
    if (!solver) {
      this._logger.warn('CAPTCHA solver disabled or unconfigured; cannot solve challenge', { botId, userId });
      return false;
    }

    await this.logCaptchaAnchors(botId, userId);

    const maxRetries = config.teamsCaptchaMaxRetries;
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        const imageBase64 = await this.captureCaptchaImage();
        if (!imageBase64) {
          this._logger.warn('Could not capture CAPTCHA image element', { attempt, botId, userId });
        } else {
          const answer = await solver.solveImage(imageBase64);
          this._logger.info('CAPTCHA solver returned an answer', { attempt, answerLength: answer.length, botId, userId });
          await this.submitCaptchaAnswer(answer);
        }
      } catch (err) {
        this._logger.warn('CAPTCHA solve attempt failed', { attempt, error: err instanceof Error ? err.message : String(err), botId, userId });
      }

      // Give the page a moment to validate, then classify what we're looking at.
      await this.page.waitForTimeout(1500);
      const outcome = await this.classifyCaptchaOutcome();

      if (outcome === 'solved') {
        this._logger.info('CAPTCHA accepted; continuing', { attempt, botId, userId });
        return true;
      }

      if (attempt >= maxRetries) break;

      if (outcome === 'rejected') {
        // Wrong answer: Teams navigated to the "could not verify you" screen.
        // Rejoin to bring the challenge back for the next attempt.
        this._logger.info('CAPTCHA rejected (retry screen); rejoining for another attempt', { attempt, botId, userId });
        if (!(await this.recoverFromRetryScreen(name, userId, botId))) {
          this._logger.warn('Could not recover from CAPTCHA retry screen; aborting', { attempt, botId, userId });
          return false;
        }
      } else {
        // Still in the dialog: refresh for a fresh challenge and try again.
        this._logger.info('CAPTCHA still present; refreshing challenge for next attempt', { attempt, botId, userId });
        await this.refreshCaptcha();
      }
    }

    // Image attempts exhausted — try the audio challenge as a last resort, but
    // only while the dialog is still up (the audio toggle lives inside it).
    if (config.teamsCaptchaAudioFallbackEnabled && typeof solver.solveAudio === 'function' && await this.isCaptchaVisible()) {
      this._logger.info('Image CAPTCHA exhausted; attempting audio fallback', { botId, userId });
      if (await this.solveCaptchaViaAudio(solver.solveAudio.bind(solver), userId, botId)) {
        return true;
      }
    }

    return false;
  }

  // Audio fallback. Switches the dialog to the audio challenge, records
  // a short clip from the system audio, transcribes it via the solver and submits.
  private async solveCaptchaViaAudio(
    solveAudio: (audioBase64: string, mimeType: string) => Promise<string>,
    userId: string,
    botId?: string,
  ): Promise<boolean> {
    if (!(await this.switchToAudioCaptcha())) {
      this._logger.warn('Could not switch to audio CAPTCHA', { botId, userId });
      return false;
    }

    const maxRetries = config.teamsCaptchaMaxRetries;
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        await this.playCaptchaAudio();
        const clip = await captureSystemAudioClip(config.teamsCaptchaAudioCaptureSeconds, this._logger);
        if (!clip) {
          this._logger.warn('Could not capture CAPTCHA audio clip', { attempt, botId, userId });
        } else {
          const answer = await solveAudio(clip.base64, clip.mimeType);
          this._logger.info('Audio CAPTCHA solver returned an answer', { attempt, answerLength: answer.length, botId, userId });
          await this.submitCaptchaAnswer(answer);
        }
      } catch (err) {
        this._logger.warn('Audio CAPTCHA solve attempt failed', { attempt, error: err instanceof Error ? err.message : String(err), botId, userId });
      }

      await this.page.waitForTimeout(1500);
      const outcome = await this.classifyCaptchaOutcome();
      if (outcome === 'solved') {
        this._logger.info('CAPTCHA accepted after audio solve', { attempt, botId, userId });
        return true;
      }
      if (outcome === 'rejected') {
        // Wrong audio answer bounced us off the dialog; the toggle is gone, so
        // stop here rather than loop against a screen that no longer has it.
        this._logger.info('Audio CAPTCHA rejected (retry screen)', { attempt, botId, userId });
        return false;
      }

      if (attempt < maxRetries) {
        await this.refreshCaptcha();
      }
    }

    return false;
  }

  // Clicks the image/audio toggle (same button both directions; label changes
  // per mode/locale). Targeted by its stable data-tid, scoped to the dialog.
  private async switchToAudioCaptcha(): Promise<boolean> {
    const dialog = this.captchaDialog();
    const toggle = dialog.locator('[data-tid="HIP-Captcha-Switch-Type"]').first();
    if (await toggle.isVisible({ timeout: 1500 }).catch(() => false)) {
      await toggle.click().catch(() => undefined);
      await this.page.waitForTimeout(1000);
      return true;
    }

    // Fallback: localized accessible name, as a button or a link.
    const byText = dialog.getByRole('button', { name: /áudio|audio|som|sound/i })
      .or(dialog.getByRole('link', { name: /áudio|audio|som|sound/i })).first();
    if (await byText.isVisible({ timeout: 1000 }).catch(() => false)) {
      await byText.click().catch(() => undefined);
      await this.page.waitForTimeout(1000);
      return true;
    }

    return false;
  }

  // Clicks the play control so the audio challenge is emitted to the system sink
  // (captured by ffmpeg from the PulseAudio monitor). Prefers the stable data-tid
  // observed on the widget; falls back to the localized accessible name.
  private async playCaptchaAudio(): Promise<void> {
    const play = this.captchaDialog().locator('[data-tid="HIP-Captcha-Audio-Play"]').first();
    const target = (await play.isVisible({ timeout: 1500 }).catch(() => false))
      ? play
      : this.captchaDialog().getByRole('button', { name: /reproduzir|tocar|play|ouvir|listen/i }).first();
    if (await target.isVisible({ timeout: 1500 }).catch(() => false)) {
      await target.click().catch(() => undefined);
      // Brief lead so the capture starts after playback begins.
      await this.page.waitForTimeout(500);
    }
  }

  // Capture ONLY the challenge image (never the full page) so meeting
  // content is never sent to the solver. Targeted by data-tid, scoped to dialog.
  private async captureCaptchaImage(): Promise<string | null> {
    const img = this.captchaDialog()
      .locator('[data-tid="HIP-Captcha-Image"], img[alt*="captcha" i]')
      .first();
    if (await img.isVisible({ timeout: 500 }).catch(() => false)) {
      const buffer = await img.screenshot({ type: 'png' });
      return buffer.toString('base64');
    }

    // Fallback: the first visible image inside the verification dialog.
    const dialogImg = this.captchaDialog().locator('img').first();
    if (await dialogImg.isVisible({ timeout: 500 }).catch(() => false)) {
      const buffer = await dialogImg.screenshot({ type: 'png' });
      return buffer.toString('base64');
    }

    return null;
  }

  private async submitCaptchaAnswer(answer: string): Promise<void> {
    await this.captchaInput().fill(answer.trim());
    // The submit button is disabled until the field has text; Playwright's click
    // auto-waits for it to become enabled after the fill above.
    const submit = this.captchaDialog().locator('[data-tid="HIP-Captcha-Submit"]').first();
    await submit.click();
  }

  private async refreshCaptcha(): Promise<void> {
    const refresh = this.captchaDialog().locator('[data-tid="HIP-Captcha-Refresh"]').first();
    if (await refresh.isVisible({ timeout: 1000 }).catch(() => false)) {
      await refresh.click();
      await this.page.waitForTimeout(1000);
    }
  }

  private async recordMeetingPageWithFFmpeg(
    { teamId, userId, eventId, botId, uploader }:
    { teamId: string, userId: string, eventId?: string, botId?: string, uploader: IUploader }
  ): Promise<void> {
    // Use config max recording duration (3 hours default) - only for safety
    const duration = config.maxRecordingDuration * 60 * 1000;
    this._logger.info(`Recording max duration set to ${duration / 60000} minutes (safety limit only)`);

    // Use the same temp folder as Google Meet bot (has proper permissions)
    const tempFolder = path.join(process.cwd(), 'dist', '_tempvideo');
    const outputPath = path.join(tempFolder, `recording-${botId || Date.now()}${config.uploaderFileExtension}`);

    this._logger.info('Starting ffmpeg recording...', { outputPath, duration });

    // Verify PulseAudio is ready before starting FFmpeg
    this._logger.info('Verifying PulseAudio status before starting FFmpeg...');
    try {
      const { stdout: paStatus } = await execAsync('pactl list sources short');
      this._logger.info('PulseAudio sources available:', paStatus.trim() || '(empty - no sources found)');

      if (!paStatus.includes('virtual_output.monitor')) {
        this._logger.error('WARNING: virtual_output.monitor not found in PulseAudio sources!');
        this._logger.info('Attempting to restart PulseAudio and recreate virtual audio device...');

        // Try to restart PulseAudio
        try {
          await execAsync('pulseaudio --kill || true');
          await new Promise(resolve => setTimeout(resolve, 1000));
          await execAsync('pulseaudio -D --exit-idle-time=-1 --log-level=info');
          await new Promise(resolve => setTimeout(resolve, 1000));
          this._logger.info('Restarted PulseAudio');

          // Recreate the null sink
          await execAsync('pactl load-module module-null-sink sink_name=virtual_output sink_properties=device.description="Virtual_Output"');
          await execAsync('pactl set-default-sink virtual_output');
          this._logger.info('Recreated virtual_output sink and monitor');

          // Verify it worked
          const { stdout: newStatus } = await execAsync('pactl list sources short');
          this._logger.info('PulseAudio sources after restart:', newStatus.trim());
        } catch (err) {
          this._logger.error('Failed to restart PulseAudio or recreate virtual audio device:', err);
        }
      }
    } catch (err) {
      this._logger.error('Error checking PulseAudio status:', err);
    }

    // Create and start ffmpeg recorder
    const recorder = new FFmpegRecorder(outputPath, this._logger);

    // Track FFmpeg status
    let ffmpegFailed = false;
    let ffmpegError: Error | null = null;
    let recordingStartedAt: number | undefined;
    // Hoisted out of the try block so the matching finally can set it to signal
    // the silence detector (declared inside the try) to stop on its next tick.
    let meetingEnded = false;

    if (botId) {
      // Unlike Zoom/Google Meet, the stop signal here is a Node-side flag, not a
      // browser bridge — no page.evaluate needed.
      registerActiveRecording(botId, async () => {
        meetingEnded = true;
      });
    }

    try {
      await recorder.start();
      recordingStartedAt = Date.now();
      const startedAt = recordingStartedAt;
      this._logger.info('FFmpeg recording started successfully');

      // Monitor FFmpeg process - if it dies, stop recording immediately
      recorder.onProcessExit((code) => {
        if (code !== 0 && code !== null) {
          this._logger.error('FFmpeg died unexpectedly during recording', { exitCode: code });
          ffmpegFailed = true;
          ffmpegError = new Error(`FFmpeg exited with code ${code} during recording`);
        }
      });

      // Set up browser-based inactivity detection (meetingEnded declared outside the try block)
      await this.page.exposeFunction('screenAppMeetEnd', () => {
        this._logger.info('Meeting ended signal received from browser');
        meetingEnded = true;
      });

      // Capture and forward browser console logs to Node.js logger
      this.page.on('console', async msg => {
        try {
          await browserLogCaptureCallback(this._logger, msg);
        } catch(err) {
          this._logger.info('Playwright chrome logger: Failed to log browser messages...', err instanceof Error ? err.message : String(err));
        }
      });

      // Start audio silence detection (runs in parallel with participant detection)
      // Convert inactivityLimit from minutes to milliseconds
      const inactivityLimitMs = config.inactivityLimit * 60 * 1000;
      const idleFallbackMs = config.idleFallbackExtraMinutes * 60 * 1000;

      const monitorAudioSilence = async () => {
        try {
          this._logger.info('Starting audio silence detection for Microsoft Teams', {
            inactivityLimitMs,
            inactivityLimitMinutes: inactivityLimitMs / 60000
          });
          let consecutiveSilentChecks = 0;
          let idleNotified = false;
          const checkIntervalSeconds = 5;
          const checksNeeded = Math.ceil(inactivityLimitMs / 1000 / checkIntervalSeconds); // e.g., 120000ms / 1000 / 5 = 24 checks
          const fallbackChecksNeeded = checksNeeded + Math.ceil(idleFallbackMs / 1000 / checkIntervalSeconds);

          const checkInterval = setInterval(async () => {
            // If the meeting ended via any other path (browser signal, page-state change,
            // browser close, error), stop polling. Without this, the interval kept running
            // for minutes after the recording was uploaded and the browser was closed.
            if (meetingEnded) {
              clearInterval(checkInterval);
              return;
            }
            try {
              // Sample audio from virtual_output.monitor and check if it's silent
              // Use parec to capture 1 second of audio and check the peak level
              const { stdout } = await execAsync(
                'timeout 1 parec --device=virtual_output.monitor --format=s16le --rate=16000 --channels=1 2>/dev/null | ' +
                'od -An -td2 -v | awk \'BEGIN{max=0} {for(i=1;i<=NF;i++) {val=($i<0)?-$i:$i; if(val>max) max=val}} END{print max}\''
              );

              // Get peak audio level (0-32767 for 16-bit audio)
              const peakLevel = parseInt(stdout.trim()) || 0;
              const silenceThreshold = 200; // Adjust this threshold as needed

              this._logger.debug('Audio level check', { peakLevel, threshold: silenceThreshold });

              // Check if audio is silent (low peak level)
              if (peakLevel < silenceThreshold) {
                consecutiveSilentChecks++;
                this._logger.info(`Silence detected: ${consecutiveSilentChecks}/${checksNeeded} checks`, { peakLevel });

                if (!idleNotified && consecutiveSilentChecks >= checksNeeded) {
                  idleNotified = true;
                  this._logger.warn('Audio silence threshold reached; notifying idle instead of ending immediately', {
                    userId,
                    teamId,
                    silenceDurationMs: inactivityLimitMs,
                    silenceDurationMinutes: inactivityLimitMs / 60000,
                    finalPeakLevel: peakLevel,
                  });
                  void notifyMeetingIdle(
                    { botId: botId ?? '', event: 'idle-started', silenceSeconds: Math.round(inactivityLimitMs / 1000) },
                    this._logger,
                  ).catch((error) => this._logger.warn('Failed to deliver idle notification', error));
                }

                if (consecutiveSilentChecks >= fallbackChecksNeeded) {
                  this._logger.warn('Silence exceeded fallback window, ending Microsoft Teams meeting', {
                    userId,
                    teamId,
                    checksNeeded: fallbackChecksNeeded,
                    checksDetected: consecutiveSilentChecks
                  });
                  clearInterval(checkInterval);
                  meetingEnded = true;
                }
              } else {
                // Reset counter if we detect audio
                if (consecutiveSilentChecks > 0) {
                  this._logger.info('Audio detected, resetting silence counter', { peakLevel });
                }
                if (idleNotified) {
                  void notifyMeetingIdle({ botId: botId ?? '', event: 'idle-cleared' }, this._logger).catch((error) =>
                    this._logger.warn('Failed to deliver idle notification', error),
                  );
                  idleNotified = false;
                }
                consecutiveSilentChecks = 0;
              }
            } catch (err) {
              this._logger.error('Error checking audio level:', err);
              // Don't fail the entire detection on a single error
            }
          }, 5000); // Check every 5 seconds

        } catch (error) {
          this._logger.error('Failed to initialize audio silence detection:', error);
          this._logger.warn('Will rely on participant detection only');
        }
      };

      // Start silence monitoring after delay
      setTimeout(() => {
        monitorAudioSilence();
      }, config.activateInactivityDetectionAfter * 60 * 1000);

      // Inject inactivity detection script
      await this.page.evaluate(
        ({ activateAfterMinutes, loneParticipantExitDelayMs, maxDuration }: { activateAfterMinutes: number, loneParticipantExitDelayMs: number, maxDuration: number }) => {
          // Max duration timeout - safety limit (3 hours default in production)
          setTimeout(() => {
            console.log(`Max recording duration (${maxDuration / 60000} minutes) reached, ending meeting`);
            (window as any).screenAppMeetEnd();
          }, maxDuration);
          console.log(`Max duration timeout set to ${maxDuration / 60000} minutes (safety limit)`);

          console.log('Activating participant count detection...');

          const recordingStartedAt = Date.now();
          const initialAloneGraceMs = activateAfterMinutes * 60 * 1000;
          let hasSeenOtherParticipant = false;
          let aloneSince: number | null = null;
          let lastParticipantDetectionLogAt = 0;

          const shouldStopForParticipantCount = (participants: number) => {
            const now = Date.now();
            if (participants >= 2) {
              hasSeenOtherParticipant = true;
              aloneSince = null;
              return false;
            }

            if (hasSeenOtherParticipant) {
              if (aloneSince === null) {
                aloneSince = now;
                console.log('Bot is alone after previously seeing participants; waiting before ending recording.');
              }
              return now - aloneSince >= loneParticipantExitDelayMs;
            }

            return now - recordingStartedAt >= initialAloneGraceMs;
          };

          const normalizeText = (text: string) => text.replace(/\s+/g, ' ').trim();

          const parseParticipantCount = (text: string): number | undefined => {
            const normalized = normalizeText(text);
            const patterns = [
              /\b(?:people|participants?|teilnehm(?:er|ende)?|personen)\D{0,30}(\d{1,3})\b/i,
              /\b(\d{1,3})\D{0,30}(?:people|participants?|teilnehm(?:er|ende)?|personen)\b/i,
            ];

            for (const pattern of patterns) {
              const match = normalized.match(pattern);
              if (match) {
                const value = Number(match[1]);
                if (Number.isFinite(value)) {
                  return value;
                }
              }
            }

            if (/^\D*\d{1,3}\D*$/.test(normalized) && normalized.length <= 16) {
              const match = normalized.match(/\d{1,3}/);
              const value = match ? Number(match[0]) : NaN;
              if (Number.isFinite(value)) {
                return value;
              }
            }

            return undefined;
          };

          const isExplicitEmptyMeetingText = (text: string): boolean => {
            const emptyMeetingPatterns = [
              /\b0\D{0,30}(?:people|participants?|teilnehm(?:er|ende)?|personen)\D{0,80}(?:in|inside|joined|here|meeting|call|besprechung|anruf)\b/i,
              /\b(?:no|zero)\D{0,30}(?:one|one else|people|participants?)\D{0,80}(?:in|inside|joined|here|meeting|call)\b/i,
              /\b(?:keine|niemand)\D{0,80}(?:teilnehm(?:er|ende)?|personen|hier|besprechung|anruf)\b/i,
            ];

            return emptyMeetingPatterns.some(pattern => pattern.test(text));
          };

          const getTeamsMeetingState = (): 'active' | 'alone' | 'empty' | 'ended' => {
            const bodyText = normalizeText(document.body.innerText || '');

            const endedPhrases = [
              'the meeting has ended',
              'this meeting has ended',
              'meeting has been ended',
              'call ended',
              'you have been removed',
              'you’ve been removed',
              'removed from the meeting',
              'besprechung wurde beendet',
              'anruf beendet',
              'sie wurden entfernt',
              'du wurdest entfernt',
            ];

            if (endedPhrases.some(phrase => bodyText.toLowerCase().includes(phrase))) {
              return 'ended';
            }

            if (isExplicitEmptyMeetingText(bodyText)) {
              return 'empty';
            }

            const alonePhrases = [
              'you\'re the only one here',
              'you’re the only one here',
              'you are the only one here',
              'you\'re the only one in this meeting',
              'you’re the only one in this meeting',
              'you are the only one in this meeting',
              'only one in this meeting',
              'only you are here',
              'no one else is here',
              'waiting for others to join',
              'sie sind der einzige',
              'du bist der einzige',
              'sie sind die einzige person',
              'du bist die einzige person',
              'warten auf andere',
            ];

            return alonePhrases.some(phrase => bodyText.toLowerCase().includes(phrase)) ? 'alone' : 'active';
          };

          const getParticipantCount = (): { count?: number; samples: string[] } => {
            const selectors = [
              'button[data-tid*="roster" i]',
              '[data-tid*="roster" i]',
              'button[id*="roster" i]',
              '[id*="roster" i]',
              'button[aria-label*="people" i]',
              '[aria-label*="people" i]',
              'button[aria-label*="participant" i]',
              '[aria-label*="participant" i]',
              'button[aria-label*="teilnehm" i]',
              '[aria-label*="teilnehm" i]',
              'button[aria-label*="personen" i]',
              '[aria-label*="personen" i]',
            ];

            const candidates = Array.from(document.querySelectorAll(selectors.join(',')));
            const samples: string[] = [];

            for (const element of candidates) {
              const searchRoots = [
                element,
                element.parentElement,
                element.parentElement?.parentElement,
              ].filter(Boolean) as Element[];

              for (const root of searchRoots) {
                const text = normalizeText([
                  root.getAttribute('aria-label') ?? '',
                  root.getAttribute('title') ?? '',
                  root.getAttribute('data-tid') ?? '',
                  root.textContent ?? '',
                ].join(' '));

                if (!text) continue;
                if (samples.length < 6) {
                  samples.push(text.slice(0, 140));
                }

                const count = parseParticipantCount(text);
                if (typeof count === 'number') {
                  return { count, samples };
                }
              }
            }

            const bodyLines = (document.body.innerText || '')
              .split(/\n+/)
              .map(normalizeText)
              .filter(text => (
                text.length > 0 &&
                /(?:people|participants?|teilnehm|personen|meeting|call|besprechung|anruf)/i.test(text)
              ));

            for (const text of bodyLines) {
              if (samples.length < 6) {
                samples.push(text.slice(0, 140));
              }

              if (isExplicitEmptyMeetingText(text)) {
                return { count: 0, samples };
              }

              if (/(?:people|participants?|teilnehm|personen)/i.test(text)) {
                const count = parseParticipantCount(text);
                if (typeof count === 'number') {
                  return { count, samples };
                }
              }
            }

            return { samples };
          };

          const interval = setInterval(() => {
            try {
              const meetingState = getTeamsMeetingState();
              if (meetingState === 'ended') {
                console.log('Teams meeting ended page state detected, ending recording.');
                clearInterval(interval);
                (window as any).screenAppMeetEnd();
                return;
              }

              const { count, samples } = getParticipantCount();
              let inferredCount = count;
              if (typeof inferredCount !== 'number') {
                if (meetingState === 'empty') {
                  inferredCount = 0;
                } else if (meetingState === 'alone') {
                  inferredCount = 1;
                }
              }

              if (typeof inferredCount !== 'number') {
                const now = Date.now();
                if (now - lastParticipantDetectionLogAt > 30000) {
                  console.log('Teams participant count not detected yet', { samples });
                  lastParticipantDetectionLogAt = now;
                }
                return;
              }

              if (!shouldStopForParticipantCount(inferredCount)) {
                return;
              }

              console.log('Bot is alone, ending Teams recording', { inferredCount, meetingState });
              clearInterval(interval);
              (window as any).screenAppMeetEnd();
            } catch (error) {
              console.error('Participant detection error:', error);
            }
          }, 2000);
        },
        {
          activateAfterMinutes: config.activateInactivityDetectionAfter,
          loneParticipantExitDelayMs: config.loneParticipantExitDelaySeconds * 1000,
          maxDuration: duration,
        }
      );

      // Wait for either timeout, meeting end, or FFmpeg failure
      while (!meetingEnded && !ffmpegFailed && (Date.now() - startedAt) < duration) {
        await new Promise(resolve => setTimeout(resolve, 1000));
      }

      this._logger.info('Recording period ended', {
        meetingEnded,
        ffmpegFailed,
        recordedDuration: Math.floor((Date.now() - startedAt) / 1000) + 's'
      });

      // If FFmpeg failed during recording, throw the error
      if (ffmpegFailed && ffmpegError) {
        throw ffmpegError;
      }

    } catch (error) {
      // If recorder.start() failed or any other error occurred, mark FFmpeg as failed
      this._logger.error('Error during recording:', error);
      ffmpegFailed = true;
      ffmpegError = error instanceof Error ? error : new Error(String(error));
      // Re-throw to be caught by outer try/catch in joinMeeting
      throw error;
    } finally {
      if (botId) clearActiveRecording(botId);

      // Signal the silence detector to stop on its next tick. Without this, the
      // detector keeps polling parec every 5s for minutes after the recording is done
      // (until silence threshold eventually fires or pod exits).
      meetingEnded = true;

      // Always stop ffmpeg
      this._logger.info('Stopping ffmpeg recording...');
      await recorder.stop();

      // Stage the recorded file for upload (the actual remote upload happens in
      // join()'s handleUpload after joinMeeting returns).
      this._logger.info('Staging recorded file for upload...', { outputPath });

      let staged = false;
      if (fs.existsSync(outputPath)) {
        if (recordingStartedAt) {
          const recordedDurationSeconds = Math.max(1, Math.round((Date.now() - recordingStartedAt) / 1000));
          uploader.setRecordingDuration(recordedDurationSeconds);
        }

        const fileBuffer = fs.readFileSync(outputPath);
        await uploader.saveDataToTempFile(fileBuffer);

        // Remove the ffmpeg output file (the staged copy now lives in the uploader temp)
        fs.unlinkSync(outputPath);
        this._logger.info('Recording staged to temp; ffmpeg output file removed');
        staged = true;
      } else {
        this._logger.error('Recording file not found!', { outputPath });
      }

      // Close browser
      this._logger.info('Closing the browser...');
      await this.page.context().browser()?.close();

      // Log final status. The real remote upload + true completion is logged in
      // join() after handleUpload: 'Recording and upload completed successfully'.
      if (ffmpegFailed) {
        this._logger.error('Recording failed due to FFmpeg error', { botId, eventId, userId, teamId });
      } else if (!staged) {
        this._logger.error('Recording file missing; nothing to upload', { botId, eventId, userId, teamId });
      } else {
        this._logger.info('Recording captured and staged; finalizing upload next...', { botId, eventId, userId, teamId });
      }
    }
  }
}
