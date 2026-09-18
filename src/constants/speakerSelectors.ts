import type { SpeakerDomSelectors } from '../lib/speakerTimeline';

// Seletores do DOM para a linha do tempo de oradores. As interfaces do Teams e
// do Meet mudam com frequência: ajuste aqui quando o log avisar que nenhum
// orador foi detectado (ele lista os data-tid presentes na página).

export const TEAMS_SPEAKER_SELECTORS: SpeakerDomSelectors = {
  speaking: [
    '[data-tid="voice-level-stream-outline"][class*="speaking" i]',
    '[data-is-speaking="true"]',
    '[data-tid*="speaking" i]',
    '[aria-label*="is speaking" i]',
    '[aria-label*="está falando" i]',
  ],
  tile: [
    '[data-tid="video-tile"]',
    '[data-cid="calling-participant-stream"]',
    '[data-stream-type]',
    '[role="listitem"]',
  ],
  name: [
    '[data-tid="participant-name"]',
    '[data-tid="display-name"]',
    '[data-tid*="name" i][title]',
    '[title]',
  ],
  participants: [
    '[data-tid="video-tile"] [data-tid="participant-name"]',
    '[data-cid="calling-participant-stream"] [data-tid="participant-name"]',
    '[data-tid^="participantsInCall-"] [title]',
  ],
};

export const MEET_SPEAKER_SELECTORS: SpeakerDomSelectors = {
  speaking: [
    '[data-participant-id] [data-is-speaking="true"]',
    '[data-participant-id][data-is-speaking="true"]',
    '[data-participant-id] [aria-label*="is speaking" i]',
    '[data-participant-id] [aria-label*="está falando" i]',
  ],
  tile: ['[data-participant-id]', '[data-requested-participant-id]'],
  name: ['[data-self-name]', '.notranslate'],
  participants: ['[data-participant-id] [data-self-name]', '[data-requested-participant-id] [data-self-name]'],
};
