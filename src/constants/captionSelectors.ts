import type { CaptionDomSelectors } from '../lib/captions';

// Seletores do DOM para as legendas nativas. Ainda não validados numa reunião
// real: ajuste aqui quando o log avisar que nenhuma legenda foi lida (ele lista
// os rótulos de legenda/idioma/menu presentes na página). Prefira atributos
// (`aria-label`, `data-tid`) a classes, que mudam a cada versão.

export const TEAMS_CAPTION_SELECTORS: CaptionDomSelectors = {
  toggle: [
    '[data-tid="closed-captions-button"]',
    '#closed-captions-button',
    'button[aria-label*="live captions" i]',
    'button[aria-label*="legendas ao vivo" i]',
    '[role="menuitem"][aria-label*="captions" i]',
    '[role="menuitem"][aria-label*="legendas" i]',
  ],
  menuPath: [
    ['#callingButtons-showMoreBtn', '[data-tid="LanguageSpeechMenuControl-id"]'],
    ['[data-tid="more-button"]', '[data-tid="LanguageSpeechMenuControl-id"]'],
    ['#callingButtons-showMoreBtn'],
    ['[data-tid="more-button"]'],
  ],
  region: [
    '[data-tid="closed-captions-renderer"]',
    '[data-tid="closed-caption-renderer-wrapper"]',
    '[aria-label*="captions" i][role="log"]',
    '[aria-label*="legendas" i][role="log"]',
  ],
  item: [
    '[data-tid="closed-caption-message"]',
    '.fui-ChatMessageCompact',
    '[role="listitem"]',
  ],
  name: ['[data-tid="author"]', '.ui-chat__message__author', '[class*="author" i]'],
  text: ['[data-tid="closed-caption-text"]', '[class*="caption-text" i]'],
};

export const MEET_CAPTION_SELECTORS: CaptionDomSelectors = {
  toggle: [
    'button[aria-label*="Turn on captions" i]',
    'button[aria-label*="Ativar legendas" i]',
    'button[aria-label*="captions" i][aria-pressed="false"]',
    'button[aria-label*="legendas" i][aria-pressed="false"]',
  ],
  menuPath: [],
  // Atalho do Meet para ligar/desligar as legendas.
  shortcut: 'c',
  region: [
    '[role="region"][aria-label*="Captions" i]',
    '[role="region"][aria-label*="Legendas" i]',
    '[jsname="dsyhDe"]',
  ],
  item: ['[role="region"] > div > div', ':scope > div > div', ':scope > div'],
  name: ['[data-self-name]', 'img + span', 'img + div'],
  text: ['[jsname="tgaKEf"]', 'span[jsname]'],
};
