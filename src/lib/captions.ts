import type { Page } from 'playwright';
import type { Logger } from 'winston';

// Legendas nativas da plataforma (Meet/Teams), coletadas do DOM durante a
// gravação: cada bloco de legenda vira { nome, texto, início, fim } medido a
// partir do início da gravação. Recurso experimental, ligado por flag da API.
// Tudo aqui é best-effort: uma falha nunca afeta a gravação.

export interface CaptionEntry {
  name: string;
  text: string;
  startMs: number;
  endMs: number;
}

/** Seletores de cada plataforma, isolados para ajustar quando a interface mudar. */
export interface CaptionDomSelectors {
  /** Botões que ligam as legendas, só clicados se ainda não estiverem ativos. */
  toggle: string[];
  /** Menus que precisam ser abertos antes do botão (Teams: "Mais" → "Idioma e fala"). */
  menuPath: string[][];
  /** Tecla de atalho que liga as legendas, tentada se nenhum botão funcionar. */
  shortcut?: string;
  /** Contêiner das legendas. */
  region: string[];
  /** Cada bloco de legenda dentro do contêiner. */
  item: string[];
  /** Nome de quem fala dentro do bloco. */
  name: string[];
  /** Texto falado dentro do bloco. */
  text: string[];
}

export interface CaptionSample {
  /** Id estável do bloco no DOM (marcado pelo coletor). */
  id: string;
  name: string;
  text: string;
}

const SAMPLE_INTERVAL_MS = 1_000;
const MAX_NAME_LENGTH = 80;
const MAX_ENTRIES = 20_000;
/** Sobreposição mínima (caracteres) para emendar um bloco que rolou para fora da tela. */
const MIN_OVERLAP = 12;

/**
 * Junta duas leituras do mesmo bloco. A legenda cresce, é corrigida enquanto a
 * pessoa fala e, em blocos longos, o começo sai da tela: a leitura nova pode
 * conter a anterior, estar contida nela, emendar pelo fim ou reescrevê-la.
 */
export function mergeCaptionText(previous: string, next: string): string {
  if (!previous) return next;
  if (!next || previous === next) return previous;
  if (next.includes(previous)) return next;
  if (previous.includes(next)) return previous;
  for (let size = Math.min(previous.length, next.length); size >= MIN_OVERLAP; size--) {
    if (previous.endsWith(next.slice(0, size))) return previous + next.slice(size);
  }
  // Correção da própria plataforma: mesmo começo, final reescrito.
  const common = commonPrefixLength(previous, next);
  if (common >= Math.min(previous.length, next.length) * 0.5) return next;
  return `${previous} ${next}`;
}

function commonPrefixLength(a: string, b: string): number {
  let index = 0;
  while (index < a.length && index < b.length && a[index] === b[index]) index++;
  return index;
}

/** Acumula leituras dos blocos visíveis; um bloco termina quando some da tela. */
export class CaptionRecorder {
  private readonly open = new Map<string, CaptionEntry>();
  private readonly closed: CaptionEntry[] = [];

  constructor(private readonly startedAt: number) {}

  sample(items: CaptionSample[], at = Date.now()): void {
    const now = Math.max(0, at - this.startedAt);
    const seen = new Set<string>();
    for (const item of items) {
      const text = cleanText(item.text);
      if (!text) continue;
      seen.add(item.id);
      const entry = this.open.get(item.id);
      if (!entry) {
        this.open.set(item.id, { name: cleanName(item.name), text, startMs: now, endMs: now });
        continue;
      }
      const merged = mergeCaptionText(entry.text, text);
      if (merged !== entry.text) {
        entry.text = merged;
        entry.endMs = now;
      }
      if (!entry.name && item.name) entry.name = cleanName(item.name);
    }
    for (const [id, entry] of this.open) {
      if (seen.has(id)) continue;
      this.close(entry);
      this.open.delete(id);
    }
  }

  finish(): CaptionEntry[] {
    for (const entry of this.open.values()) this.close(entry);
    this.open.clear();
    return [...this.closed].sort((a, b) => a.startMs - b.startMs);
  }

  private close(entry: CaptionEntry): void {
    if (this.closed.length >= MAX_ENTRIES) return;
    this.closed.push(entry);
  }
}

function cleanName(value: string): string {
  return value.replace(/\s+/g, ' ').trim().slice(0, MAX_NAME_LENGTH);
}

function cleanText(value: string): string {
  return value.replace(/\s+/g, ' ').trim();
}

export interface CaptionCollector {
  stop(): CaptionEntry[];
}

/**
 * Liga as legendas e lê os blocos visíveis a cada segundo. Sem nenhuma legenda
 * depois de dois minutos, registra uma vez os rótulos presentes na página, para
 * ajustar os seletores.
 */
export function startCaptionCollector(options: {
  page: Page;
  startedAt: number;
  selectors: CaptionDomSelectors;
  logger: Logger;
  platform: string;
}): CaptionCollector {
  const { page, selectors, logger, platform } = options;
  const recorder = new CaptionRecorder(options.startedAt);
  let stopped = false;
  let running = false;
  let captionSeen = false;
  let warned = false;
  let failures = 0;
  let enableAttempts = 0;
  let enabled = false;

  const tick = async (): Promise<void> => {
    if (stopped || running) return;
    running = true;
    try {
      // Liga uma vez só (clicar de novo desligaria); enquanto o botão não
      // aparece, tenta a cada 20 s, e o atalho de teclado fica por último.
      if (!enabled && !captionSeen && enableAttempts < 6 && (Date.now() - options.startedAt) / 20_000 >= enableAttempts) {
        enableAttempts += 1;
        enabled = await enableCaptions(page, selectors, logger, platform, enableAttempts >= 3);
      }
      const sample = await page.evaluate(readCaptionsFromDom, selectors);
      recorder.sample(sample.items);
      if (sample.items.length > 0 && !captionSeen) {
        captionSeen = true;
        logger.info('Captions: first caption read', { platform });
      }
      if (!captionSeen && !warned && Date.now() - options.startedAt > 120_000) {
        warned = true;
        logger.warn('Captions: nothing read yet; check the caption selectors', { platform, hints: sample.hints });
      }
      failures = 0;
    } catch (error) {
      failures += 1;
      if (failures === 5) {
        logger.warn('Caption sampling keeps failing', { platform, error: error instanceof Error ? error.message : String(error) });
      }
    } finally {
      running = false;
    }
  };

  const timer = setInterval(() => void tick(), SAMPLE_INTERVAL_MS);
  logger.info('Caption collector started', { platform });

  return {
    stop() {
      stopped = true;
      clearInterval(timer);
      const captions = recorder.finish();
      if (captions.length === 0) logger.warn('Captions are empty', { platform });
      else logger.info('Captions collected', { platform, entries: captions.length });
      return captions;
    },
  };
}

/** true quando as legendas já estavam ligadas ou algo foi acionado para ligá-las. */
async function enableCaptions(page: Page, selectors: CaptionDomSelectors, logger: Logger, platform: string, useShortcut: boolean): Promise<boolean> {
  const active = await page.evaluate(captionsAlreadyOn, selectors).catch(() => false);
  if (active) return true;
  for (const path of [[], ...selectors.menuPath]) {
    try {
      for (const step of path) {
        const menu = page.locator(step).first();
        if (!(await menu.isVisible({ timeout: 500 }).catch(() => false))) throw new Error(`menu não encontrado: ${step}`);
        await menu.click({ timeout: 2_000 });
      }
      for (const selector of selectors.toggle) {
        const button = page.locator(selector).first();
        if (!(await button.isVisible({ timeout: 500 }).catch(() => false))) continue;
        if ((await button.getAttribute('aria-pressed').catch(() => null)) === 'true') return true;
        await button.click({ timeout: 2_000 });
        logger.info('Captions: toggle clicked', { platform, selector });
        return true;
      }
    } catch {
      // Caminho de menu indisponível: tenta o próximo.
    }
    // Fecha um menu aberto por engano antes do próximo caminho.
    if (path.length > 0) await page.keyboard.press('Escape').catch(() => undefined);
  }
  if (useShortcut && selectors.shortcut) {
    await page.keyboard.press(selectors.shortcut).catch(() => undefined);
    logger.info('Captions: shortcut pressed', { platform, shortcut: selectors.shortcut });
    return true;
  }
  return false;
}

// Roda no navegador: não pode referenciar nada de fora da função.
function captionsAlreadyOn(selectors: CaptionDomSelectors): boolean {
  return selectors.region.some((selector) => document.querySelector(selector) != null);
}

// Roda no navegador: não pode referenciar nada de fora da função.
function readCaptionsFromDom(selectors: CaptionDomSelectors): { items: { id: string; name: string; text: string }[]; hints: string[] } {
  const w = window as unknown as { __callfredCaptionSeq?: number };
  const first = (root: Element, list: string[]): Element | null => {
    for (const selector of list) {
      const found = root.querySelector(selector);
      if (found) return found;
    }
    return null;
  };
  const textOf = (element: Element | null): string => (element ? (element as HTMLElement).innerText || element.textContent || '' : '').trim();

  let region: Element | null = null;
  for (const selector of selectors.region) {
    region = document.querySelector(selector);
    if (region) break;
  }

  const items: { id: string; name: string; text: string }[] = [];
  if (region) {
    let blocks: Element[] = [];
    for (const selector of selectors.item) {
      blocks = Array.from(region.querySelectorAll(selector));
      if (blocks.length > 0) break;
    }
    for (const block of blocks) {
      let id = block.getAttribute('data-callfred-caption');
      if (!id) {
        w.__callfredCaptionSeq = (w.__callfredCaptionSeq ?? 0) + 1;
        id = String(w.__callfredCaptionSeq);
        block.setAttribute('data-callfred-caption', id);
      }
      let name = textOf(first(block, selectors.name));
      let text = textOf(first(block, selectors.text));
      // Sem seletores que casem: primeira linha é o nome, o resto é a fala.
      if (!text) {
        const lines = textOf(block).split('\n').map((line) => line.trim()).filter(Boolean);
        if (lines.length >= 2) {
          name = name || lines[0];
          text = lines.slice(1).join(' ');
        }
      }
      if (text) items.push({ id, name: name.split('\n')[0] ?? '', text });
    }
  }

  const hints = Array.from(new Set(Array.from(document.querySelectorAll('[aria-label], [data-tid]'))
    .map((element) => element.getAttribute('data-tid') ?? element.getAttribute('aria-label') ?? '')
    .filter((label) => /caption|legend|subtit|language|idioma|more|mais/i.test(label))))
    .slice(0, 60);

  return { items, hints };
}
