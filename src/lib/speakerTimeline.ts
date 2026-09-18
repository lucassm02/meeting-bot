import type { Page } from 'playwright';
import type { Logger } from 'winston';

// Linha do tempo de quem fala durante a gravação, lida do DOM da chamada.
// O Node amostra a página periodicamente e mede os tempos a partir do início
// da gravação no mesmo relógio que a iniciou, então os intervalos se alinham
// ao arquivo gravado. Tudo aqui é best-effort: uma falha nunca afeta a gravação.

export interface SpeakerInterval {
  name: string;
  startMs: number;
  endMs: number;
}

export interface SpeakerTimelineData {
  speakerTimeline: SpeakerInterval[];
  participants: string[];
}

/** Seletores de cada plataforma, isolados para ajustar quando a interface mudar. */
export interface SpeakerDomSelectors {
  /** Elementos que marcam o participante que está falando agora. */
  speaking: string[];
  /** Contêiner do participante (tile/linha do roster) a partir do marcador. */
  tile: string[];
  /** Elemento com o nome dentro do contêiner (texto, `title` ou `aria-label`). */
  name: string[];
  /** Elementos com nomes de participantes visíveis na chamada. */
  participants: string[];
}

const SAMPLE_INTERVAL_MS = 750;
/** Pausas menores que isso entre falas do mesmo nome viram um intervalo só. */
export const MERGE_GAP_MS = 1_500;
const MAX_NAME_LENGTH = 80;

/** Acumula amostras "quem fala agora" em intervalos fundidos. */
export class SpeakerTimelineRecorder {
  private readonly open = new Map<string, SpeakerInterval>();
  private readonly closed: SpeakerInterval[] = [];
  private readonly participants = new Set<string>();

  constructor(private readonly startedAt: number) {}

  sample(speaking: string[], participants: string[], at = Date.now()): void {
    const now = Math.max(0, at - this.startedAt);
    const active = new Set(speaking.map(cleanName).filter(Boolean));
    for (const name of participants.map(cleanName).filter(Boolean)) this.participants.add(name);
    for (const name of active) this.participants.add(name);

    for (const [name, interval] of this.open) {
      if (active.has(name)) {
        interval.endMs = now;
      } else if (now - interval.endMs > MERGE_GAP_MS) {
        this.closed.push(interval);
        this.open.delete(name);
      }
    }
    for (const name of active) {
      if (!this.open.has(name)) this.open.set(name, { name, startMs: now, endMs: now });
    }
  }

  finish(): SpeakerTimelineData {
    const speakerTimeline = [...this.closed, ...this.open.values()]
      .filter((interval) => interval.endMs > interval.startMs)
      .sort((a, b) => a.startMs - b.startMs);
    return { speakerTimeline, participants: [...this.participants] };
  }
}

function cleanName(value: string): string {
  return value.replace(/\s+/g, ' ').trim().slice(0, MAX_NAME_LENGTH);
}

export interface SpeakerTimelineCollector {
  stop(): SpeakerTimelineData;
}

/**
 * Amostra a página a cada 750 ms. Sem nenhum orador detectado depois de um
 * minuto com participantes visíveis, registra uma vez os `data-tid` e
 * atributos presentes, para ajustar os seletores.
 */
export function startSpeakerTimelineCollector(options: {
  page: Page;
  startedAt: number;
  selectors: SpeakerDomSelectors;
  logger: Logger;
  platform: string;
}): SpeakerTimelineCollector {
  const { page, selectors, logger, platform } = options;
  const recorder = new SpeakerTimelineRecorder(options.startedAt);
  let stopped = false;
  let running = false;
  let speakerSeen = false;
  let warned = false;
  let failures = 0;

  const tick = async (): Promise<void> => {
    if (stopped || running) return;
    running = true;
    try {
      const sample = await page.evaluate(readSpeakersFromDom, selectors);
      recorder.sample(sample.speaking, sample.participants);
      if (sample.speaking.length > 0) speakerSeen = true;
      if (!speakerSeen && !warned && Date.now() - options.startedAt > 60_000 && sample.participants.length > 0) {
        warned = true;
        logger.warn('Speaker timeline: no active speaker detected yet; check the selectors', { platform, hints: sample.hints });
      }
      failures = 0;
    } catch (error) {
      failures += 1;
      if (failures === 5) {
        logger.warn('Speaker timeline sampling keeps failing', { platform, error: error instanceof Error ? error.message : String(error) });
      }
    } finally {
      running = false;
    }
  };

  const timer = setInterval(() => void tick(), SAMPLE_INTERVAL_MS);
  logger.info('Speaker timeline collector started', { platform });

  return {
    stop() {
      stopped = true;
      clearInterval(timer);
      const data = recorder.finish();
      if (data.speakerTimeline.length === 0) {
        logger.warn('Speaker timeline is empty; transcription will keep generic speakers', { platform, participants: data.participants.length });
      } else {
        logger.info('Speaker timeline collected', { platform, intervals: data.speakerTimeline.length, participants: data.participants.length });
      }
      return data;
    },
  };
}

// Roda no navegador: não pode referenciar nada de fora da função.
function readSpeakersFromDom(selectors: SpeakerDomSelectors): { speaking: string[]; participants: string[]; hints: string[] } {
  const textOf = (element: Element | null): string => {
    if (!element) return '';
    const direct = (element.getAttribute('title') || element.getAttribute('aria-label') || element.textContent || '').trim();
    return direct.split('\n')[0]?.trim() ?? '';
  };
  const nameIn = (container: Element): string => {
    for (const selector of selectors.name) {
      const found = container.matches(selector) ? container : container.querySelector(selector);
      const text = textOf(found);
      if (text) return text;
    }
    return '';
  };
  const unique = (values: string[]): string[] => [...new Set(values.filter(Boolean))];

  const speaking: string[] = [];
  for (const selector of selectors.speaking) {
    for (const marker of Array.from(document.querySelectorAll(selector))) {
      let tile: Element | null = null;
      for (const tileSelector of selectors.tile) {
        tile = marker.closest(tileSelector);
        if (tile) break;
      }
      const name = nameIn(tile ?? marker);
      if (name) speaking.push(name);
    }
  }

  const participants: string[] = [];
  for (const selector of selectors.participants) {
    for (const element of Array.from(document.querySelectorAll(selector))) participants.push(textOf(element));
  }

  const hints = unique(Array.from(document.querySelectorAll('[data-tid], [data-participant-id]'))
    .slice(0, 400)
    .map((element) => element.getAttribute('data-tid') ?? 'data-participant-id')).slice(0, 60);

  return { speaking: unique(speaking), participants: unique(participants), hints };
}
