// Registro em memória de, no máximo, uma gravação ativa por vez neste processo —
// reaproveita a mesma garantia de concorrência única do globalJobStore/isBusy.

interface ActiveRecording {
  botId: string;
  requestStop: () => Promise<void>;
}

let activeRecording: ActiveRecording | null = null;

export const registerActiveRecording = (botId: string, requestStop: () => Promise<void>): void => {
  activeRecording = { botId, requestStop };
};

export const clearActiveRecording = (botId: string): void => {
  if (activeRecording?.botId === botId) {
    activeRecording = null;
  }
};

export const requestLeaveNow = async (
  botId: string
): Promise<{ ok: boolean; detail: string; status: 200 | 404 | 409 }> => {
  if (!activeRecording) {
    return { ok: false, detail: 'no active recording', status: 404 };
  }

  if (activeRecording.botId !== botId) {
    return { ok: false, detail: 'botId does not match the active recording', status: 409 };
  }

  try {
    await activeRecording.requestStop();
    return { ok: true, detail: 'stop requested', status: 200 };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return { ok: false, detail: `failed to request stop: ${message}`, status: 409 };
  }
};
