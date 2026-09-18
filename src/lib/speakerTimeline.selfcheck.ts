// Fusão de amostras em intervalos da linha do tempo. Rodar com:
//   npx tsx src/lib/speakerTimeline.selfcheck.ts
import assert from 'node:assert/strict';
import { SpeakerTimelineRecorder } from './speakerTimeline';

const start = 1_000_000;
const recorder = new SpeakerTimelineRecorder(start);
const at = (ms: number) => start + ms;

recorder.sample(['Ana'], ['Ana', 'Bruno', ' Callfred  Bot '], at(0));
recorder.sample(['Ana'], [], at(750));
recorder.sample([], [], at(1_500)); // pausa curta: continua o mesmo intervalo
recorder.sample(['Ana'], [], at(2_250));
recorder.sample(['Bruno'], [], at(3_000));
recorder.sample(['Bruno'], [], at(6_000));
recorder.sample([], [], at(9_000)); // pausa longa: fecha Bruno
recorder.sample(['Bruno'], [], at(9_750));
recorder.sample(['Bruno'], [], at(10_500));

const data = recorder.finish();
assert.deepEqual(data.speakerTimeline, [
  { name: 'Ana', startMs: 0, endMs: 2_250 },
  { name: 'Bruno', startMs: 3_000, endMs: 6_000 },
  { name: 'Bruno', startMs: 9_750, endMs: 10_500 },
]);
assert.deepEqual(data.participants, ['Ana', 'Bruno', 'Callfred Bot']);

console.log('speakerTimeline selfcheck ok');
