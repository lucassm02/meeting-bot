// Junção das leituras de legenda em blocos com tempo. Rodar com:
//   npx tsx src/lib/captions.selfcheck.ts
import assert from 'node:assert/strict';
import { CaptionRecorder, mergeCaptionText } from './captions';

// Crescimento, leitura menor do mesmo bloco, emenda pelo fim e correção.
assert.equal(mergeCaptionText('vamos aprovar', 'vamos aprovar o orçamento'), 'vamos aprovar o orçamento');
assert.equal(mergeCaptionText('vamos aprovar o orçamento', 'aprovar o orçamento'), 'vamos aprovar o orçamento');
assert.equal(
  mergeCaptionText('primeira parte da fala que já saiu da tela', 'que já saiu da tela e continua aqui'),
  'primeira parte da fala que já saiu da tela e continua aqui',
);
assert.equal(mergeCaptionText('a gente vai fazer isso amanha', 'a gente vai fazer isso amanhã.'), 'a gente vai fazer isso amanhã.');
assert.equal(mergeCaptionText('bloco um', 'outra coisa'), 'bloco um outra coisa');

const recorder = new CaptionRecorder(1_000);
recorder.sample([{ id: '1', name: 'Ana', text: 'Bom' }], 2_000);
recorder.sample([{ id: '1', name: 'Ana', text: 'Bom dia a todos' }, { id: '2', name: '', text: ' ' }], 3_000);
// Bloco 1 some da tela (fecha); bloco 2 começa com nome.
recorder.sample([{ id: '2', name: 'Bruno', text: 'Obrigado,  Ana' }], 5_000);
recorder.sample([{ id: '2', name: 'Bruno', text: 'Obrigado, Ana' }], 6_000);
assert.deepEqual(recorder.finish(), [
  { name: 'Ana', text: 'Bom dia a todos', startMs: 1_000, endMs: 2_000 },
  { name: 'Bruno', text: 'Obrigado, Ana', startMs: 4_000, endMs: 4_000 },
]);

console.log('captions selfcheck ok');
