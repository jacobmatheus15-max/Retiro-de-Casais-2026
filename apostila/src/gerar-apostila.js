// Gera a apostila do Retiro de Casais 2026 em Word (.docx)
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, BorderStyle, WidthType, ShadingType, PageBreak, Header, Footer,
  PageNumber, VerticalAlign, HeightRule, TableLayoutType,
} = require('docx');

// ---------- Paleta e medidas ----------
const GREEN = '2F4A3A', GREEN2 = '4E6B57', GOLD = 'B08D57', CREAM = 'F7F2E8', SAND = 'EFE6D4', INK = '33403A', GRAY = 'B9B2A3';
const HEAD = 'Georgia', BODY = 'Calibri';
const W = 9638; // largura útil (A4 com margens de 2 cm)

const NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE };
const thin = (c = GRAY, s = 4) => ({ style: BorderStyle.SINGLE, size: s, color: c });

// ---------- Helpers de texto ----------
const run = (text, o = {}) => new TextRun({ text, font: o.font || BODY, size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || INK, allCaps: o.caps, characterSpacing: o.spacing });

function P(children, o = {}) {
  if (typeof children === 'string') children = [run(children, o)];
  return new Paragraph({
    children, alignment: o.align, keepNext: o.keepNext, keepLines: o.keepLines,
    spacing: { before: o.before ?? 0, after: o.after ?? 120, line: o.line ?? 276 },
    indent: o.indent, border: o.border,
  });
}
const spacer = (after = 120) => P([run('')], { after, before: 0 });

// Rótulo em negrito + texto
const LP = (label, text, o = {}) => P([run(label + ' ', { bold: true, color: GREEN }), run(text)], { after: 140, ...o });

const kicker = (text) => P([run(text, { size: 17, bold: true, color: GOLD, caps: true, spacing: 40 })], { after: 120, keepNext: true });

function H1(text, sub) {
  const out = [P([run(text, { font: HEAD, size: 38, color: GREEN })], { after: sub ? 60 : 200, keepNext: true, line: 300 })];
  if (sub) out.push(P([run(sub, { font: HEAD, size: 22, italics: true, color: GREEN2 })], { after: 220, keepNext: true }));
  out.push(new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: GOLD, space: 1 } }, spacing: { after: 240 }, children: [] }));
  return out;
}
const H2 = (text) => P([run(text, { font: HEAD, size: 28, color: GREEN })], { before: 280, after: 140, keepNext: true });
const H3 = (text) => P([run(text, { size: 20, bold: true, color: GOLD, caps: true, spacing: 20 })], { before: 220, after: 100, keepNext: true });

const cellM = { top: 120, bottom: 120, left: 180, right: 180 };

// Caixa de destaque (uma célula) com borda esquerda dourada
function box(paragraphs, o = {}) {
  const color = o.color || GOLD;
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    borders: noBorders,
    rows: [new TableRow({ cantSplit: o.cantSplit ?? true, children: [new TableCell({
      width: { size: W, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, color: 'auto', fill: o.fill || CREAM },
      margins: { top: 200, bottom: 160, left: 300, right: 300 },
      borders: { top: o.full ? thin(color, 6) : NONE, bottom: o.full ? thin(color, 6) : NONE, right: o.full ? thin(color, 6) : NONE, left: { style: BorderStyle.SINGLE, size: 24, color } },
      children: paragraphs,
    })] })],
  });
}
const boxTitle = (t) => P([run(t, { bold: true, color: GREEN, caps: true, size: 19, spacing: 20 })], { after: 100 });
const quote = (t) => P([run('“' + t + '”', { font: HEAD, italics: true, size: 22, color: GREEN2 })], { after: 140 });

// Linhas para escrever (tabela só com bordas inferiores)
function lines(n, width = W, h = 440) {
  return new Table({
    width: { size: width, type: WidthType.DXA }, columnWidths: [width], borders: noBorders,
    rows: Array.from({ length: n }, () => new TableRow({ height: { value: h, rule: HeightRule.EXACT }, cantSplit: true, children: [new TableCell({
      width: { size: width, type: WidthType.DXA },
      borders: { top: NONE, left: NONE, right: NONE, bottom: thin(GRAY, 4) },
      children: [P('', { after: 0 })],
    })] })),
  });
}
// Pergunta + linhas
const ask = (q, n = 3, o = {}) => [P([run(q, { bold: o.bold ?? false })], { after: 40, keepNext: true }), lines(n), spacer(160)];

// Quadro de anotações
function notesBox(title = 'ANOTAÇÕES', n = 6) {
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    borders: noBorders,
    rows: [new TableRow({ cantSplit: true, children: [new TableCell({
      width: { size: W, type: WidthType.DXA },
      borders: { top: thin(GOLD, 6), bottom: thin(GOLD, 6), left: thin(GOLD, 6), right: thin(GOLD, 6) },
      margins: { top: 160, bottom: 200, left: 260, right: 260 },
      children: [boxTitle('✎  ' + title), lines(n, W - 520)],
    })] })],
  });
}

// Tabela de dados genérica
function table(widths, header, rows, o = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const mk = (content, i, isHead, rowIdx) => new TableCell({
    width: { size: widths[i], type: WidthType.DXA },
    shading: isHead ? { type: ShadingType.CLEAR, color: 'auto', fill: GREEN } : (rowIdx % 2 ? { type: ShadingType.CLEAR, color: 'auto', fill: 'FBF8F2' } : undefined),
    margins: { top: 90, bottom: 90, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    borders: { top: thin('D9D1C1'), bottom: thin('D9D1C1'), left: thin('D9D1C1'), right: thin('D9D1C1') },
    children: Array.isArray(content) ? content : [P([run(String(content), isHead ? { bold: true, color: 'FFFFFF', size: 19 } : { size: 20 })], { after: 0, align: (o.center || []).includes(i) || isHead ? AlignmentType.CENTER : AlignmentType.LEFT })],
  });
  const tr = [];
  if (header) tr.push(new TableRow({ tableHeader: true, cantSplit: true, children: header.map((h, i) => mk(h, i, true)) }));
  rows.forEach((r, ri) => tr.push(new TableRow({ cantSplit: true, children: r.map((c, i) => mk(c, i, false, ri)) })));
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: tr, layout: TableLayoutType.FIXED });
}
// Célula com título em negrito + descrição
const tdTitle = (t, d) => [P([run(t, { bold: true, size: 20, color: GREEN })], { after: 20 }), P([run(d, { size: 18 })], { after: 0 })];
const CHECK = '☐';

const pageBreak = () => new Paragraph({ children: [new PageBreak()] });

// Página divisória de parte
function divider(num, title, sub) {
  return [
    P('', { before: 3200, after: 0 }),
    P([run(num, { size: 20, bold: true, color: GOLD, caps: true, spacing: 80 })], { align: AlignmentType.CENTER, after: 200 }),
    P([run(title, { font: HEAD, size: 56, color: GREEN })], { align: AlignmentType.CENTER, after: 200, line: 240 }),
    P([run('— ◆ —', { color: GOLD, size: 22 })], { align: AlignmentType.CENTER, after: 200 }),
    P([run(sub, { font: HEAD, italics: true, size: 26, color: GREEN2 })], { align: AlignmentType.CENTER, after: 0 }),
  ];
}

// Cabeçalho de atividade
function activityHead(tag, title, sub) {
  return [kicker(tag), ...H1(title, sub)];
}

// ======================================================================
// CONTEÚDO
// ======================================================================
const C = [];

// ---------- Apresentação / Sumário ----------
C.push(kicker('Retiro de Casais 2026'));
C.push(...H1('Boas-vindas', 'Um caminho de maturidade a dois'));
C.push(P('Esta apostila foi elaborada para servir como manual prático e pedagógico para os casais participantes. O retiro é composto por 3 palestras, divididas intencionalmente entre o sábado (diagnóstico e ferramentas de comunicação) e o domingo (sustentabilidade no cotidiano).', { after: 200 }));
C.push(P('Guarde este material com carinho: ele foi pensado para ser escrito, rabiscado, conversado e revisitado depois do retiro.', { after: 300 }));

C.push(H3('Nesta apostila'));
const tocRows = [
  ['Parte I', 'Guia e Roteiro do Retiro'],
  ['', 'Visão Geral do Retiro'],
  ['', 'Palestra 1 — Do Analfabetismo Relacional ao Amor Verdadeiro'],
  ['', 'Palestra 2 — Pacificação Real e Ferramentas Práticas de Comunicação'],
  ['', 'Palestra 3 — Sustentando o Amor no Cotidiano: A Treliça do Lar'],
  ['Parte II', 'Atividades de Apoio'],
  ['', 'Estudo de Caso: A História de Jéssica  ·  Palestra 1'],
  ['', 'Avaliação de Maturidade Emocional Individual  ·  Palestra 1'],
  ['', 'Declaração de Direitos Relacionais no Casamento  ·  Palestra 1'],
  ['', 'Mapeando Suposições e Alergias Emocionais  ·  Palestra 2'],
  ['', 'Avaliação das 4 Áreas da Vida da Família  ·  Palestra 3'],
];
C.push(new Table({
  width: { size: W, type: WidthType.DXA }, columnWidths: [1500, W - 1500], borders: noBorders,
  rows: tocRows.map(([a, b]) => new TableRow({ children: [
    new TableCell({ width: { size: 1500, type: WidthType.DXA }, borders: noBorders, children: [P([run(a, { bold: true, color: GOLD, caps: true, size: 19, spacing: 20 })], { after: 60, before: a ? 160 : 0 })] }),
    new TableCell({ width: { size: W - 1500, type: WidthType.DXA }, borders: { top: NONE, left: NONE, right: NONE, bottom: a ? NONE : thin('E4DCCB') }, children: [P([run(b, a ? { font: HEAD, size: 26, color: GREEN } : { size: 21 })], { after: 60, before: a ? 160 : 60 })] }),
  ] })),
}));

// ---------- PARTE I ----------
C.push(pageBreak(), ...divider('Parte I', 'Guia e Roteiro', 'Roteiro e guia de exercícios para as mesas'));
C.push(pageBreak());

C.push(kicker('Parte I · Guia e Roteiro'));
C.push(...H1('Visão Geral do Retiro', 'Espiritualidade Profunda e Maturidade Relacional — baseado na obra de Peter Scazzero'));
C.push(P('Este documento foi elaborado para servir como manual prático e pedagógico para os casais participantes. O retiro é composto por 3 palestras, divididas intencionalmente entre o sábado (diagnóstico e ferramentas de comunicação) e o domingo (sustentabilidade no cotidiano).', { after: 240 }));
C.push(table([2300, 3300, 4038], ['Sessão', 'Tema Central', 'Foco do Aprendizado & Capítulo'], [
  [tdTitle('Palestra 1', 'Sábado · Manhã'), 'Do Analfabetismo Relacional ao Amor Verdadeiro', 'Maturidade emocional, relação Eu-Tu, Declaração de Direitos (Cap. 9)'],
  [tdTitle('Palestra 2', 'Sábado · Noite'), 'Pacificação Real e Comunicação', 'Falsa paz vs. verdadeira paz, escuta ativa, suposições, alergias (Cap. 9)'],
  [tdTitle('Palestra 3', 'Domingo · Manhã'), 'Sustentando o Amor no Cotidiano', 'A Regra de Vida (Treliça), descanso bíblico, ritmos intencionais (Cap. 10)'],
]));
C.push(spacer(300));
C.push(box([
  boxTitle('Como usar esta apostila'),
  P('Cada palestra traz o roteiro do que será exposto e um guia de exercícios para as mesas, sempre em três passos: (A) reflexão individual, (B) conversa do casal e (C) partilha entre os dois casais da mesa. As atividades de apoio da Parte II aprofundam cada tema.', { after: 0 }),
]));

// ---- Palestra 1
function palestraHead(n, title, when) {
  return [pageBreak(), kicker(`Palestra ${n} · ${when}`), ...H1(title)];
}
C.push(...palestraHead(1, 'Do Analfabetismo Relacional ao Amor Verdadeiro', 'Sábado Manhã · Capítulo 9'));
C.push(H2('1. Roteiro de Exposição Verbal'));
C.push(LP('Abertura — Introdução e Estudo de Caso (Jéssica):', 'Conheça Jéssica. Dedicada às disciplinas espirituais há 15 anos, mas incapaz de lidar com conflitos e sentimentos no trabalho. Conclusão: espiritualidade sem maturidade relacional produz relacionamentos disfuncionais.'));
C.push(LP('Fundamento Bíblico:', 'João 4.20 e Mateus 22.37-40. Amar a Deus e amar ao próximo são dimensões inseparáveis. Não podemos amar a Deus, a quem não vemos, se coisificamos o cônjuge, a quem vemos.'));
C.push(LP('Os 4 Estágios da Maturidade:', 'Bebê (só pensa em si), Criança (exige gratificação instantânea), Adolescente (defensivo, guarda mágoas, cobra retribuição) e Adulto Emocional (assume responsabilidade, respeita o espaço alheio).'));
C.push(LP('Filosofia Eu-Tu vs. Eu-Coisa:', 'Conceito de Martin Buber. A relação Eu-Coisa trata o outro como extensão de mim ou utilitário. A relação Eu-Tu reconhece o cônjuge como ser humano singular criado à imagem de Deus.'));
C.push(LP('A Declaração de Direitos Relacionais:', 'Tradução prática da relação Eu-Tu. O cônjuge tem direito de ser diferente, de discordar, de cometer erros, de ser ouvido e de ter privacidade.'));
C.push(spacer(80));
C.push(box([
  boxTitle('Notas e citações do livro'),
  quote('Não é possível ser emocionalmente maduro enquanto se permanece espiritualmente imaturo. Mas a tragédia é que muitos são espiritualmente atuantes enquanto permanecem bebês emocionais.'),
  quote('Quando tratamos nosso cônjuge como um ‘Tu’, o espaço entre nós torna-se um Espaço Sagrado onde a presença de Deus se manifesta com poder.'),
]));
C.push(H2('2. Guia de Exercícios para as Mesas'));
C.push(LP('Passo A — Autorreflexão Individual:', 'Em qual área do seu casamento você tende a agir como criança ou adolescente emocional? Onde/quando você tem praticado a falsa paz para evitar atrito (paz ou razão)?', { after: 40, keepNext: true }));
C.push(lines(5), spacer(200));
C.push(LP('Passo B — Compartilhamento do Casal:', 'Compartilhem focando exclusivamente nas suas vulnerabilidades, não nas falhas do cônjuge.'));
C.push(LP('Passo C — Diálogo na Mesa:', 'Diálogo aberto entre os dois casais da mesa sobre o desafio de sair da imaturidade.'));
C.push(spacer(120), notesBox('Anotações', 14));

// ---- Palestra 2
C.push(...palestraHead(2, 'Pacificação Real e Ferramentas Práticas de Comunicação', 'Sábado à Noite · Capítulo 9'));
C.push(H2('1. Roteiro de Exposição Verbal'));
C.push(LP('Fundamento Bíblico:', 'Mt 5.9 e Êx 20.16. Pacificar não é apaziguar (varrer problemas para debaixo do tapete). A verdadeira paz exige falar a verdade EM AMOR.'));
C.push(LP('Exemplos do Livro:', 'Karl (que tolerava atrasos), Bob (que rachava contas caras contrariado) e Sharon (que tolerava irresponsabilidade) praticavam a falsa paz por medo do confronto.'));
C.push(LP('Habilidade 1 — Falar e Ouvir Ativamente:', 'Falar em 1ª pessoa (“Eu sinto” em vez de “Você é”); quem ouve deve parafrasear antes de responder, para garantir compreensão sem atitude defensiva.'));
C.push(LP('Habilidade 2 — Verificar Suposições:', 'Não inventar histórias na mente. Usar o modelo: “Tenho sua permissão para verificar uma suposição?”'));
C.push(LP('Habilidade 3 — Alergias Emocionais:', 'Reconhecer que feridas ou padrões da família de origem geram reações desproporcionais no presente.'));
C.push(spacer(80));
C.push(box([
  boxTitle('Notas e citações do livro'),
  quote('A falsa paz finge que está tudo bem quando não está. A verdadeira paz de Cristo passa pelo conflito para alcançar uma reconciliação fundamentada na verdade.'),
  quote('Uma expectativa só é saudável quando é consciente, realista, falada e mutuamente acordada. Todas as outras geram ressentimento.'),
]));
C.push(H2('2. Guia de Exercícios para as Mesas'));
C.push(LP('Passo A — Autorreflexão Individual:', 'Complete para si mesmo:', { keepNext: true }));
C.push(...ask('A alergia emocional que uma atitude sua desperta em mim é:', 2));
C.push(...ask('Quando isso acontece, me lembro de (uma situação do passado):', 2));
C.push(...ask('As palavras de aceitação/cura que eu gostaria de ter ouvido no passado são:', 2));
C.push(LP('Passo B — Treinamento a Dois:', 'Usem a regra “Falar e Ouvir”. O cônjuge não resolve nem se defende, apenas acolhe e valida.'));
C.push(LP('Passo C — Oração na Mesa:', 'Orem em grupo (mesa) pela cura do passado uns dos outros. Ninguém precisa expor a alergia para a mesa, apenas orar.'));
C.push(spacer(120), notesBox('Anotações', 14));

// ---- Palestra 3
C.push(...palestraHead(3, 'Sustentando o Amor no Cotidiano — A Treliça do Lar', 'Domingo Manhã · Capítulo 10'));
C.push(H2('1. Roteiro de Exposição Verbal'));
C.push(LP('Fundamento Bíblico:', 'Jo 15.1-5 (A Videira) e Daniel 6 (o ritmo de Daniel na Babilônia). Sem uma estrutura (treliça), a videira cai na terra e seus frutos apodrecem.'));
C.push(LP('O Conceito da Treliça:', 'A Regra de Vida (Trellis) é um padrão intencional de disciplinas espirituais e hábitos relacionais para manter Cristo no centro da vida diária.'));
C.push(LP('Os 4 Pilares da Treliça Conjugal:', 'Oração (silêncio, Lectio Divina, Ofício Divino); Descanso (Sábado de 24h: parar, descansar, divertir-se, contemplar); Trabalho (limites saudáveis); Relacionamentos (saúde emocional e comunidade).'));
C.push(LP('O Princípio da Gradualidade:', 'Começar devagar. Escolher apenas 1 ou 2 novos hábitos nos primeiros 3 meses para consolidar a prática sem esgotamento.'));
C.push(spacer(80));
C.push(box([
  boxTitle('Notas e citações do livro'),
  quote('Uma Regra de Vida não é um chicote para nos punir, mas uma treliça de amor que sustenta nossa união com Cristo diante das tempestades da rotina.'),
  quote('O Sábado bíblico não é apenas um dia sem trabalho; é um ato de confiança de 24 horas onde declaramos que o mundo continua girando sob a providência de Deus.'),
]));
C.push(H2('2. Guia de Exercícios para as Mesas'));
C.push(LP('Passo A — Mapeamento Individual:', 'Avaliar as 4 áreas (Oração, Descanso, Trabalho, Relacionamentos) e identificar onde a rotina mais tem caído no piloto automático.', { after: 40, keepNext: true }));
C.push(lines(4), spacer(200));
C.push(LP('Passo B — Construindo a Treliça:', 'O casal escolhe juntamente 1 hábito de espiritualidade e 1 hábito de descanso para iniciarem na semana seguinte.', { after: 40, keepNext: true }));
C.push(lines(4), spacer(200));
C.push(LP('Passo C — Aliança na Mesa:', 'Aliança de prestação de contas entre os dois casais da mesa para acompanhamento após o retiro.'));
C.push(spacer(120), notesBox('Anotações', 14));

// ---------- PARTE II ----------
C.push(pageBreak(), ...divider('Parte II', 'Atividades de Apoio', 'Exercícios, avaliações e dinâmicas para o casal'));

// ---- Atividade 1: Jéssica
C.push(pageBreak(), ...activityHead('Atividade de Apoio · Palestra 1', 'Estudo de Caso: A História de Jéssica', 'O analfabetismo relacional espiritualizado e a maturidade no casamento'));
C.push(H2('1. A História de Jéssica (Peter Scazzero)'));
C.push(box([
  P('Jéssica é uma competente gerente de sua empresa. Ela é cristã há quinze anos e ama passar tempo com Deus em suas disciplinas espirituais. Certo dia, o vice-presidente de sua empresa estava programando as viagens dos clientes fora da cidade e pediu que Jéssica escolhesse as semanas de sua preferência para viajar nos três meses seguintes.'),
  P('Ela enviou por e-mail suas datas preferidas e esperou ansiosamente pela confirmação. Não houve retorno. Na semana seguinte, Jéssica telefonou para o escritório do vice-presidente. A assistente dele respondeu: “De acordo com a programação que tenho aqui, as próximas três semanas estão completas. Suponho que isso significa que ele não precisa de você agora. Mas obrigada por telefonar.” Jéssica ficou atônita e respondeu apenas um “obrigada” mecânico antes de desligar.'),
  P('Nas duas semanas seguintes, Jéssica lutou com Deus e consigo mesma. Pediu perdão a Deus pela raiva que sentia, tentou imaginar por que o vice-presidente havia mudado de ideia, humilhou-se em oração pedindo amor pelos colaboradores e perdeu o sono. Finalmente, concluiu de forma resignada que “Deus estava tratando sua persistente teimosia”.'),
  P('Ao longo do tempo, Jéssica se distanciou do vice-presidente e dos outros gerentes, evitando-os sempre que possível. Trabalhou duro por dois anos, mas sentiu-se estagnada. Por fim, pediu demissão e assumiu um cargo em outra empresa. Jéssica ama a Deus e pratica disciplinas espirituais, mas seu compromisso com Cristo não incluía relacionar-se com as pessoas de forma emocionalmente adulta.', { after: 0 }),
], { cantSplit: false }));
C.push(H2('2. Diagnóstico: O que Aconteceu com Jéssica?'));
C.push(P('O exemplo de Jéssica ilumina um sintoma muito comum na vida cristã: o analfabetismo emocional espiritualizado. Jéssica usou a linguagem da piedade (orar, pedir perdão pela raiva, resignar-se à “vontade de Deus”) para evitar o desconforto de uma conversa madura e direta.'));
C.push(table([2900, W - 2900], null, [
  [[P([run('Suposições Não Checadas', { bold: true, color: GREEN })], { after: 0 })], 'Ela deduziu o motivo do silêncio do vice-presidente sem jamais perguntar diretamente a ele o que havia ocorrido.'],
  [[P([run('Comunicação Passivo-Agressiva', { bold: true, color: GREEN })], { after: 0 })], 'Aceitou a resposta da assistente de forma mecânica e depois se isolou, guardando ressentimento.'],
  [[P([run('Espiritualização do Problema', { bold: true, color: GREEN })], { after: 0 })], 'Em vez de resolver o conflito interpessoal, “espiritualizou” a situação, concluindo que Deus estava punindo sua teimosia.'],
  [[P([run('Falsa Pacificação', { bold: true, color: GREEN })], { after: 0 })], 'Manteve uma aparência de paz externa enquanto o relacionamento profissional morria por dentro.'],
]));
C.push(H2('3. Dinâmica Prática do Casal na Mesa'));
C.push(box([boxTitle('Momento a dois na mesa'), P('O casal conversa a dois na mesa. Releem a história de Jéssica e identificam momentos de infantilidade emocional no casamento, pedindo perdão mútuo.', { after: 0 })], { fill: SAND }));
C.push(H2('4. Espaço para Anotações do Casal'));
C.push(P('Utilizem o espaço abaixo para registrar as percepções, os momentos de infantilidade relacional identificados e os compromissos de perdão e restauração assumidos hoje.'));
const half = (W - 240) / 2;
const sideCol = (title) => new TableCell({
  width: { size: half, type: WidthType.DXA }, borders: noBorders,
  children: [P([run(title, { bold: true, color: GOLD, caps: true, size: 19, spacing: 20 })], { after: 40 }), P([run('Áreas onde percebo comportamentos de infantilidade emocional em mim:', { size: 20 })], { after: 40 }), lines(6, half)],
});
C.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [half, 240, half], borders: noBorders, rows: [new TableRow({ cantSplit: true, children: [
  sideCol('Reflexões do Esposo'), new TableCell({ width: { size: 240, type: WidthType.DXA }, borders: noBorders, children: [P('')] }), sideCol('Reflexões da Esposa'),
] })] }));
C.push(spacer(200));
C.push(H3('O nosso pedido de perdão e compromisso mútuo como casal'));
C.push(P('Orando juntos e perdoando-nos mutuamente por momentos de infantilidade e falta de comunicação sincera:', { after: 40, keepNext: true }), lines(5));

// ---- Atividade 2: Maturidade
C.push(pageBreak(), ...activityHead('Atividade de Apoio · Palestra 1', 'Avaliação de Maturidade Emocional Individual', 'Diagnóstico dos 4 Estágios do Desenvolvimento Relacional (Capítulo 9)'));
C.push(box([
  boxTitle('Instruções de preenchimento'),
  P('Responda com sinceridade diante de Deus, avaliando como você reage sob estresse, cansaço, contrariedade ou conflito no casamento. Não responda com base no que você gostaria de ser, mas na sua prática real do dia a dia.'),
  P('Atribua uma nota de 1 a 4 para cada uma das afirmações abaixo, conforme a escala:'),
  P([run('1 = Quase Nunca (Raro)   ·   2 = Às Vezes (Ocasional)   ·   3 = Frequentemente (Comum)   ·   4 = Quase Sempre (Habitual)', { bold: true, color: GREEN, size: 20 })], { after: 0 }),
]));
const blocos = [
  ['Bloco A: O Bebê Emocional', 'Busca gratificação instantânea, opera em autopreservação e trata o outro como objeto/coisa.', [
    'Quando me sinto contrariado ou estressado, espero que minhas necessidades sejam atendidas imediatamente pelo meu cônjuge.',
    'Tenho pouca percepção sobre como minhas palavras ríspidas ou silêncios afetam a dor e o estado emocional do outro.',
    'Entro em pânico, crise ou desespero com facilidade quando as coisas não acontecem exatamente do meu jeito no lar.',
    'Sem perceber, uso meu cônjuge como um utensílio ou meio para alcançar meus próprios planos e ter conforto.',
    'Diante de dificuldades no casamento, busco formas imediatas de alívio rápido (comida, telas, compras, isolamento) em vez de encarar a realidade.',
  ]],
  ['Bloco B: A Criança Emocional', 'Só fica feliz quando recebe o que quer; interpreta discordâncias como rejeição e magoa-se facilmente.', [
    'Só consigo me sentir feliz, satisfeito e em paz quando meu cônjuge concorda comigo e faz o que desejo.',
    'Interpreto opiniões ou preferências diferentes do meu cônjuge como afronta pessoal, falta de amor ou rejeição ao meu valor.',
    'Quando me sinto magoado, costumo me amuar, emburrar ou fechar a cara para punir o outro ou demonstrar meu descontentamento.',
    'Tenho grande dificuldade para receber críticas ou correções sem me sentir profundamente vitimado ou injustiçado.',
    'Culpo frequentemente meu cônjuge pelas minhas reações de raiva ou tristeza (ex.: “eu só gritei porque você me provocou”).',
  ]],
  ['Bloco C: O Adolescente Emocional', 'Atitude defensiva diante de falhas, guarda contas mentais e evita conflitos através da falsa paz.', [
    'Reajo quase imediatamente com justificativas e atitude defensiva sempre que meu cônjuge aponta um erro meu.',
    'Guardo uma “conta mental” secreta de tudo o que fiz pelo cônjuge para poder cobrar retribuição ou usar como barganha depois.',
    'Fujo de conversas difíceis ou finjo que está tudo bem para evitar climas chatos, varrendo conflitos para debaixo do tapete (falsa paz).',
    'Guardo mágoas passadas por muito tempo e as trago de volta em discussões presentes como munição.',
    'Comparo meu cônjuge ou nosso casamento com o de outros casais, sentindo inveja, ressentimento ou amargura.',
  ]],
  ['Bloco D: O Adulto Emocional em Cristo', 'Assume responsabilidade, comunica-se com clareza e honra a dignidade do outro como “Tu” sagrado.', [
    'Assumo a responsabilidade pelas minhas próprias emoções, escolhas e erros sem culpar meu cônjuge ou as circunstâncias.',
    'Consigo expressar minhas necessidades, limites e expectativas em 1ª pessoa, com clareza, serenidade e respeito.',
    'Reconheço e respeito meu cônjuge como um “Tu” sagrado (imagem de Deus), concedendo-lhe o direito de ser diferente e discordar de mim.',
    'Escuto críticas do meu cônjuge com postura calma, buscando extrair o aprendizado em vez de me defender ou contra-atacar.',
    'Enfrento conflitos necessários com coragem e amor (verdadeira paz), buscando falar a verdade bíblica sem fingimentos.',
  ]],
];
for (const [t, marca, itens] of blocos) {
  C.push(H2(t));
  C.push(P([run('Marca principal: ', { bold: true, color: GOLD }), run(marca, { italics: true })], { keepNext: true }));
  C.push(table([600, W - 600 - 1300, 1300], ['#', 'Comportamento / Atitude Relacional', 'Nota (1-4)'], itens.map((x, i) => [String(i + 1), x, '']), { center: [0] }));
}
C.push(pageBreak());
C.push(H2('Apuração e Interpretação dos Resultados'));
C.push(table([3000, W - 3000 - 1900, 1900], ['Estágio de Maturidade', 'Descrição da Tendência', 'Pontuação Total'], [
  ['Bloco A: Bebê Emocional', 'Soma das notas de 1 a 5 (faixa: 5 a 20 pontos)', '_____ pontos'],
  ['Bloco B: Criança Emocional', 'Soma das notas de 1 a 5 (faixa: 5 a 20 pontos)', '_____ pontos'],
  ['Bloco C: Adolescente Emocional', 'Soma das notas de 1 a 5 (faixa: 5 a 20 pontos)', '_____ pontos'],
  ['Bloco D: Adulto Emocional', 'Soma das notas de 1 a 5 (faixa: 5 a 20 pontos)', '_____ pontos'],
], { center: [2] }));
C.push(spacer(240));
C.push(box([
  boxTitle('Como interpretar seu diagnóstico'),
  LP('Maior pontuação nos Blocos A, B ou C:', 'indica a sua área de “regressão emocional sob estresse”. É onde a sua natureza humana e velhos hábitos da infância/família de origem tendem a assumir o controle quando você se sente ameaçado ou estressado.'),
  LP('Alta pontuação no Bloco D (Adulto Emocional):', 'demonstra a presença do fruto do Espírito e o aprendizado ativo de novas habilidades relacionais. O objetivo bíblico é que o Bloco D seja a sua resposta predominante e consistente no lar.'),
  LP('Aviso pastoral:', 'este teste não serve para condenar você nem para apontar os erros do seu cônjuge. Ele serve como um espelho espiritual para que você identifique onde precisa da graça de Cristo e do auxílio do Espírito Santo para amadurecer.', { after: 0 }),
]));
C.push(H2('Plano de Ação e Oração de Entrega'));
C.push(...ask('1. Qual o principal sinal de infantilidade emocional (Bebê, Criança ou Adolescente) que me comprometo a confessar a Deus e pedir perdão ao meu cônjuge hoje?', 3));
C.push(...ask('2. Qual hábito do Adulto Emocional (ex.: escutar sem defender-se, checar suposições, honrar o “Tu”) quero praticar a partir desta semana?', 3));
C.push(box([boxTitle('Oração de humildade e crescimento'), quote('Senhor Deus, revela-me onde tenho agido como criança ou adolescente no meu casamento. Concede-me a graça da humildade para abandonar o egoísmo e aprender a amar o meu cônjuge como Tu nos amaste em Cristo. Amém.')], { fill: SAND }));

// ---- Atividade 3: Declaração de Direitos
C.push(pageBreak(), ...activityHead('Atividade de Apoio · Palestra 1', 'Declaração de Direitos Relacionais no Casamento', 'Guia prático de consulta e reflexão para o casal — Cap. 9'));
C.push(box([
  boxTitle('O que significa respeitar o cônjuge no cotidiano?'),
  P('Respeito não é um sentimento vago — é a forma concreta como tratamos o outro no dia a dia. Tratar o cônjuge como um “Tu” (alguém feito à imagem de Deus com dignidade única) e não como uma “Coisa” (um utensílio para suprir nossas necessidades ou planos) exige conceder-lhe direitos relacionais fundamentais. Este guia apresenta os 11 itens da Declaração de Direitos com perguntas para reflexão conjunta.', { after: 0 }),
]));
C.push(H2('Os 11 Direitos Relacionais no Lar'));
const direitos = [
  ['Direito a Espaço e Privacidade', 'Respeitar a necessidade mútua por momentos de silêncio, ter seu próprio tempo e espaço pessoal. Inclui atitudes simples como bater na porta antes de entrar, não abrir a correspondência ou mensagens privadas do outro sem autorização e honrar os momentos em que o outro precisa se recolher para orar ou descansar.', 'Em que momentos da rotina semanal eu posso demonstrar maior respeito pelo seu espaço e tempo a sós?'],
  ['Direito de Ser Diferente', 'Permitir e celebrar as preferências individuais do cônjuge — gostos culinários, ritmo de vida, escolhas de lazer, estilo de vestuário e forma de se expressar. O casamento saudável conecta duas pessoas distintas, sem exigir que um se torne cópia do outro.', 'Existe alguma diferença individual sua que eu tenho tentado mudar em vez de aceitar e celebrar?'],
  ['Direito de Discordar', 'Dar espaço para que cada pessoa pense e veja a vida de forma diferente. Discordância de opinião não é ataque pessoal nem falta de amor. Adultos emocionais conseguem debater ideias sem se tornarem antagônicos ou defensivos.', 'Como podemos reagir com mais maturidade e serenidade quando divergirmos em opiniões?'],
  ['Direito de Ser Ouvido', 'Escutar com atenção genuína os desejos, opiniões, sentimentos e preocupações do cônjuge. Significa desligar distrações (telas, TV), adiar a própria resposta e colocar-se inteiramente presente para entrar no mundo do outro.', 'Você sente que eu realmente paro para ouvir seus sentimentos sem te interromper ou dar conselhos apressados?'],
  ['Direito de Ser Levado a Sério', 'Validar a dor, o cansaço ou as preocupações expressas pelo outro, mesmo quando não compreendemos totalmente a razão. Nunca desacreditar ou ridicularizar os sentimentos do cônjuge com frases como “você está exagerando” ou “isso é bobagem”.', 'Já houve momentos em que minimizei suas preocupações? Como posso te validar melhor?'],
  ['Direito ao Benefício da Dúvida (Verificar Suposições)', 'Antes de julgar as intenções do cônjuge diante de um atraso, esquecimento ou ruído de comunicação, assumir a melhor intenção e perguntar em vez de presumir. “Não dar falso testemunho” (9º Mandamento) inclui não inventar mentiras sobre o cônjuge na própria mente.', 'Que suposição sobre você eu costumo fazer sem confirmar antes?'],
  ['Direito de Ser Informado da Verdade', 'Confiar que o lar é um ambiente de transparência e verdade falada em amor. Rejeitar a “falsa paz” (varrer problemas para debaixo do tapete) e cultivar a coragem de comunicar fatos e sentimentos com honestidade e gentileza.', 'Existe algum assunto em nosso relacionamento que precisamos trazer à luz com amor e verdade?'],
  ['Direito de Ser Consultado', 'Não tomar decisões unilaterais que afetem a vida, o orçamento, o tempo ou a rotina da família. Toda decisão com impacto mútuo deve ser conversada e alinhada antes de ser concretizada.', 'Em quais áreas da nossa vida a dois precisamos reforçar a prática da consulta mútua?'],
  ['Direito de Ser Imperfeito e Cometer Erros', 'Dar ao cônjuge a graça de falhar, esquecer algo, quebrar um objeto ou cometer equívocos sem ser punido com sarcasmo, cara feia ou rejeição. Lembrar que Deus nos aceita em nossas imperfeições.', 'Eu tenho oferecido um ambiente seguro para você errar e aprender sem medo de julgamento?'],
  ['Direito a Tratamento Cortês e Respeitoso', 'Usar palavras gentis, tom de voz brando e educação básica no dia a dia. Eliminar palavras que magoam, ironia, gritos, ofensas ou apelidos pejorativos. O respeito verbal reflete a dignidade da Imago Dei no cônjuge.', 'Quais palavras ou tons de voz precisamos banir do nosso diálogo cotidiano?'],
  ['Direito de Ser Respeitado nos Sentimentos', 'Levar em conta o impacto emocional que nossas palavras, escolhas e hábitos causam no cônjuge. Amar ativamente é importar-se com a saúde emocional de quem caminha ao nosso lado.', 'Como posso demonstrar de forma prática que a sua paz emocional é uma prioridade para mim?'],
];
direitos.forEach(([t, d, q], i) => {
  C.push(new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [900, W - 900], borders: noBorders,
    rows: [new TableRow({ cantSplit: true, children: [
      new TableCell({ width: { size: 900, type: WidthType.DXA }, borders: noBorders, verticalAlign: VerticalAlign.TOP,
        children: [P([run(String(i + 1).padStart(2, '0'), { font: HEAD, size: 40, color: GOLD })], { after: 0 })] }),
      new TableCell({ width: { size: W - 900, type: WidthType.DXA }, borders: { top: NONE, left: NONE, right: NONE, bottom: thin('E4DCCB') }, margins: { bottom: 160 },
        children: [
          P([run(t, { font: HEAD, size: 26, color: GREEN })], { after: 60 }),
          P([run(d, { size: 21 })], { after: 80 }),
          P([run('Pergunta para o casal: ', { bold: true, color: GOLD, size: 21 }), run(q, { italics: true, size: 21 })], { after: 0 }),
        ] }),
    ] })],
  }));
  C.push(spacer(160));
});
C.push(H2('Compromisso de Respeito e Aliança do Casal'));
C.push(P('Diante de Deus e como fruto da revelação do Seu amor por nós, declaramos nosso compromisso de cultivar um lar onde o respeito, a verdade e a graça prevaleçam. Reconhecemos o nosso cônjuge como uma bênção singular e imagem de Deus, comprometendo-nos a honrar seus direitos relacionais dia a dia.', { after: 600, keepNext: true }));
const sig = (label) => new TableCell({ width: { size: half, type: WidthType.DXA }, borders: { top: thin(GREEN, 6), bottom: NONE, left: NONE, right: NONE },
  children: [P([run(label, { size: 19, color: GREEN2, caps: true, spacing: 20 })], { align: AlignmentType.CENTER, before: 80, after: 0 })] });
C.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [half, 240, half], borders: noBorders, rows: [new TableRow({ cantSplit: true, children: [
  sig('Assinatura do Esposo'), new TableCell({ width: { size: 240, type: WidthType.DXA }, borders: noBorders, children: [P('')] }), sig('Assinatura da Esposa'),
] })] }));

// ---- Atividade 4: Suposições e Alergias
C.push(pageBreak(), ...activityHead('Atividade de Apoio · Palestra 2', 'Mapeando Suposições e Alergias Emocionais', 'Exercício prático para as mesas · Sábado à noite · Atividade individual e a dois'));
C.push(H2('1. Fundamentação e Conceitos-Chave'));
C.push(P('Para construir um casamento maduro e um ambiente de paz verdadeira, Peter Scazzero nos ensina a identificar dois grandes vilões da comunicação conjugal:'));
C.push(table([3000, W - 3000], ['Conceito', 'O que é no casamento?'], [
  [[P([run('Suposição Não Confirmada', { bold: true, color: GREEN })], { after: 0 })], 'Ocorre quando criamos em nossa mente uma narrativa ou julgamento sobre as intenções ou pensamentos do cônjuge sem nunca perguntar a ele(a). É uma forma indireta de dar falso testemunho contra o parceiro na própria mente.'],
  [[P([run('Alergia Emocional (Gatilho do Passado)', { bold: true, color: GREEN })], { after: 0 })], 'Reação emocional intensa, desproporcional e exagerada no presente diante de um comportamento do cônjuge, ativada por feridas, dinâmicas familiares ou memórias dolorosas da infância e juventude.'],
]));
C.push(H2('2. Etapa 1: Autorreflexão Individual (silêncio · 5 min)'));
C.push(P([run('Instrução: ', { bold: true, color: GREEN }), run('em momentos de silêncio e reflexão perante Deus, cada cônjuge responde às perguntas abaixo antes do diálogo a dois.')]));
C.push(...ask('1. Qual suposição recente fiz sobre as atitudes ou intenções do meu cônjuge que eu jamais parei para confirmar com ele(a)?', 4));
C.push(...ask('2. Qual “alergia emocional” da minha infância/família de origem faz com que eu reaja com irritação ou fechamento desproporcional?', 4));
C.push(H2('3. Etapa 2: Prática Guiada a Dois (12 min)'));
C.push(box([
  boxTitle('Roteiro prático: ferramenta “Verificar Suposições”'),
  LP('Passo 1 — Pedir permissão:', '“Tenho sua permissão para verificar uma suposição que estou fazendo?”'),
  LP('Passo 2 — Expressar a suposição:', '“Quando aconteceu [X], eu supus que você pensou/pretendeu [Y]. Isso está correto?”'),
  LP('Passo 3 — Ouvir e acolher:', 'o outro cônjuge responde com verdade e sem atitude defensiva. Quem ouve apenas parafraseia para confirmar o entendimento.'),
  LP('Passo 4 — Revelar a alergia:', '“Percebo que reagi de forma intensa porque isso toca em uma alergia emocional minha do passado, que é…”', { after: 0 }),
], { fill: SAND }));
C.push(H2('4. Registro de Alinhamento e Pedido de Perdão'));
C.push(P('Anotem abaixo os aprendizados e combinados desta conversa, seguidos de um tempo de oração juntos e pedido de perdão pelas suposições falsas alimentadas:', { after: 40, keepNext: true }));
C.push(lines(6), spacer(240));
C.push(P([run('“O amor paciente e maduro aprende a perguntar em vez de julgar, e acolhe a história do outro com compaixão.”', { font: HEAD, italics: true, size: 24, color: GREEN2 })], { align: AlignmentType.CENTER }));

// ---- Atividade 5: Piloto automático
C.push(pageBreak(), ...activityHead('Atividade de Apoio · Palestra 3', 'Avaliação das 4 Áreas da Vida da Família', 'Identificando o piloto automático na rotina do casal'));
C.push(box([
  boxTitle('A treliça e o perigo do piloto automático (Capítulo 10)'),
  P('Como ensina Peter Scazzero, a maioria dos casais não decide conscientemente esfriar na fé ou se distanciar emocionalmente. Nós simplesmente caímos no “piloto automático”, engolidos pelas urgências do cotidiano. A Regra de Vida é a nossa “treliça” (João 15.1-5) — uma estrutura consciente de ritmos de oração, descanso, trabalho e relacionamentos que sustenta a videira do nosso lar para dar frutos que permanecem.', { after: 0 }),
]));
C.push(H2('Escala de Avaliação de 1 a 5'));
const sw = Math.floor(W / 5);
C.push(table([sw, sw, sw, sw, W - 4 * sw], null, [
  [['1', 'Totalmente Automático'], ['2', 'Quase Automático'], ['3', 'Moderado / Oscilante'], ['4', 'Saudável com Ajustes'], ['5', 'Totalmente Intencional']].map(([n, t]) => [
    P([run(n, { font: HEAD, size: 32, color: GOLD })], { align: AlignmentType.CENTER, after: 0 }),
    P([run(t, { bold: true, size: 18, color: GREEN })], { align: AlignmentType.CENTER, after: 0 })]),
  ['Negligenciado; a rotina corre sem direção.', 'Acontece só por reação ou quando surge crise.', 'Funciona às vezes; engolido pela rotina.', 'Prioridade mantida, mas com pequenas falhas.', 'Ritmo claro e consciente praticado com alegria.'].map((t) => [P([run(t, { size: 17 })], { align: AlignmentType.CENTER, after: 0 })]),
]));
C.push(H2('Etapa 1: Autoavaliação Individual (5 min)'));
C.push(P('Em silêncio, dê uma nota de 1 a 5 para cada item abaixo, indicando o quanto essa prática está no “piloto automático” na sua vida e no seu lar.'));
const areas = [
  ['Área 1: Oração & Espiritualidade', [
    ['Silêncio e Quietude Pessoal', 'Prática de momentos a sós com Deus sem distrações tecnológicas.'],
    ['Lectio Divina / Meditação nas Escrituras', 'Ler a Palavra para ser transformado e ouvir a voz de Deus, não apenas cumprir tarefa.'],
    ['Ofício Divino / Pausas Diárias', 'Pausar o trabalho ao longo do dia para reorientar o coração em Deus.'],
    ['Oração Mútua do Casal', 'Orar juntos com frequência, apresentando a família e o casamento a Deus.'],
  ]],
  ['Área 2: Descanso & Lazer (Sábado Bíblico)', [
    ['Descanso', 'Praticar as 4 marcas: parar o trabalho, descansar o corpo/mente, divertir-se e contemplar a Deus.'],
    ['Simplicidade e Desaceleração', 'Dizer “não” ao excesso de compromissos e ao consumismo que esgotam a família.'],
    ['Brincar e Recreação no Lar', 'Cultivar o riso, hobbies e momentos de alegria pura sem utilitarismo.'],
    ['Desconexão Digital', 'Estabelecer limites para telas e redes sociais nos momentos de descanso e refeições.'],
  ]],
  ['Área 3: Trabalho & Atividade', [
    ['Cuidado com o Corpo Físico', 'Sono adequado, alimentação saudável e exercícios entendidos como ato espiritual de mordomia.'],
    ['Limites no Trabalho Profissional', 'Evitar que o trabalho invada as noites, fins de semana e a atenção devida à família.'],
    ['Serviço e Ministério com Limites', 'Servir na igreja/comunidade de forma saudável, sem sacrificar a saúde do casamento.'],
    ['Estresse e Carga Mental', 'Gerenciar a ansiedade cotidiana sem descarregar a irritação no cônjuge.'],
  ]],
  ['Área 4: Relacionamentos & Comunidade', [
    ['Saúde Emocional Conjugal', 'Tempo diário de conversa de qualidade para cultivar a intimidade e alinhar sentimentos.'],
    ['Companheiros de Viagem / Mentoria', 'Manter amizades profundas e prestação de contas com outros casais de fé.'],
    ['Hospitalidade e Serviço em Família', 'Abrir o lar e servir aos outros juntos como expressão do amor de Deus.'],
    ['Resolução de Conflitos', 'Enfrentar divergências com a verdade em amor sem acumular ressentimentos.'],
  ]],
];
const nw = 620, pw = W - 5 * nw;
for (const [t, itens] of areas) {
  C.push(H3(t));
  C.push(table([pw, nw, nw, nw, nw, nw], ['Prática da Regra de Vida', '1', '2', '3', '4', '5'],
    itens.map(([a, b]) => [tdTitle(a, b), CHECK, CHECK, CHECK, CHECK, CHECK]), { center: [1, 2, 3, 4, 5] }));
}
C.push(pageBreak());
C.push(H2('Etapa 2: Compartilhamento e Diálogo a Dois (12 min)'));
C.push(P([run('Instrução para o casal: ', { bold: true, color: GREEN }), run('voltando-se um para o outro na mesa, compartilhem os seus diagnósticos individuais com humildade e acolhimento.')]));
C.push(box([
  boxTitle('Perguntas para guiar a conversa do casal'),
  P('1. Em qual das 4 áreas obtivemos as pontuações mais baixas (mais no piloto automático)? O que tem alimentado essa distração?'),
  P('2. Como a correria do dia a dia tem afetado a nossa capacidade de amar a Deus e cuidar um do outro?'),
  P('3. Qual é o principal obstáculo interno (medo, controle, ansiedade) que nos impede de parar e guardar o Sábado bíblico ou momentos de oração?'),
  P('4. Qual 1 hábito de oração e qual 1 hábito de descanso vamos escolher juntos para ser a nossa nova “treliça” a partir da próxima semana?', { after: 0 }),
], { fill: SAND }));
C.push(H2('Nosso Plano da Treliça do Lar'));
const planCol = (title) => new TableCell({ width: { size: half, type: WidthType.DXA },
  borders: { top: thin(GOLD, 6), bottom: thin(GOLD, 6), left: thin(GOLD, 6), right: thin(GOLD, 6) }, margins: { top: 160, bottom: 160, left: 200, right: 200 },
  children: [P([run(title, { bold: true, color: GREEN, size: 20 })], { after: 60 }), lines(5, half - 400)] });
C.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [half, 240, half], borders: noBorders, rows: [new TableRow({ cantSplit: true, children: [
  planCol('1 hábito de oração / espiritualidade escolhido:'), new TableCell({ width: { size: 240, type: WidthType.DXA }, borders: noBorders, children: [P('')] }), planCol('1 hábito de descanso / lazer escolhido:'),
] })] }));
C.push(H2('Etapa 3: Compartilhamento na Mesa (2 casais · 13 min)'));
C.push(P([run('Pergunta para debate entre os dois casais da mesa: ', { bold: true, color: GREEN }), run('“Por que é tão difícil manter ritmos saudáveis de descanso e oração sem uma comunidade de fé que nos apoie e nos cobre com amor? Como podemos nos encorajar mutuamente após o retiro?”', { italics: true })], { keepNext: true }));
C.push(lines(5));

// ---------- Página final ----------
C.push(pageBreak());
C.push(P('', { before: 2400 }));
C.push(...H1('Anotações'));
C.push(lines(24));

// ======================================================================
// DOCUMENTO
// ======================================================================
const capa = fs.readFileSync(path.join(__dirname, 'capa.png'));
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
  new TextRun({ text: 'Retiro de Casais 2026   ·   ', font: BODY, size: 16, color: GREEN2 }),
  new TextRun({ children: [PageNumber.CURRENT], font: BODY, size: 16, color: GOLD, bold: true }),
] })] });
const header = new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
  new TextRun({ text: 'ESPIRITUALIDADE PROFUNDA E MATURIDADE RELACIONAL', font: BODY, size: 14, color: GOLD, characterSpacing: 30 }),
] })] });

const doc = new Document({
  creator: 'Retiro de Casais 2026', title: 'Apostila — Retiro de Casais 2026',
  styles: { default: { document: { run: { font: BODY, size: 22, color: INK } } } },
  sections: [
    { // Capa sem margens
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 0, bottom: 0, left: 0, right: 0, header: 0, footer: 0 } } },
      children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [new ImageRun({ type: 'png', data: capa, transformation: { width: 793, height: 1120 } })] })],
    },
    {
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1250, bottom: 1134, left: 1134, right: 1134, header: 600, footer: 500 }, pageNumbers: { start: 2 } } },
      headers: { default: header }, footers: { default: footer },
      children: C,
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  const out = path.join(__dirname, '..', 'Apostila_Retiro_de_Casais_2026.docx');
  fs.writeFileSync(out, buf);
  console.log('OK', out);
});
