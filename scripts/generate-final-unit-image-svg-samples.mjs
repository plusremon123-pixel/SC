import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outputDir = path.resolve(__dirname, '..', 'exports');

const COLUMN_STEP = 1060;
const COLUMN_X_BASE = 552;
const HEADER_X_BASE = 544;
const COLUMN_WIDTH = 928;
const HEADER_COLUMN_WIDTH = 936;
const BAND_HEIGHT = 3290;

const samples = [
  {
    grade: 1,
    subject: '국어',
    output: '1학년_국어_단원별이미지_최종샘플.svg',
    bands: [
      [
        unit('3', '[9월 작업] 3단원', 'active', [item('그림일기를 써요', ['1_국어_3_1', '1_국어_3_2'])]),
        unit('4', '4단원', 'default', [item('감동을 나누어요', ['1_국어_4_1', '1_국어_4_2'])]),
        unit('5', '5단원', 'default', [item('생각을 키워요', ['1_국어_5_1', '1_국어_5_2'])]),
      ],
      [
        unit('6', '6단원', 'default', [item('문장을 읽고 써요', ['1_국어_6_1', '1_국어_6_2'])]),
        unit('7', '7단원', 'default', [item('무엇이 중요할까요', ['1_국어_7_1', '1_국어_7_2'])]),
        unit('8', '8단원', 'default', [item('느끼고 표현해요', ['1_국어_8_1', '1_국어_8_2'])]),
      ],
    ],
  },
  {
    grade: 4,
    subject: '수학',
    output: '4학년_수학_단원별이미지_최종샘플.svg',
    bands: [
      [
        unit('2', '[9월 작업] 2단원', 'active', [
          item('삼각형', ['4_수학_2_1_1', '4_수학_2_1_2']),
          item('사각형', ['4_수학_2_2_1', '4_수학_2_2_2']),
          item('분수의 덧셈과 뺄셈', ['4_수학_2_3_1', '4_수학_2_3_2']),
        ]),
        unit('3', '3단원', 'default', [
          item('소수의 덧셈과 뺄셈', ['4_수학_3_1_1', '4_수학_3_1_2']),
          item('규칙 찾기와 식 만들기', ['4_수학_3_2_1', '4_수학_3_2_2']),
          item('사각형', ['4_수학_2_2_1', '4_수학_2_2_2'], true),
        ]),
      ],
      [
        unit('4', '4단원', 'default', [
          item('사각형', ['4_수학_2_2_1', '4_수학_2_2_2'], true),
          item('소수의 덧셈과 뺄셈', ['4_수학_3_1_1', '4_수학_3_1_2'], true),
          item('다각형', ['4_수학_4_1_1', '4_수학_4_1_2']),
        ]),
        unit('5', '5단원', 'default', [
          item(['꺾은선그래프', '자료와 꺾은선그래프'], ['4_수학_5_1_1', '4_수학_5_1_2']),
          item('다각형', ['4_수학_4_1_1', '4_수학_4_1_2'], true),
        ]),
      ],
      [
        unit('6', '6단원', 'default', [
          item('꺾은선그래프', ['4_수학_5_1_1', '4_수학_5_1_2'], true),
          item('다각형', ['4_수학_4_1_1', '4_수학_4_1_2'], true),
          item('평면도형의 이동', ['4_수학_6_1_1', '4_수학_6_1_2']),
        ]),
      ],
    ],
  },
  {
    grade: 5,
    subject: '수학',
    output: '5학년_수학_단원별이미지_최종샘플.svg',
    bands: [
      [
        unit('2', '[9월 작업] 2단원', 'active', [item('분수의 곱셈', ['5_수학_2_1', '5_수학_2_2'])]),
        unit('3', '3단원', 'default', [item('합동과 대칭', ['5_수학_3_1', '5_수학_3_2'])]),
      ],
      [
        unit('4', '4단원', 'default', [item('소수의 곱셈', ['5_수학_4_1', '5_수학_4_2'])]),
        unit('5', '5단원', 'default', [item(['직육면체와 정육면체', '직육면체'], ['5_수학_5_1', '5_수학_5_2'])]),
        unit('6', '6단원', 'default', [item('평균과 가능성', ['5_수학_6_1', '5_수학_6_2'])]),
      ],
    ],
  },
  {
    grade: 3,
    subject: '사회',
    output: '3학년_사회_단원별이미지_최종샘플.svg',
    bands: [
      [
        unit('1-2', '[9월 작업] 1-2단원', 'active', [item([
          '다양한 문화에 대한 이해와 존중',
          '다양한 문화의 모습과 존중',
          '다양한 문화의 이해와 존중',
        ], ['3_사회_1-2_1', '3_사회_1-2_2'])]),
        unit('2-1', '2-1단원', 'default', [item([
          '옛날과 오늘날의 풍습',
          '옛날과 오늘날의 달라진 풍습',
          '옛날의 풍습과 오늘날의 변화 모습',
        ], ['3_사회_2-1_1', '3_사회_2-1_2'])]),
      ],
      [
        unit('2-2', '2-2단원', 'default', [item([
          '교통의 변화로 달라진 생활 모습',
          '교통 발달에 따른 생활 모습의 변화',
          '교통의 발달과 생활 모습의 변화',
          '교통의 변화와 달라진 생활 모습',
          '교통의 변화와 생활 모습',
        ], ['3_사회_2-2_1', '3_사회_2-2_2'])]),
        unit('2-3', '2-3단원', 'default', [item([
          '통신수단의 변화로 달라진 생활 모습',
          '통신수단 발달에 따른 생활 모습의 변화',
          '통신수단의 발달과 생활 모습의 변화',
          '통신수단의 변화와 달라진 생활 모습',
        ], ['3_사회_2-3_1', '3_사회_2-3_2'])]),
      ],
    ],
  },
];

function item(names, files, reuse = false) {
  return {
    names: Array.isArray(names) ? names : [names],
    files,
    reuse,
  };
}

function unit(number, label, tone, items) {
  return { number, label, tone, items };
}

const colors = {
  active: '#ff6b00',
  reuse: '#8f3fab',
  default: '#666666',
};

const esc = (value) => String(value)
  .replaceAll('&', '&amp;')
  .replaceAll('<', '&lt;')
  .replaceAll('>', '&gt;')
  .replaceAll('"', '&quot;')
  .replaceAll("'", '&apos;');

function bandMetrics(bandIndex) {
  const offset = bandIndex * BAND_HEIGHT;
  return {
    headerY: 119 + offset,
    nameY: 349 + offset,
    imageY: [590 + offset, 1922 + offset],
    dividerY: 1833.5 + offset,
  };
}

function renderNameText({ id, x, y, lines }) {
  const lineCount = lines.length;
  const fontSize = lineCount === 1 ? 68 : lineCount === 2 ? 42 : lineCount === 3 ? 32 : 25;
  const lineHeight = lineCount === 1 ? 0 : lineCount === 2 ? 52 : lineCount === 3 ? 39 : 30;
  const blockHeight = lineCount === 1 ? fontSize : fontSize + (lineCount - 1) * lineHeight;
  const startY = y + (176 - blockHeight) / 2 + fontSize * 0.78;
  const tspans = lines.map((line, index) => (
    `<tspan id="${id}-line-${index + 1}" x="${x}" dy="${index === 0 ? 0 : lineHeight}">${esc(line)}</tspan>`
  )).join('');
  return `<text id="${id}" x="${x}" y="${startY}" text-anchor="middle" style="font-size:${fontSize}px;font-weight:700">${tspans}</text>`;
}

function renderSlot({ unitId, itemIndex, slotIndex, columnIndex, y, filename, reuse }) {
  const x = COLUMN_X_BASE + columnIndex * COLUMN_STEP;
  return `
      <g id="${unitId}-item-${itemIndex + 1}-thumbnail-${slotIndex + 1}" data-name="${filename}">
        <rect id="${unitId}-item-${itemIndex + 1}-thumbnail-${slotIndex + 1}-placeholder" x="${x}" y="${y}" width="${COLUMN_WIDTH}" height="611" fill="#d9d9d9"/>
        <text id="${unitId}-item-${itemIndex + 1}-thumbnail-${slotIndex + 1}-filename" x="${x}" y="${y + 735}" class="filename${reuse ? ' reuse-filename' : ''}">${esc(filename)}</text>
        <text id="${unitId}-item-${itemIndex + 1}-thumbnail-${slotIndex + 1}-request" x="${x}" y="${y + 855}" class="request-label">요청내용</text>
      </g>`;
}

function renderItem(unitId, itemData, itemIndex, columnIndex, metrics) {
  const x = COLUMN_X_BASE + columnIndex * COLUMN_STEP;
  const headerX = HEADER_X_BASE + columnIndex * COLUMN_STEP;
  const itemName = itemData.names.join(' / ');
  return `
    <g id="${unitId}-item-${itemIndex + 1}" data-name="${esc(itemName)}">
      <desc>출판사별 단원명: ${esc(itemName)}</desc>
      <g id="${unitId}-item-${itemIndex + 1}-name-title" data-name="${esc(itemName)} 타이틀">
        <rect id="${unitId}-item-${itemIndex + 1}-name-background" x="${x}" y="${metrics.nameY}" width="${COLUMN_WIDTH}" height="176" rx="20" fill="#e9e9e9"/>
        ${renderNameText({
          id: `${unitId}-item-${itemIndex + 1}-name`,
          x: headerX + HEADER_COLUMN_WIDTH / 2,
          y: metrics.nameY,
          lines: [itemData.names[0]],
        })}
      </g>
      ${renderSlot({ unitId, itemIndex, slotIndex: 0, columnIndex, y: metrics.imageY[0], filename: itemData.files[0], reuse: itemData.reuse })}
      <line id="${unitId}-item-${itemIndex + 1}-divider" x1="${headerX - 15}" y1="${metrics.dividerY}" x2="${headerX + HEADER_COLUMN_WIDTH + 39}" y2="${metrics.dividerY}" stroke="#b7b7b7" stroke-width="4"/>
      ${renderSlot({ unitId, itemIndex, slotIndex: 1, columnIndex, y: metrics.imageY[1], filename: itemData.files[1], reuse: itemData.reuse })}
    </g>`;
}

function renderUnit(unitData, bandIndex, startColumn) {
  const metrics = bandMetrics(bandIndex);
  const unitId = `band-${bandIndex + 1}-unit-${unitData.number.replaceAll('-', '_')}`;
  const x = HEADER_X_BASE + startColumn * COLUMN_STEP;
  const width = unitData.items.length * HEADER_COLUMN_WIDTH + (unitData.items.length - 1) * 124;
  return `
  <g id="${unitId}" data-name="${esc(unitData.label)}">
    <g id="${unitId}-title" data-name="${esc(unitData.label)} 타이틀">
      <rect id="${unitId}-title-background" x="${x}" y="${metrics.headerY}" width="${width}" height="176" rx="20" fill="${colors[unitData.tone]}"/>
      <text id="${unitId}-title-text" x="${x + width / 2}" y="${metrics.headerY + 113}" text-anchor="middle" class="unit-title">${esc(unitData.label)}</text>
    </g>
    ${unitData.items.map((itemData, itemIndex) => renderItem(unitId, itemData, itemIndex, startColumn + itemIndex, metrics)).join('')}
  </g>`;
}

function renderBand(band, bandIndex) {
  let startColumn = 0;
  const output = [];
  for (const unitData of band) {
    output.push(renderUnit(unitData, bandIndex, startColumn));
    startColumn += unitData.items.length;
  }
  return output.join('');
}

function renderLabels(bandIndex) {
  const metrics = bandMetrics(bandIndex);
  return `
    <g id="band-${bandIndex + 1}-labels" data-name="${bandIndex + 1}번째 행 라벨">
      <text id="band-${bandIndex + 1}-unit-label" x="230" y="${metrics.headerY + 116}" class="row-label">단원</text>
      <text id="band-${bandIndex + 1}-thumbnail-1-label" x="165" y="${metrics.imageY[0] + 350}" class="row-label">섬네일1</text>
      <text id="band-${bandIndex + 1}-thumbnail-2-label" x="165" y="${metrics.imageY[1] + 350}" class="row-label">섬네일2</text>
    </g>`;
}

function renderSample(sample) {
  const maxColumns = Math.max(...sample.bands.map((band) => band.reduce((sum, unitData) => sum + unitData.items.length, 0)), 5);
  const width = 6387 + (maxColumns - 5) * COLUMN_STEP;
  const height = sample.bands.length * BAND_HEIGHT + 237;
  const dividers = sample.bands.slice(0, -1).map((_, index) => (
    `<line id="band-${index + 1}-major-divider" x1="191" y1="${3254.5 + index * BAND_HEIGHT}" x2="${width - 327}" y2="${3254.5 + index * BAND_HEIGHT}" stroke="#b7b7b7" stroke-width="4"/>`
  )).join('');

  return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" role="img" aria-labelledby="svg-title svg-desc">
  <title id="svg-title">${sample.grade}학년 ${sample.subject} 단원별 이미지 최종 샘플</title>
  <desc id="svg-desc">확정 배치와 파일명 규칙을 적용한 Figma 편집용 SVG 샘플</desc>
  <style>
    text { font-family: 'NanumGothic', sans-serif; fill: #111111; letter-spacing: 0; }
    .unit-title { fill: #ffffff; font-size: 72px; font-weight: 700; }
    .filename { font-size: 52px; font-weight: 700; }
    .reuse-filename { fill: #8f3fab; }
    .request-label { font-size: 46px; font-weight: 400; }
    .row-label { font-size: 52px; font-weight: 700; }
  </style>
  <g id="board" data-name="작업표 배경">
    <rect id="board-background" x="0.5" y="0.5" width="${width - 1}" height="${height - 1}" rx="29.5" fill="#ffffff" stroke="#111111"/>
  </g>
  <g id="labels" data-name="행 라벨">
    ${sample.bands.map((_, bandIndex) => renderLabels(bandIndex)).join('')}
  </g>
  ${dividers}
  ${sample.bands.map((band, bandIndex) => renderBand(band, bandIndex)).join('')}
</svg>
`;
}

fs.mkdirSync(outputDir, { recursive: true });
for (const sample of samples) {
  const outputPath = path.join(outputDir, sample.output);
  fs.writeFileSync(outputPath, renderSample(sample), 'utf8');
  console.log(outputPath);
}
