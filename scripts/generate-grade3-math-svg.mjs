import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outputDir = path.resolve(__dirname, '..', 'exports');
const outputPath = path.join(outputDir, '3학년_수학_단원별이미지_샘플.svg');

const WIDTH = 6387;
const HEIGHT = 6817;
const COLUMN_X = [552, 1612, 2672, 3732, 4792];
const COLUMN_WIDTH = 928;
const UNIT_HEADER_X = [544, 1604, 2664, 3724, 4784];
const UNIT_HEADER_WIDTH = 936;
const TOP = {
  headerY: 119,
  nameY: 349,
  imageY: [590, 1922],
  dividerY: 1833.5,
};
const BOTTOM = {
  headerY: 3409,
  nameY: 3639,
  imageY: [3880, 5212],
  dividerY: 5123.5,
};

const units = [
  {
    id: 'unit-2',
    label: '[9월 작업] 2단원',
    tone: 'active',
    row: TOP,
    startColumn: 0,
    columnSpan: 2,
    items: [
      { name: '나눗셈', files: ['3_수학_2_1_1', '3_수학_2_1_2'] },
      { name: '원', files: ['3_수학_2_2_1', '3_수학_2_2_2'] },
    ],
  },
  {
    id: 'unit-3',
    label: '[앞단원 재사용] 3단원',
    tone: 'reuse',
    row: TOP,
    startColumn: 2,
    columnSpan: 2,
    items: [
      { name: '원', files: ['3_수학_2_2_1', '3_수학_2_2_2'], reuse: true },
      { name: '나눗셈', files: ['3_수학_2_1_1', '3_수학_2_1_2'], reuse: true },
    ],
  },
  {
    id: 'unit-4',
    label: '4단원',
    tone: 'default',
    row: BOTTOM,
    startColumn: 0,
    columnSpan: 2,
    items: [
      { name: '분수', files: ['3_수학_4_1_1', '3_수학_4_1_2'] },
      { name: '들이와 무게', files: ['3_수학_4_2_1', '3_수학_4_2_2'] },
    ],
  },
  {
    id: 'unit-5',
    label: '[앞단원 재사용] 5단원',
    tone: 'reuse',
    row: BOTTOM,
    startColumn: 2,
    columnSpan: 2,
    items: [
      { name: '분수', files: ['3_수학_4_1_1', '3_수학_4_1_2'], reuse: true },
      { name: '들이와 무게', files: ['3_수학_4_2_1', '3_수학_4_2_2'], reuse: true },
    ],
  },
  {
    id: 'unit-6',
    label: '6단원',
    tone: 'default',
    row: BOTTOM,
    startColumn: 4,
    columnSpan: 1,
    items: [
      { name: '그림그래프', files: ['3_수학_6_1_1', '3_수학_6_1_2'] },
    ],
  },
];

const toneColors = {
  active: '#ff6b00',
  reuse: '#8f3fab',
  default: '#666666',
};

const escapeXml = (value) => String(value)
  .replaceAll('&', '&amp;')
  .replaceAll('<', '&lt;')
  .replaceAll('>', '&gt;')
  .replaceAll('"', '&quot;')
  .replaceAll("'", '&apos;');

function renderThumbnail({ unitId, itemIndex, slotIndex, columnIndex, y, filename, reuse }) {
  const x = COLUMN_X[columnIndex];
  const filenameColor = reuse ? '#8f3fab' : '#111111';

  return `
      <g id="${unitId}-item-${itemIndex + 1}-thumbnail-${slotIndex + 1}" data-name="${filename}">
        <rect id="${unitId}-item-${itemIndex + 1}-thumbnail-${slotIndex + 1}-placeholder" x="${x}" y="${y}" width="${COLUMN_WIDTH}" height="611" fill="#d9d9d9"/>
        <text id="${unitId}-item-${itemIndex + 1}-thumbnail-${slotIndex + 1}-filename" x="${x}" y="${y + 735}" class="filename" fill="${filenameColor}">${escapeXml(filename)}</text>
        <text id="${unitId}-item-${itemIndex + 1}-thumbnail-${slotIndex + 1}-request" x="${x}" y="${y + 855}" class="request-label">요청내용</text>
      </g>`;
}

function renderItem(unit, item, itemIndex) {
  const columnIndex = unit.startColumn + itemIndex;
  const x = COLUMN_X[columnIndex];
  const headerX = UNIT_HEADER_X[columnIndex];

  return `
    <g id="${unit.id}-item-${itemIndex + 1}" data-name="${unit.label} ${item.name}">
      <g id="${unit.id}-item-${itemIndex + 1}-name-title" data-name="${item.name} 타이틀">
        <rect id="${unit.id}-item-${itemIndex + 1}-name-background" x="${x}" y="${unit.row.nameY}" width="${COLUMN_WIDTH}" height="176" rx="20" fill="#e9e9e9"/>
        <text id="${unit.id}-item-${itemIndex + 1}-name" x="${headerX + UNIT_HEADER_WIDTH / 2}" y="${unit.row.nameY + 113}" text-anchor="middle" class="content-name">${escapeXml(item.name)}</text>
      </g>
      ${renderThumbnail({
        unitId: unit.id,
        itemIndex,
        slotIndex: 0,
        columnIndex,
        y: unit.row.imageY[0],
        filename: item.files[0],
        reuse: item.reuse,
      })}
      <line id="${unit.id}-item-${itemIndex + 1}-divider" x1="${headerX - 15}" y1="${unit.row.dividerY}" x2="${headerX + UNIT_HEADER_WIDTH + 39}" y2="${unit.row.dividerY}" stroke="#b7b7b7" stroke-width="4"/>
      ${renderThumbnail({
        unitId: unit.id,
        itemIndex,
        slotIndex: 1,
        columnIndex,
        y: unit.row.imageY[1],
        filename: item.files[1],
        reuse: item.reuse,
      })}
    </g>`;
}

function renderUnit(unit) {
  const x = UNIT_HEADER_X[unit.startColumn];
  const width = unit.columnSpan * UNIT_HEADER_WIDTH + (unit.columnSpan - 1) * 124;
  const color = toneColors[unit.tone];

  return `
  <g id="${unit.id}" data-name="${unit.label}">
    <g id="${unit.id}-title" data-name="${unit.label} 타이틀">
      <rect id="${unit.id}-title-background" x="${x}" y="${unit.row.headerY}" width="${width}" height="176" rx="20" fill="${color}"/>
      <text id="${unit.id}-title-text" x="${x + width / 2}" y="${unit.row.headerY + 113}" text-anchor="middle" class="unit-title">${escapeXml(unit.label)}</text>
    </g>
    ${unit.items.map((item, itemIndex) => renderItem(unit, item, itemIndex)).join('')}
  </g>`;
}

const svg = `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="${WIDTH}" height="${HEIGHT}" viewBox="0 0 ${WIDTH} ${HEIGHT}" role="img" aria-labelledby="svg-title svg-desc">
  <title id="svg-title">3학년 수학 단원별 이미지 작업표</title>
  <desc id="svg-desc">11.svg 배치를 기준으로 만든 Figma 편집용 SVG. 단원 타이틀, 단원명, 섬네일 자리, 파일명이 개별 그룹으로 구성되어 있습니다.</desc>
  <style>
    text { font-family: 'NanumGothic', sans-serif; letter-spacing: 0; }
    .unit-title { fill: #ffffff; font-size: 72px; font-weight: 700; }
    .content-name { fill: #111111; font-size: 68px; font-weight: 700; }
    .filename { font-size: 52px; font-weight: 700; }
    .request-label { fill: #111111; font-size: 46px; font-weight: 400; }
    .row-label { fill: #111111; font-size: 52px; font-weight: 700; }
  </style>

  <g id="board" data-name="작업표 배경">
    <rect id="board-background" x="0.5" y="0.5" width="6386" height="6816" rx="29.5" fill="#ffffff" stroke="#111111"/>
  </g>

  <g id="labels" data-name="행 라벨">
    <text id="unit-label-top" x="230" y="235" class="row-label">단원</text>
    <text id="thumbnail-1-label-top" x="165" y="940" class="row-label">섬네일1</text>
    <text id="thumbnail-2-label-top" x="165" y="2272" class="row-label">섬네일2</text>
    <text id="unit-label-bottom" x="230" y="3525" class="row-label">단원</text>
    <text id="thumbnail-1-label-bottom" x="165" y="4230" class="row-label">섬네일1</text>
    <text id="thumbnail-2-label-bottom" x="165" y="5562" class="row-label">섬네일2</text>
  </g>

  <line id="major-divider" x1="191" y1="3254.5" x2="6060" y2="3254.5" stroke="#b7b7b7" stroke-width="4"/>
  ${units.map(renderUnit).join('')}
</svg>
`;

fs.mkdirSync(outputDir, { recursive: true });
fs.writeFileSync(outputPath, svg, 'utf8');
console.log(outputPath);
