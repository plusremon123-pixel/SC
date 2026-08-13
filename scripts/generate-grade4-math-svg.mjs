import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outputDir = path.resolve(__dirname, '..', 'exports');
const outputPath = path.join(outputDir, '4학년_수학_단원별이미지_샘플.svg');

// v_unit_image_schedule_active_detail의 4학년 수학 데이터만 사용한다.
const rows = [
  {
    y: 58,
    units: [
      {
        number: '2',
        title: '[9월 작업] 2단원',
        tone: 'work',
        items: [
          { name: '삼각형', files: ['4_수학_2_1_1', '4_수학_2_1_2'] },
          { name: '사각형', files: ['4_수학_2_2_1', '4_수학_2_2_2'] },
          { name: '분수의 덧셈과 뺄셈', files: ['4_수학_2_3_1', '4_수학_2_3_2'] },
        ],
      },
      {
        number: '3',
        title: '3단원',
        tone: 'normal',
        items: [
          { name: '소수의 덧셈과 뺄셈', files: ['4_수학_3_1_1', '4_수학_3_1_2'] },
          { name: '사각형', files: ['', '4_수학_2_2_2'], reuse: true },
        ],
      },
    ],
  },
  {
    y: 770,
    units: [
      {
        number: '4',
        title: '[앞단원 재사용] 4단원',
        tone: 'reuse',
        items: [
          { name: '사각형', files: ['', ''], reuse: true },
        ],
      },
      {
        number: '5',
        title: '5단원',
        tone: 'normal',
        items: [
          { name: '꺾은선그래프', files: ['', ''] },
          { name: '다각형', files: ['', ''] },
        ],
      },
      {
        number: '6',
        title: '6단원',
        tone: 'normal',
        items: [
          { name: '꺾은선그래프', files: ['', ''], reuse: true },
          { name: '다각형', files: ['', ''], reuse: true },
          { name: '평면도형의 이동', files: ['', ''] },
        ],
      },
    ],
  },
];

const SVG_WIDTH = 1600;
const SVG_HEIGHT = 1520;
const RAIL_X = 24;
const RAIL_Y = 34;
const RAIL_WIDTH = 205;
const BOARD_X = 262;
const BOARD_Y = 24;
const BOARD_WIDTH = 1310;
const BOARD_HEIGHT = 1470;
const CONTENT_X = 315;
const CONTENT_WIDTH = 1200;
const COLUMN_GAP = 20;
const COLUMN_WIDTH = (CONTENT_WIDTH - COLUMN_GAP * 5) / 6;
const UNIT_GAP = 34;
const UNIT_HEADER_HEIGHT = 36;
const NAME_GAP = 10;
const NAME_HEIGHT = 34;
const IMAGE_GAP = 12;
const IMAGE_HEIGHT = 128;
const SLOT_ROW_GAP = 96;

const colors = {
  work: '#ff6b00',
  reuse: '#9145aa',
  normal: '#666666',
  name: '#e7e7e7',
  image: '#d7d7d7',
  imageStroke: '#cdcdcd',
  board: '#ffffff',
  canvas: '#f4f4f4',
  rail: '#d9d9d9',
  line: '#c9c9c9',
  text: '#111111',
};

const esc = (value) => String(value)
  .replaceAll('&', '&amp;')
  .replaceAll('<', '&lt;')
  .replaceAll('>', '&gt;')
  .replaceAll('"', '&quot;')
  .replaceAll("'", '&apos;');

const totalColumns = (unit) => unit.items.length;

function titleGroup(unit, x, y, width) {
  return `
    <g id="unit-${unit.number}-title" data-name="${esc(unit.title)} 타이틀">
      <rect id="unit-${unit.number}-title-background" x="${x}" y="${y}" width="${width}" height="${UNIT_HEADER_HEIGHT}" rx="4" fill="${colors[unit.tone]}"/>
      <text id="unit-${unit.number}-title-text" x="${x + width / 2}" y="${y + 23}" text-anchor="middle" class="unit-title" fill="#ffffff">${esc(unit.title)}</text>
    </g>`;
}

function nameGroup(unit, item, x, y, index) {
  return `
    <g id="unit-${unit.number}-content-${index + 1}-title" data-name="${esc(item.name)} 타이틀">
      <rect id="unit-${unit.number}-content-${index + 1}-title-background" x="${x}" y="${y}" width="${COLUMN_WIDTH}" height="${NAME_HEIGHT}" rx="3" fill="${colors.name}"/>
      <text id="unit-${unit.number}-content-${index + 1}-title-text" x="${x + COLUMN_WIDTH / 2}" y="${y + 22}" text-anchor="middle" class="content-title">${esc(item.name)}</text>
    </g>`;
}

function slotGroup(unit, item, x, y, itemIndex, slotIndex) {
  const filename = item.files[slotIndex - 1] || '';
  const filenameClass = item.reuse && filename ? 'filename reuse-filename' : 'filename';
  return `
    <g id="unit-${unit.number}-content-${itemIndex + 1}-slot-${slotIndex}" data-name="${esc(item.name)} 섬네일${slotIndex}">
      <rect id="unit-${unit.number}-content-${itemIndex + 1}-slot-${slotIndex}-image" x="${x}" y="${y}" width="${COLUMN_WIDTH}" height="${IMAGE_HEIGHT}" fill="${colors.image}" stroke="${colors.imageStroke}" stroke-width="0.8"/>
      <text id="unit-${unit.number}-content-${itemIndex + 1}-slot-${slotIndex}-filename-label" x="${x}" y="${y + IMAGE_HEIGHT + 22}" class="meta-label">파일명</text>
      ${filename ? `<text id="unit-${unit.number}-content-${itemIndex + 1}-slot-${slotIndex}-filename" x="${x}" y="${y + IMAGE_HEIGHT + 41}" class="${filenameClass}">${esc(filename)}</text>` : ''}
      <text id="unit-${unit.number}-content-${itemIndex + 1}-slot-${slotIndex}-request" x="${x}" y="${y + IMAGE_HEIGHT + 59}" class="request-label">${filename ? '이미지 활용' : '요청내용'}</text>
    </g>`;
}

function renderUnit(unit, unitX, rowY) {
  const width = totalColumns(unit) * COLUMN_WIDTH + (totalColumns(unit) - 1) * COLUMN_GAP;
  const titleY = rowY;
  const nameY = titleY + UNIT_HEADER_HEIGHT + NAME_GAP;
  const slot1Y = nameY + NAME_HEIGHT + IMAGE_GAP;
  const slot2Y = slot1Y + IMAGE_HEIGHT + 60 + SLOT_ROW_GAP;
  const cards = unit.items.map((item, index) => {
    const x = unitX + index * (COLUMN_WIDTH + COLUMN_GAP);
    return `
      <g id="unit-${unit.number}-content-${index + 1}" data-name="${unit.number}단원 ${esc(item.name)}">
        ${nameGroup(unit, item, x, nameY, index)}
        ${slotGroup(unit, item, x, slot1Y, index, 1)}
        ${slotGroup(unit, item, x, slot2Y, index, 2)}
      </g>`;
  }).join('');

  return `
  <g id="unit-${unit.number}" data-name="${unit.number}단원">
    ${titleGroup(unit, unitX, titleY, width)}
    ${cards}
  </g>`;
}

function renderRow(row) {
  let cursorX = CONTENT_X;
  const output = [];
  for (const unit of row.units) {
    output.push(renderUnit(unit, cursorX, row.y));
    cursorX += totalColumns(unit) * COLUMN_WIDTH + (totalColumns(unit) - 1) * COLUMN_GAP + UNIT_GAP;
  }
  return output.join('');
}

const svg = `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="${SVG_WIDTH}" height="${SVG_HEIGHT}" viewBox="0 0 ${SVG_WIDTH} ${SVG_HEIGHT}" role="img" aria-labelledby="svg-title svg-desc">
  <title id="svg-title">4학년 수학 단원별 이미지 작업표</title>
  <desc id="svg-desc">Supabase에 저장된 4학년 수학 단원 데이터를 확정 화면 구성으로 배치한 Figma 편집용 SVG</desc>
  <style>
    text { font-family: 'NanumGothic', sans-serif; fill: ${colors.text}; }
    .unit-title { font-size: 14px; font-weight: 700; }
    .content-title { font-size: 14px; font-weight: 700; }
    .row-label { font-size: 14px; font-weight: 700; }
    .grade-label { font-size: 27px; font-weight: 700; }
    .section-label { font-size: 13px; font-weight: 700; }
    .meta-label { font-size: 10px; fill: #555555; }
    .filename { font-size: 11px; font-weight: 700; }
    .reuse-filename { fill: #8b2ca3; }
    .request-label { font-size: 10px; fill: #222222; }
  </style>

  <g id="background" data-name="배경">
    <rect id="canvas-background" x="0" y="0" width="${SVG_WIDTH}" height="${SVG_HEIGHT}" fill="${colors.canvas}"/>
    <rect id="grade-background" x="${RAIL_X}" y="${RAIL_Y}" width="${RAIL_WIDTH}" height="142" fill="${colors.rail}"/>
    <rect id="board-background" x="${BOARD_X}" y="${BOARD_Y}" width="${BOARD_WIDTH}" height="${BOARD_HEIGHT}" rx="5" fill="${colors.board}" stroke="#bdbdbd" stroke-width="1"/>
  </g>

  <g id="grade-title" data-name="4학년 타이틀">
    <text id="grade-title-text" x="${RAIL_X + RAIL_WIDTH / 2}" y="${RAIL_Y + 79}" text-anchor="middle" class="grade-label">4학년</text>
  </g>

  <g id="layout-labels" data-name="행 구분">
    <text id="unit-label" x="285" y="74" class="section-label">단원</text>
    <text id="thumbnail-1-label" x="300" y="222" text-anchor="end" class="row-label">섬네일1</text>
    <text id="thumbnail-2-label" x="300" y="505" text-anchor="end" class="row-label">섬네일2</text>
  </g>

  <line id="top-thumbnail-divider" x1="${CONTENT_X}" y1="450" x2="${CONTENT_X + CONTENT_WIDTH}" y2="450" stroke="${colors.line}" stroke-width="1"/>
  <line id="section-divider" x1="285" y1="735" x2="${CONTENT_X + CONTENT_WIDTH}" y2="735" stroke="${colors.line}" stroke-width="1"/>
  <line id="bottom-thumbnail-divider" x1="${CONTENT_X}" y1="1162" x2="${CONTENT_X + CONTENT_WIDTH}" y2="1162" stroke="${colors.line}" stroke-width="1"/>

  ${rows.map(renderRow).join('')}
</svg>
`;

fs.mkdirSync(outputDir, { recursive: true });
fs.writeFileSync(outputPath, svg.replace(/[ \t]+$/gm, ''), 'utf8');
console.log(outputPath);
