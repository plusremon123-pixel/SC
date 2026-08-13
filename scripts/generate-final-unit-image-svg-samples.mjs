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
const DEFAULT_COLUMNS_PER_BAND = 5;
const MAX_MATH_COLUMNS_PER_BAND = 6;

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
    grade: 2,
    subject: '국어',
    output: '2학년_국어_단원별이미지_최종샘플.svg',
    bands: [
      [
        unit('2', '[9월 작업] 2단원', 'active', [item('서로 존중해요', ['2_국어_2_1', ''])]),
        unit('3', '3단원', 'default', [item('내용을 살펴요', ['2_국어_3_1', ''])]),
        unit('4', '4단원', 'default', [item('마음을 전해요', ['2_국어_4_1', '2_국어_4_2'])]),
        unit('5', '5단원', 'default', [item('바른 말로 이야기 나누어요', ['2_국어_5_1', '2_국어_5_2'])]),
        unit('6', '6단원', 'default', [item('매체를 경험해요', ['2_국어_6_1', '2_국어_6_2'])]),
      ],
      [
        unit('7', '7단원', 'default', [item('내 생각은 이래요', ['2_국어_7_1', '2_국어_7_2'])]),
        unit('8', '8단원', 'default', [item('나도 작가', ['2_국어_8_1', '2_국어_8_2'])]),
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
    grade: 6,
    subject: '수학',
    output: '6학년_수학_단원별이미지_최종샘플.svg',
    bands: [
      [
        unit('2', '[9월 작업] 2단원', 'active', [
          item('공간과 입체', ['6_수학_2_1_1', '6_수학_2_1_2']),
          item('소수의 나눗셈', ['6_수학_2_2_1', '6_수학_2_2_2']),
        ]),
        unit('3', '[앞단원 재사용] 3단원', 'reuse', [
          item('공간과 입체', ['6_수학_2_1_1', '6_수학_2_1_2'], true),
        ]),
        unit('4', '4단원', 'default', [
          item('비례식과 비례배분', ['6_수학_4_1_1', '6_수학_4_1_2']),
          item('원의 둘레와 넓이', ['6_수학_4_2_1', '6_수학_4_2_2']),
        ]),
      ],
      [
        unit('5', '5단원', 'default', [
          item(['원의 넓이', '원주율과 원의 넓이'], ['6_수학_5_1_1', '6_수학_5_1_2']),
          item('원의 둘레와 넓이', ['6_수학_4_2_1', '6_수학_4_2_2'], true),
        ]),
        unit('6', '6단원', 'default', [
          item('원기둥, 원뿔, 구', ['6_수학_6_1', '6_수학_6_2']),
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
    grade: 1,
    subject: '바슬즐',
    output: '1학년_바슬즐_단원별이미지_최종샘플.svg',
    bands: [
      [
        unit('2', '[9월 작업] 2단원', 'active', [item('약속', ['1_바슬즐_2_1', '1_바슬즐_2_2'])]),
        unit('3', '3단원', 'default', [item('상상', ['1_바슬즐_3_1', '1_바슬즐_3_2'])]),
        unit('4', '4단원', 'default', [item('이야기', ['1_바슬즐_4_1', '1_바슬즐_4_2'])]),
      ],
    ],
  },
  {
    grade: 3,
    subject: '과학',
    output: '3학년_과학_단원별이미지_최종샘플.svg',
    bands: [
      [
        unit('2', '[9월 작업] 2단원', 'active', [item('지구와 바다', ['3_과학_2_1', '3_과학_2_2'])]),
        unit('3', '3단원', 'default', [item('소리의 성질', ['3_과학_3_1', '3_과학_3_2'])]),
        unit('4', '4단원', 'default', [item('감염병과 건강한 생활', ['3_과학_4_1', '3_과학_4_2'])]),
      ],
    ],
  },
  {
    grade: 6,
    subject: '영어',
    output: '6학년_영어_단원별이미지_최종샘플.svg',
    bands: [
      [
        unit('9', '[9월 작업] 9단원', 'active', [item('How Can I Get to the Museum?', ['6_영어_9_1', ''])]),
        unit('10', '10단원', 'default', [item('Who Painted the Picture?', ['6_영어_10_1', '6_영어_10_2'])]),
        unit('11', '11단원', 'default', [item('Why Are You Happy?', ['6_영어_11_1', '6_영어_11_2'])]),
        unit('12', '12단원', 'default', [item('How About Turning Off the Light?', ['6_영어_12_1', '6_영어_12_2'])]),
        unit('13', '13단원', 'default', [item('What Do You Think?', ['6_영어_13_1', '6_영어_13_2'])]),
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

function reflowBands(sample) {
  const units = sample.bands.flat();
  const hasMultiItemUnit = units.some((unitData) => unitData.items.length > 1);
  const maxColumns = sample.subject === '수학' && hasMultiItemUnit
    ? MAX_MATH_COLUMNS_PER_BAND
    : DEFAULT_COLUMNS_PER_BAND;
  const bands = [[]];
  let availableColumns = maxColumns;

  for (const unitData of units) {
    if (unitData.items.length <= maxColumns && unitData.items.length > availableColumns) {
      bands.push([]);
      availableColumns = maxColumns;
    }

    let itemIndex = 0;

    while (itemIndex < unitData.items.length) {
      if (availableColumns === 0) {
        bands.push([]);
        availableColumns = maxColumns;
      }

      const itemCount = Math.min(availableColumns, unitData.items.length - itemIndex);
      bands.at(-1).push({
        ...unitData,
        items: unitData.items.slice(itemIndex, itemIndex + itemCount),
      });
      itemIndex += itemCount;
      availableColumns -= itemCount;
    }
  }

  return bands.filter((band) => band.length > 0);
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
  const text = lines.join(' ');
  const widthUnits = [...text].reduce((sum, character) => {
    if (/\s/u.test(character)) return sum + 0.32;
    if (/[\u3131-\uD79D]/u.test(character)) return sum + 0.96;
    return sum + 0.58;
  }, 0);
  const fontSize = Math.max(24, Math.min(68, Math.floor(820 / Math.max(widthUnits, 1))));
  const baselineY = y + 88 + fontSize * 0.35;
  return `<text id="${id}" x="${x}" y="${baselineY}" text-anchor="middle" style="font-size:${fontSize}px;font-weight:700">${esc(text)}</text>`;
}

function wrapDisplayName(value) {
  return [String(value || '').trim()];
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
          lines: wrapDisplayName(itemData.names[0]),
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
  const layoutBands = reflowBands(sample);
  const maxColumns = Math.max(...layoutBands.map((band) => band.reduce((sum, unitData) => sum + unitData.items.length, 0)), DEFAULT_COLUMNS_PER_BAND);
  const width = 6387 + (maxColumns - 5) * COLUMN_STEP;
  const height = layoutBands.length * BAND_HEIGHT + 237;
  const dividers = layoutBands.slice(0, -1).map((_, index) => (
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
    ${layoutBands.map((_, bandIndex) => renderLabels(bandIndex)).join('')}
  </g>
  ${dividers}
  ${layoutBands.map((band, bandIndex) => renderBand(band, bandIndex)).join('')}
</svg>
`;
}

fs.mkdirSync(outputDir, { recursive: true });
for (const sample of samples) {
  const outputPath = path.join(outputDir, sample.output);
  fs.writeFileSync(outputPath, renderSample(sample), 'utf8');
  console.log(outputPath);
}
