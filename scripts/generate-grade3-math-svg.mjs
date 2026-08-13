import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outputDir = path.resolve(__dirname, '..', 'exports');
const outputPath = path.join(outputDir, '3학년_수학_단원별이미지_샘플.svg');

const units = [
  {
    unit: '2',
    label: '[9월 작업] 2단원',
    tone: 'active',
    x: 300,
    y: 80,
    items: [
      {
        key: 'division',
        name: '나눗셈',
        publishers: ['동아출판', '비상', '아이스크림', '천재(박)', '천재(한)'],
        files: ['3_수학_2_1_1', '3_수학_2_1_2'],
      },
      {
        key: 'circle',
        name: '원',
        publishers: ['디딤돌', '미래엔', '와이비엠', '지학사'],
        files: ['3_수학_2_2_1', '3_수학_2_2_2'],
      },
    ],
  },
  {
    unit: '3',
    label: '3단원 · 2단원 이미지 재사용',
    tone: 'reuse',
    x: 860,
    y: 80,
    items: [
      {
        key: 'circle-reuse',
        name: '원',
        publishers: ['동아출판', '비상', '아이스크림', '천재(박)', '천재(한)'],
        files: ['3_수학_2_2_1', '3_수학_2_2_2'],
        reuse: true,
      },
    ],
  },
  {
    unit: '4',
    label: '4단원',
    tone: 'default',
    x: 300,
    y: 750,
    items: [
      {
        key: 'fraction',
        name: '분수',
        publishers: ['동아출판', '아이스크림', '와이비엠', '천재(한)'],
        files: ['3_수학_4_1_1', '3_수학_4_1_2'],
      },
      {
        key: 'capacity-weight',
        name: '들이와 무게',
        publishers: ['디딤돌', '미래엔', '지학사', '천재(박)'],
        files: ['3_수학_4_2_1', '3_수학_4_2_2'],
      },
    ],
  },
  {
    unit: '5',
    label: '5단원 · 4단원 이미지 재사용',
    tone: 'reuse',
    x: 860,
    y: 750,
    items: [
      {
        key: 'capacity-weight-reuse',
        name: '들이와 무게',
        publishers: ['동아출판', '비상', '아이스크림', '와이비엠', '천재(한)'],
        files: ['3_수학_4_2_1', '3_수학_4_2_2'],
        reuse: true,
      },
    ],
  },
  {
    unit: '6',
    label: '6단원',
    tone: 'default',
    x: 1140,
    y: 750,
    items: [
      {
        key: 'picture-graph',
        name: '그림그래프',
        publishers: ['동아출판', '디딤돌', '미래엔', '비상', '아이스크림', '와이비엠', '지학사', '천재(박)', '천재(한)'],
        files: ['3_수학_6_1_1', '3_수학_6_1_2'],
      },
    ],
  },
];

const CARD_WIDTH = 250;
const CARD_GAP = 30;
const UNIT_HEADER_HEIGHT = 42;
const NAME_HEIGHT = 42;
const IMAGE_WIDTH = 250;
const IMAGE_HEIGHT = 150;

const escapeXml = (value) => String(value)
  .replaceAll('&', '&amp;')
  .replaceAll('<', '&lt;')
  .replaceAll('>', '&gt;')
  .replaceAll('"', '&quot;')
  .replaceAll("'", '&apos;');

const toneColors = {
  active: { header: '#f97316', panel: '#fff7ed', border: '#fdba74', text: '#ffffff' },
  reuse: { header: '#7e22ce', panel: '#faf5ff', border: '#c084fc', text: '#ffffff' },
  default: { header: '#60646b', panel: '#ffffff', border: '#d5d9df', text: '#ffffff' },
};

function wrapText(text, maxChars) {
  const words = text.split(', ');
  const lines = [];
  let current = '';
  for (const word of words) {
    const candidate = current ? `${current}, ${word}` : word;
    if (candidate.length > maxChars && current) {
      lines.push(current);
      current = word;
    } else {
      current = candidate;
    }
  }
  if (current) lines.push(current);
  return lines;
}

function textLines({ id, x, y, lines, className, lineHeight = 20, anchor = 'start' }) {
  const tspans = lines.map((line, index) => (
    `<tspan id="${id}-line-${index + 1}" x="${x}" dy="${index === 0 ? 0 : lineHeight}">${escapeXml(line)}</tspan>`
  )).join('');
  return `<text id="${id}" class="${className}" x="${x}" y="${y}" text-anchor="${anchor}">${tspans}</text>`;
}

function renderCard(unit, item, index) {
  const x = unit.x + index * (CARD_WIDTH + CARD_GAP);
  const y = unit.y + UNIT_HEADER_HEIGHT + 12;
  const colors = item.reuse ? toneColors.reuse : toneColors.default;
  const publisherLines = wrapText(item.publishers.join(', '), 24);
  const slot1Y = y + NAME_HEIGHT + 12;
  const slot2Y = slot1Y + 278;

  const slot = (slotIndex, slotY, filename) => {
    const isReuse = Boolean(item.reuse);
    const filenameClass = isReuse ? 'filename reuse-text' : 'filename';
    return `
      <g id="unit-${unit.unit}-${item.key}-image-${slotIndex}" data-name="${unit.unit}단원 ${item.name} 이미지 ${slotIndex}">
        <rect id="unit-${unit.unit}-${item.key}-image-${slotIndex}-placeholder" x="${x}" y="${slotY}" width="${IMAGE_WIDTH}" height="${IMAGE_HEIGHT}" rx="3" class="image-placeholder"/>
        <text id="unit-${unit.unit}-${item.key}-image-${slotIndex}-placeholder-label" x="${x + IMAGE_WIDTH / 2}" y="${slotY + IMAGE_HEIGHT / 2 + 5}" text-anchor="middle" class="placeholder-label">IMAGE ${slotIndex}</text>
        <text id="unit-${unit.unit}-${item.key}-image-${slotIndex}-filename-label" x="${x}" y="${slotY + IMAGE_HEIGHT + 25}" class="caption-label">파일명</text>
        <text id="unit-${unit.unit}-${item.key}-image-${slotIndex}-filename" x="${x}" y="${slotY + IMAGE_HEIGHT + 46}" class="${filenameClass}">${escapeXml(filename)}</text>
        ${isReuse ? `<text id="unit-${unit.unit}-${item.key}-image-${slotIndex}-reuse-label" x="${x}" y="${slotY + IMAGE_HEIGHT + 66}" class="reuse-caption">앞 단원 이미지 재사용</text>` : ''}
      </g>`;
  };

  return `
    <g id="unit-${unit.unit}-${item.key}" data-name="${unit.unit}단원 ${item.name}">
      <rect id="unit-${unit.unit}-${item.key}-panel" x="${x - 8}" y="${y - 8}" width="${CARD_WIDTH + 16}" height="600" rx="5" fill="${colors.panel}" stroke="${colors.border}" stroke-width="1"/>
      <g id="unit-${unit.unit}-${item.key}-name-title" data-name="${item.name} 타이틀">
        <rect id="unit-${unit.unit}-${item.key}-name-bg" x="${x}" y="${y}" width="${CARD_WIDTH}" height="${NAME_HEIGHT}" rx="3" class="name-bar"/>
        <text id="unit-${unit.unit}-${item.key}-name" x="${x + CARD_WIDTH / 2}" y="${y + 27}" text-anchor="middle" class="unit-name">${escapeXml(item.name)}</text>
      </g>
      ${slot(1, slot1Y, item.files[0])}
      <line id="unit-${unit.unit}-${item.key}-divider" x1="${x}" y1="${slot2Y - 24}" x2="${x + CARD_WIDTH}" y2="${slot2Y - 24}" class="divider"/>
      ${slot(2, slot2Y, item.files[1])}
      ${textLines({
        id: `unit-${unit.unit}-${item.key}-publishers`,
        x,
        y: slot2Y + IMAGE_HEIGHT + 92,
        lines: publisherLines,
        className: 'publishers',
        lineHeight: 18,
      })}
    </g>`;
}

function renderUnit(unit) {
  const width = unit.items.length * CARD_WIDTH + (unit.items.length - 1) * CARD_GAP;
  const colors = toneColors[unit.tone];
  return `
  <g id="unit-${unit.unit}" data-name="${unit.unit}단원">
    <g id="unit-${unit.unit}-title" data-name="${unit.label} 타이틀">
      <rect id="unit-${unit.unit}-header" x="${unit.x}" y="${unit.y}" width="${width}" height="${UNIT_HEADER_HEIGHT}" rx="4" fill="${colors.header}"/>
      <text id="unit-${unit.unit}-header-label" x="${unit.x + width / 2}" y="${unit.y + 27}" text-anchor="middle" class="unit-header" fill="${colors.text}">${escapeXml(unit.label)}</text>
    </g>
    ${unit.items.map((item, index) => renderCard(unit, item, index)).join('')}
  </g>`;
}

const svg = `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="1600" height="1480" viewBox="0 0 1600 1480" role="img" aria-labelledby="svg-title svg-desc">
  <title id="svg-title">3학년 수학 단원별 이미지 작업표</title>
  <desc id="svg-desc">Figma에서 개별 편집할 수 있도록 모든 단원, 콘텐츠, 이미지 자리, 파일명을 독립된 그룹과 객체로 구성한 SVG 샘플</desc>
  <style>
    text { font-family: Arial, 'Malgun Gothic', sans-serif; fill: #172033; }
    .page-title { font-size: 28px; font-weight: 700; }
    .page-subtitle { font-size: 16px; font-weight: 600; fill: #5b6472; }
    .unit-header { font-size: 17px; font-weight: 700; }
    .unit-name { font-size: 16px; font-weight: 700; }
    .image-placeholder { fill: #d9d9dc; stroke: #c6c9ce; stroke-width: 1; }
    .placeholder-label { font-size: 16px; font-weight: 700; fill: #9a9da4; letter-spacing: 1px; }
    .name-bar { fill: #ececef; }
    .caption-label { font-size: 12px; fill: #717783; }
    .filename { font-size: 15px; font-weight: 700; fill: #252b36; }
    .reuse-text { fill: #7e22ce; }
    .reuse-caption { font-size: 12px; font-weight: 700; fill: #7e22ce; }
    .publishers { font-size: 12px; fill: #5f6672; }
    .divider { stroke: #d1d5db; stroke-width: 1; }
    .rail-grade { font-size: 30px; font-weight: 700; }
    .rail-subject { font-size: 20px; font-weight: 700; fill: #4b5563; }
    .legend { font-size: 13px; font-weight: 700; fill: #5f6672; }
  </style>

  <g id="background" data-name="배경">
    <rect id="canvas-background" x="0" y="0" width="1600" height="1480" fill="#f4f5f7"/>
    <rect id="grade-rail-background" x="30" y="30" width="220" height="1400" rx="4" fill="#dedfe2"/>
    <rect id="content-board-background" x="270" y="30" width="1300" height="1400" rx="8" fill="#ffffff" stroke="#c9cdd3" stroke-width="1"/>
  </g>

  <g id="grade-rail" data-name="학년 과목">
    <text id="grade-label" x="140" y="105" text-anchor="middle" class="rail-grade">3학년</text>
    <text id="subject-label" x="140" y="145" text-anchor="middle" class="rail-subject">수학 · 2학기</text>
    <line id="grade-rail-divider" x1="70" y1="175" x2="210" y2="175" class="divider"/>
    <text id="legend-title" x="70" y="215" class="page-subtitle">표시 기준</text>
    <rect id="legend-new-color" x="70" y="238" width="18" height="18" rx="3" fill="#f97316"/>
    <text id="legend-new-label" x="98" y="252" class="legend">신규 작업 단원</text>
    <rect id="legend-reuse-color" x="70" y="270" width="18" height="18" rx="3" fill="#7e22ce"/>
    <text id="legend-reuse-label" x="98" y="284" class="legend">앞 단원 재사용</text>
    <rect id="legend-default-color" x="70" y="302" width="18" height="18" rx="3" fill="#60646b"/>
    <text id="legend-default-label" x="98" y="316" class="legend">일반 단원</text>
  </g>

  <g id="page-heading" data-name="제목">
    <text id="page-title" x="300" y="62" class="page-title">3학년 수학 단원별 이미지 작업표</text>
    <text id="page-subtitle" x="930" y="60" class="page-subtitle">파일명 규칙: 학년_과목_단원_내용구분_핑퐁번호</text>
  </g>

  <line id="row-divider" x1="300" y1="720" x2="1530" y2="720" class="divider"/>
  ${units.map(renderUnit).join('')}
</svg>
`;

fs.mkdirSync(outputDir, { recursive: true });
fs.writeFileSync(outputPath, svg, 'utf8');
console.log(outputPath);
