
'use strict';

// HWPX / DOCX 문서 뷰어
// JSZip으로 압축 해제, DOMParser로 XML 읽기

const $ = (id) => document.getElementById(id);
const PX = 1 / 75;

const MIME = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  webp: 'image/webp'
};

// XML 헬퍼
const kids = (el, n) =>
  el ? Array.from(el.children).filter(c => c.localName === n) : [];

const first = (el, n) => kids(el, n)[0];

const all = (el, n) =>
  el ? Array.from(el.getElementsByTagNameNS('*', n)) : [];

const num = (v) => Number(v) || 0;

function parseXml(text, label) {
  const doc = new DOMParser().parseFromString(
    text, 'application/xml'
  );

  if (doc.getElementsByTagName('parsererror').length) {
    throw new Error(`${label} 파싱 실패`);
  }

  return doc;
}

const findEntry = (zip, path) => {
  const p = path.replace(/^\.?\//, '').toLowerCase();

  return Object.values(zip.files).find(
    f => !f.dir && f.name.toLowerCase() === p
  );
};

async function readXml(zip, path) {
  const f = findEntry(zip, path);

  if (!f) {
    throw new Error(`${path} 파일이 없습니다`);
  }

  return parseXml(await f.async('string'), path);
}

const baseName = (p) =>
  p.split('/').pop().replace(/\.[^.]+$/, '');

function missing(msg) {
  const s = document.createElement('span');
  s.className = 'missing';
  s.textContent = msg;
  return s;
}

// 이미지 URL 관리
let currentUrls = [];

function clearUrls() {
  for (const url of currentUrls) {
    URL.revokeObjectURL(url);
  }

  currentUrls = [];
}

async function imageUrl(file) {
  const ext = file.name.split('.').pop().toLowerCase();

  const blob = new Blob(
    [await file.async('uint8array')],
    { type: MIME[ext] || 'application/octet-stream' }
  );

  const url = URL.createObjectURL(blob);
  currentUrls.push(url);

  return url;
}

// ========================================
// HWPX 처리
// ========================================

// 1. HWPX 스타일 분석
function parseHeader(doc) {
  const charPr = new Map();
  const paraPr = new Map();

  for (const el of all(doc, 'charPr')) {
    const ul = first(el, 'underline');
    const st = first(el, 'strikeout');

    charPr.set(el.getAttribute('id'), {
      size: num(el.getAttribute('height')) / 100,
      color: el.getAttribute('textColor'),
      bold: !!first(el, 'bold'),
      italic: !!first(el, 'italic'),
      underline: !!ul && ul.getAttribute('type') !== 'NONE',
      strike: !!st && st.getAttribute('shape') !== 'NONE'
    });
  }

  for (const el of all(doc, 'paraPr')) {
    const align = all(el, 'align')[0]?.getAttribute('horizontal');
    const ls = all(el, 'lineSpacing')[0];

    paraPr.set(el.getAttribute('id'), {
      align: {
        CENTER: 'center',
        RIGHT: 'right',
        JUSTIFY: 'justify',
        LEFT: 'left'
      }[align],

      lineHeight:
        ls && ls.getAttribute('type') === 'PERCENT'
          ? num(ls.getAttribute('value')) / 100
          : 0
    });
  }

  return { charPr, paraPr };
}

// 2. HWPX 내부 파일 목록
async function loadManifest(zip) {
  const f = findEntry(zip, 'Contents/content.hpf');

  const items = new Map();
  let spine = [];

  if (f) {
    const doc = parseXml(
      await f.async('string'),
      'content.hpf'
    );

    for (const it of all(doc, 'item')) {
      items.set(
        it.getAttribute('id'),
        it.getAttribute('href')
      );
    }

    spine = all(doc, 'itemref')
      .map(r => items.get(r.getAttribute('idref')))
      .filter(h => /section\d+\.xml$/i.test(h || ''));
  }

  if (!spine.length) {
    spine = Object.keys(zip.files)
      .filter(n => /^contents\/section\d+\.xml$/i.test(n))
      .sort();
  }

  return { items, spine };
}

// 3. HWPX 이미지 로드
async function loadImages(zip, manifest) {
  const urls = new Map();

  for (const f of Object.values(zip.files)) {
    if (f.dir || !/^bindata\//i.test(f.name)) {
      continue;
    }

    urls.set(baseName(f.name), await imageUrl(f));
  }

  for (const [id, href] of manifest.items) {
    const url = href && urls.get(baseName(href));

    if (url) {
      urls.set(id, url);
    }
  }

  return urls;
}

// 4. 텍스트 출력
function renderText(t, span) {
  for (const n of t.childNodes) {
    if (n.nodeType === Node.TEXT_NODE) {
      span.append(n.data);
    } else if (n.localName === 'lineBreak') {
      span.append(document.createElement('br'));
    } else if (n.localName === 'tab') {
      span.append('\u2003');
    } else if (n.localName === 'nbSpace') {
      span.append('\u00A0');
    } else if (n.localName === 'fwSpace') {
      span.append('\u3000');
    }
  }
}

// 5. 글자 스타일 적용
function applyChar(span, c) {
  if (!c) return;

  if (c.size) span.style.fontSize = c.size + 'pt';

  const color = (c.color || '').replace(/^#?/, '#');

  if (/^#[0-9a-f]{6}$/i.test(color)) {
    span.style.color = color;
  }

  if (c.bold) span.style.fontWeight = 'bold';
  if (c.italic) span.style.fontStyle = 'italic';

  const deco = [
    c.underline && 'underline',
    c.strike && 'line-through'
  ].filter(Boolean).join(' ');

  if (deco) span.style.textDecoration = deco;
}

// 6. HWPX 이미지 출력
function renderPic(pic, ctx) {
  const ref = all(pic, 'img')[0]
    ?.getAttribute('binaryItemIDRef');

  const url = ctx.urls.get(ref);

  if (!url) {
    return missing(`[이미지 없음: ${ref}]`);
  }

  const w =
    num(first(pic, 'curSz')?.getAttribute('width')) ||
    num(first(pic, 'sz')?.getAttribute('width'));

  const img = new Image();

  img.src = url;
  img.alt = '문서 이미지';
  img.style.maxWidth = '100%';

  if (w) {
    img.style.width = w * PX + 'px';
  }

  img.onerror = () =>
    img.replaceWith(missing('[표시할 수 없는 이미지 형식]'));

  return img;
}

// 7. HWPX 표 출력
function renderTable(tbl, ctx) {
  const t = document.createElement('table');
  t.className = 'hwpx-tbl';

  const sz = first(tbl, 'sz');

  if (sz) {
    t.style.width = num(sz.getAttribute('width')) * PX + 'px';
  }

  for (const tr of kids(tbl, 'tr')) {
    const row = t.insertRow();

    for (const tc of kids(tr, 'tc')) {
      const td = row.insertCell();
      const span = first(tc, 'cellSpan');

      if (span) {
        td.colSpan = num(span.getAttribute('colSpan')) || 1;
        td.rowSpan = num(span.getAttribute('rowSpan')) || 1;
      }

      const cw = num(
        first(tc, 'cellSz')?.getAttribute('width')
      );

      if (cw) {
        td.style.width = cw * PX + 'px';
      }

      for (const p of kids(first(tc, 'subList') || tc, 'p')) {
        td.append(renderParagraph(p, ctx));
      }
    }
  }

  return t;
}

// 8. HWPX 문단 출력
function renderParagraph(p, ctx) {
  const box = document.createElement('div');
  box.className = 'hwpx-p';

  const pp = ctx.styles.paraPr.get(
    p.getAttribute('paraPrIDRef')
  );

  if (pp?.align) {
    box.style.textAlign = pp.align;
  }

  if (pp?.lineHeight) {
    box.style.lineHeight = pp.lineHeight;
  }

  for (const run of kids(p, 'run')) {
    const cp = ctx.styles.charPr.get(
      run.getAttribute('charPrIDRef')
    );

    for (const node of run.children) {
      if (node.localName === 't') {
        const span = document.createElement('span');
        applyChar(span, cp);
        renderText(node, span);
        box.append(span);
      } else if (node.localName === 'tbl') {
        box.append(renderTable(node, ctx));
      } else if (node.localName === 'pic') {
        box.append(renderPic(node, ctx));
      }
    }
  }

  if (!box.childNodes.length) {
    box.append(document.createElement('br'));
  }

  return box;
}

// 9. HWPX 섹션 출력
function renderSection(doc, ctx) {
  const sheet = document.createElement('section');
  sheet.className = 'sheet';

  const pg = all(doc, 'pagePr')[0];
  const mg = first(pg, 'margin');

  if (pg) {
    sheet.style.width =
      num(pg.getAttribute('width')) * PX + 'px';
  }

  if (mg) {
    sheet.style.padding =
      ['top', 'right', 'bottom', 'left']
        .map(k => num(mg.getAttribute(k)) * PX + 'px')
        .join(' ');
  }

  for (const p of kids(doc.documentElement, 'p')) {
    sheet.append(renderParagraph(p, ctx));
  }

  return sheet;
}

// 10. HWPX 파일 열기
async function openHwpx(zip, viewer) {
  const styles = parseHeader(
    await readXml(zip, 'Contents/header.xml')
  );

  const manifest = await loadManifest(zip);
  const urls = await loadImages(zip, manifest);

  const ctx = { styles, urls };

  for (const path of manifest.spine) {
    const doc = await readXml(zip, path);
    viewer.append(renderSection(doc, ctx));
  }

  return `HWPX - 섹션 ${manifest.spine.length}개`;
}

// ========================================
// DOCX 처리 (추가 기능)
// ========================================

// DOCX의 XML 속성을 네임스페이스와 관계없이 읽기
function attr(el, name) {
  if (!el) return null;

  for (const a of Array.from(el.attributes)) {
    if (a.localName === name) return a.value;
  }

  return null;
}

// DOCX 단위 변환
// twip = 1/1440 inch
const twipPx = (v) => num(v) * 96 / 1440;

// EMU = 1/914400 inch
const emuPx = (v) => num(v) * 96 / 914400;

// DOCX 이미지 관계 정보 읽기
async function loadDocxImages(zip) {
  const urls = new Map();

  const relFile = findEntry(
    zip,
    'word/_rels/document.xml.rels'
  );

  if (!relFile) return urls;

  const relDoc = parseXml(
    await relFile.async('string'),
    'document.xml.rels'
  );

  for (const rel of all(relDoc, 'Relationship')) {
    const id = rel.getAttribute('Id');
    const target = rel.getAttribute('Target');
    const type = rel.getAttribute('Type') || '';

    if (!type.endsWith('/image') || !target) continue;
    if (/^[a-z]+:/i.test(target)) continue;

    const parts = ['word'];

    for (const part of target.replace(/\\/g, '/').split('/')) {
      if (part === '..') parts.pop();
      else if (part !== '.' && part) parts.push(part);
    }

    const path = parts.join('/');
    const f = findEntry(zip, path);

    if (f) {
      urls.set(id, await imageUrl(f));
    }
  }

  return urls;
}

// DOCX 스타일 파일 읽기
async function loadDocxStyles(zip) {
  const result = new Map();
  const f = findEntry(zip, 'word/styles.xml');

  if (!f) return result;

  const doc = parseXml(
    await f.async('string'),
    'styles.xml'
  );

  for (const style of all(doc, 'style')) {
    const id = attr(style, 'styleId');
    const rPr = first(style, 'rPr');
    const pPr = first(style, 'pPr');

    result.set(id, { rPr, pPr });
  }

  return result;
}

// DOCX 글자 스타일 적용
function applyDocxStyle(span, rPr) {
  if (!rPr) return;

  const bold = first(rPr, 'b');

  if (bold && !['0', 'false', 'off'].includes(
    attr(bold, 'val')
  )) {
    span.style.fontWeight = 'bold';
  }

  const italic = first(rPr, 'i');

  if (italic && !['0', 'false', 'off'].includes(
    attr(italic, 'val')
  )) {
    span.style.fontStyle = 'italic';
  }

  const underline = first(rPr, 'u');

  if (underline && attr(underline, 'val') !== 'none') {
    span.style.textDecoration = 'underline';
  }

  const strike = first(rPr, 'strike');

  if (strike && attr(strike, 'val') !== '0') {
    span.style.textDecoration =
      (span.style.textDecoration + ' line-through').trim();
  }

  const color = attr(first(rPr, 'color'), 'val');

  if (color && /^[0-9a-f]{6}$/i.test(color)) {
    span.style.color = '#' + color;
  }

  const size = attr(first(rPr, 'sz'), 'val');

  if (size) {
    span.style.fontSize = num(size) / 2 + 'pt';
  }

  const highlight = attr(
    first(rPr, 'highlight'),
    'val'
  );

  if (highlight && highlight !== 'none') {
    span.style.backgroundColor = highlight;
  }
}

// DOCX 문단 스타일 적용
function applyDocxParagraphStyle(box, pPr) {
  if (!pPr) return;

  const align = attr(first(pPr, 'jc'), 'val');

  const alignMap = {
    center: 'center',
    right: 'right',
    left: 'left',
    both: 'justify'
  };

  if (alignMap[align]) {
    box.style.textAlign = alignMap[align];
  }

  const spacing = first(pPr, 'spacing');

  if (spacing) {
    const before = attr(spacing, 'before');
    const after = attr(spacing, 'after');

    if (before) {
      box.style.marginTop = twipPx(before) + 'px';
    }

    if (after) {
      box.style.marginBottom = twipPx(after) + 'px';
    }
  }

  const ind = first(pPr, 'ind');

  if (ind) {
    const left = attr(ind, 'left');
    const firstLine = attr(ind, 'firstLine');

    if (left) {
      box.style.marginLeft = twipPx(left) + 'px';
    }

    if (firstLine) {
      box.style.textIndent = twipPx(firstLine) + 'px';
    }
  }
}

// DOCX 그림 출력
function renderDocxImage(drawing, ctx) {
  const blip = all(drawing, 'blip')[0];
  const id = attr(blip, 'embed');

  const url = ctx.urls.get(id);

  if (!url) {
    return missing('[이미지를 찾을 수 없습니다]');
  }

  const img = new Image();

  img.src = url;
  img.alt = 'DOCX 이미지';
  img.style.maxWidth = '100%';
  img.style.height = 'auto';

  const extent =
    all(drawing, 'extent')[0] ||
    all(drawing, 'ext')[0];

  if (extent) {
    const w = attr(extent, 'cx');

    if (w) {
      img.style.width = emuPx(w) + 'px';
    }
  }

  img.onerror = () =>
    img.replaceWith(missing('[이미지 표시 오류]'));

  return img;
}

// DOCX 글자 구간 출력
function renderDocxRun(run, ctx) {
  const span = document.createElement('span');

  const styleId = attr(
    first(first(run, 'rPr'), 'rStyle'),
    'val'
  );

  const namedStyle = ctx.styles.get(styleId);

  if (namedStyle) {
    applyDocxStyle(span, namedStyle.rPr);
  }

  applyDocxStyle(span, first(run, 'rPr'));

  for (const node of run.children) {
    switch (node.localName) {
      case 't':
        span.append(node.textContent);
        break;

      case 'tab':
        span.append('\u2003');
        break;

      case 'br':
      case 'cr':
        span.append(document.createElement('br'));
        break;

      case 'drawing':
      case 'pict':
        span.append(renderDocxImage(node, ctx));
        break;

      case 'sym':
        span.append('□');
        break;
    }
  }

  return span;
}

// DOCX 문단 출력
function renderDocxParagraph(p, ctx) {
  const box = document.createElement('div');
  box.className = 'hwpx-p';

  const pPr = first(p, 'pPr');
  const styleId = attr(first(pPr, 'pStyle'), 'val');

  const namedStyle = ctx.styles.get(styleId);

  if (namedStyle) {
    applyDocxParagraphStyle(box, namedStyle.pPr);
  }

  applyDocxParagraphStyle(box, pPr);

  for (const node of p.children) {
    if (node.localName === 'r') {
      box.append(renderDocxRun(node, ctx));
    } else if (
      node.localName === 'hyperlink' ||
      node.localName === 'sdt' ||
      node.localName === 'ins'
    ) {
      for (const run of all(node, 'r')) {
        box.append(renderDocxRun(run, ctx));
      }
    }
  }

  if (!box.childNodes.length) {
    box.append(document.createElement('br'));
  }

  return box;
}

// DOCX 표 출력
function renderDocxTable(tbl, ctx) {
  const table = document.createElement('table');
  table.className = 'hwpx-tbl';

  for (const tr of kids(tbl, 'tr')) {
    const row = table.insertRow();

    for (const tc of kids(tr, 'tc')) {
      const cell = row.insertCell();

      const tcPr = first(tc, 'tcPr');
      const gridSpan = first(tcPr, 'gridSpan');

      if (gridSpan) {
        cell.colSpan = num(attr(gridSpan, 'val')) || 1;
      }

      const width = first(tcPr, 'tcW');

      if (width && attr(width, 'type') === 'dxa') {
        cell.style.width =
          twipPx(attr(width, 'w')) + 'px';
      }

      for (const node of tc.children) {
        if (node.localName === 'p') {
          cell.append(renderDocxParagraph(node, ctx));
        } else if (node.localName === 'tbl') {
          cell.append(renderDocxTable(node, ctx));
        }
      }
    }
  }

  return table;
}

// DOCX 페이지 생성
function createDocxSheet(doc) {
  const sheet = document.createElement('section');
  sheet.className = 'sheet';

  const sectPr = all(doc, 'sectPr')[0];
  const pgSz = first(sectPr, 'pgSz');
  const pgMar = first(sectPr, 'pgMar');

  if (pgSz) {
    sheet.style.width =
      twipPx(attr(pgSz, 'w')) + 'px';
  } else {
    sheet.style.width = '794px';
  }

  if (pgMar) {
    const sides = ['top', 'right', 'bottom', 'left'];

    sheet.style.padding = sides.map(side =>
      twipPx(attr(pgMar, side)) + 'px'
    ).join(' ');
  } else {
    sheet.style.padding = '75px';
  }

  return sheet;
}

// DOCX 파일 열기
async function openDocx(zip, viewer) {
  const doc = await readXml(
    zip,
    'word/document.xml'
  );

  const styles = await loadDocxStyles(zip);
  const urls = await loadDocxImages(zip);

  const ctx = { styles, urls };

  const sheet = createDocxSheet(doc);
  const body = all(doc, 'body')[0];

  if (!body) {
    throw new Error('DOCX 본문을 찾을 수 없습니다');
  }

  let paragraphCount = 0;
  let tableCount = 0;

  for (const node of body.children) {
    if (node.localName === 'p') {
      sheet.append(renderDocxParagraph(node, ctx));
      paragraphCount++;
    } else if (node.localName === 'tbl') {
      sheet.append(renderDocxTable(node, ctx));
      tableCount++;
    }
  }

  viewer.append(sheet);

  return `DOCX - 문단 ${paragraphCount}개, 표 ${tableCount}개, 이미지 ${urls.size}개`;
}

// ========================================
// 파일 자동 구분 및 UI
// ========================================

async function openDocument(file) {
  const status = $('status');
  const viewer = $('viewer');

  status.className = '';
  status.textContent = `${file.name} 여는 중...`;

  clearUrls();
  viewer.replaceChildren();

  $('structure').hidden = true;
  $('tree').textContent = '';

  try {
    if (typeof JSZip === 'undefined') {
      throw new Error('JSZip 라이브러리를 불러오지 못했습니다');
    }

    const ext = file.name.split('.').pop().toLowerCase();

    if (ext === 'hwp') {
      throw new Error(
        '.hwp 형식은 지원하지 않습니다. .hwpx로 저장해 주세요.'
      );
    }

    if (ext !== 'hwpx' && ext !== 'docx') {
      throw new Error(
        '.hwpx 또는 .docx 파일만 지원합니다'
      );
    }

    const zip = await JSZip.loadAsync(file);

    // 파일 확장자뿐 아니라 내부 XML 구조도 확인
    const isHwpx = !!findEntry(
      zip, 'Contents/header.xml'
    );

    const isDocx = !!findEntry(
      zip, 'word/document.xml'
    );

    let result;

    if (ext === 'hwpx' && isHwpx) {
      result = await openHwpx(zip, viewer);
    } else if (ext === 'docx' && isDocx) {
      result = await openDocx(zip, viewer);
    } else {
      throw new Error(
        '파일 확장자와 실제 문서 형식이 일치하지 않습니다'
      );
    }

    // ZIP 내부 구조 표시
    $('tree').textContent = Object.values(zip.files)
      .map(f => {
        if (f.dir) return f.name;

        const size =
          f._data?.uncompressedSize ?? '?';

        return `${f.name}  (${size} B)`;
      })
      .join('\n');

    $('structure').hidden = false;

    status.textContent = `${file.name} - ${result}`;

  } catch (e) {
    status.className = 'error';
    status.textContent = `열 수 없습니다: ${e.message}`;

    viewer.replaceChildren();
    clearUrls();
  }
}

// 파일 선택 버튼
$('file').addEventListener('change', (e) => {
  const file = e.target.files[0];

  if (file) {
    openDocument(file);
  }
});

// 드래그 앤 드롭
const drop = $('drop');

['dragenter', 'dragover'].forEach(ev => {
  drop.addEventListener(ev, e => {
    e.preventDefault();
    drop.classList.add('over');
  });
});

['dragleave', 'drop'].forEach(ev => {
  drop.addEventListener(ev, e => {
    e.preventDefault();
    drop.classList.remove('over');
  });
});

drop.addEventListener('drop', e => {
  e.preventDefault();

  const file = e.dataTransfer.files[0];

  if (file) {
    openDocument(file);
  }
});
