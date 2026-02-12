import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const manualInputCandidates = [
  path.join(repoRoot, 'docs', 'operations', 'operation_manual_vehicle_lease_renewal.md'),
  path.join(repoRoot, 'docs', 'operation_manual_vehicle_lease_renewal.md'),
];
const outputPath = path.join(repoRoot, 'dist', 'operation_manual_vehicle_lease_renewal.html');

function resolveExistingPath(candidates, label) {
  const found = candidates.find((candidate) => fs.existsSync(candidate));
  if (!found) {
    throw new Error(`${label}が見つかりません: ${candidates.join(', ')}`);
  }
  return found;
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function normalizeLinkHref(href) {
  return String(href).replace(/\.md(?=($|[?#]))/i, '.html');
}

function renderInline(escaped) {
  // escaped text in -> safe html out
  const codeTokens = [];
  const withCodeTokens = escaped.replace(/`([^`]+)`/g, (_, code) => {
    const id = codeTokens.length;
    codeTokens.push(`<code>${code}</code>`);
    return `@@CODE_TOKEN_${id}@@`;
  });

  let rendered = withCodeTokens
    .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, (_, label, href) => {
      const normalizedHref = normalizeLinkHref(href.trim());
      const externalAttrs = normalizedHref.startsWith('#')
        ? ''
        : ' target="_blank" rel="noopener noreferrer"';
      return `<a href="${normalizedHref}"${externalAttrs}>${label}</a>`;
    })
    .replace(/&lt;br\s*\/?&gt;/gi, '<br>');

  for (let i = 0; i < codeTokens.length; i += 1) {
    rendered = rendered.replace(`@@CODE_TOKEN_${i}@@`, codeTokens[i]);
  }

  return rendered;
}

function createHeadingId(text, headingIdMap) {
  const plainText = String(text)
    .replace(/`([^`]+)`/g, '$1')
    .replace(/\*\*([^*]+)\*\*/g, '$1')
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '$1');

  const base = plainText
    .toLowerCase()
    .replace(/[!"#$%&'()*+,./:;<=>?@[\\\]^_`{|}~]/g, '')
    .replace(/[、。！？」「『』（）［］【】：]/g, '')
    .trim()
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-');

  const normalized = base || 'section';
  const count = headingIdMap.get(normalized) || 0;
  headingIdMap.set(normalized, count + 1);
  return count === 0 ? normalized : `${normalized}-${count + 1}`;
}

function splitTableRow(line) {
  const trimmed = line.trim();
  const inner = trimmed.startsWith('|') ? trimmed.slice(1) : trimmed;
  const body = inner.endsWith('|') ? inner.slice(0, -1) : inner;
  return body.split('|').map((cell) => cell.trim());
}

function isTableDividerRow(cells) {
  if (!cells || cells.length === 0) return false;
  return cells.every((cell) => /^:?-{3,}:?$/.test(cell.replace(/\s+/g, '')));
}

function markdownToHtml(md) {
  const lines = String(md).replace(/\r\n/g, '\n').split('\n');

  let html = '';
  let inFence = false;
  let fenceType = 'code';
  let fenceLang = '';
  let inUl = false;
  let inOl = false;
  let inP = false;
  const tableRows = [];
  const headingIdMap = new Map();

  const closeParagraph = () => {
    if (inP) {
      html += '</p>\n';
      inP = false;
    }
  };
  const closeLists = () => {
    if (inUl) {
      html += '</ul>\n';
      inUl = false;
    }
    if (inOl) {
      html += '</ol>\n';
      inOl = false;
    }
  };
  const flushTable = () => {
    if (tableRows.length === 0) return;

    const header = tableRows[0];
    const hasDivider = tableRows.length > 1 && isTableDividerRow(tableRows[1]);
    const body = hasDivider ? tableRows.slice(2) : tableRows.slice(1);

    html += '<table>\n<thead><tr>';
    for (const cell of header) {
      html += `<th>${renderInline(escapeHtml(cell))}</th>`;
    }
    html += '</tr></thead>\n<tbody>\n';

    for (const row of body) {
      html += '<tr>';
      for (const cell of row) {
        html += `<td>${renderInline(escapeHtml(cell))}</td>`;
      }
      html += '</tr>\n';
    }
    html += '</tbody>\n</table>\n';
    tableRows.length = 0;
  };

  for (const rawLine of lines) {
    const line = rawLine ?? '';

    if (line.trim().startsWith('```mermaid')) {
      closeParagraph();
      closeLists();
      flushTable();
      html += '<pre class="mermaid">';
      inFence = true;
      fenceType = 'mermaid';
      fenceLang = '';
      continue;
    }

    if (!inFence && line.trim().startsWith('```')) {
      const mFence = /^```(\S+)?$/.exec(line.trim());
      closeParagraph();
      closeLists();
      flushTable();
      inFence = true;
      fenceType = 'code';
      fenceLang = mFence?.[1] || '';
      const langClass = fenceLang ? ` class="language-${escapeHtml(fenceLang)}"` : '';
      html += `<pre><code${langClass}>`;
      continue;
    }

    if (inFence) {
      if (line.trim().startsWith('```')) {
        html += fenceType === 'mermaid' ? '</pre>\n' : '</code></pre>\n';
        inFence = false;
        fenceType = 'code';
        fenceLang = '';
      } else {
        html += fenceType === 'mermaid' ? `${line}\n` : `${escapeHtml(line)}\n`;
      }
      continue;
    }

    const trimmed = line.trim();
    if (!trimmed) {
      closeParagraph();
      closeLists();
      flushTable();
      continue;
    }

    const isTableLine = trimmed.startsWith('|') && trimmed.includes('|');
    if (isTableLine) {
      closeParagraph();
      closeLists();
      tableRows.push(splitTableRow(trimmed));
      continue;
    }
    flushTable();

    if (/^(-{3,}|\*{3,})$/.test(trimmed)) {
      closeParagraph();
      closeLists();
      html += '<hr>\n';
      continue;
    }

    const mHeading = /^(#{1,6})\s+(.+)$/.exec(trimmed);
    if (mHeading) {
      closeParagraph();
      closeLists();
      const level = mHeading[1].length;
      const text = mHeading[2];
      const id = createHeadingId(text, headingIdMap);
      html += `<h${level} id="${id}">${renderInline(escapeHtml(text))}</h${level}>\n`;
      continue;
    }

    const mUl = /^[-*]\s+(.+)$/.exec(trimmed);
    if (mUl) {
      closeParagraph();
      if (!inUl) {
        closeLists();
        html += '<ul>\n';
        inUl = true;
      }
      html += `<li>${renderInline(escapeHtml(mUl[1]))}</li>\n`;
      continue;
    }

    const mOl = /^(\d+)[.)]\s+(.+)$/.exec(trimmed);
    if (mOl) {
      closeParagraph();
      if (!inOl) {
        closeLists();
        html += '<ol>\n';
        inOl = true;
      }
      html += `<li>${renderInline(escapeHtml(mOl[2]))}</li>\n`;
      continue;
    }

    // paragraph (supports "two spaces at EOL" -> <br>)
    closeLists();
    if (!inP) {
      html += '<p>';
      inP = true;
    } else {
      html += '\n';
    }
    const hasHardBreak = /\s\s$/.test(line);
    const escaped = renderInline(escapeHtml(line.replace(/\s+$/, '')));
    html += escaped + (hasHardBreak ? '<br>' : '');
  }

  closeParagraph();
  closeLists();
  flushTable();
  if (inFence) html += fenceType === 'mermaid' ? '</pre>\n' : '</code></pre>\n';

  return html;
}

function buildHtmlDocument(bodyHtml) {
  return `<!doctype html>
<html>
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
      :root { color-scheme: light; }
      body { font-family: system-ui, -apple-system, "Segoe UI", Roboto, "Noto Sans JP", sans-serif; line-height: 1.6; padding: 16px; }
      h1 { font-size: 20px; margin: 0 0 12px; }
      h2 { font-size: 16px; margin: 18px 0 8px; border-top: 1px solid #eee; padding-top: 12px; }
      h3 { font-size: 14px; margin: 14px 0 6px; }
      p { margin: 8px 0; }
      ul, ol { margin: 8px 0 8px 20px; padding: 0; }
      li { margin: 4px 0; }
      table { width: 100%; border-collapse: collapse; margin: 10px 0; }
      th, td { border: 1px solid #ddd; padding: 6px 8px; vertical-align: top; text-align: left; }
      th { background: #f8f8f8; }
      code { background: #f6f8fa; padding: 1px 4px; border-radius: 4px; }
      pre { background: #f6f8fa; padding: 10px; border-radius: 8px; overflow: auto; }
      pre code { background: none; padding: 0; }
      a { color: #0b57d0; text-decoration: underline; }
      .footer { margin-top: 18px; font-size: 12px; color: #666; }
      .mermaid { background: #f6f8fa; padding: 10px; border-radius: 8px; text-align: center; }
    </style>
    <script src="https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"></script>
    <script>
      mermaid.initialize({ startOnLoad: true, theme: 'default' });
    </script>
  </head>
  <body>
${bodyHtml}
    <div class="footer">このページはMarkdownから自動生成されています。</div>
  </body>
</html>
`;
}

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
const manualInputPath = resolveExistingPath(manualInputCandidates, '運用マニュアルMarkdown');
const manualMd = fs.readFileSync(manualInputPath, 'utf8');
const bodyHtml = markdownToHtml(manualMd);
const doc = buildHtmlDocument(bodyHtml);
fs.writeFileSync(outputPath, doc, 'utf8');
console.log(`generated: ${path.relative(repoRoot, outputPath)}`);
