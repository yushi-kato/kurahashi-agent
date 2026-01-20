import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const inputPath = path.join(repoRoot, 'docs', 'operation_manual_vehicle_lease_renewal.md');
const outputPath = path.join(repoRoot, 'dist', 'operation_manual_vehicle_lease_renewal.html');

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function renderInline(escaped) {
  // escaped text in -> safe html out
  return escaped
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
}

function markdownToHtml(md) {
  const lines = String(md).replace(/\r\n/g, '\n').split('\n');

  let html = '';
  let inCode = false;
  let inUl = false;
  let inOl = false;
  let inP = false;

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

  for (const rawLine of lines) {
    const line = rawLine ?? '';

    if (line.trim().startsWith('```')) {
      if (!inCode) {
        closeParagraph();
        closeLists();
        html += '<pre><code>';
        inCode = true;
      } else {
        html += '</code></pre>\n';
        inCode = false;
      }
      continue;
    }

    if (inCode) {
      html += `${escapeHtml(line)}\n`;
      continue;
    }

    const trimmed = line.trim();
    if (!trimmed) {
      closeParagraph();
      closeLists();
      continue;
    }

    if (/^(-{3,}|\*{3,})$/.test(trimmed)) {
      closeParagraph();
      closeLists();
      html += '<hr>\n';
      continue;
    }

    const mH1 = /^#\s+(.+)$/.exec(trimmed);
    if (mH1) {
      closeParagraph();
      closeLists();
      html += `<h1>${renderInline(escapeHtml(mH1[1]))}</h1>\n`;
      continue;
    }
    const mH2 = /^##\s+(.+)$/.exec(trimmed);
    if (mH2) {
      closeParagraph();
      closeLists();
      html += `<h2>${renderInline(escapeHtml(mH2[1]))}</h2>\n`;
      continue;
    }
    const mH3 = /^###\s+(.+)$/.exec(trimmed);
    if (mH3) {
      closeParagraph();
      closeLists();
      html += `<h3>${renderInline(escapeHtml(mH3[1]))}</h3>\n`;
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
  if (inCode) html += '</code></pre>\n';

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
      code { background: #f6f8fa; padding: 1px 4px; border-radius: 4px; }
      pre { background: #f6f8fa; padding: 10px; border-radius: 8px; overflow: auto; }
      pre code { background: none; padding: 0; }
      .footer { margin-top: 18px; font-size: 12px; color: #666; }
    </style>
  </head>
  <body>
${bodyHtml}
    <div class="footer">このページはMarkdownから自動生成されています。</div>
  </body>
</html>
`;
}

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
const md = fs.readFileSync(inputPath, 'utf8');
const bodyHtml = markdownToHtml(md);
const doc = buildHtmlDocument(bodyHtml);
fs.writeFileSync(outputPath, doc, 'utf8');
console.log(`generated: ${path.relative(repoRoot, outputPath)}`);
