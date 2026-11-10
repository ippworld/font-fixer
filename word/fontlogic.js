// Word (.docx) font logic: fonts live in <w:rFonts> attributes (ascii/hAnsi = Latin, eastAsia, cs), styles, theme, charts.
(function (g) {
  const A_FONT = /<a:(latin|ea|cs)\b[^>]*?\/>/g;
  const W_FONT = /<w:rFonts\b[^>]*?\/>/g;
  const PARTS = /^word\/(?:document|styles|stylesWithEffects|numbering|footnotes|endnotes|comments|header\d*|footer\d*)\.xml$|^word\/(?:theme|charts|diagrams)\/[^/]+\.xml$/;
  const SYMBOL = /^(wingdings|webdings|symbol|marlett|segoe (ui symbol|mdl2)|ms reference)/i;
  const W_MAP = [['ascii', 'latin', 'asciiTheme'], ['hAnsi', 'latin', 'hAnsiTheme'], ['eastAsia', 'ea', 'eastAsiaTheme'], ['cs', 'cs', 'cstheme']];
  const faceOf = t => (t.match(/typeface="([^"]*)"/) || [, ''])[1];
  const esc = s => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
  const unesc = s => s.replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&amp;/g, '&');
  const area = p => /\/theme\//.test(p) ? 'Theme' : /styles/.test(p) ? 'Styles' : /header|footer/.test(p) ? 'Headers/footers'
    : /footnotes|endnotes/.test(p) ? 'Notes' : /comments/.test(p) ? 'Comments' : /charts/.test(p) ? 'Charts' : /diagrams/.test(p) ? 'SmartArt' : 'Body';
  const explicit = tag => W_MAP.filter(([a, , th]) => !new RegExp(`\\sw:${th}=`).test(tag))
    .map(([a, slot]) => { const m = tag.match(new RegExp(`\\sw:${a}="([^"]*)"`)); return m ? [slot, unesc(m[1])] : null; }).filter(Boolean);

  // Language presets: which font slot carries the script, and the theme <a:font script="..."> code.
  const LANGS = {
    ja:  { label: 'Japanese',             slot: 'ea', script: 'Jpan', fonts: ['Yu Gothic', 'Meiryo', 'MS PGothic', 'Noto Sans JP', 'Noto Serif JP'] },
    zhs: { label: 'Chinese (Simplified)', slot: 'ea', script: 'Hans', fonts: ['Microsoft YaHei', 'SimSun', 'PingFang SC', 'Noto Sans SC', 'Noto Serif SC'] },
    zht: { label: 'Chinese (Traditional)', slot: 'ea', script: 'Hant', fonts: ['Microsoft JhengHei', 'PMingLiU', 'PingFang TC', 'Noto Sans TC', 'Noto Serif TC'] },
    ko:  { label: 'Korean',               slot: 'ea', script: 'Hang', fonts: ['Malgun Gothic', 'Batang', 'Apple SD Gothic Neo', 'Noto Sans KR', 'Noto Serif KR'] },
    th:  { label: 'Thai',                 slot: 'cs', script: 'Thai', fonts: ['Leelawadee UI', 'Tahoma', 'Cordia New', 'Noto Sans Thai', 'Noto Serif Thai'] },
    ar:  { label: 'Arabic',               slot: 'cs', script: 'Arab', fonts: ['Segoe UI', 'Arial', 'Traditional Arabic', 'Noto Sans Arabic'] },
    he:  { label: 'Hebrew',               slot: 'cs', script: 'Hebr', fonts: ['Arial', 'Segoe UI', 'David', 'Noto Sans Hebrew'] },
    hi:  { label: 'Hindi / Devanagari',   slot: 'cs', script: 'Deva', fonts: ['Nirmala UI', 'Mangal', 'Noto Sans Devanagari'] },
    lat: { label: 'Other (Latin / Cyrillic / Greek)', slot: null, script: null, fonts: ['Arial', 'Segoe UI', 'Calibri'] }
  };

  async function scan(zip) {
    const out = new Map();
    const add = (f, p) => { if (!f || f[0] === '+') return; const e = out.get(f) || { n: 0, areas: new Set() }; e.n++; e.areas.add(area(p)); out.set(f, e); };
    for (const p of Object.keys(zip.files).filter(p => PARTS.test(p))) {
      const xml = await zip.file(p).async('string');
      for (const m of xml.matchAll(A_FONT)) add(unesc(faceOf(m[0])), p);
      for (const m of xml.matchAll(W_FONT)) explicit(m[0]).forEach(([, f]) => add(f, p));
    }
    return out;
  }

  function rewrite(path, xml, cfg) {
    const { target, slots, sel, force, lang } = cfg;
    const theme = /\/theme\//.test(path); let n = 0;
    // DrawingML fonts (theme, charts, SmartArt)
    xml = xml.replace(A_FONT, (m, slot) => {
      if (!slots[slot]) return m;
      const f = unesc(faceOf(m));
      if (f[0] === '+' || !((f === '' && !slots.extra[slot]) || sel.has(f))) return m;
      n++; return `<a:${slot} typeface="${esc(target)}"/>`;
    });
    if (theme && lang.script) {
      const re = new RegExp(`<a:font script="${lang.script}" typeface="[^"]*"\\s*/>`);
      const tag = `<a:font script="${lang.script}" typeface="${esc(target)}"/>`;
      xml = xml.replace(/<a:(major|minor)Font>[\s\S]*?<\/a:\1Font>/g, blk => { n++; return re.test(blk) ? blk.replace(re, () => tag) : blk.replace(/<\/a:(major|minor)Font>/, m => tag + m); });
    }
    if (theme && cfg.everywhere) {      // other per-script theme entries (Arab, Hebr, Viet...) that still name a selected font
      xml = xml.replace(/<a:font script="\w+" typeface="([^"]*)"\s*\/>/g, (m, face) => {
        if (!sel.has(unesc(face))) return m;
        n++; return m.replace(/typeface="[^"]*"/, `typeface="${esc(target)}"`);
      });
    }
    // WordprocessingML fonts: only explicit attributes (theme-linked ones follow the updated theme)
    xml = xml.replace(W_FONT, tag => {
      let t = tag;
      for (const [attr, slot, th] of W_MAP) {
        if (!slots[slot] || new RegExp(`\\sw:${th}=`).test(t)) continue;
        const re = new RegExp(`(\\sw:${attr}=")([^"]*)(")`), m = t.match(re);
        if (m) { const f = unesc(m[2]); if ((f === '' && !slots.extra[slot]) || sel.has(f)) { t = t.replace(re, (_, a, b, c) => a + esc(target) + c); n++; } }
        else if (force) { t = t.replace(/\s*\/>$/, ` w:${attr}="${esc(target)}"/>`); n++; }
      }
      return t;
    });
    // Force: give every run-properties block without a font an explicit one.
    if (force && !theme) {
      const attrs = W_MAP.filter(([, slot]) => slots[slot]).map(([a]) => `w:${a}="${esc(target)}"`).join(' ');
      xml = xml.replace(/<w:rPr>([\s\S]*?)<\/w:rPr>/g, (m, inner) => {
        if (/<w:rFonts\b/.test(inner) || /<w:rPrChange\b/.test(inner) || !attrs) return m;
        n++;
        const s = inner.match(/^<w:rStyle\b[^>]*\/>/);
        return '<w:rPr>' + (s ? s[0] : '') + `<w:rFonts ${attrs}/>` + inner.slice(s ? s[0].length : 0) + '</w:rPr>';
      });
    }
    return { xml, n };
  }

  async function apply(zip, cfg) {
    const base = { latin: cfg.setLatin, ea: cfg.lang.slot === 'ea', cs: cfg.lang.slot === 'cs' };
    cfg.slots = Object.assign({}, base, { extra: {} });
    if (cfg.everywhere) for (const s of ['ea', 'cs']) if (!base[s]) { cfg.slots[s] = true; cfg.slots.extra[s] = true; }
    let files = 0, changes = 0;
    for (const p of Object.keys(zip.files).filter(p => PARTS.test(p))) {
      const src = await zip.file(p).async('string'), r = rewrite(p, src, cfg);
      if (r.xml !== src) { zip.file(p, r.xml); files++; changes += r.n; }
    }
    return { files, changes };
  }

  const api = { scan, apply, rewrite, LANGS, SYMBOL };
  if (typeof module !== 'undefined') module.exports = api; else g.FontFix = api;
})(typeof window !== 'undefined' ? window : globalThis);
