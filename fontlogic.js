// Pure OOXML font logic (no Office dependency, so it can be unit-tested in Node).
(function (g) {
  const FONT_TAG = /<a:(latin|ea|cs)\b[^>]*?\/>/g;           // text-run font slots (bullets <a:buFont> and <a:sym> are deliberately untouched)
  const PARTS = /^ppt\/(slides|slideLayouts|slideMasters|notesSlides|notesMasters|handoutMasters|theme|charts|diagrams)\/[^/]+\.xml$|^ppt\/presentation\.xml$/;
  const SYMBOL = /^(wingdings|webdings|symbol|marlett|segoe (ui symbol|mdl2)|ms reference)/i;
  const faceOf = t => (t.match(/typeface="([^"]*)"/) || [, ''])[1];
  const esc = s => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
  const unesc = s => s.replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&amp;/g, '&');
  const area = p => /\/theme\//.test(p) ? 'Theme'
    : /slideMasters|slideLayouts|handoutMasters|notesMasters|presentation\.xml/.test(p) ? 'Masters/layouts'
    : /notesSlides/.test(p) ? 'Notes' : /charts/.test(p) ? 'Charts' : /diagrams/.test(p) ? 'SmartArt' : 'Slides';

  // Language presets: which font slot carries the script, and the theme <a:font script="..."> code.
  const LANGS = {
    ja:  { label: 'Japanese',             slot: 'ea', script: 'Jpan', fonts: ['Yu Gothic', 'Meiryo', 'MS PGothic', 'Noto Sans JP'] },
    zhs: { label: 'Chinese (Simplified)', slot: 'ea', script: 'Hans', fonts: ['Microsoft YaHei', 'SimSun', 'PingFang SC', 'Noto Sans SC'] },
    zht: { label: 'Chinese (Traditional)', slot: 'ea', script: 'Hant', fonts: ['Microsoft JhengHei', 'PMingLiU', 'PingFang TC', 'Noto Sans TC'] },
    ko:  { label: 'Korean',               slot: 'ea', script: 'Hang', fonts: ['Malgun Gothic', 'Batang', 'Apple SD Gothic Neo', 'Noto Sans KR'] },
    th:  { label: 'Thai',                 slot: 'cs', script: 'Thai', fonts: ['Leelawadee UI', 'Tahoma', 'Cordia New', 'Noto Sans Thai'] },
    ar:  { label: 'Arabic',               slot: 'cs', script: 'Arab', fonts: ['Segoe UI', 'Arial', 'Traditional Arabic', 'Noto Sans Arabic'] },
    he:  { label: 'Hebrew',               slot: 'cs', script: 'Hebr', fonts: ['Arial', 'Segoe UI', 'David', 'Noto Sans Hebrew'] },
    hi:  { label: 'Hindi / Devanagari',   slot: 'cs', script: 'Deva', fonts: ['Nirmala UI', 'Mangal', 'Noto Sans Devanagari'] },
    lat: { label: 'Other (Latin / Cyrillic / Greek)', slot: null, script: null, fonts: ['Arial', 'Segoe UI', 'Calibri'] }
  };

  async function scan(zip) {
    const out = new Map();
    for (const p of Object.keys(zip.files).filter(p => PARTS.test(p))) {
      const xml = await zip.file(p).async('string');
      for (const m of xml.matchAll(FONT_TAG)) {
        const f = unesc(faceOf(m[0]));
        if (!f || f[0] === '+') continue;                      // blank or theme reference (+mn-lt etc.) – follows theme
        const e = out.get(f) || { n: 0, areas: new Set() };
        e.n++; e.areas.add(area(p)); out.set(f, e);
      }
    }
    return out;
  }

  function rewrite(path, xml, cfg) {
    const { target, slots, sel, force, lang } = cfg;
    let n = 0;
    // 1. Replace explicit typefaces in the enabled slots (selected fonts, plus blank ones).
    xml = xml.replace(FONT_TAG, (m, slot) => {
      if (!slots[slot]) return m;
      const f = unesc(faceOf(m));
      if (f[0] === '+' || !(f === '' || sel.has(f))) return m;
      n++; return `<a:${slot} typeface="${esc(target)}"/>`;
    });
    // 2. Theme: per-script font entries (e.g. <a:font script="Jpan" .../>) in major and minor font sets.
    if (/\/theme\//.test(path) && lang.script) {
      const re = new RegExp(`<a:font script="${lang.script}" typeface="[^"]*"\\s*/>`);
      const tag = `<a:font script="${lang.script}" typeface="${esc(target)}"/>`;
      xml = xml.replace(/<a:(major|minor)Font>[\s\S]*?<\/a:\1Font>/g, blk => {
        n++;
        return re.test(blk) ? blk.replace(re, () => tag) : blk.replace(/<\/a:(major|minor)Font>/, m => tag + m);
      });
    }
    // 3. Optional: force the font onto every run/default/end-paragraph property block that lacks one.
    if (force && !/\/theme\//.test(path)) {
      xml = xml.replace(/<a:(rPr|defRPr|endParaRPr)\b([^>]*?)(\/>|>([\s\S]*?)<\/a:\1>)/g, (m, tag, attrs, tail, inner) => {
        inner = inner || '';
        let add = '';
        for (const s of ['latin', 'ea', 'cs'])
          if (slots[s] && !new RegExp(`<a:${s}\\b`).test(inner)) add += `<a:${s} typeface="${esc(target)}"/>`;
        if (!add) return m;
        n++;
        const i = inner.search(/<a:(sym|hlinkClick|hlinkMouseOver|rtl|extLst)\b/);   // schema order: latin, ea, cs come before these
        inner = i < 0 ? inner + add : inner.slice(0, i) + add + inner.slice(i);
        return `<a:${tag}${attrs}>${inner}</a:${tag}>`;
      });
    }
    return { xml, n };
  }

  async function apply(zip, cfg) {
    cfg.slots = { latin: cfg.setLatin, ea: cfg.lang.slot === 'ea', cs: cfg.lang.slot === 'cs' };
    let files = 0, changes = 0;
    for (const p of Object.keys(zip.files).filter(p => PARTS.test(p))) {
      const src = await zip.file(p).async('string');
      const r = rewrite(p, src, cfg);
      if (r.xml !== src) { zip.file(p, r.xml); files++; changes += r.n; }
    }
    return { files, changes };
  }

  const api = { scan, apply, rewrite, LANGS, SYMBOL };
  if (typeof module !== 'undefined') module.exports = api; else g.FontFix = api;
})(typeof window !== 'undefined' ? window : globalThis);
