// Pure OOXML font logic (no Office dependency, so it can be unit-tested in Node).
(function (g) {
  const FONT_TAG = /<a:(latin|ea|cs|sym|buFont)\b[^>]*?\/>/g;   // text font slots plus symbol (<a:sym>) and bullet (<a:buFont>) fonts
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

  const own = (o, k) => Object.prototype.hasOwnProperty.call(o, k);
  const SERIF_NAMES = /simsun|ming|mincho|batang|myungjo|times|georgia|garamond|didot|bodoni|bookman|bookmania|palatino|cambria|baskerville|caslon|minion|merriweather|playfair|lora\b|cormorant|angsana|david|constantia|century|song/i;
  // Best guess of a font's style from its name: 'sans', 'serif', or 'keep' for symbol fonts. The user can change every guess in the pane.
  function guessRole(name) {
    if (SYMBOL.test(name)) return 'keep';
    if (/serif/i.test(name) && !/sans/i.test(name)) return 'serif';
    if (/sans/i.test(name)) return 'sans';
    return SERIF_NAMES.test(name) ? 'serif' : 'sans';
  }

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
    const { target, slots, map, force, lang } = cfg, theme = /\/theme\//.test(path); let n = 0;
    const pick = f => f === '' ? target : (own(map, f) ? map[f] : null);
    // T = the replacement chosen for this run's Latin font. Asian/complex-script text follows it so a serif run stays serif.
    const tagRep = (m, slot, T) => {
      if (!slots[slot]) return m;
      const f = unesc(faceOf(m)); let to;
      if (T && (slot === 'ea' || slot === 'cs')) to = T;
      else { if (f[0] === '+' || (f === '' && slots.extra[slot])) return m; to = pick(f); }
      if (!to || to === f) return m;
      n++; return `<a:${slot} typeface="${esc(to)}"/>`;
    };
    if (theme) xml = xml.replace(FONT_TAG, (m, slot) => tagRep(m, slot, null));
    else {
      xml = xml.replace(/<a:(rPr|defRPr|endParaRPr)\b((?:\s[^>]*[^/>])?)>([\s\S]*?)<\/a:\1>/g, (m, tag, attrs, inner) => {
        const lat = inner.match(/<a:latin\b[^>]*?\/>/), f = lat ? unesc(faceOf(lat[0])) : '';
        const T = (cfg.pair && f && f[0] !== '+' && own(map, f)) ? map[f] : null;
        let out = inner.replace(FONT_TAG, (t, slot) => tagRep(t, slot, T));
        if (T) {
          const needEa = slots.ea && !/<a:ea\b/.test(out), needCs = slots.cs && !/<a:cs\b/.test(out);
          const E = `<a:ea typeface="${esc(T)}"/>`, C = `<a:cs typeface="${esc(T)}"/>`;
          if (needEa) { out = out.replace(/<a:latin\b[^>]*?\/>/, x => x + E + (needCs ? C : '')); n += needCs ? 2 : 1; }
          else if (needCs) { const re = /<a:ea\b[^>]*?\/>/.test(out) ? /<a:ea\b[^>]*?\/>/ : /<a:latin\b[^>]*?\/>/; out = out.replace(re, x => x + C); n++; }
        }
        return out === inner ? m : `<a:${tag}${attrs}>${out}</a:${tag}>`;
      });
      xml = xml.replace(/<a:buFont\b[^>]*?\/>/g, m => tagRep(m, 'buFont', null));
    }
    // Theme: the chosen language's per-script entry (e.g. <a:font script="Hans" .../>) in the major and minor font sets.
    if (theme && lang.script) {
      const re = new RegExp(`<a:font script="${lang.script}" typeface="[^"]*"\\s*/>`);
      const tag = `<a:font script="${lang.script}" typeface="${esc(target)}"/>`;
      xml = xml.replace(/<a:(major|minor)Font>[\s\S]*?<\/a:\1Font>/g, blk => {
        const nb = re.test(blk) ? blk.replace(re, () => tag) : blk.replace(/<\/a:(major|minor)Font>/, m => tag + m);
        if (nb !== blk) n++; return nb;
      });
    }
    // "Everywhere": other per-script theme entries (Arab, Hebr, Viet...) that still name a mapped font.
    if (theme && cfg.everywhere) {
      xml = xml.replace(/<a:font script="\w+" typeface="([^"]*)"\s*\/>/g, (m, face) => {
        const f = unesc(face), to = own(map, f) ? map[f] : null;
        if (!to || to === f) return m;
        n++; return m.replace(/typeface="[^"]*"/, `typeface="${esc(to)}"`);
      });
    }
    // Optional: force the font onto every run/default/end-paragraph property block that lacks one.
    if (force && !theme) {
      xml = xml.replace(/<a:(rPr|defRPr|endParaRPr)\b([^>]*?)(\/>|>([\s\S]*?)<\/a:\1>)/g, (m, tag, attrs, tail, inner) => {
        inner = inner || '';
        let add = '';
        for (const s of ['latin', 'ea', 'cs'])
          if (slots[s] && !new RegExp(`<a:${s}\\b`).test(inner)) add += `<a:${s} typeface="${esc(target)}"/>`;
        if (!add) return m;
        n++;
        const i = inner.search(/<a:(sym|hlinkClick|hlinkMouseOver|rtl|extLst)\b/);
        inner = i < 0 ? inner + add : inner.slice(0, i) + add + inner.slice(i);
        return `<a:${tag}${attrs}>${inner}</a:${tag}>`;
      });
    }
    return { xml, n };
  }

  async function apply(zip, cfg) {
    const base = { latin: cfg.setLatin, ea: cfg.lang.slot === 'ea', cs: cfg.lang.slot === 'cs', sym: false, buFont: false };
    cfg.slots = Object.assign({}, base, { extra: {} });
    if (cfg.everywhere) for (const s of ['ea', 'cs', 'sym', 'buFont']) if (!base[s]) { cfg.slots[s] = true; cfg.slots.extra[s] = true; }
    if (!cfg.map) cfg.map = Object.fromEntries([...(cfg.sel || [])].map(f => [f, cfg.target]));   // one font for every selected font
    if (cfg.pair === undefined) cfg.pair = true;
    let files = 0, changes = 0;
    for (const p of Object.keys(zip.files).filter(p => PARTS.test(p))) {
      const src = await zip.file(p).async('string');
      const r = rewrite(p, src, cfg);
      if (r.xml !== src) { zip.file(p, r.xml); files++; changes += r.n; }
    }
    return { files, changes };
  }

  const api = { scan, apply, rewrite, LANGS, SYMBOL, guessRole };
  if (typeof module !== 'undefined') module.exports = api; else g.FontFix = api;
})(typeof window !== 'undefined' ? window : globalThis);
