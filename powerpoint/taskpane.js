const $ = id => document.getElementById(id);
const log = m => { $('log').textContent += m + '\n'; $('log').scrollTop = 1e9; };
const busy = b => {
  document.body.classList.toggle('busy', b);
  ['scan', 'apply'].forEach(i => $(i).disabled = b || (i === 'apply' && !$('backup').checked));
};
const SAMPLES = { ja: '日本語のフォント見本', zhs: '简体中文字体示例', zht: '繁體中文字體範例', ko: '한국어 글꼴 미리보기', th: 'ตัวอย่างฟอนต์ภาษาไทย', ar: 'معاينة الخط العربي', he: 'תצוגה מקדימה של גופן', hi: 'हिन्दी फ़ॉन्ट पूर्वावलोकन', lat: 'The quick brown fox' };
function specimen() {
  const f = $('target').value.trim(), l = $('lang').value;
  $('specText').textContent = l ? SAMPLES[l] + ' Aa 0123' : 'Aa 0123';
  $('specText').dir = (l === 'ar' || l === 'he') ? 'rtl' : 'ltr';
  $('specText').style.fontFamily = f ? JSON.stringify(f) + ', sans-serif' : '';
  $('specNote').textContent = f ? `Preview in ${f} (shows only if the font is installed)` : 'Type a font name to preview it';
}
let scanned = false;

// --- read the open presentation as a .pptx (works on Windows, Mac and web) ---
function getFile() {
  return new Promise((res, rej) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, r => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return rej(r.error);
      const f = r.value, parts = []; let i = 0;
      (function next() {
        f.getSliceAsync(i, s => {
          if (s.status !== Office.AsyncResultStatus.Succeeded) { f.closeAsync(); return rej(s.error); }
          parts[i] = s.value.data;
          if (++i < f.sliceCount) return next();
          f.closeAsync();
          const all = new Uint8Array(parts.reduce((a, p) => a + p.length, 0)); let o = 0;
          parts.forEach(p => { all.set(p, o); o += p.length; });
          res(all);
        });
      })();
    });
  });
}
const load = async () => JSZip.loadAsync(await getFile());

// --- write back: insert the fixed deck (keeping its own masters/theme), then delete the originals ---
async function replaceSlides(b64) {
  await PowerPoint.run(async ctx => {
    const slides = ctx.presentation.slides; slides.load('items/id'); await ctx.sync();
    const old = slides.items.map(s => s.id);
    ctx.presentation.insertSlidesFromBase64(b64, { formatting: 'KeepSourceFormatting', targetSlideId: old[old.length - 1] });
    await ctx.sync();
    old.forEach(id => ctx.presentation.slides.getItem(id).delete());
    await ctx.sync();
  });
}

async function doScan() {
  const zip = await load(), found = await FontFix.scan(zip), box = $('fonts');
  box.textContent = '';
  const max = Math.max(1, ...[...found.values()].map(v => v.n));
  [...found].sort((a, b) => b[1].n - a[1].n).forEach(([name, v]) => {
    const row = document.createElement('label'); row.className = 'row';
    const cb = document.createElement('input'); cb.type = 'checkbox'; cb.value = name; cb.checked = !FontFix.SYMBOL.test(name);
    const nm = document.createElement('span'); nm.className = 'nm'; nm.textContent = name; nm.style.fontFamily = JSON.stringify(name) + ', sans-serif';
    const sm = document.createElement('small'); sm.textContent = `${v.n} use${v.n > 1 ? 's' : ''} in ${[...v.areas].join(', ')}`;
    const bar = document.createElement('i'); bar.className = 'bar'; const fill = document.createElement('b'); fill.style.width = Math.max(4, 100 * v.n / max) + '%'; bar.append(fill);
    row.append(cb, nm, sm, bar); box.append(row);
  });
  $('summary').textContent = found.size ? `${found.size} font${found.size > 1 ? 's' : ''} found. Untick any you want to keep.` : 'No explicit fonts found.';
  $('s1').classList.toggle('done', found.size > 0);
  scanned = found.size > 0;
  log(found.size ? `Found ${found.size} font(s). Untick any you want to keep (symbol fonts are unticked by default).` : 'No explicit fonts found (everything inherits theme fonts).');
}

Office.onReady(() => {
  const sel = $('lang');
  const ph = new Option('Select language…', '', true, true); ph.disabled = true; sel.add(ph);
  Object.entries(FontFix.LANGS)                     // A–Z by name, "Other" last
    .sort(([ak, a], [bk, b]) => (ak === 'lat') - (bk === 'lat') || a.label.localeCompare(b.label))
    .forEach(([k, v]) => sel.add(new Option(v.label, k)));
  const suggest = () => {
    $('suggest').textContent = '';
    const L = FontFix.LANGS[sel.value];
    $('target').placeholder = L ? 'e.g. ' + L.fonts[0] : 'Select a language first';
    if (!L) { $('latin').disabled = false; return; }
    L.fonts.forEach(f => $('suggest').append(new Option(f)));
    $('latin').checked = true; $('latin').disabled = !L.slot;
  };
  sel.onchange = () => { suggest(); specimen(); }; suggest(); specimen(); $('target').oninput = specimen;
  $('backup').onchange = () => busy(false);

  $('scan').onclick = async () => { busy(true); try { await doScan(); } catch (e) { log('Error: ' + (e.message || e)); } busy(false); };
  $('apply').onclick = async () => {
    busy(true);
    try {
      const target = $('target').value.trim();
      if (!$('lang').value) throw new Error('Choose the language of the translated file first.');
      if (!target) throw new Error('Enter a target font.');
      if (!scanned) throw new Error('Run "Scan presentation" first.');
      const sel = new Set([...document.querySelectorAll('#fonts input:checked')].map(i => i.value));
      const zip = await load();
      const r = await FontFix.apply(zip, { target, lang: FontFix.LANGS[$('lang').value], sel, setLatin: $('latin').checked || !FontFix.LANGS[$('lang').value].slot, force: $('force').checked, everywhere: $('everywhere').checked });
      if (!r.changes) { log('Nothing needed changing.'); }
      else {
        log(`Updated ${r.changes} font setting(s) in ${r.files} part(s). Re-inserting slides…`);
        await replaceSlides(await zip.generateAsync({ type: 'base64', compression: 'DEFLATE' }));
        log('Done. Re-scanning to verify…'); await doScan();
      }
    } catch (e) { log('Error: ' + (e.message || e)); }
    busy(false);
  };
});
