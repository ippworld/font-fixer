const $ = id => document.getElementById(id);
const log = m => { $('log').textContent += m + '\n'; $('log').scrollTop = 1e9; };
const mode = () => $('mode').value;
const refresh = b => { $('scan').disabled = b; $('apply').disabled = b || (mode() === 'replace' && !$('backup').checked); };
const busy = b => { document.body.classList.toggle('busy', b); refresh(b); };
let scanned = false;
const BUILD = { app: 'word', version: '1.6', date: '11 Nov 2026' };
const cmpV = (a, b) => { const x = a.split('.').map(Number), y = b.split('.').map(Number); for (let i = 0; i < Math.max(x.length, y.length); i++) { const d = (x[i] || 0) - (y[i] || 0); if (d) return d; } return 0; };
function reloadLatest() { const u = new URL(location.href); u.searchParams.set('r', Date.now()); location.replace(u.toString()); }   // keeps Office's own query string
async function checkVersion() {
  try {
    const r = await fetch('../version.json?t=' + Date.now(), { cache: 'no-store' });
    const latest = (await r.json())[BUILD.app], chk = $('chk');
    if (!latest) return;
    chk.hidden = false;
    if (cmpV(latest, BUILD.version) > 0) {
      chk.textContent = 'Update available'; chk.classList.add('old');
      const box = $('upd'); box.textContent = ''; box.hidden = false;
      const t = document.createElement('span'); t.textContent = `A newer version (v${latest}) is available. You have v${BUILD.version}.`;
      const b = document.createElement('button'); b.type = 'button'; b.textContent = 'Reload now'; b.onclick = reloadLatest;
      box.append(t, b);
    } else { chk.textContent = '✓ Latest'; chk.title = 'You have the newest version'; }
  } catch (e) { /* offline or blocked: show no status */ }
}

const SAMPLES = { ja: '日本語のフォント見本', zhs: '简体中文字体示例', zht: '繁體中文字體範例', ko: '한국어 글꼴 미리보기', th: 'ตัวอย่างฟอนต์ภาษาไทย', ar: 'معاينة الخط العربي', he: 'תצוגה מקדימה של גופן', hi: 'हिन्दी फ़ॉन्ट पूर्वावलोकन', lat: 'The quick brown fox' };
const MODE_TEXT = {
  new: 'Your original stays untouched. A fixed copy opens in a new window; check it, then save it under a new name.',
  replace: 'This edits the open document in place, and Undo (⌘Z / Ctrl+Z) may not reverse it. Save a copy first (File > Save As). If the result looks wrong, close without saving and reopen the original.'
};
function specimen() {
  const f = $('target').value.trim(), f2 = $('target2').value.trim() || f, l = $('lang').value;
  const t = l ? SAMPLES[l] + ' Aa 0123' : 'Aa 0123', dir = (l === 'ar' || l === 'he') ? 'rtl' : 'ltr';
  [['specText', f], ['specText2', f2]].forEach(([id, ff]) => { const e = $(id); e.textContent = t; e.dir = dir; e.style.fontFamily = ff ? JSON.stringify(ff) + ', sans-serif' : ''; });
  document.querySelectorAll('#chips button').forEach(b => b.classList.toggle('on', b.textContent === f));
  document.querySelectorAll('#chips2 button').forEach(b => b.classList.toggle('on', b.textContent === $('target2').value.trim()));
  $('specNote').textContent = f ? 'Top line: sans-serif text. Bottom line: serif text. (Shown only if the fonts are installed.)' : 'Pick or type a font to preview it';
}

// Read the open document as a .docx (Windows, Mac and web)
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

async function writeDoc(b64) {
  await Word.run(async ctx => {
    if (mode() === 'new') {
      const d = ctx.application.createDocument(b64); await ctx.sync(); d.open(); await ctx.sync();
    } else {
      ctx.document.insertFileFromBase64(b64, 'Replace', { importTheme: true, importStyles: true, importParagraphSpacing: true, importPageColor: true, importChangeTrackingMode: 'All', importCustomProperties: true, importCustomXmlParts: true, importDifferentOddEvenPages: true });
      await ctx.sync();
    }
  });
}

async function doScan() {
  const zip = await load(), found = await FontFix.scan(zip), box = $('fonts');
  box.textContent = '';
  const max = Math.max(1, ...[...found.values()].map(v => v.n));
  [...found].sort((a, b) => b[1].n - a[1].n).forEach(([name, v]) => {
    const row = document.createElement('div'); row.className = 'row'; row.dataset.font = name;
    const nm = document.createElement('span'); nm.className = 'nm'; nm.textContent = name; nm.style.fontFamily = JSON.stringify(name) + ', sans-serif';
    const sm = document.createElement('small'); sm.textContent = `${v.n} use${v.n > 1 ? 's' : ''} in ${[...v.areas].join(', ')}`;
    const bar = document.createElement('i'); bar.className = 'bar'; const fill = document.createElement('b'); fill.style.width = Math.max(4, 100 * v.n / max) + '%'; bar.append(fill);
    const role = document.createElement('select'); role.className = 'role'; role.setAttribute('aria-label', 'Style of ' + name);
    [['sans', 'Sans'], ['serif', 'Serif'], ['keep', 'Keep']].forEach(([val, txt]) => role.add(new Option(txt, val)));
    role.value = FontFix.guessRole(name);
    row.append(nm, sm, bar, role); box.append(row);
  });
  $('summary').textContent = found.size ? `${found.size} font${found.size > 1 ? 's' : ''} found. Each is set to Sans, Serif or Keep: please check the guesses.` : 'No explicit fonts found.';
  $('s1').classList.toggle('done', found.size > 0);
  scanned = found.size > 0;
  log(found.size ? 'Scan complete.' : 'No explicit fonts found (everything inherits theme fonts).');
}

Office.onReady(() => {
  checkVersion();
  const sel = $('lang');
  const ph = new Option('Select language…', '', true, true); ph.disabled = true; sel.add(ph);
  Object.entries(FontFix.LANGS)                     // A–Z by name, "Other" last
    .sort(([ak, a], [bk, b]) => (ak === 'lat') - (bk === 'lat') || a.label.localeCompare(b.label))
    .forEach(([k, v]) => sel.add(new Option(v.label, k)));
  const suggest = () => {
    $('suggest').textContent = '';
    const L = FontFix.LANGS[sel.value];
    $('target').placeholder = L ? 'e.g. ' + (L.fonts.find(f => FontFix.guessRole(f) === 'sans') || L.fonts[0]) : 'Select a language first';
    $('target2').placeholder = L ? 'e.g. ' + (L.fonts.find(f => FontFix.guessRole(f) === 'serif') || 'same as above') : 'Leave empty to use the same font';
    const chips = $('chips'); chips.textContent = ''; $('chips2').textContent = '';
    $('hint').textContent = L ? 'Tap a suggested font under each box, or type the name of any font installed on your computer. Serif is optional.' : 'First choose the language, then pick a suggested font or type your own.';
    if (!L) { $('latin').disabled = false; return; }
    const sansL = L.fonts.filter(f => FontFix.guessRole(f) === 'sans'), serifL = L.fonts.filter(f => FontFix.guessRole(f) === 'serif');
    const mk = (box, list, field) => (list.length ? list : L.fonts).forEach(f => { const b = document.createElement('button'); b.type = 'button'; b.textContent = f; b.onclick = () => { $(field).value = f; specimen(); }; box.append(b); });
    mk(chips, sansL, 'target'); mk($('chips2'), serifL, 'target2');
    L.fonts.forEach(f => $('suggest').append(new Option(f)));
    $('latin').checked = true; $('latin').disabled = !L.slot;
  };
  const modeUi = () => { $('modeText').textContent = MODE_TEXT[mode()]; $('backupRow').hidden = mode() !== 'replace'; refresh(false); };
  sel.onchange = () => { suggest(); specimen(); }; suggest(); specimen(); $('target').oninput = specimen; $('target2').oninput = specimen;
  $('mode').onchange = modeUi; $('backup').onchange = () => refresh(false); modeUi();

  $('scan').onclick = async () => { busy(true); try { await doScan(); } catch (e) { log('Error: ' + (e.message || e)); } busy(false); };
  $('apply').onclick = async () => {
    busy(true);
    try {
      const target = $('target').value.trim();
      if (!$('lang').value) throw new Error('Choose the language of the translated file first.');
      if (!target) throw new Error('Enter the sans-serif replacement font first.');
      if (!scanned) throw new Error('Run "Scan presentation" first.'.replace('presentation', 'document'));
      const target2 = $('target2').value.trim() || target, map = {};
      document.querySelectorAll('#fonts .row').forEach(r => { const v = r.querySelector('select').value; if (v !== 'keep') map[r.dataset.font] = v === 'serif' ? target2 : target; });
      if (!Object.keys(map).length) throw new Error('Every font is set to Keep. Set at least one font to Sans or Serif.');
      const lang = FontFix.LANGS[$('lang').value], zip = await load();
      const r = await FontFix.apply(zip, { target, lang, map, pair: $('pair').checked, setLatin: $('latin').checked || !lang.slot, force: $('force').checked, everywhere: $('everywhere').checked });
      if (!r.changes) log('Nothing needed changing.');
      else {
        log(`Updated ${r.changes} font setting(s) in ${r.files} part(s).`);
        await writeDoc(await zip.generateAsync({ type: 'base64', compression: 'DEFLATE' }));
        log(mode() === 'new' ? 'Done. The fixed copy opened in a new window.' : 'Done. Re-scanning to verify…');
        if (mode() === 'replace') await doScan();
      }
    } catch (e) { log('Error: ' + (e.message || e)); }
    busy(false);
  };
});
