// Builds ACE-Design-Review.pptx — the PowerPoint twin of deck/index.html.
// Usage: npm i pptxgenjs@3 && node pptx/build-pptx.js
// Featured pair: FEATURED below (default Big Sky + Heartland Steel).
const path = require('path');
const pptxgen = require('pptxgenjs');

const ROOT = path.resolve(__dirname, '..');
const A = p => path.join(ROOT, p);
const OUT = A('ACE-Design-Review.pptx');

const THEME = {
  name: 'ACE Gold',
  headFontFace: 'Cambria',
  bodyFontFace: 'Calibri',
  colors: {
    dk1: '1D1810', lt1: 'FAF6EE', dk2: '15110B', lt2: 'F5EFE3',
    accent1: 'D6A64A', accent2: 'A9792A', accent3: '6B604C', accent4: 'E2D6BF',
    accent5: 'B9AE98', accent6: '2A2116', hlink: 'A9792A', folHlink: '6B604C'
  }
};
const HEX = THEME.colors;

const CONCEPTS = {
  bigsky: {
    n: '01', name: 'Big Sky', feel: 'Warm, editorial, timeless.',
    accent: 'Gold · prairie heritage', colors: ['D6A64A', '2A1D10', 'F4ECDC'],
    img: 'deck/assets/hero_bigsky.jpg', matrix: 'deck/assets/matrix_bigsky.jpg',
    feeling: 'A family company that’s been doing this right for a long time.',
    speaks: 'Farm, ranch, and long-time accounts',
    risk: 'Can read soft or lifestyle to hard-nosed fleet buyers.',
    why: [
      ['The gold is ours. ', 'Its palette is the ACE mark — the logo looks native, not placed.'],
      ['Editorial serif = credibility. ', 'Reads established without reading corporate.'],
      ['“Delivered like family” ', 'says the family-company promise without a history lesson.']
    ],
    notes: 'Name the feeling in one sentence: a family company that has been doing this right for a long time.|Point out how the gold in the site is the gold in our mark — nothing fights the logo.|Ask the room: cover the logo for two seconds. Still ACE?'
  },
  piney: {
    n: '02', name: 'Piney Woods', feel: 'Rooted, calm, organic.',
    accent: 'Forest green · handshake trust', colors: ['C9B27C', '24301F', 'EEF0E6'],
    img: 'deck/assets/hero_piney.jpg', matrix: 'deck/assets/matrix_piney.jpg',
    feeling: 'The handshake you can trust.',
    speaks: 'Relationship buyers who value calm and trust',
    risk: 'Green pulls away from the gold mark; forest isn’t oilfield TX-LA-OK.',
    why: [],
    notes: 'The handshake you can trust. Calm, rooted.|Honest risk: the green pulls away from our gold, and the forest feels East Texas more than our corridors.'
  },
  openwater: {
    n: '03', name: 'Open Water', feel: 'Crisp, corporate, confident.',
    accent: 'Teal / sand · B2B polish', colors: ['D8A35A', '14303A', 'F2EDE2'],
    img: 'deck/assets/hero_openwater.jpg', matrix: 'deck/assets/matrix_openwater.jpg',
    feeling: 'A supplier that runs like a serious business.',
    speaks: 'Procurement and larger B2B accounts',
    risk: 'Offshore rig reads as drilling, not distribution. Coolest temperature.',
    why: [],
    notes: 'A supplier that runs like a serious business. Built for procurement.|Risk: the offshore rig says drilling, not distribution. Imagery would need to change.'
  },
  heartland: {
    n: '04', name: 'Heartland Steel', feel: 'Bold, cinematic, industrial.',
    accent: 'Orange on dark · ops drama', colors: ['E08A2C', '141210', 'EFE7DA'],
    img: 'deck/assets/hero_heartland.jpg', matrix: 'deck/assets/matrix_heartland.jpg',
    feeling: 'Heavy-duty product and people who answer the phone.',
    speaks: 'Fleet, oilfield, and operations buyers',
    risk: 'Loudest option. The orange accent needs tuning toward our gold.',
    why: [
      ['Built for the buyer who runs trucks. ', 'Speaks ops, not lifestyle.'],
      ['Dark stage lets the mark pop. ', 'The gold logo has never looked stronger.'],
      ['Most distinct in the market. ', 'Nobody in our space looks like this.']
    ],
    notes: 'Heavy-duty product and people who answer the phone. Built for the long haul.|This is the same promise as Big Sky at a different volume.|Fair risk: it is loud, and the orange would move toward our gold in the build.'
  },
  hill: {
    n: '05', name: 'Hill Country', feel: 'Warm, friendly, welcoming.',
    accent: 'Terracotta · Texas hospitality', colors: ['B8553A', '2C2118', 'F3EAD9'],
    img: 'deck/assets/hero_hill.jpg', matrix: 'deck/assets/matrix_hill.jpg',
    feeling: 'Pull up a chair — we’ll take care of you.',
    speaks: 'Smaller and first-time accounts',
    risk: 'Most consumer-feeling; terracotta competes with the gold mark.',
    why: [],
    notes: 'Pull up a chair — Texas hospitality.|Risk: it is the most consumer-feeling of the five, and terracotta competes with our gold.'
  }
};
const ORDER = ['bigsky', 'piney', 'openwater', 'heartland', 'hill'];
// Live sites live in designs/ next to the .pptx; links are relative so they work from a local copy of the repo.
const liveUrl = k => `designs/${k}.html`;
ORDER.forEach(k => { CONCEPTS[k].key = k; });
const FEATURED = ['bigsky', 'heartland'];
const ALTS = ORDER.filter(k => !FEATURED.includes(k));

const notes = s => s.split('|').join('\n\n');
const HERO_RATIO = 1100 / 644;

(async () => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5
  pres.title = 'ACE Design Review';
  pres.company = 'A.C.E. Distribution';
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  const W = 13.333, L = 0.97, CW = W - 2 * L;
  const LOGO_RATIO = 1200 / 255;

  const numDark = { x: W - L - 1, y: 6.95, w: 1, h: 0.3, align: 'right', fontSize: 10, color: '8C826F' };
  const numLight = { x: W - L - 1, y: 6.95, w: 1, h: 0.3, align: 'right', fontSize: 10, color: HEX.accent5 };
  const smallLogo = (file) => ({ image: { path: A(file), x: L, y: 0.42, w: 0.26 * LOGO_RATIO, h: 0.26 } });
  const kickerPh = (gold) => ({
    placeholder: {
      options: { name: 'kicker', type: 'body', align: 'left', x: L, y: 1.15, w: CW, h: 0.35, fontSize: 11, bold: true, charSpacing: 3, color: gold, margin: 0, valign: 'top' },
      text: ''
    }
  });
  const titlePh = (color, h = 1.5) => ({
    placeholder: {
      options: { name: 'title', type: 'title', align: 'left', x: L, y: 1.5, w: CW, h, fontFace: THEME.headFontFace, fontSize: 44, bold: true, color, margin: 0, valign: 'top' },
      text: ''
    }
  });

  pres.defineSlideMaster({
    title: 'ACE Title', background: { path: A('pptx/assets/bg-dark.png') }, objects: []
  });
  pres.defineSlideMaster({
    title: 'ACE Dark', background: { path: A('pptx/assets/bg-dark.png') },
    objects: [smallLogo('deck/assets/ace-logo-gold.png'), kickerPh(C.accent1), titlePh(C.background2)],
    slideNumber: numDark
  });
  pres.defineSlideMaster({
    title: 'ACE Light', background: { path: A('pptx/assets/bg-light.png') },
    objects: [smallLogo('pptx/assets/ace-logo-deepgold.png'), kickerPh(C.accent2), titlePh(C.text1)],
    slideNumber: numLight
  });
  pres.defineSlideMaster({
    title: 'ACE Light Bare', background: { path: A('pptx/assets/bg-light.png') },
    objects: [smallLogo('pptx/assets/ace-logo-deepgold.png')],
    slideNumber: numLight
  });
  pres.defineSlideMaster({
    title: 'ACE Statement', background: { path: A('pptx/assets/bg-dark.png') },
    objects: [
      smallLogo('deck/assets/ace-logo-gold.png'),
      { placeholder: { options: { name: 'kicker', type: 'body', align: 'left', x: L, y: 2.15, w: CW, h: 0.35, fontSize: 11, bold: true, charSpacing: 3, color: C.accent1, margin: 0, valign: 'top' }, text: '' } },
      { placeholder: { options: { name: 'title', type: 'title', align: 'left', x: L, y: 2.55, w: 9.6, h: 2.3, fontFace: THEME.headFontFace, fontSize: 58, bold: true, color: C.background2, margin: 0, valign: 'top' }, text: '' } },
      { placeholder: { options: { name: 'body', type: 'body', align: 'left', x: L, y: 5.0, w: 8.2, h: 1.0, fontSize: 18, color: 'BDB4A2', margin: 0, valign: 'top' }, text: '' } }
    ],
    slideNumber: numDark
  });
  pres.defineSlideMaster({
    title: 'ACE Showcase', background: { path: A('pptx/assets/bg-stage.png') }, objects: []
  });

  // ---------- helpers ----------
  const tb = (slide, text, opts) => slide.addText(text, { isTextBox: true, margin: 0, valign: 'top', ...opts });
  const kicker = (slide, text) => slide.addText(text.toUpperCase(), { placeholder: 'kicker' });

  function browser(slide, c, x, y, w, name, barH = 0.24) {
    const h = w / HERO_RATIO;
    slide.addShape(pres.shapes.RECTANGLE, {
      objectName: name + ' frame', x, y, w, h: h + barH, fill: { color: '1F1A14' }, line: { color: '1F1A14', width: 0 },
      shadow: { type: 'outer', color: '000000', opacity: 0.45, blur: 18, offset: 6, angle: 90 }
    });
    const dot = barH * 0.32;
    for (let i = 0; i < 3; i++) {
      slide.addShape(pres.shapes.OVAL, {
        objectName: name + ' dot ' + (i + 1), x: x + barH * 0.45 + i * dot * 1.7, y: y + (barH - dot) / 2, w: dot, h: dot,
        fill: { color: '3B342A' }, line: { color: '3B342A', width: 0 }
      });
    }
    slide.addImage({ path: A(c.img), x, y: y + barH, w, h, altText: c.name + ' homepage concept', objectName: name + ' screenshot' });
    return y + barH + h;
  }

  function showcase(slide, c, tag) {
    const w = 9.6;
    const x = (W - w) / 2;
    const bottom = browser(slide, c, x, 0.42, w, c.name);
    const y = bottom + 0.28;
    tb(slide, c.n, { x, y: y + 0.1, w: 0.5, h: 0.3, fontSize: 12, bold: true, color: C.accent1, charSpacing: 2 });
    tb(slide, c.name, { x: x + 0.55, y, w: 2.85, h: 0.5, fontFace: THEME.headFontFace, fontSize: 26, bold: true, color: C.background2 });
    tb(slide, c.feel, { x: x + 3.45, y: y + 0.06, w: 2.8, h: 0.4, fontSize: 15, color: 'BDB4A2' });
    slide.addText([{ text: 'OPEN LIVE ↗', options: { hyperlink: { url: liveUrl(c.key), tooltip: 'Open the ' + c.name + ' live site' } } }], {
      isTextBox: true, x: x + w - 1.45, y: y + 0.04, w: 1.45, h: 0.36, align: 'center', valign: 'middle', margin: 0,
      fontSize: 10, bold: true, charSpacing: 2, color: C.text2, fill: { color: C.accent1 }, objectName: c.name + ' live link'
    });
    slide.addText(tag.toUpperCase(), {
      isTextBox: true, x: x + w - 3.25, y: y + 0.04, w: 1.65, h: 0.36, align: 'center', valign: 'middle', margin: 0,
      fontSize: 10, bold: true, charSpacing: 2, color: C.accent1, line: { color: C.accent2, width: 0.75 }
    });
  }

  function whySlide(slide, c, letter) {
    const bw = 5.9;
    const bh = bw / HERO_RATIO + 0.22;
    const by = (7.5 - bh) / 2 + 0.1;
    browser(slide, c, L, by, bw, c.name, 0.22);
    slide.addText([{ text: 'OPEN THE LIVE SITE ↗', options: { hyperlink: { url: liveUrl(c.key), tooltip: 'Open the ' + c.name + ' live site' } } }], {
      isTextBox: true, x: L, y: by + bh + 0.3, w: 2.3, h: 0.38, align: 'center', valign: 'middle', margin: 0,
      fontSize: 10, bold: true, charSpacing: 2, color: C.text2, fill: { color: C.accent1 }, objectName: c.name + ' live link'
    });
    const x = 7.35, w = W - L - x;
    tb(slide, `FEATURED ${letter} · ${c.n}`, { x, y: 1.2, w, h: 0.3, fontSize: 11, bold: true, charSpacing: 3, color: C.accent2 });
    tb(slide, c.name, { x, y: 1.55, w, h: 0.75, fontFace: THEME.headFontFace, fontSize: 38, bold: true, color: C.text1 });
    tb(slide, `“${c.feeling}”`, { x, y: 2.4, w, h: 0.85, fontFace: THEME.headFontFace, italic: true, fontSize: 18, color: C.accent2 });
    const runs = [];
    c.why.forEach(([b, rest]) => {
      runs.push({ text: b, options: { bold: true, bullet: { indent: 16 } } });
      runs.push({ text: rest, options: { breakLine: true } });
    });
    runs.push({ text: 'Watch: ', options: { bold: true, color: C.accent3, bullet: { indent: 16 } } });
    runs.push({ text: c.risk, options: { color: C.accent3 } });
    tb(slide, runs, { x, y: 3.4, w, h: 2.75, fontSize: 15, color: C.text1, paraSpaceAfter: 9 });
    c.colors.forEach((hex, i) => slide.addShape(pres.shapes.OVAL, {
      objectName: 'swatch ' + (i + 1), x: x + i * 0.42, y: 6.3, w: 0.32, h: 0.32, fill: { color: hex }, line: { color: 'CFC4AE', width: 0.5 }
    }));
    tb(slide, c.accent.toUpperCase(), { x: x + 1.4, y: 6.37, w: w - 1.4, h: 0.25, fontSize: 10, bold: true, charSpacing: 2, color: C.accent3 });
  }

  const hRule = (slide, x, y, w, color, name) =>
    slide.addShape(pres.shapes.LINE, { objectName: name, x, y, w, h: 0, line: { color, width: 0.75 } });

  // ================= OPENING =================
  pres.addSection({ title: 'Opening' });

  let s = pres.addSlide({ masterName: 'ACE Title', sectionTitle: 'Opening' });
  s.addImage({ path: A('deck/assets/ace-logo-gold.png'), x: L, y: 1.05, w: 1.0 * LOGO_RATIO, h: 1.0, altText: 'A.C.E. Distribution', objectName: 'title logo' });
  tb(s, 'WEBSITE DESIGN REVIEW', { x: L, y: 2.95, w: CW, h: 0.35, fontSize: 12, bold: true, charSpacing: 3, color: C.accent1 });
  tb(s, [
    { text: 'Picking ACE’s', options: { breakLine: true } },
    { text: 'public face', options: { italic: true, bold: false, color: C.accent1 } }
  ], { x: L, y: 3.35, w: CW, h: 2.1, fontFace: THEME.headFontFace, fontSize: 66, bold: true, color: C.background2, lineSpacingMultiple: 0.95 });
  tb(s, 'Five directions. Two to sit with. One decision — and the names that carry it.', { x: L, y: 5.65, w: 8, h: 0.8, fontSize: 18, color: 'BDB4A2' });
  s.addNotes(notes('We’re not here to admire mockups forever. We’re here to pick a face and own what happens next.|Keep this short. The room knows the business. Say what the next stretch is: look at the work, choose, put names on it.'));

  s = pres.addSlide({ masterName: 'ACE Light', sectionTitle: 'Opening' });
  kicker(s, 'Why we’re here');
  s.addText([{ text: 'Pick a face. Leave with ' }, { text: 'owners', options: { italic: true, bold: false } }, { text: ', not vibes' }], { placeholder: 'title' });
  [
    ['01', 'Choose the direction the website ships in', 'Five built. Two featured. All on the table.'],
    ['02', 'Put one name on the build', 'One owner, not a committee.'],
    ['03', 'Leave with Rocks and To-Dos', 'Written down before we stand up.']
  ].forEach(([n, t, sub], i) => {
    const y = 3.0 + i * 1.15;
    hRule(s, L, y, 9.7, C.accent4, 'row rule ' + (i + 1));
    tb(s, n, { x: L, y: y + 0.32, w: 0.6, h: 0.3, fontSize: 12, bold: true, color: C.accent2, charSpacing: 2 });
    tb(s, [{ text: t, options: { fontSize: 21, color: C.text1, breakLine: true } }, { text: sub, options: { fontSize: 14, color: C.accent3 } }],
      { x: L + 0.75, y: y + 0.25, w: 8.9, h: 0.8 });
  });
  hRule(s, L, 3.0 + 3 * 1.15, 9.7, C.accent4, 'row rule end');
  s.addNotes(notes('Three outcomes. If we leave with a direction but no owner, we didn’t finish the meeting.|The site is the first thing a new fleet manager or farm buyer sees before they ever call Pleasanton. That’s why it’s worth a real decision, not a vibe check.'));

  s = pres.addSlide({ masterName: 'ACE Statement', sectionTitle: 'Opening' });
  kicker(s, 'The lens for today · EOS');
  s.addText([{ text: 'One name on every ' }, { text: 'Rock', options: { italic: true, bold: false, color: C.accent1 } }], { placeholder: 'title' });
  s.addText('Hold that while we look at the designs. Liking one is easy. Owning it to launch is the job.', { placeholder: 'body' });
  s.addNotes(notes('EOS is blunt on purpose: one name on the Rock. That’s how websites actually ship.|Plant the lens now. We’ll come back to it after the designs, when it’s time to assign.'));

  // ================= DESIGNS =================
  pres.addSection({ title: 'Designs' });

  s = pres.addSlide({ masterName: 'ACE Dark', sectionTitle: 'Designs' });
  kicker(s, 'The work');
  s.addText([{ text: 'Five directions. ' }, { text: 'Two', options: { italic: true, bold: false, color: C.accent1 } }, { text: ' to sit with' }], { placeholder: 'title' });
  {
    const gap = 0.3, tw = (CW - 4 * gap) / 5, th = tw / HERO_RATIO;
    ORDER.forEach((k, i) => {
      const c = CONCEPTS[k], f = FEATURED.includes(k), x = L + i * (tw + gap), y = 3.0;
      if (f) s.addShape(pres.shapes.RECTANGLE, { objectName: c.name + ' highlight', x: x - 0.05, y: y - 0.05, w: tw + 0.1, h: th + 0.1, fill: { color: C.accent1 }, line: { color: C.accent1, width: 0 } });
      s.addImage({ path: A(c.img), x, y, w: tw, h: th, transparency: f ? 0 : 55, altText: c.name, objectName: c.name + ' thumb', hyperlink: { url: liveUrl(k), tooltip: 'Open the ' + c.name + ' live site' } });
      tb(s, c.n + (f ? ' · FEATURED' : ''), { x, y: y + th + 0.25, w: tw, h: 0.25, fontSize: 10, bold: true, charSpacing: 2, color: C.accent1 });
      tb(s, c.name, { x, y: y + th + 0.52, w: tw, h: 0.4, fontFace: THEME.headFontFace, fontSize: 18, bold: true, color: f ? C.background2 : '8C826F' });
      tb(s, c.feel, { x, y: y + th + 0.92, w: tw, h: 0.4, fontSize: 12, color: f ? 'BDB4A2' : '6E6656' });
    });
  }
  s.addNotes(notes('Five directions were built end to end — homepage, services, booking flow — not just mood boards.|The two highlighted are the ones I want us to sit with. The other three are real alternates; we’ll see them side by side after.|Try this with each one: cover the logo for two seconds. Does it still feel like ACE?'));

  FEATURED.forEach((k, i) => {
    const c = CONCEPTS[k], letter = 'AB'[i];
    let sl = pres.addSlide({ masterName: 'ACE Showcase', sectionTitle: 'Designs' });
    showcase(sl, c, 'Featured ' + letter);
    sl.addNotes(notes(c.notes));
    sl = pres.addSlide({ masterName: 'ACE Light Bare', sectionTitle: 'Designs' });
    whySlide(sl, c, letter);
    sl.addNotes(notes('Three reasons it works for ACE, one honest risk.|' + c.why.map(w => w.join('')).join('|') + '|Watch: ' + c.risk));
  });

  s = pres.addSlide({ masterName: 'ACE Light', sectionTitle: 'Designs' });
  kicker(s, 'Also on the table');
  s.addText('The other three', { placeholder: 'title' });
  {
    const gap = 0.45, cw = (CW - 2 * gap) / 3;
    ALTS.forEach((k, i) => {
      const c = CONCEPTS[k], x = L + i * (cw + gap);
      const b = browser(s, c, x, 2.45, cw, c.name, 0.16);
      tb(s, `${c.n} · ${c.accent}`.toUpperCase(), { x, y: b + 0.3, w: cw, h: 0.25, fontSize: 10, bold: true, charSpacing: 1.5, color: C.accent2 });
      tb(s, c.name, { x, y: b + 0.58, w: cw, h: 0.45, fontFace: THEME.headFontFace, fontSize: 22, bold: true, color: C.text1 });
      tb(s, `${c.feel} ${c.feeling}`, { x, y: b + 1.05, w: cw, h: 0.75, fontSize: 14, color: C.accent3 });
    });
  }
  s.addNotes(notes('These are not throwaways. Each one is a complete direction and any of them could ship.|If someone in the room is pulling for one of these, now is the time to say so — we can swap it into the final two.'));

  s = pres.addSlide({ masterName: 'ACE Light', sectionTitle: 'Designs' });
  kicker(s, 'All five, side by side');
  s.addText('Feeling, audience, watch-outs', { placeholder: 'title' });
  {
    const head = ['DIRECTION', 'FEELING', 'SPEAKS TO', 'WATCH OUT FOR'].map(t => ({
      text: t, options: { bold: true, fontSize: 10, charSpacing: 2, color: C.accent2, border: [{ type: 'none' }, { type: 'none' }, { pt: 0.75, color: HEX.accent4 }, { type: 'none' }] }
    }));
    const rows = [head];
    ORDER.forEach(k => {
      const c = CONCEPTS[k], f = FEATURED.includes(k);
      const base = { fontSize: 14, color: C.text1, valign: 'middle', border: [{ type: 'none' }, { type: 'none' }, { pt: 0.75, color: HEX.accent4 }, { type: 'none' }], ...(f ? { fill: { color: 'F1E4C6' } } : {}) };
      rows.push([
        { text: c.name + (f ? '  ★' : ''), options: { ...base, fontFace: THEME.headFontFace, bold: true, fontSize: 16 } },
        { text: c.feel, options: { ...base } },
        { text: c.speaks, options: { ...base } },
        { text: c.risk, options: { ...base, color: C.accent3 } }
      ]);
    });
    s.addTable(rows, { x: L, y: 2.45, w: CW, colW: [2.5, 2.4, 2.9, 3.59], rowH: [0.4, 0.72, 0.72, 0.72, 0.72, 0.72], margin: [0.06, 0.12, 0.06, 0.12], objectName: 'comparison table' });
    tb(s, '★ featured pair', { x: L, y: 6.55, w: 3, h: 0.25, fontSize: 11, color: C.accent3 });
  }
  s.addNotes(notes('One line each so we’re comparing the same things: feeling, who it speaks to, what it costs us.|The watch-out column is where the real conversation is.'));

  s = pres.addSlide({ masterName: 'ACE Light', sectionTitle: 'Designs' });
  kicker(s, 'How the site converts');
  s.addText('Booking is built in — and it wears whichever face we pick', { placeholder: 'title' });
  {
    const gap = 0.25, mw = (CW - 4 * gap) / 5, mh = 2.15;
    ORDER.forEach((k, i) => {
      const c = CONCEPTS[k], x = L + i * (mw + gap), y = 3.45;
      s.addImage({ path: A(c.matrix), x, y, w: mw, h: mh, sizing: { type: 'crop', x: 0, y: 0, w: mw, h: mh }, altText: c.name + ' booking matrix', objectName: c.name + ' matrix',
        shadow: { type: 'outer', color: '281C0A', opacity: 0.3, blur: 12, offset: 4, angle: 90 } });
      tb(s, `${c.n} · ${c.name}`.toUpperCase(), { x, y: y + mh + 0.2, w: mw, h: 0.25, fontSize: 10, bold: true, charSpacing: 1.5, color: C.accent3 });
    });
  }
  s.addNotes(notes('One slide on this — it’s supporting, not the story.|Whichever face we pick, the booking matrix already wears it. The site doesn’t just look good; it turns a visit into a scheduled delivery window.'));

  // ================= ACCOUNTABILITY =================
  pres.addSection({ title: 'Accountability (EOS)' });

  s = pres.addSlide({ masterName: 'ACE Statement', sectionTitle: 'Accountability (EOS)' });
  tb(s, 'EOS', { x: 7.9, y: 4.6, w: 5.0, h: 2.6, fontFace: THEME.headFontFace, fontSize: 170, color: C.accent1, transparency: 85, align: 'right', objectName: 'chapter mark' });
  kicker(s, 'EOS · Traction');
  s.addText('Accountability', { placeholder: 'title' });
  s.addText('Picking a design isn’t the finish line. It’s the first thing someone has to own.', { placeholder: 'body' });
  s.addNotes(notes('Shift gears. The designs are the easy part. This is the part that decides whether the site exists in ninety days.'));

  s = pres.addSlide({ masterName: 'ACE Dark', sectionTitle: 'Accountability (EOS)' });
  kicker(s, 'One seat · one name');
  s.addText([{ text: 'If everyone owns it, ' }, { text: 'no one', options: { italic: true, bold: false, color: C.accent1 } }, { text: ' owns it' }], { placeholder: 'title', fontSize: 52 });
  {
    // simple seat diagram: many dashed names vs one solid seat
    const cy = 3.55;
    ['Everyone', 'The team', 'Someone'].forEach((t, i) => {
      s.addText(t, { isTextBox: true, x: L + i * 1.75, y: cy, w: 1.55, h: 0.5, align: 'center', valign: 'middle', margin: 0, fontSize: 14, color: '8C826F',
        shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, line: { color: '5A5142', width: 1, dashType: 'dash' }, objectName: 'vague owner ' + (i + 1) });
    });
    tb(s, '→', { x: L + 5.3, y: cy - 0.02, w: 0.6, h: 0.5, fontSize: 26, color: C.accent1, align: 'center' });
    s.addText('One name', { isTextBox: true, x: L + 6.0, y: cy, w: 1.9, h: 0.5, align: 'center', valign: 'middle', margin: 0, fontSize: 15, bold: true, color: C.text2,
      shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, fill: { color: C.accent1 }, line: { color: C.accent1, width: 1 }, objectName: 'one owner' });
    const colW = (CW - 0.8) / 2;
    [['ACCOUNTABILITY CHART', 'Names who owns each outcome — not who reports to whom. The website gets a seat, and the seat gets one name.'],
     ['NOT BLAME', 'Clarity + ownership + follow-through. Clarity kills drift. Nobody wonders who’s driving.']].forEach(([h, p], i) => {
      const x = L + i * (colW + 0.8);
      tb(s, h, { x, y: 4.85, w: colW, h: 0.3, fontSize: 11, bold: true, charSpacing: 3, color: C.accent1 });
      tb(s, p, { x, y: 5.25, w: colW, h: 1.2, fontSize: 17, color: 'BDB4A2' });
    });
  }
  s.addNotes(notes('In EOS every seat has one person accountable. Not two. Not “the team.”|The Accountability Chart isn’t the org chart. It’s who owns which outcome. The website needs a seat and a name in it.|Accountability isn’t blame. It’s clarity — so nobody has to guess who’s driving.'));

  s = pres.addSlide({ masterName: 'ACE Dark', sectionTitle: 'Accountability (EOS)' });
  kicker(s, 'Commitments');
  s.addText('What you say, you do', { placeholder: 'title' });
  {
    const colW = CW / 2;
    hRule(s, L, 2.75, CW, '3A3226', 'top rule');
    s.addShape(pres.shapes.LINE, { objectName: 'divider', x: L + colW, y: 2.75, w: 0, h: 3.6, line: { color: '3A3226', width: 0.75 } });
    [['Rocks', '90 DAYS · FEW · IMPORTANT', ['The handful of things that must be true by the end of the quarter.', 'You take a Rock → you drive it to done.']],
     ['To-Dos', '7 DAYS · SHORT · SPECIFIC', ['This week’s promises to the team.', 'Done — or flag it early. Never quietly late.']]].forEach(([h, m, ps], i) => {
      const x = L + i * colW + (i ? 0.55 : 0), w = colW - 0.55;
      tb(s, h, { x, y: 3.15, w, h: 0.8, fontFace: THEME.headFontFace, fontSize: 40, bold: true, color: C.background2 });
      tb(s, m, { x, y: 4.0, w, h: 0.3, fontSize: 11, bold: true, charSpacing: 3, color: C.accent1 });
      tb(s, ps.map((p, j) => ({ text: p, options: { breakLine: j < ps.length - 1 } })), { x, y: 4.5, w, h: 1.6, fontSize: 18, color: 'BDB4A2', paraSpaceAfter: 8 });
    });
  }
  s.addNotes(notes('Rocks are the few things that matter this quarter. You take one, you drive it to done.|To-Dos are the weekly promises. You do them — or you raise the flag early, not at the deadline.|Both are commitments to yourself and to the team. Same standard.'));

  s = pres.addSlide({ masterName: 'ACE Statement', sectionTitle: 'Accountability (EOS)' });
  kicker(s, 'Drive');
  s.addText([{ text: 'Finish what you ' }, { text: 'gave yourself', options: { italic: true, bold: false, color: C.accent1 } }], { placeholder: 'title' });
  s.addText('Ambition without ownership is noise. Credit goes to the finish — not the intention.', { placeholder: 'body' });
  s.addNotes(notes('This is the personal one. You are accountable for the things you give yourself.|Ambition without ownership is noise. Everyone in this room has had a “we should redo the website” moment. You don’t get credit for the intention to rebuild the site. You get credit for the finish.'));

  s = pres.addSlide({ masterName: 'ACE Light', sectionTitle: 'Accountability (EOS)' });
  kicker(s, 'This quarter · proposed Rocks');
  s.addText('The website, as Rocks', { placeholder: 'title' });
  {
    const bd = [{ type: 'none' }, { type: 'none' }, { pt: 0.75, color: HEX.accent4 }, { type: 'none' }];
    const rows = [[
      { text: 'ROCK', options: { bold: true, fontSize: 10, charSpacing: 2, color: C.accent2, border: bd } },
      { text: 'OWNER', options: { bold: true, fontSize: 10, charSpacing: 2, color: C.accent2, border: bd } }
    ]];
    [['Launch the homepage in the chosen direction', 'Live on the domain, mobile-checked, real photos.'],
     ['Booking flow taking real requests', 'Matrix live, requests landing with whoever confirms windows.'],
     ['Content & photos ready to publish', 'Services, coverage, contact — written and approved.']].forEach(([t, sub]) => {
      rows.push([
        { text: [{ text: t, options: { fontSize: 18, color: C.text1, breakLine: true } }, { text: sub, options: { fontSize: 13, color: C.accent3 } }], options: { border: bd, valign: 'middle' } },
        { text: '', options: { border: bd, fontFace: THEME.headFontFace, fontSize: 18, color: C.text1, valign: 'middle' } }
      ]);
    });
    s.addTable(rows, { x: L, y: 2.5, w: CW, colW: [7.6, 3.79], rowH: [0.4, 0.95, 0.95, 0.95], margin: [0.06, 0.12, 0.06, 0.0], objectName: 'rocks table' });
    tb(s, 'Type owner names into the Owner column during the meeting.', { x: L, y: 6.4, w: 8, h: 0.3, fontSize: 12, italic: true, color: C.accent3 });
  }
  s.addNotes(notes('What it means for ACE this quarter. These are proposed — adjust them in the room.|Type owner names straight into the Owner column.'));

  // ================= DECISION =================
  pres.addSection({ title: 'Decision' });

  s = pres.addSlide({ masterName: 'ACE Dark', sectionTitle: 'Decision' });
  kicker(s, 'Decision');
  s.addText('Which face ships?', { placeholder: 'title' });
  {
    const gap = 0.22, cw = (CW - 4 * gap) / 5;
    ORDER.forEach((k, i) => {
      const c = CONCEPTS[k], f = FEATURED.includes(k), x = L + i * (cw + gap), y = 2.55;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: c.name + ' card', x, y, w: cw, h: 1.2, rectRadius: 0.08,
        fill: { color: '2A2116', transparency: f ? 0 : 40 }, line: { color: f ? HEX.accent1 : '3A3226', width: f ? 1.25 : 0.75 } });
      tb(s, c.n, { x: x + 0.2, y: y + 0.18, w: cw - 0.4, h: 0.25, fontSize: 10, bold: true, charSpacing: 2, color: C.accent1 });
      tb(s, c.name, { x: x + 0.2, y: y + 0.45, w: cw - 0.4, h: 0.4, fontFace: THEME.headFontFace, fontSize: 17, bold: true, color: C.background2, fit: 'shrink' });
      tb(s, c.feel, { x: x + 0.2, y: y + 0.82, w: cw - 0.4, h: 0.3, fontSize: 11, color: '9C9381' });
    });
    const bd = [{ type: 'none' }, { type: 'none' }, { pt: 0.75, color: '3A3226' }, { type: 'none' }];
    const hd = t => ({ text: t, options: { bold: true, fontSize: 10, charSpacing: 2, color: C.accent1, border: bd } });
    const rows = [[hd('TO-DO · THIS WEEK'), hd('OWNER'), hd('DUE')]];
    ['Confirm the direction (or final two + who decides)', 'Kick off the build with the chosen design', 'Gather photos & copy owners need'].forEach(t => {
      rows.push([
        { text: t, options: { fontSize: 16, color: C.background2, border: bd, valign: 'middle' } },
        { text: '', options: { fontSize: 16, color: C.background2, border: bd, valign: 'middle', fontFace: THEME.headFontFace } },
        { text: '', options: { fontSize: 16, color: C.background2, border: bd, valign: 'middle' } }
      ]);
    });
    s.addTable(rows, { x: L, y: 4.2, w: CW, colW: [6.4, 3.0, 1.99], rowH: [0.38, 0.6, 0.6, 0.6], margin: [0.06, 0.12, 0.06, 0.0], objectName: 'todo table' });
  }
  s.addNotes(notes('Ask it plainly: which face ships? If the room can’t land on one, land on two and name who decides by when.|Fill the To-Do owners and dates before anyone stands up.'));

  s = pres.addSlide({ masterName: 'ACE Title', sectionTitle: 'Decision' });
  s.addImage({ path: A('deck/assets/ace-logo-gold.png'), x: L, y: 1.05, w: 0.78 * LOGO_RATIO, h: 0.78, altText: 'A.C.E. Distribution', objectName: 'close logo' });
  tb(s, [
    { text: 'Own it.', options: { breakLine: true } },
    { text: 'Finish it.', options: { italic: true, bold: false, color: C.accent1 } }
  ], { x: L, y: 2.55, w: CW, h: 2.4, fontFace: THEME.headFontFace, fontSize: 66, bold: true, color: C.background2, lineSpacingMultiple: 0.95 });
  tb(s, 'We picked a face. Now we put the names on it and drive it to done.', { x: L, y: 5.2, w: 8, h: 0.8, fontSize: 18, color: 'BDB4A2' });
  s.addNotes(notes('Close on the commitment, not the design. Read the names back.'));

  // ================= APPENDIX =================
  pres.addSection({ title: 'Appendix' });

  s = pres.addSlide({ masterName: 'ACE Statement', sectionTitle: 'Appendix' });
  kicker(s, 'Appendix');
  s.addText('The alternates, full size', { placeholder: 'title' });
  s.addText('Swap any of these into the final two if the room is pulling for it.', { placeholder: 'body' });
  s.addNotes(notes('Backup slides: a full view of each alternate, in case the room wants to swap one into the final two.'));

  ALTS.forEach(k => {
    const c = CONCEPTS[k];
    const sl = pres.addSlide({ masterName: 'ACE Showcase', sectionTitle: 'Appendix' });
    showcase(sl, c, 'Alternate');
    sl.addNotes(notes(c.notes));
  });

  await pres.writeFile({ fileName: OUT });
  const { applyTheme } = require('./apply_theme.js');
  await applyTheme(OUT, THEME);
  console.log('wrote', OUT);
})();
