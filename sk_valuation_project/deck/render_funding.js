// 펀딩 덱 렌더러 — funding_slides.html을 Chromium으로 고해상도 PNG 렌더 → pptx에 full-bleed 삽입
// 레퍼런스(Tesla ESS) 방식: 폰트를 이미지에 구워넣어 어떤 PC에서도 동일한 SK 룩.
// 실행: NODE_PATH=$(npm root -g) node render_funding.js
const path = require('path');
const groot = require('child_process').execSync('npm root -g').toString().trim();
const { chromium } = require(groot + '/playwright');
const pptxgen = require(groot + '/pptxgenjs');

const SLIDES = ['s1','s2','s3','s4','s5'];
const W = 1920, H = 1358, SCALE = 2;          // → 3840×2716 PNG (A4 가로 √2 비율)
const HTML = 'file://' + path.join(__dirname, 'funding_slides.html');
const OUTDIR = path.join(__dirname, '..', 'output');
const IMGDIR = path.join(OUTDIR, 'funding_png');

(async () => {
  require('fs').mkdirSync(IMGDIR, { recursive: true });
  const browser = await chromium.launch({ args: ['--no-sandbox','--force-color-profile=srgb'] });
  const page = await browser.newPage({ viewport: { width: W, height: H }, deviceScaleFactor: SCALE });
  await page.goto(HTML, { waitUntil: 'networkidle' });
  await page.evaluate(() => document.fonts.ready);
  await page.waitForTimeout(300);

  const files = [];
  for (const id of SLIDES) {
    const el = await page.$('#' + id);
    const f = path.join(IMGDIR, id + '.png');
    await el.screenshot({ path: f });
    files.push(f);
    console.log('rendered', id);
  }
  await browser.close();

  // ── pptx 조립 (A4 가로, 각 슬라이드 = full-bleed 이미지) ──
  const pres = new pptxgen();
  const PW = 11.69, PH = 8.27;
  pres.defineLayout({ name: 'A4L', width: PW, height: PH });
  pres.layout = 'A4L';
  for (const f of files) {
    const s = pres.addSlide();
    s.background = { color: 'FAF7F1' };
    s.addImage({ path: f, x: 0, y: 0, w: PW, h: PH });
  }
  const out = path.join(OUTDIR, 'SK_Funding_v3.pptx');
  await pres.writeFile({ fileName: out });
  console.log('WROTE', out);
})();
