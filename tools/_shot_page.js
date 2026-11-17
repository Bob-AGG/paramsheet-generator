const puppeteer = require('/Users/bob/.openclaw/workspace/genset-selector/node_modules/puppeteer-core');

const CHROME = '/Users/bob/Library/Caches/ms-playwright/chromium-1234/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing';
const OUT = '/Users/bob/.openclaw/workspace/projects/参数单生成器/docs/screenshots';

(async () => {
  const browser = await puppeteer.launch({
    executablePath: CHROME,
    headless: 'new',
    args: ['--no-sandbox', '--force-device-scale-factor=2'],
    protocolTimeout: 180000,
  });
  const page = await browser.newPage();
  await page.setViewport({ width: 1680, height: 1050, deviceScaleFactor: 2 });
  page.on('dialog', async d => { try { await d.accept(); } catch(e){} });

  await page.goto('http://127.0.0.1:8080/', { waitUntil: 'networkidle2', timeout: 60000 });
  await new Promise(r => setTimeout(r, 1500));
  await page.screenshot({ path: `${OUT}/01_初始页.png` });
  console.log('shot 01 ok');

  // 尝试搜索 AF11D5
  const typed = await page.evaluate(() => {
    const cand = [...document.querySelectorAll('input')].find(i => (i.placeholder||'').includes('搜索') || (i.placeholder||'').includes('型号') || (i.type==='search'));
    if (!cand) return null;
    cand.value = 'AF11D5';
    cand.dispatchEvent(new Event('input', { bubbles: true }));
    cand.dispatchEvent(new Event('keyup', { bubbles: true }));
    if (cand.form) cand.form.dispatchEvent(new Event('submit', { bubbles: true }));
    return cand.placeholder || 'found';
  });
  console.log('search box:', typed);
  await new Promise(r => setTimeout(r, 2500));
  await page.screenshot({ path: `${OUT}/02_搜索AF11D5.png` });
  console.log('shot 02 ok');

  // 点击第一条结果
  const clicked = await page.evaluate(() => {
    const items = [...document.querySelectorAll('[class*=item],[class*=row],li,tr')].filter(e => (e.innerText||'').includes('AF11D5') && e.getBoundingClientRect().height > 10);
    if (!items.length) return null;
    items[0].click();
    return items[0].innerText.slice(0, 60);
  });
  console.log('clicked:', clicked);
  await new Promise(r => setTimeout(r, 3000));
  await page.screenshot({ path: `${OUT}/03_详情编辑.png` });
  console.log('shot 03 ok');

  await browser.close();
})().catch(e => { console.error('ERR', e.message); process.exit(1); });
