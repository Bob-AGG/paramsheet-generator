// 参数单生成器 — 真实用户流程验收（浏览器级）
const puppeteer = require('puppeteer-core');
const CHROME = "/Users/bob/Library/Caches/ms-playwright/chromium-1234/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing";
const BASE = process.argv[2] || 'http://127.0.0.1:8080';
const R = [];
const ok = (n,v,d) => R.push({n, pass: !!v, d});

(async () => {
  const b = await puppeteer.launch({executablePath: CHROME, headless:'new', args:['--no-sandbox'], protocolTimeout:240000});
  const p = await b.newPage();
  await p.setViewport({width:1600,height:1000});
  const errs=[];
  p.on('pageerror', e=>errs.push('PAGEERR: '+e.message));
  p.on('console', m=>{ if(m.type()==='error' && !m.text().includes('favicon') && !m.text().includes('404')) errs.push('CONSOLE: '+m.text()); });
  // 拦截 alert/prompt，避免卡死
  await p.evaluateOnNewDocument(()=>{ window.alert=()=>{}; });

  await p.goto(BASE+'/', {waitUntil:'networkidle2', timeout:60000});
  await new Promise(r=>setTimeout(r,1200));

  // === 1. 标准库基础状态 ===
  ok('列表默认加载', (await p.$eval('#listCnt', e=>e.textContent)).includes('1945'));
  const vOpts = await p.$eval('#fVolt', e=>[...e.options].map(o=>o.value));
  ok('电压下拉有值', vOpts.length>30, vOpts.length+' 项');
  const pendTxt = await p.$eval('#pendBtn', e=>[e.style.display,e.textContent]);
  ok('待分类按钮可见', pendTxt[0]!=='none' && pendTxt[1].includes('46'), pendTxt.join('|'));

  // === 2. 电压筛选（真实下拉选择）===
  await p.select('#fVolt','400/230V');
  await new Promise(r=>setTimeout(r,1500));
  const cnt400 = await p.$eval('#listCnt', e=>e.textContent);
  ok('400/230V 筛选 1322 条', cnt400.includes('1322'), cnt400);

  // === 3. 待分类切换 ===
  await p.select('#fVolt','');
  await new Promise(r=>setTimeout(r,1200));
  await p.click('#pendBtn');
  await new Promise(r=>setTimeout(r,1500));
  const cntPend = await p.$eval('#listCnt', e=>e.textContent);
  ok('待分类筛选 46 条', cntPend.includes('46'), cntPend);
  const pendBtn2 = await p.$eval('#pendBtn', e=>e.textContent);
  ok('待分类按钮显示激活态', pendBtn2.includes('✕'), pendBtn2);
  await p.click('#pendBtn');  // 关闭
  await new Promise(r=>setTimeout(r,1500));
  ok('关闭待分类回到 1945', (await p.$eval('#listCnt', e=>e.textContent)).includes('1945'));

  // === 4. 关键词搜索 ===
  await p.type('#fKw','AF11D5');
  await p.keyboard.press('Enter');
  await new Promise(r=>setTimeout(r,1500));
  const cntKw = await p.$eval('#listCnt', e=>e.textContent);
  ok('关键词搜索有结果', /[1-9]\d* 条/.test(cntKw), cntKw);

  // === 5. 打开详情 ===
  const firstItem = await p.$('#list .it, #list [onclick]');
  ok('列表项存在', !!firstItem);
  if (firstItem) {
    await firstItem.click();
    await new Promise(r=>setTimeout(r,2500));
    const nm = await p.$eval('#editName', e=>e.textContent);
    ok('详情打开有名称', nm && nm.length>3, nm);
    const editorHtml = await p.$eval('#editor', e=>e.innerHTML.length);
    ok('编辑器有内容', editorHtml>2000, editorHtml+' chars');
    const pvHtml = await p.$eval('#preview', e=>e.innerHTML.length);
    ok('预览区渲染', pvHtml>100, pvHtml+' chars');
  }

  // === 6. 切到非标库 ===
  await p.evaluate(()=>setKind('nonstd'));
  await new Promise(r=>setTimeout(r,2000));
  ok('非标库 112 条', (await p.$eval('#listCnt', e=>e.textContent)).includes('112'),
     await p.$eval('#listCnt', e=>e.textContent));
  ok('非标模式显示派生按钮', (await p.$eval('#deriveBtn', e=>e.style.display))!=='none');
  ok('非标模式隐藏新增按钮', (await p.$eval('#newBtn', e=>e.style.display))==='none');
  const pendHidden = await p.$eval('#pendBtn', e=>e.style.display);
  ok('非标模式隐藏待分类', pendHidden==='none', pendHidden);

  // 非标筛选：动力品牌
  const famOpts = await p.$eval('#fFamily', e=>[...e.options].filter(o=>o.value).map(o=>o.value));
  ok('非标库有动力品牌筛选', famOpts.length>0, famOpts.join(',').slice(0,60));
  if (famOpts.length) {
    await p.select('#fFamily', famOpts[0]);
    await new Promise(r=>setTimeout(r,1500));
    const n = await p.$eval('#listCnt', e=>e.textContent);
    ok('非标按品牌筛选生效', /^\d+ 条$/.test(n), famOpts[0]+' -> '+n);
    await p.select('#fFamily','');
    await new Promise(r=>setTimeout(r,1200));
  }

  // === 7. 打开一条非标单，检查差异条 ===
  const ni = await p.$('#list .it, #list [onclick]');
  if (ni) {
    await ni.click();
    await new Promise(r=>setTimeout(r,3000));
    const nm = await p.$eval('#editName', e=>e.textContent);
    ok('非标详情打开', nm && nm.length>3, nm);
    const bar = await p.$eval('#nstdbar', e=>[e.style.display, e.textContent.slice(0,80)]);
    ok('非标提示条显示', bar[0]==='block' && bar[1].includes('非标'), bar.join(' | '));
  }

  // === 8. 切回标准库，确认筛选被重置 ===
  await p.evaluate(()=>setKind('std'));
  await new Promise(r=>setTimeout(r,2000));
  const back = await p.$eval('#listCnt', e=>e.textContent);
  ok('切回标准库回到全量', back.includes('1945'), back);
  const vReset = await p.$eval('#fVolt', e=>e.value);
  ok('切库后电压筛选已重置', vReset==='', vReset||'(空)');

  ok('无 JS 报错', errs.filter(e=>!e.includes('favicon')).length===0, errs.slice(0,3).join(' || ')||'无');

  await b.close();
  const pass = R.filter(r=>r.pass).length;
  console.log(`\n===== ${BASE} =====`);
  R.forEach(r=>console.log(`${r.pass?'✅':'❌'} ${r.n}${r.d?'  ['+r.d+']':''}`));
  console.log(`\n结果: ${pass}/${R.length} 通过`);
  process.exit(pass===R.length?0:1);
})().catch(e=>{ console.error('FAIL:', e.message); process.exit(2); });
