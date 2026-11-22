const fs = require('fs');
const path = require('path');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const projectRoot = path.resolve(__dirname, '..');
(async () => {
 const browser = await chromium.launch({channel: process.env.UI_BROWSER_CHANNEL || 'msedge', headless:true});
 const page = await browser.newPage({viewport:{width:380,height:850}});
 const errors=[]; page.on('pageerror',e=>errors.push(e.message));
 await page.route('**/*', async route => {
  const url=new URL(route.request().url());
  if(url.pathname.endsWith('/office-js/office.js')) return route.fulfill({contentType:'application/javascript',body:`window.Office={onReady:fn=>fn(),EventType:{DocumentSelectionChanged:'selection'},context:{document:{settings:{get:()=>null,set:()=>{},remove:()=>{},saveAsync:()=>{}},addHandlerAsync:()=>{}}}};window.Excel={run:()=>Promise.reject(new Error('Preview: Excel unavailable'))};`});
  if(url.pathname.startsWith('/api/')) return route.fulfill({status:404,body:'{}',contentType:'application/json'});
  const name=url.pathname==='/'?'taskpane.html':url.pathname.slice(1);
  const assetPath=path.join(projectRoot, 'dist', name);
  if(fs.existsSync(assetPath)) return route.fulfill({body:fs.readFileSync(assetPath),contentType:name.endsWith('.js')?'application/javascript':name.endsWith('.html')?'text/html':'application/octet-stream'});
  return route.abort();
 });
 await page.goto('https://excel-agent.test/taskpane.html');
 await page.locator('.welcome-panel').waitFor({timeout:10000}).catch(async e=>{console.log(await page.locator('body').innerText(),errors);throw e;});
 if(await page.locator('.chat-list').evaluate(el=>el.scrollTop)!==0) throw Error('Welcome scrolled away');
 const settleAnimations = () => page.evaluate(() => document.getAnimations().forEach(animation => { if (animation.effect?.getTiming().iterations !== Infinity) animation.finish(); }));
 await settleAnimations();
 await page.screenshot({path:path.join(__dirname, 'workspace-home.png')});
 await page.getByRole('button',{name:'全部任务'}).click();
 await page.getByRole('textbox',{name:'搜索快捷任务'}).fill('找不到的任务');
 await page.getByText('没有匹配的任务，试试其他关键词。').waitFor();
 await page.getByRole('textbox',{name:'搜索快捷任务'}).press('Escape');
 // 退场动画播完才卸载，等待 DOM 移除而不是立刻断言
 await page.locator('.task-catalog').waitFor({state:'detached',timeout:3000}).catch(()=>{throw Error('Escape did not dismiss');});
 await page.getByRole('button',{name:'全部任务'}).click();
 await page.getByRole('textbox',{name:'搜索快捷任务'}).fill('公式');
 if(await page.locator('.task-results button').count()<1) throw Error('Search result mismatch');
 await page.getByRole('button',{name:'写公式',exact:false}).click();
 await page.locator('.task-catalog').waitFor({state:'detached'});
 if(!(await page.getByRole('textbox',{name:'任务指令'}).inputValue()).includes('公式')) throw Error('Prompt not filled');
 await page.getByRole('textbox',{name:'任务指令'}).fill('这是需要多行显示的一段任务说明。'.repeat(20));
 await page.waitForFunction(() => document.querySelector('.input-row textarea').clientHeight >= 120);
 const height=await page.getByRole('textbox',{name:'任务指令'}).evaluate(el=>el.clientHeight);
 if(height<120 || height>160) throw Error('Auto height '+height);
 await page.getByRole('textbox',{name:'任务指令'}).fill('');
 await page.setViewportSize({width:300,height:650});
 await page.screenshot({path:path.join(__dirname, 'workspace-narrow.png')});
 const overflow=await page.evaluate(()=>document.documentElement.scrollWidth>innerWidth);
 if(overflow) throw Error('Horizontal overflow');
 if(errors.length) throw Error(errors.join('\n'));
 await page.setViewportSize({width:380,height:850});
 await page.evaluate(() => {
  const meta={id:'ui-preview',title:'销售汇总 · 界面示例',createdAt:Date.now(),updatedAt:Date.now()};
  localStorage.setItem('excel-ai-agent-conv-index',JSON.stringify({currentId:meta.id,items:[meta]}));
  localStorage.setItem('excel-ai-agent-conv:'+meta.id,JSON.stringify({...meta,messages:[],items:[
   {kind:'user',id:'demo-user',text:'按地区汇总销售额，找出表现最好的地区。'},
   {kind:'tool',id:'demo-tool',label:'读取销售表',name:'read_range',status:'ok',argsSummary:'销售明细 · A1:F128',detail:'界面示例数据，未执行 Excel 操作。'},
   {kind:'assistant',id:'demo-reply',text:'已按地区整理汇总。\n\n| 地区 | 销售额 | 占比 |\n| --- | ---: | ---: |\n| 华东 | 128,600 | 46% |\n| 华南 | 96,400 | 34% |\n| 华北 | 55,000 | 20% |\n\n**华东表现最好**，销售额占总额的 46%。\n\n以上为界面展示用的示例数据。'}
  ]}));
 });
 await page.reload();
 await page.locator('.assistant-bubble').waitFor();
 await settleAnimations();
 await page.screenshot({path:path.join(__dirname,'workspace-chat.png')});
 console.log('PASS: task search, prompt fill, textarea resize, 300px layout; Office mocked.');
 await browser.close();
})().catch(e=>{console.error(e);process.exit(1)});


