// Run with: node test_frontend.cjs (no browser dependencies).
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const elements = new Map();
function element(id) {
  if (!elements.has(id)) elements.set(id, {
    value: '0', textContent: '', innerHTML: '', disabled: false, hidden: false, checked: false, files: [],
    listeners: {}, dataset: {}, classList: {toggle() {}, remove() {}, add() {}},
    setAttribute() {}, addEventListener(name, callback) {this.listeners[name] = callback;},
  });
  return elements.get(id);
}
for (const id of ['search','compare-search']) element(id).value = '';
element('font').value = 'Target'; element('size').value = '10';
element('view').value = 'list';element('compare-filter').value = 'all';
const cell = {address:'A1',row:1,col:1,value:'<changed>',colors:Array(7).fill('none'),font:'Target',size:10};
const oldCell = {...cell,value:'before'};
const makeBook = (name, entry) => ({name,title:'',sheets:[{name:'Sheet',rows:1,cols:1,cells:[entry]}]});
const change = {kind:'修改',old:oldCell,new:cell,reference:null,basis:'区段顺序推测',format_changes:[{label:'字号',field:'size',old:10,new:12}],content_changed:true,moved:false,review:true,conclusion:'匹配需确认'};
const summary = {total:1,content:1,format:1,moved:0,review:1,added:0,deleted:0,unmatched_sheets:0,sheet_review:0};
const payload = {reports:[makeBook('old.xlsx',oldCell),makeBook('new.xlsx',cell)],comparison:{old_file:'old.xlsx',new_file:'new.xlsx',summary,notes:['差异不等于错误'],sheets:[{old_name:'Sheet',new_name:'Sheet',pairing:'同名工作表',status:'已配对',summary,metadata:[],changes:[change]}]}};
let rejectRequest = false, sent;
const context = {console, Blob, URL, setTimeout, clearTimeout, btoa,
 document:{getElementById:element,querySelectorAll:()=>[],querySelector:()=>({content:'token'}),addEventListener(){}},
 fetch:async (url, options) => {
   assert.equal(url,'/api/compare');assert.equal(element('old-file').disabled,true);
   sent=JSON.parse(options.body);
   return {ok:!rejectRequest,json:async()=>rejectRequest?{error:'无法读取工作簿'}:payload};
 }};
vm.createContext(context);
for (const path of ['web/js/app.js','web/js/comparison.js']) vm.runInContext(fs.readFileSync(path,'utf8'),context);
(async () => {
  await element('compare-run').onclick();
  assert.match(element('compare-error').textContent,/请选择/);
  for (const [id,name] of [['old-file','old.xlsx'],['new-file','new.xlsx']]) element(id).files=[{name,size:1,arrayBuffer:async()=>new Uint8Array([1]).buffer}];
  const drop = (version, files, items = []) => {
    let prevented = false;
    element(version + '-drop').listeners.drop({preventDefault(){prevented=true;},dataTransfer:{files,items}});
    assert.equal(prevented,true);
  };
  const previous = element('old-file').files[0];
  element('old-file').files = [];
  drop('old',[previous]);
  element('new-file').onchange();
  assert.match(element('old-file-name').textContent,/old.xlsx/);
  assert.match(element('new-file-name').textContent,/new.xlsx/);
  await element('compare-run').onclick();
  assert.equal(sent.old.name,'old.xlsx');assert.equal(sent.new.name,'new.xlsx');
  assert.equal(element('compare-output').hidden,false);
  assert.equal(element('compare-run').disabled,false);
  assert.match(element('compare-rows').innerHTML,/&lt;changed&gt;/);
  assert.match(element('results').innerHTML,/&lt;changed&gt;/);
  assert.match(element('compare-rows').innerHTML,/字号：10 → 12/);
  element('compare-filter').value='moved';element('compare-filter').listeners.change();
  assert.match(element('compare-rows').innerHTML,/没有符合条件/);
  element('compare-filter').value='review';element('compare-filter').listeners.change();
  assert.match(element('compare-rows').innerHTML,/匹配需确认/);
  drop('old',[previous,previous]);assert.match(element('compare-error').textContent,/只放入一个/);
  drop('old',[{name:'wrong.txt',size:1}]);assert.match(element('compare-error').textContent,/xlsx/);
  drop('old',[{name:'huge.xlsx',size:31*1024*1024}]);assert.match(element('compare-error').textContent,/30 MB/);
  drop('old',[],[{webkitGetAsEntry:()=>({isDirectory:true})}]);assert.match(element('compare-error').textContent,/文件夹/);
  vm.runInContext('running=true',context);
  drop('old',[{...previous,name:'ignored.xlsx'}]);
  vm.runInContext('running=false',context);
  assert.match(element('old-file-name').textContent,/old.xlsx/);
  drop('old',[previous]);assert.equal(element('compare-output').hidden,true);
  rejectRequest=true;await element('compare-run').onclick();
  assert.equal(element('compare-output').hidden,true);
  assert.equal(element('compare-export').disabled,true);
  assert.match(element('compare-error').textContent,/无法读取/);
  element('new-file').onchange();assert.equal(element('compare-error').textContent,'');
  console.log('Frontend workflow passed: validation, version upload, escaped results, seven-rule retention, filters, errors, drag-and-drop validation, busy-state protection and stale-result reset.');
})().catch(error=>{console.error(error);process.exitCode=1;});
