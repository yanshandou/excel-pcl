'use strict';
// Version comparison is separate from the original seven-rule inspector.
(() => {
  let result = null;
  let mapping = Object.create(null);
  let comparisonPage = 0;
  let comparisonReports = [];
  const versionFiles = {old: null, new: null};
  const pageSize = 100;
  const labels = {total:'差异条目',content:'内容变化',format:'格式差异',moved:'位置变化',review:'需确认条目',unmatched_sheets:'未配对工作表',sheet_review:'工作表需关注'};
  const text = value => value === null || value === undefined ? '未指定' : typeof value === 'object' ? JSON.stringify(value) : String(value);
  const snapshot = cell => cell ? `<b>${esc(cell.address)}</b><small>${esc(cell.value) || '（空白）'}</small>` : '—';

  function renderComparison() {
    if (!result) return;
    const sheet = result.sheets[+$('compare-sheet').value];
    const filter = $('compare-filter').value;
    const query = $('compare-search').value.toLowerCase();
    const changes = (sheet?.changes || []).filter(change =>
      (filter === 'all' || filter === 'review' && change.review || filter === 'format' && change.format_changes.length || filter === 'content' && change.content_changed || filter === 'moved' && change.moved)
      && (!query || JSON.stringify(change).toLowerCase().includes(query)));
    const pages = Math.max(1, Math.ceil(changes.length / pageSize));
    comparisonPage = Math.min(comparisonPage, pages - 1);
    $('compare-page').textContent = `${comparisonPage + 1} / ${pages} · ${changes.length} 条`;
    $('compare-prev').disabled = comparisonPage === 0;
    $('compare-next').disabled = comparisonPage >= pages - 1;
    $('compare-metadata').textContent = sheet ? `${sheet.pairing}。${sheet.metadata.map(item => `${item.label}：${text(item.old)} → ${text(item.new)}`).join('；')}${sheet.status !== '已配对' ? '。该工作表未完成逐格比较，请手工配对，或人工核实是否新增 / 删除。' : ''}` : '';
    $('compare-rows').innerHTML = changes.slice(comparisonPage * pageSize, (comparisonPage + 1) * pageSize).map(change => {
      const formats = change.format_changes.length ? '<ul>' + change.format_changes.map(item => `<li>${esc(item.label)}：${esc(text(item.old))} → ${esc(text(item.new))}</li>`).join('') + '</ul>' : '未发现已比较属性的变化';
      const reference = change.reference ? `<small>建议参考 ${esc(sheet.old_name)}!${esc(change.reference.address)}：${esc(change.reference.value)}</small>` : '';
      return `<tr class="${change.review ? 'needs-review' : ''}"><td>${esc(change.kind)}<small>${esc(change.conclusion)}</small></td><td>${snapshot(change.old)}</td><td>${snapshot(change.new)}</td><td>${formats}</td><td>${esc(change.basis)}${reference}</td><td>${change.old ? '<button data-open-version="0">上一版</button>' : ''}${change.new ? '<button data-open-version="1">这一版</button>' : ''}</td></tr>`;
    }).join('') || '<tr><td colspan="6">没有符合条件的差异；未配对工作表请查看上方提示。</td></tr>';
  }

  function showResult(payload) {
    result = payload.comparison;
    comparisonReports = payload.reports;
    // Feed both versions into the original seven-rule viewer without replacing its rules.
    reports = comparisonReports;
    failures = [];
    $('book').innerHTML = reports.map((book, index) => `<option value="${index}">${index ? '这一版' : '上一版'} · ${esc(book.name)}</option>`).join('');
    $('book').value = '1';
    $('books').textContent = '2';
    $('failed').textContent = '0';
    $('sheets').textContent = reports.reduce((sum, book) => sum + book.sheets.length, 0);
    $('cells').textContent = reports.reduce((sum, book) => sum + book.sheets.reduce((count, sheet) => count + sheet.cells.length, 0), 0).toLocaleString();
    $('status').textContent = '版本比较完成，可在此查看两个版本的七项检查。';
    $('error').textContent = '';
    $('search').value = ''; $('marked').checked = false;
    changeBook();
    $('compare-output').hidden = false;
    $('compare-export').disabled = false;
    $('compare-summary').innerHTML = Object.entries(labels).map(([key, label]) => `<span class="chip">${label} <b>${result.summary[key]}</b></span>`).join('');
    $('compare-notes').textContent = result.notes.join(' ') + ' 各项统计可能重叠，不相加计算。';
    $('compare-sheets').innerHTML = result.sheets.map((sheet, index) => `<tr data-comparison-sheet="${index}"><td>${esc(sheet.old_name || '—')}</td><td>${esc(sheet.new_name || '—')}</td><td>${esc(sheet.pairing)}<small> ${esc(sheet.status)}</small></td><td>${sheet.summary.total}</td><td>${sheet.summary.format}</td><td>${sheet.summary.review}${sheet.review_pairing ? '；配对需确认' : ''}${sheet.metadata.length ? '；属性变化' : ''}</td></tr>`).join('');
    $('compare-sheet').innerHTML = result.sheets.map((sheet, index) => `<option value="${index}">${esc(sheet.old_name || '—')} → ${esc(sheet.new_name || '—')}</option>`).join('');
    for (const [id, book] of [['pair-old', reports[0]], ['pair-new', reports[1]]]) {
      $(id).innerHTML = book.sheets.map(sheet => `<option value="${esc(sheet.name)}">${esc(sheet.name)}</option>`).join('');
    }
    $('pair-status').textContent = Object.entries(mapping).map(([newName, oldName]) => `${oldName} → ${newName}`).join('；') || '尚无手工配对。';
    comparisonPage = 0;
    renderComparison();
  }

  async function compareVersions() {
    if (running) return;
    const oldFile = versionFiles.old || $('old-file').files[0];
    const newFile = versionFiles.new || $('new-file').files[0];
    const font = $('font').value.trim(), size = +$('size').value;
    if (!oldFile || !newFile) { $('compare-error').textContent = '请选择上一版和这一版文件'; return; }
    if (!font || !Number.isFinite(size) || size <= 0 || size > 409) { $('compare-error').textContent = '请填写有效的目标字体和字号，用于保留的七项检查'; return; }
    if ([oldFile, newFile].some(file => file.size > 30 * 1024 * 1024 || !/\.(xlsx|xlsm)$/i.test(file.name))) { $('compare-error').textContent = '每份文件需为 xlsx / xlsm，且不超过 30 MB'; return; }
    busy(true);
    $('cancel').disabled = true;
    $('compare-export').disabled = true;
    $('compare-output').hidden = true;
    result = null;
    $('compare-error').textContent = '';
    $('compare-status').textContent = '正在读取和比较两个版本，请稍候…';
    try {
      const response = await fetch('/api/compare', {
        method:'POST', headers:{'Content-Type':'application/json','X-Inspector-Token':document.querySelector('meta[name=inspector-token]').content},
        body:JSON.stringify({old:{name:oldFile.name,data:await base64(oldFile)},new:{name:newFile.name,data:await base64(newFile)},font,size,mapping})
      });
      const payload = await response.json();
      if (!response.ok) throw Error(payload.error || '比较失败');
      showResult(payload);
      $('compare-status').textContent = `${payload.comparison.old_file} → ${payload.comparison.new_file}：比较完成。请先看总览，再审查需确认项目。`;
    } catch (error) {
      $('compare-error').textContent = error.message;
      $('compare-status').textContent = '比较未完成，请检查文件后重试。';
    } finally { busy(false); }
  }

  $('compare-run').onclick = compareVersions;
  function selectVersion(version, files) {
    if (running) return;
    if (files.length !== 1) {
      $('compare-error').textContent = '每个版本区域请只放入一个 Excel 文件';
      return;
    }
    const file = files[0];
    if (!/\.(xlsx|xlsm)$/i.test(file.name) || file.name.startsWith('~$')) {
      $('compare-error').textContent = '请选择 .xlsx / .xlsm 文件，不支持文件夹或 Excel 临时文件';
      return;
    }
    if (file.size > 30 * 1024 * 1024) {
      $('compare-error').textContent = '文件超过 30 MB，请选择较小的工作簿';
      return;
    }
    versionFiles[version] = file;
    $(version + '-file-name').textContent = `${file.name} · ${(file.size / 1024 / 1024).toFixed(2)} MB`;
    $(version + '-drop').classList.add('has-file');
    mapping = Object.create(null); result = null; comparisonReports = [];
    $('compare-output').hidden = true;
    $('compare-export').disabled = true;
    $('compare-error').textContent = '';
    $('compare-status').textContent = '文件选择已变化，请点击“比较两个版本”。';
  }

  for (const version of ['old', 'new']) {
    const input = $(version + '-file'), zone = $(version + '-drop');
    let dragDepth = 0;
    input.onchange = () => {
      if (input.files.length) selectVersion(version, Array.from(input.files));
      // Store the File separately so clicking can select the same file again after a drop.
      input.value = '';
    };
    zone.addEventListener('dragenter', event => {
      event.preventDefault();
      if (!running) { dragDepth++; zone.classList.add('drag-over'); }
    });
    zone.addEventListener('dragover', event => {
      event.preventDefault();
      if (event.dataTransfer) event.dataTransfer.dropEffect = running ? 'none' : 'copy';
    });
    zone.addEventListener('dragleave', event => {
      event.preventDefault();
      dragDepth = Math.max(0, dragDepth - 1);
      if (!dragDepth) zone.classList.remove('drag-over');
    });
    zone.addEventListener('drop', event => {
      event.preventDefault();
      dragDepth = 0; zone.classList.remove('drag-over');
      if (running) return;
      const items = Array.from(event.dataTransfer?.items || []);
      if (items.some(item => item.webkitGetAsEntry?.()?.isDirectory)) {
        $('compare-error').textContent = '请拖入 Excel 文件，不能拖入文件夹';
        return;
      }
      selectVersion(version, Array.from(event.dataTransfer?.files || []));
    });
  }
  // Prevent the browser from navigating away when a file misses its drop zone.
  for (const name of ['dragover', 'drop']) document.addEventListener(name, event => {
    if (Array.from(event.dataTransfer?.types || []).includes('Files')) event.preventDefault();
  });
  $('pair-apply').onclick = () => {
    const newName = $('pair-new').value, oldName = $('pair-old').value;
    for (const [key, value] of Object.entries(mapping)) if (value === oldName) delete mapping[key];
    mapping[newName] = oldName;
    compareVersions();
  };
  $('pair-reset').onclick = () => { mapping = Object.create(null); compareVersions(); };
  for (const id of ['compare-sheet', 'compare-filter', 'compare-search']) $(id).addEventListener(id === 'compare-search' ? 'input' : 'change', () => { comparisonPage = 0; renderComparison(); });
  $('compare-prev').onclick = () => { comparisonPage = Math.max(0, comparisonPage - 1); renderComparison(); };
  $('compare-next').onclick = () => { comparisonPage++; renderComparison(); };
  $('compare-sheets').onclick = event => {
    const row = event.target.closest('[data-comparison-sheet]');
    if (row) { $('compare-sheet').value = row.dataset.comparisonSheet; comparisonPage = 0; renderComparison(); }
  };
  $('compare-rows').onclick = event => {
    const button = event.target.closest('[data-open-version]');
    if (!button || !result) return;
    const version = +button.dataset.openVersion;
    const sheet = result.sheets[+$('compare-sheet').value];
    reports = comparisonReports;
    $('book').innerHTML = reports.map((book, index) => `<option value="${index}">${index ? '这一版' : '上一版'} · ${esc(book.name)}</option>`).join('');
    $('book').value = String(version); changeBook();
    const name = version ? sheet.new_name : sheet.old_name;
    $('sheet').value = String(reports[version].sheets.findIndex(item => item.name === name));
    const row = button.closest('tr');
    const address = row.children[version ? 2 : 1].querySelector('b').textContent;
    $('search').value = `地址:${address}`; $('marked').checked = false; page = 0; render();
    $('results').querySelector(`[data-cell="${address}"]`)?.click();
    $('results').scrollIntoView({behavior:'smooth',block:'center'});
  };
  $('compare-export').onclick = () => {
    if (!result) return;
    const url = URL.createObjectURL(new Blob([JSON.stringify(result, null, 2)], {type:'application/json'}));
    const link = document.createElement('a'); link.href = url; link.download = 'excel-version-comparison.json'; link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };
})();
