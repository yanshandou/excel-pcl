"""Compare inspected workbook versions; uncertain matches stay reviewable."""
from collections import defaultdict
from bisect import bisect_left
from difflib import SequenceMatcher
from openpyxl.worksheet.cell_range import CellRange

FIELDS = {
    'format': '显示格式', 'horizontal': '水平对齐', 'vertical': '垂直对齐',
    'wrap': '自动换行', 'shrink': '缩小填充', 'merge_shape': '合并尺寸',
    'font': '字体', 'size': '字号', 'bold': '粗体', 'font_color': '字体颜色',
    'italic': '斜体', 'underline': '下划线', 'strike': '删除线',
    'rotation': '文字旋转', 'indent': '缩进',
}


def token(cell):
    return (cell.get('data_type', 'f' if cell.get('formula') else 's'), cell['value'])


def signature(sheet):
    return tuple((c['row'], c['col'], token(c)) for c in sheet['cells'] if c['value'] != '')


def axis_map(old_cells, new_cells, axis):
    """Align occupied rows/columns using content, preserving gaps in coordinates."""
    other = 'col' if axis == 'row' else 'row'
    def groups(cells):
        result = defaultdict(list)
        for cell in cells:
            if cell['value'] != '':
                result[cell[axis]].append(cell)
        return [(index, tuple(token(c) for c in sorted(values, key=lambda c:c[other])))
                for index, values in sorted(result.items())]
    before, after = groups(old_cells), groups(new_cells)
    matcher = SequenceMatcher(None, [x[1] for x in before], [x[1] for x in after], autojunk=True)
    result = {}
    for tag, a, b, c, d in matcher.get_opcodes():
        if tag == 'equal':
            for left, right in zip(before[a:b], after[c:d]):
                result[right[0]] = (left[0], '内容序列匹配')
        elif tag == 'replace' and b-a == d-c:
            for left, right in zip(before[a:b], after[c:d]):
                result[right[0]] = (left[0], '区段顺序推测')
    return result


def style(cell):
    if cell is None:
        return {}
    result = {key: cell.get(key) for key in FIELDS if key != 'merge_shape'}
    merged = cell.get('merge')
    region = CellRange(merged) if merged else None
    result['merge_shape'] = (region.size['rows'], region.size['columns']) if region else (1, 1)
    return result


def differences(before, after):
    a, b = style(before), style(after)
    return [{'field': key, 'label': label, 'old': a[key], 'new': b[key]}
            for key, label in FIELDS.items() if a[key] != b[key]]


def compare_sheet(before, after, pairing):
    old = {c['address']: c for c in before['cells']}
    new = after['cells']
    old_positions = {(c['row'], c['col']): c for c in before['cells']}
    row_map = axis_map(before['cells'], new, 'row')
    col_map = axis_map(before['cells'], new, 'col')
    # Completely blank axes are matched by coordinate only when dimensions are stable.
    old_rows = {c['row'] for c in before['cells'] if c['value'] != ''}
    new_rows = {c['row'] for c in new if c['value'] != ''}
    old_cols = {c['col'] for c in before['cells'] if c['value'] != ''}
    new_cols = {c['col'] for c in new if c['value'] != ''}
    old_values, new_values = defaultdict(list), defaultdict(list)
    for cell in before['cells']:
        if cell['value'] != '': old_values[token(cell)].append(cell)
    for cell in new:
        if cell['value'] != '': new_values[token(cell)].append(cell)
    used, matches = set(), {}
    # Unique unchanged contents can follow a cell moved across row/column boundaries.
    for cell in new:
        key = token(cell)
        if cell['value'] != '' and len(old_values[key]) == len(new_values[key]) == 1:
            source = old_values[key][0]
            row=row_map.get(cell['row']); col=col_map.get(cell['col'])
            aligned=bool(row and col and row[0]==source['row'] and col[0]==source['col']
                         and row[1]=='内容序列匹配' and col[1]=='内容序列匹配')
            uncertain=source['address']!=cell['address'] and not aligned
            matches[cell['address']] = (source, '唯一内容匹配' + ('（移动对应需确认）' if uncertain else ''), uncertain)
            used.add(source['address'])
    for cell in new:
        if cell['address'] in matches: continue
        row = row_map.get(cell['row'])
        col = col_map.get(cell['col'])
        if row is None and before['rows'] == after['rows'] and cell['row'] not in old_rows|new_rows:
            row = (cell['row'], '空白行坐标参考')
        if col is None and before['cols'] == after['cols'] and cell['col'] not in old_cols|new_cols:
            col = (cell['col'], '空白列坐标参考')
        if row and col:
            source = old_positions.get((row[0], col[0]))
            if source and source['address'] not in used:
                uncertain = '推测' in row[1] or '推测' in col[1]
                if source['address'] != cell['address'] and len(old_values[token(source)]) > 1:
                    uncertain = True
                matches[cell['address']] = (source, f'{row[1]} / {col[1]}', uncertain)
                used.add(source['address'])
    changes = []
    references = defaultdict(list)
    for cell in before['cells']:
        if cell['value'] != '': references[(cell['col'],cell.get('data_type'))].append(cell)
    reference_rows={key:[c['row'] for c in values] for key,values in references.items()}
    for cell in new:
        match = matches.get(cell['address'])
        if match:
            source, basis, uncertain = match
            diffs = differences(source, cell)
            content = token(source) != token(cell)
            moved = source['address'] != cell['address']
            if not (diffs or content or moved): continue
            changes.append({'kind': '修改' if content or diffs else '移动', 'old': source, 'new': cell,
                            'reference': None, 'basis': basis, 'format_changes': diffs,
                            'content_changed': content, 'moved': moved,
                            'review': bool(diffs) or uncertain or pairing not in ('同名工作表','用户指定工作表'),
                            'conclusion': '匹配需确认' if uncertain else '格式变化需确认' if diffs else '沿用原格式'})
        else:
            mapped_col = col_map.get(cell['col'])
            key=(mapped_col[0],cell.get('data_type')) if mapped_col else None
            candidates=references.get(key,[])
            index=bisect_left(reference_rows.get(key,[]),cell['row'])
            nearby=candidates[max(0,index-1):index+1]
            reference=min(nearby,key=lambda c:abs(c['row']-cell['row'])) if nearby else None
            reference_diffs=differences(reference,cell) if reference else []
            changes.append({'kind': '新增或未匹配', 'old': None, 'new': cell, 'reference': reference,
                            'basis': '老版同列、同数据类型的邻近单元格（建议参考）' if reference else '没有可靠对应位置',
                            'format_changes': reference_diffs,
                            'content_changed': True, 'moved': False, 'review': True,
                            'conclusion': '建议参考格式存在差异' if reference_diffs else '需要人工确认'})
    for address, cell in old.items():
        if address not in used:
            changes.append({'kind': '删除或未匹配', 'old': cell, 'new': None, 'reference': None,
                            'basis': '新版中未找到可靠对应位置', 'format_changes': [],
                            'content_changed': True, 'moved': False, 'review': True,
                            'conclusion': '需要人工确认'})
    metadata = []
    for key,label in [('print_area','打印范围'),('state','可见性'),('rows','最大行'),('cols','最大列')]:
        if before[key] != after[key]: metadata.append({'label':label,'old':before[key],'new':after[key]})
    return {'old_name':before['name'], 'new_name':after['name'], 'pairing':pairing,
            'status':'已配对', 'changes':changes, 'metadata':metadata,
            'summary':summarize(changes), 'review_pairing':pairing not in ('同名工作表','用户指定工作表')}


def summarize(changes):
    return {'total':len(changes), 'content':sum(c['content_changed'] for c in changes),
            'format':sum(bool(c['format_changes']) for c in changes),
            'moved':sum(c['moved'] for c in changes), 'review':sum(c['review'] for c in changes),
            'added':sum(c['kind']=='新增或未匹配' for c in changes),
            'deleted':sum(c['kind']=='删除或未匹配' for c in changes)}


def compare_reports(before, after, mapping=None):
    old = {s['name']:s for s in before['sheets']}
    new = {s['name']:s for s in after['sheets']}
    paired, used = {}, set()
    mapping = mapping or {}
    if not isinstance(mapping, dict) or len(set(mapping.values())) != len(mapping):
        raise ValueError('工作表配对必须一一对应')
    for new_name, old_name in mapping.items():
        if new_name not in new or old_name not in old: raise ValueError('配对工作表不存在')
        paired[new_name] = (old[old_name], '用户指定工作表')
        used.add(old_name)
    for name in new:
        if name not in paired and name in old and name not in used: paired[name]=(old[name],'同名工作表');used.add(name)
    # Renamed sheets are paired only if the complete nonempty content signature is unique.
    for name, sheet in new.items():
        if name in paired: continue
        sig = signature(sheet)
        candidates=[s for key,s in old.items() if key not in used and sig and signature(s)==sig]
        competing=[s for key,s in new.items() if key not in paired and signature(s)==sig]
        if len(candidates)==len(competing)==1:
            paired[name]=(candidates[0],'内容相同，疑似工作表改名');used.add(candidates[0]['name'])
    sheets=[]
    for name,sheet in new.items():
        if name in paired:
            source,basis=paired[name];sheets.append(compare_sheet(source,sheet,basis))
        else:
            sheets.append({'old_name':None,'new_name':name,'pairing':'无法自动配对，可手工指定',
                           'status':'新增或未匹配工作表','changes':[],'metadata':[],
                           'summary':summarize([]),'review_pairing':True})
    for name in old:
        if name in used: continue
        sheets.append({'old_name':name,'new_name':None,'pairing':'无法自动配对，可手工指定',
                       'status':'删除或未匹配工作表','changes':[],'metadata':[],
                       'summary':summarize([]),'review_pairing':True})
    all_changes=[c for s in sheets for c in s['changes']]
    return {'old_file':before['name'],'new_file':after['name'],'sheets':sheets,
            'summary':{**summarize(all_changes),'unmatched_sheets':sum(s['status']!='已配对' for s in sheets),
                       'sheet_review':sum(s['review_pairing'] or bool(s['metadata']) for s in sheets)},
            'notes':['差异不等于错误；老版是参考，不是明确的合格标准。',
                     '行列按内容序列和唯一内容匹配；重复内容、同时增删及大幅改版可能无法可靠对应，推测结果需人工确认。',
                     '新增内容只建议参考同列同类型邻近单元格，不自动判定合格。',
                     '比较覆盖单元格内容及七项属性；不比较图形、边框、填充色、行高列宽、条件格式和宏；公式不重算。']}
