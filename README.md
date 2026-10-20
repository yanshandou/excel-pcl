# Excel Inspector

本机运行的 Excel 检查工具：HTML / JavaScript + Python + openpyxl。不需要安装 Excel，不修改原文件，不向外部服务发送工作簿。

## Windows 启动

安装 Python 3.10 或更新版本，解压后双击 `start.bat`。首次运行需要联网安装 openpyxl，随后会打开 http://127.0.0.1:8765 。命令行窗口需保持打开；关闭窗口或 Ctrl+C 停止服务。

也可手动运行：

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
python -m pip install -r requirements.txt
python app.py
```

端口冲突时：`python app.py --port 8766`。

## 使用

1. 选择多个文件，或选择文件夹（包含子文件夹）。后一次选择替换前一次选择。
2. 填写目标字体和字号，点击开始检查。每个文件单独处理，失败文件不会中断后续文件。
3. 选择文件和工作表，切换七项规则；搜索、仅看有颜色、点击单元格查看属性。
4. 可切换当前页网格，或导出完整 JSON。导出数据可能包含原文件内容。

默认字体是建议值，原 VBA 从操作表读取字体、字号，源码没有固定值；请按原工具配置调整。

## 已实现与差异

- 七种分类的优先级按 `reference/original.vba` 移植；蓝色不一定异常，红色也不一定异常。
- 字体：字号匹配标蓝；字体不匹配覆盖为红；粗体/非黑色显示网纹。
- 格式：仅 `0_`、`#,##0_`、`0.0_` 标绿，其他非标准/文本格式标黄；没有自行扩大数字格式范围。
- 垂直 distributed 按代码标蓝，虽然 VBA 说明文字称“其他红”。
- 提取文档标题、工作表、打印范围、名称定义及 DrawingML 图形名称。图形按 XML 部件列出，暂未映射回工作表；VML/ActiveX/嵌入对象等没有完整覆盖。
- 打印范围外依照 VBA 只比较已用区域的右下角；不检查左/上边界，多区域为“不明”。这是原规则限制，不代表严格的内容越界检测。
- 稀疏扫描有数据或样式的单元格，不铺满整个 UsedRange，也不额外生成外围 10 行/列的空单元格。Excel COM 的 UsedRange、继承样式、合并区从属格与 openpyxl 可能不同。
- 预览不复刻 Excel 的列宽行高、边框、原背景或图形；保留合并区域信息，网格仅展示当前页。
- NumberFormatLocal 在 openpyxl 中没有等价接口，使用 OOXML number_format。公式不重算，条件格式不计算；主题颜色识别是近似。
- 最大单文件 30 MB，解压后总计 150 MB，每表已存储单元格 250000 个。一次仅处理一个文件，结果保存在浏览器内存，刷新清空；大批次需分批。
- 原文件缺少真实测试样本，因此尚未证明与 Excel 原工具逐项等价。

## 验证

```bash
python -m unittest -v
```

测试覆盖规则优先级、换行优先、合并、公式、格式分类、打印边界、无打印区域、多打印区域。此版本还未完成真实 Windows 浏览器人工验收。
