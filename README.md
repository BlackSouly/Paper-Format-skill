# Paper Format Skill

## 快速入口

- [最简使用流程](</C:/Users/admin/Desktop/JJBand/docs/最简使用流程.md>)
- [表格规则说明](</C:/Users/admin/Desktop/JJBand/docs/表格规则说明.md>)
- [性能基准说明](</C:/Users/admin/Desktop/JJBand/docs/性能基准说明.md>)

### Benchmark

```powershell
$env:PYTHONPATH='C:\Users\admin\Desktop\JJBand\src'
python -c "from paper_format_normalizer.cli import main; main()" benchmark --input <文档路径> --rules C:\Users\admin\Desktop\JJBand\规则化文件夹 --output-dir C:\Users\admin\Desktop\JJBand\输出文档 --repeat 3
```

一个面向论文格式规范化的 Word 处理工具。它基于结构化规则表，对论文文档进行严格、可审计的格式重设，并输出：

- 规范化后的 `.docx`
- 严格修改报告 `.csv`
- 红字标注版 `.docx`

当前原生支持 `.docx`，并支持 `.doc` 自动转换为 `.docx` 后继续处理。

## 当前能力

- 规则驱动规范化标题、正文、摘要、关键词、参考文献等对象
- 支持字体、字号、加粗、行距、缩进、段前段后、页边距等格式重设
- 支持页码格式、页码起始值、页脚页码对齐等分节页码规则
- 支持分节起始方式，如 `new_page / odd_page / even_page / continuous`
- 支持中英文混排字体分流
- 支持页眉和表格对象
- 支持严格修改报告
- 支持红字标注版输出
- 支持单文件处理和批量处理
- 支持 `.doc` 自动转入处理链路

## 表格规则能力

表格规则目前支持这些属性：

- `font_name`
- `font_size`
- `bold`
- `alignment`
- `vertical_alignment`
- `border`

支持的显式定位方式包括：

- 整表
- `header_row_*`
- `body_rows_*`
- `column[n]`
- `column_range[start:end]`
- `column_by_header[表头文本]`
- `row[n]`
- `row_range[start:end]`
- `cell[row,col]`
- `cell_range[row_start:row_end,col_start:col_end]`

详细说明见 [表格规则说明.md](/C:/Users/admin/Desktop/JJBand/docs/表格规则说明.md)。

## 目录结构

```text
输入文档/             待处理的 .doc / .docx
规则化文件夹/         当前使用的规则 CSV
输出文档/             处理结果
src/                 主程序
rules/               已整理好的规则包
templates/           规则模板
tests/               自动化测试
运行批量规范化.ps1    最简批量入口
```

## 最简使用方式

### 批量处理

1. 把待处理文档放进 [输入文档](</C:/Users/admin/Desktop/JJBand/输入文档>)
2. 把规则 CSV 放进 [规则化文件夹](</C:/Users/admin/Desktop/JJBand/规则化文件夹>)
3. 运行 [运行批量规范化.ps1](</C:/Users/admin/Desktop/JJBand/运行批量规范化.ps1>)
4. 到 [输出文档](</C:/Users/admin/Desktop/JJBand/输出文档>) 查看结果

每份文档会生成：

- `xxx_规范化.docx`
- `xxx_规范化_修改报告.csv`
- `xxx_规范化_红字标注版.docx`

另外还会生成：

- `批量规范化汇总.csv`

### 单文件处理

```powershell
$env:PYTHONPATH='C:\Users\admin\Desktop\JJBand\src'
python -c "from paper_format_normalizer.cli import main; main()" normalize --input "C:\Users\admin\Desktop\JJBand\输入文档\你的文档.docx"
```

### 批量命令行处理

```powershell
$env:PYTHONPATH='C:\Users\admin\Desktop\JJBand\src'
python -c "from paper_format_normalizer.cli import main; main()" normalize-batch
```

默认会使用：

- [输入文档](</C:/Users/admin/Desktop/JJBand/输入文档>)
- [规则化文件夹](</C:/Users/admin/Desktop/JJBand/规则化文件夹>)
- [输出文档](</C:/Users/admin/Desktop/JJBand/输出文档>)

## 规则文件

当前规则目录使用 6 个 CSV：

- `document_rules.csv`
- `numbering_rules.csv`
- `paragraph_rules.csv`
- `report_schema.csv`
- `special_object_rules.csv`
- `table_rules.csv`

仓库中提供了：

- `templates/paper-format-rules/`：通用模板
- `rules/njnu_news_phase1/`：南师大新闻与传播方向规则包
- `rules/zhuhuanqin_phase1/`：另一套已整理规则包

## 开发与验证

```powershell
$env:PYTHONPATH='C:\Users\admin\Desktop\JJBand\src'
pytest tests -q
```

当前验证结果：

```text
78 passed
```

## 边界说明

- `.docx`：原生支持
- `.doc`：自动转 `.docx` 后继续处理
- `PDF`：当前仍未接入自动转换链路

## 说明

这套工具坚持规则驱动和可审计输出，不依赖“看起来像标题”之类的弱启发式判断。对规则未定义或结构不明确的对象，优先在报告中显式暴露，而不是偷偷做不可验证的修补。
