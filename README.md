# Paper Format Skill

一个面向论文格式规范化的 Word 处理工具，支持按结构化规则批量处理论文文档，并输出：

- 规范化后的 `.docx`
- 严格修改报告 `.csv`
- 红字标注版 `.docx`

当前阶段原生处理 `.docx`，并支持把 `.doc` 自动转换为 `.docx` 后继续处理。

## 当前能力

- 按规则表统一规范标题、正文、摘要、关键词、参考文献等对象
- 支持页眉、表格、段落、字体、字号、行距、缩进等格式重设
- 支持中英混排字体分流
- 支持严格修改报告
- 支持红字标注版输出
- 支持单文件处理和批量处理
- 支持 `.doc` 自动转换后继续规范化

## 目录结构

```text
输入文档/           待处理的 .doc / .docx
规则化文件夹/       当前使用的规则 CSV
输出文档/           处理结果
src/               主程序
rules/             已整理好的规则包
templates/         规则模板
tests/             自动化测试
运行批量规范化.ps1  最简批量入口
```

## 最简使用方式

### 批量处理

1. 把待处理文档放进 `输入文档`
2. 把要使用的规则 CSV 放进 `规则化文件夹`
3. 运行 `运行批量规范化.ps1`
4. 到 `输出文档` 查看结果

输出包括：

- `xxx_规范化.docx`
- `xxx_规范化_修改报告.csv`
- `xxx_规范化_红字标注版.docx`
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

- `输入文档`
- `规则化文件夹`
- `输出文档`

## 规则文件

当前规则目录采用 6 个 CSV：

- `document_rules.csv`
- `numbering_rules.csv`
- `paragraph_rules.csv`
- `report_schema.csv`
- `special_object_rules.csv`
- `table_rules.csv`

仓库里提供了：

- `templates/paper-format-rules/`：通用模板
- `rules/njnu_news_phase1/`：南师大新闻传播方向规则包
- `rules/zhuhuanqin_phase1/`：另一套已整理规则包

## 开发与验证

安装依赖后可直接运行测试：

```powershell
$env:PYTHONPATH='C:\Users\admin\Desktop\JJBand\src'
pytest tests -q
```

当前仓库上传前验证结果为：

```text
62 passed
```

## 边界说明

- `.docx`：原生支持
- `.doc`：自动转 `.docx` 后继续处理
- `PDF`：当前仍未接入自动转换链路

## 说明

这套工具坚持规则驱动和可审计输出，不依赖“看起来像标题”之类的弱启发式判断。对于规则未定义或结构不明确的对象，会优先在报告中显式暴露，而不是偷偷做不可靠修补。
