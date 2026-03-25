# Latex2Word 使用说明

## 1. 工具目标

本项目用于把 LaTeX 工程转换为可编辑的 Word 文档，并尽量保留：

- 标题层级与目录基础
- 图片/表格及题注
- 公式与公式引用
- 参考文献
- 图表公式交叉引用可维护性

这是“高保真初稿流水线”，不是“零人工无损转换器”。  
最终交付前仍需要在 Word 里做人工校对与收口。

## 2. 当前目录与入口

仓库根目录关键结构：

```text
Latex2Word/
  src/
    scripts/
      precheck.py
      normalize_tex.py
      convert_with_pandoc.py
      postcheck_docx.py
      build_manual_fix_list.py
    templates/
      reference.docx
  doc/
  examples/
  readme.md
```

说明：

- 当前仅维护 Python 入口，不再使用 `.sh` 入口。
- 推荐在仓库根目录执行命令：`python src/scripts/<script>.py ...`

## 3. 环境准备

最小依赖：

- Python 3.9+
- Pandoc（命令行可用）
- Microsoft Word（用于人工终审）

建议前提：

- 原始 LaTeX 工程可以正常编译 PDF（便于比对和定位问题）。

## 4. `reference.docx` 样式约定

样式模板路径：

- `src/templates/reference.docx`

建议至少配置：

- 标题样式（Heading 1/2/3）及多级编号
- 图题样式（如 `ImageCaption`）
- 表题样式（如 `TableCaption`）
- 显示公式段落样式（`MTDisplayEquation`）
- 三线表样式（表格样式，类型为 table）

当前脚本行为：

- 会自动识别 `MTDisplayEquation` 并应用到显示公式段落。
- 会优先识别三线表风格（样式名/ID含“`三线表` / `threeline` / `booktabs`”）并应用到普通表格。

## 5. 标准五步流程（手动执行）

下面以 Windows 为例，假设：

- LaTeX 工程：`D:\work\my-paper`
- 主文件：`book.tex`
- 工作目录：`D:\work\my-paper__latex_to_word_work`

### 第一步：预检查

```powershell
python src/scripts/precheck.py `
  --project-root "D:/work/my-paper" `
  --main-tex book.tex `
  --work-root "D:/work/my-paper__latex_to_word_work"
```

主要输出：

- `stage_precheck/precheck-report.json`
- `stage_precheck/precheck-report.md`

判定：

- `PASS` / `PASS_WITH_WARNINGS`：可继续
- `FAIL`：先修复阻塞问题

### 第二步：规范化

```powershell
python src/scripts/normalize_tex.py `
  --project-root "D:/work/my-paper" `
  --main-tex book.tex `
  --work-root "D:/work/my-paper__latex_to_word_work" `
  --force
```

主要输出：

- `stage_normalize/source_snapshot/`（规范化副本）
- `stage_normalize/normalization-report.json`
- `stage_normalize/normalization-report.md`

### 第三步：Pandoc 主转换 + DOCX 后处理

```powershell
python src/scripts/convert_with_pandoc.py `
  --work-root "D:/work/my-paper__latex_to_word_work"
```

主要输出：

- `stage_convert/output.docx`
- `stage_convert/pandoc-conversion.log`
- `stage_convert/pandoc-conversion-report.json`
- `stage_convert/pandoc-conversion-report.md`
- `stage_convert/docx-postprocess-report.json`

### 第四步：DOCX 后检查

```powershell
python src/scripts/postcheck_docx.py `
  --work-root "D:/work/my-paper__latex_to_word_work"
```

主要输出：

- `stage_postcheck/postcheck-report.json`
- `stage_postcheck/postcheck-report.md`

### 第五步：生成人工修复清单（并发布用户视图）

```powershell
python src/scripts/build_manual_fix_list.py `
  --work-root "D:/work/my-paper__latex_to_word_work"
```

主要输出：

- `stage_checklist/manual-fix-checklist.json`
- `stage_checklist/manual-fix-checklist.md`
- `README_RUN.md`
- `deliverables/`
- `reports/`
- `logs/`
- `debug/`

## 6. 结果目录（当前实现）

完整流程完成后，工作目录根部通常如下：

```text
<work-root>/
  stage_precheck/
  stage_normalize/
  stage_convert/
  stage_postcheck/
  stage_checklist/
  deliverables/
  reports/
  logs/
  debug/
  README_RUN.md
  run-manifest.json
```

建议用户优先查看：

- `deliverables/output.docx`
- `deliverables/manual-fix-checklist.md`
- `README_RUN.md`

## 7. 推荐人工收口顺序

1. 在 Word 中全选并更新全部字段（目录、交叉引用、编号）。
2. 校对标题层级与目录。
3. 校对图题、表题、公式编号与引用跳转。
4. 校对参考文献与引文链接。
5. 校对复杂表格、算法、子图等高风险区域。

## 8. 常见问题

### Q1：主转换失败

优先看：

- `stage_convert/pandoc-conversion.log`

常见原因：

- Pandoc 未安装或 PATH 不可用
- 资源路径过长/资源目录过多
- 源工程存在 Pandoc 不支持结构

### Q2：转换成功但格式不理想

先看：

- `stage_postcheck/postcheck-report.md`
- `stage_checklist/manual-fix-checklist.md`

然后按清单修复，不建议直接通篇手改。

### Q3：想切换模板样式后重新生成

更新 `src/templates/reference.docx` 后，重新执行第 3-5 步。  
若改动较大，建议从第 1 步完整重跑。

## 9. 一页命令版

```powershell
python src/scripts/precheck.py --project-root "D:/work/my-paper" --main-tex book.tex --work-root "D:/work/my-paper__latex_to_word_work"
python src/scripts/normalize_tex.py --project-root "D:/work/my-paper" --main-tex book.tex --work-root "D:/work/my-paper__latex_to_word_work" --force
python src/scripts/convert_with_pandoc.py --work-root "D:/work/my-paper__latex_to_word_work"
python src/scripts/postcheck_docx.py --work-root "D:/work/my-paper__latex_to_word_work"
python src/scripts/build_manual_fix_list.py --work-root "D:/work/my-paper__latex_to_word_work"
```
