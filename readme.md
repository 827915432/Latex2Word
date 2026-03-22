# LaTeX 转 Word 用户使用手册（简版）

=======================

1\. 这套工具是做什么的
-------------

这套工具用于把一个 **LaTeX 工程** 转成 **可编辑的 Word 文档（.docx）**，并尽量保留：

*   图片与图题
*   表格与表题
*   数学公式
*   标题层级与目录基础
*   参考文献
*   图表公式等交叉引用的可恢复结构

它不是“一键完全无损转换器”。它的工作方式是：

1.  先检查工程有没有明显问题
2.  再生成一个规范化副本
3.  用 Pandoc 完成主转换
4.  再检查 Word 结果
5.  最后生成一份**人工修复清单**

也就是说，目标是把人工工作压缩到少量高风险点，而不是完全消灭人工修正。

* * *

2\. 适用对象
--------

适合以下场景：

*   学术论文
*   技术报告
*   多文件 LaTeX 工程
*   含图片、表格、公式、参考文献的文档
*   需要交付 Word 初稿并做后续编辑的情况

不适合的情况：

*   工程里有大量极复杂自定义宏、极复杂 TikZ、极复杂表格，但又要求一次零人工修复
*   希望完全复刻 PDF 的排版细节，而不关心 Word 的可编辑性

* * *

3\. 使用前准备
---------

### 3.1 你的系统

本手册默认你使用的是：

*   **Windows 11**

### 3.2 需要安装的软件

你本机至少要有：

*   **Python 3**
*   **Pandoc**
*   **Word**
*   一个能运行 `.sh` 的 Bash 环境  
    例如：
    *   Git Bash
    *   MSYS2
    *   WSL

### 3.3 skill 目录

你已经创建好的目录应当类似这样：

```
latex-to-word/
├─ SKILL.md
├─ scripts/
│  ├─ precheck.py
│  ├─ normalize_tex.py
│  ├─ convert_with_pandoc.sh
│  ├─ postcheck_docx.py
│  └─ build_manual_fix_list.py
├─ templates/
│  └─ reference.docx
├─ rules/
│  ├─ supported_envs.md
│  ├─ downgrade_policy.md
│  └─ acceptance_criteria.md
└─ examples/
   └─ sample_project/
```

### 3.4 准备 LaTeX 工程

你的 LaTeX 工程至少应包含：

*   主文件，例如 `main.tex`
*   所有子文件
*   图片资源
*   `.bib`
*   自定义宏文件（如有）

建议保证工程本身能正常编译出 PDF，再开始转 Word。

* * *

4\. 整体使用流程
----------

用户视角下，标准流程只有五步：

1.  **预检查**
2.  **规范化**
3.  **Pandoc 主转换**
4.  **docx 后检查**
5.  **生成人工修复清单**

你不需要跳步，也不要直接从第三步开始。

* * *

5\. 具体操作步骤
----------

下面假设：

*   LaTeX 工程目录为  
    `D:\work\my-paper`
*   skill 目录为  
    `D:\tools\latex-to-word`

你可以先进入 skill 目录，再执行各脚本。

* * *

### 第一步：预检查

作用：

*   检查主文件是否存在
*   检查 `\input / \include` 链是否闭合
*   检查图片是否缺失
*   检查 `.bib` 是否存在
*   检查明显的 `\ref / \cite` 问题
*   检查高风险对象，如复杂表格、TikZ、算法、代码块等

命令：

```
python scripts/precheck.py --project-root "D:/work/my-paper" --main-tex main.tex
```

如果你的主文件不是 `main.tex`，把它改成实际文件名。

输出文件默认在工程目录下：

*   `precheck-report.json`
*   `precheck-report.md`

你需要先看这一步结果。

### 预检查结果怎么理解

*   `PASS`：可以继续
*   `PASS_WITH_WARNINGS`：可以继续，但后面要重点核查
*   `FAIL`：先不要继续，先修复错误

如果这里已经 `FAIL`，后面即使勉强跑通，也不应相信结果质量。

* * *

### 第二步：规范化

作用：

*   生成一个**独立工作副本**
*   不修改原始工程
*   对低风险、确定性高的地方做保守整理
*   让 Pandoc 更稳定

命令：

```
python scripts/normalize_tex.py --project-root "D:/work/my-paper" --main-tex main.tex --force
```

说明：

*   `--force` 表示如果工作目录已经存在，就删除后重建
*   规范化结果不会覆盖原始工程

默认会生成一个类似这样的工作目录：

```
D:\work\my-paper__latex_to_word_work
```

输出文件在工作目录下：

*   `normalization-report.json`
*   `normalization-report.md`

这一步结束后，后续操作都围绕 **工作目录** 展开，而不是原始工程目录。

* * *

### 第三步：Pandoc 主转换

作用：

*   调用 Pandoc 完成 LaTeX → docx 主转换
*   使用 `reference.docx` 作为 Word 样式模板
*   输出 Word 初稿与转换日志

命令：

```
bash scripts/convert_with_pandoc.sh --work-root "D:/work/my-paper__latex_to_word_work"
```

如果你在 Windows 上用的是 Git Bash，路径建议尽量写成正斜杠形式。

输出文件默认在工作目录下：

*   `output.docx`
*   `pandoc-conversion.log`
*   `pandoc-conversion-report.json`
*   `pandoc-conversion-report.md`

如果这里失败，先看：

*   `pandoc-conversion.log`

不要直接跳到 Word 里硬修。

* * *

### 第四步：docx 后检查

作用：

*   检查生成的 Word 文档结构是否基本可用
*   检查图片、表格、标题、公式、内部链接、字段等
*   判断是否达到“可交付初稿”标准

命令：

```
python scripts/postcheck_docx.py --work-root "D:/work/my-paper__latex_to_word_work"
```

输出文件默认在工作目录下：

*   `postcheck-report.json`
*   `postcheck-report.md`

### 后检查结果怎么理解

*   `PASS`：结果整体较稳定
*   `PASS_WITH_WARNINGS`：能用，但还有若干重点问题
*   `FAIL`：当前 docx 不满足最低可用标准

这一步是判断“当前 Word 初稿能不能继续修”的关键依据。

* * *

### 第五步：生成人工修复清单

作用：

*   汇总前四步结果
*   形成一份真正面向用户执行的修复列表
*   告诉你应该先修什么、后修什么

命令：

```
python scripts/build_manual_fix_list.py --work-root "D:/work/my-paper__latex_to_word_work"
```

输出文件默认在工作目录下：

*   `manual-fix-checklist.json`
*   `manual-fix-checklist.md`

这一步完成后，你真正需要打开的两个核心文件通常是：

*   `output.docx`
*   `manual-fix-checklist.md`

* * *

6\. 用户实际修文档时的顺序
---------------

不要一打开 Word 就从第一页往后乱查。  
正确顺序是：

### 先做全局动作

1.  打开 `output.docx`
2.  全选全文
3.  更新全部字段

### 再按清单修

优先顺序建议如下：

1.  目录和标题层级
2.  图题、表题
3.  图表公式交叉引用
4.  参考文献
5.  复杂表格
6.  子图、算法、代码块
7.  局部版式微调

### 最后做终审

重点抽查：

*   关键图是否都在
*   关键表是否可读
*   关键公式是否还能编辑
*   文内引用是否对应正确
*   文末参考文献是否完整

* * *

7\. 每一步的输入和输出关系
---------------

你可以把它记成下面这条链：

### 输入

原始 LaTeX 工程

### 第一步输出

*   `precheck-report.*`

### 第二步输出

*   规范化工作目录
*   `normalization-report.*`

### 第三步输出

*   `output.docx`
*   `pandoc-conversion-report.*`
*   `pandoc-conversion.log`

### 第四步输出

*   `postcheck-report.*`

### 第五步输出

*   `manual-fix-checklist.*`

最终用户主要看：

*   `output.docx`
*   `manual-fix-checklist.md`

* * *

8\. 常见问题
--------

### 8.1 预检查失败怎么办

先看 `precheck-report.md`。

常见原因：

*   主文件写错
*   子文件缺失
*   图片缺失
*   `.bib` 缺失
*   工程本身引用系统不完整

这类问题应该优先回源工程修，而不是继续硬转。

* * *

### 8.2 Pandoc 转换失败怎么办

先看：

*   `pandoc-conversion.log`

常见原因：

*   本机没有安装 Pandoc
*   路径问题
*   LaTeX 工程中存在 Pandoc 处理不了的结构
*   bibliography 或图片路径有问题

* * *

### 8.3 Word 打开了，但很多东西不对怎么办

先不要立即大面积手改。

先做三件事：

1.  看 `postcheck-report.md`
2.  看 `manual-fix-checklist.md`
3.  在 Word 中先更新全部字段

很多目录、编号、交叉引用类问题，更新字段后会改善一部分。

* * *

### 8.4 复杂表格还是乱了怎么办

这是正常高风险项之一。

处理原则：

*   普通表格应该尽量自动正确
*   极复杂表格允许人工修
*   如果复杂表格特别关键，且 Word 中已经严重变形，通常更值得回源工程简化结构后重新转换

* * *

### 8.5 TikZ 图变成图片是不是错误

不一定。

按照这套规则，**极复杂 TikZ 图允许降级为图片**。  
前提是：

*   图内容还在
*   图题还在
*   报告里有记录
*   你知道它已经失去 Word 原生可编辑性

* * *

9\. 最推荐的使用习惯
------------

### 建议 1

先确认 LaTeX 工程本身能正常编译 PDF，再开始转 Word。

### 建议 2

每次都从第一步开始跑完整流水线，不要跳步骤。

### 建议 3

不要直接在原始工程目录里乱改，让规范化工作目录承接中间产物。

### 建议 4

遇到源工程层面的错误，优先回源修，而不是在 Word 中补丁式硬修。

### 建议 5

把这条流水线当成“生成高保真 Word 初稿”的工具，而不是“完全自动替代人工校对”的工具。

* * *

10\. 一页版快速操作
------------

如果你只想记住最核心的命令，记下面这五条就够了。

### 1）预检查

```
python scripts/precheck.py --project-root "D:/work/my-paper" --main-tex main.tex
```

### 2）规范化

```
python scripts/normalize_tex.py --project-root "D:/work/my-paper" --main-tex main.tex --force
```

### 3）主转换

```
bash scripts/convert_with_pandoc.sh --work-root "D:/work/my-paper__latex_to_word_work"
```

### 4）后检查

```
python scripts/postcheck_docx.py --work-root "D:/work/my-paper__latex_to_word_work"
```

### 5）生成人工修复清单

```
python scripts/build_manual_fix_list.py --work-root "D:/work/my-paper__latex_to_word_work"
```

最后打开：

*   `D:/work/my-paper__latex_to_word_work/output.docx`
*   `D:/work/my-paper__latex_to_word_work/manual-fix-checklist.md`

