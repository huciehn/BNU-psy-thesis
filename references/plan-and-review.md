# 计划与验收文件

## 执行命令

`S`表示技能目录，`P`表示用户论文项目目录；执行时换成实际路径。产物全部留在P，迭代使用新文件名。

```text
python "S/scripts/inventory_docx.py" "P/input.docx" --out "P/inventory.json"
python "S/scripts/format_docx.py" "P/input.docx" --plan "P/plan.json" --output "P/formatted.docx" --report "P/changes.json"
python "S/scripts/validate_docx.py" "P/formatted.docx" --plan "P/plan.json" --original "P/input.docx" --report "P/validation.json"
python "S/scripts/render_office.py" "P/formatted.docx" --backend wps --plan "P/plan.json" --output-dir "P/render"
```

每次任务先运行 `python "S/scripts/update_skill.py" --auto`。发布配置在`references/update.json`，更新包必须附带同名技能压缩包和manifest.json，哈希及回归测试通过后才覆盖，旧版自动备份。

首次校内检测运行 `python "S/scripts/site_check.py" --setup`，由用户输入账号、邮箱和入口。登录状态用 `--check` 验证。获得用户对当前文件的明确上传授权后运行：

```text
python "S/scripts/site_check.py" --upload "P/final.docx" --report-dir "P/site-report" --language 中文 --yes
```

新写作需要封面时，默认使用`assets/cover-template.docx`首页模板，用户也可提供新版官方样例替换。以下命令提取首页校徽并在内容文件前加入检测兼容封面；已有封面须审阅后显式使用`--replace-cover`：

```text
python "S/scripts/make_cover.py" --content "P/content.json" --output "P/content-with-cover.json" --replace-cover --title "论文题目" --student-id "12位学号" --student-name "姓名" --advisor "导师姓名" --advisor-title "导师职称"
```

清点后须由代理审阅全文并编制plan，不能直接将suggested_role视为已确认角色。未决修订先由用户决定接受/拒绝范围，保留原件。旧DOC先转新DOCX并核对；扫描件须恢复结构并记录OCR不确定处。

渲染使用Word时将backend改为word。读取render-summary.json的final_directory，对其中updated.docx再次运行检查器（输出新报告）；它与rendered.pdf才是本轮复核对象。LibreOffice需明确可执行路径，目前仅做PDF渲染，另核对DOCX域缓存是否保存。目录缺域时由代理插入真实TOC，不手填页码。

反复修改时回到本轮原始输入调整plan，再生成新版本；不要直接将含旧BNU-psy-thesis定位书签的输出重新输入格式器。文本格式修正须先记录前后文本，在新副本执行，再重新清点并绑定新哈希。

完成逐页、引用和内容复核后，按本文件的记录格式填写review.json，以最终DOCX哈希绑定，执行：

```text
python "S/scripts/validate_docx.py" "P/final.docx" --plan "P/plan.json" --original "P/input.docx" --review "P/review.json" --report "P/acceptance.json"
```

## 段落计划

`inventory_docx.py` 为 DOCX 每一段（含表格、文本框）生成稳定清点编号。建议角色仅用于协助判断；审阅全文后另写plan，不能直接把建议复制成已确认计划。

```json
{
  "profile": "bnu-psychology-2021",
  "input_sha256": "原始DOCX的SHA-256",
  "paragraphs": {
    "p00000": {"role": "cover_title"},
    "p00020": {"role": "heading1"},
    "p00021": {"role": "body", "spans": [{"start": 8, "end": 9, "italic": true}]},
    "p00080": {"role": "reference", "spans": [{"start": 40, "end": 54, "italic": true}]}
  },
  "tables": {"t00000": {"kind": "data", "header_rows": 1}}
}
```

全部有内容段落均需role，数据表需显式分类；空段落可保留。支持的role见profile.styles；data为三线统计表，form为开题表单，preserve保留自定义表布局并单独复核。字符区间是当前段落中纯文本的零基左闭右开范围，按Unicode字符计数，不包含域代码。先检查真实文本，再确定斜体字段或统计变量范围。对复杂超链接/域/公式段不能安全拆run时，引擎拒绝spans；用专门OOXML编辑保持所有非文本子节点，或在排版软件中修改，再重新清点与验证。

引擎添加 `bnu_psy_thesis_p00000` 等书签供检查器定位；这些是内部定位，不能用于伪造内容合格。输入哈希不一致说明计划已过期，重建计划。文本内容修改与格式修改分开：引擎不重写文本。如确需修改标题空格、标点等，记录明确的修改清单，用新副本修改，然后把该副本作为本次排版的原始输入，保存前后对照。

## 检查器

```text
python scripts/validate_docx.py final.docx --plan plan.json --original input.docx --report validation.json
```

返回0表示完整验收通过，1表示结构/内容保真错误，2表示仍待复核。`--structural-only`只适用于开发测试：可在结构通过时返回0，但report.ready_to_submit仍为false。`--mode partial`只能验证片段，不能宣称完整本科论文合格；`--mode proposal`不要求毕业论文前置部分。

## 新写作的内容文件

仅把用户确认的真实内容放入blocks，不自动补研究内容：

```json
{
  "blocks": [
    {"role": "cover_title", "text": "用户给定论文题目"},
    {"role": "toc_title", "text": "目  录"},
    {"type": "toc"},
    {"role": "heading1", "text": "1  前言"},
    {"role": "body", "text": "经过核对的正文内容。"},
    {"type": "table", "kind": "data", "header_rows": 1, "rows": [["变量", "结果"], ["用户提供的变量", "用户提供的真实值"]]}
  ]
}
```

```text
python scripts/build_docx.py content.json --output input.docx --plan plan.json
```

上述只是结构示意，缺完整论文组件时检查器会报缺项。段落支持spans；image需要真实本地path、width_cm；formula接收已有可编辑OMML，不能用未解析的LaTeX冒充Word公式。

## Office更新后的保真

检查器允许Office重建自动目录缓存、重新编号样式及移动内部书签，也允许加入不改变数学表达式的空格式节点；这些需要按内容语义比较。排版器显式禁止Office压缩图片，并在Office保存后恢复原稿内嵌图片字节，同时保留更新后的目录和分页。正文、统计数值、公式运算结构、符号和图片字节仍独立核对。不能把重生成目录的正常文本变化误判为正文丢失，也不能用“软件正常化”解释真实数据改变。

## 实际渲染和复核

排版软件更新域可能重建目录缓存。渲染后重新清点缓存条目、核对其字体与段落；必要时再次排版并重渲染。不要以先前的目录缓存或先前截图验证修改后的文件。

复核记录示例，仅在实际完成检查后填写：

```json
{
  "docx_sha256": "最终DOCX的SHA-256",
  "rendered_pdf_sha256": "该版本渲染PDF的SHA-256",
  "reviewer": "实际复核者及日期",
  "checks": {
    "official_cover": true,
    "all_pages_visual": true,
    "physical_odd_pages": true,
    "toc_page_accuracy": true,
    "fonts_available": true,
    "citations_bijective": true,
    "reference_metadata_and_italics": true,
    "statistics_preserved": true,
    "figures_and_tables": true,
    "content_complete": true
  },
  "notes": "页数、核对方式、已修问题；无图表时注明为何该项不适用"
}
```

再次执行检查器加 `--review review.json`，生成新的验收报告。复核者可以是实际完成检查的代理或人；不得把上述示例作为全true默认值。原稿缺内容、封面或引用信息，就维持待复核/缺项状态并说明所需材料。
