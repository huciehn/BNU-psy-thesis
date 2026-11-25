# BNU-psy-thesis

## 环境与技术栈

- 项目位置：当前技能目录；Codex 可通过技能链接调用。
- Python 3.9+、python-docx、lxml、requests；每次任务检查 GitHub Releases，网络失败时使用本地版。校内检测需要用户另行提供机构账号和邮箱。
- 规范：北师大心理学部用户提供的 2021 修订材料；2025 排版样例仅作辅助。
- 本次改造、来源快照、原技能备份、验证记录全部保存在本目录。
- `_local/` 为本机私有工作区，不进入开源包；公开包仅使用白名单文件。

## 常用命令

在项目目录运行，路径有空格时加引号：

```powershell
python -m pip install -r requirements.txt
python scripts/inventory_docx.py input.docx --out project/inventory.json
python scripts/format_docx.py input.docx --plan project/plan.json --output project/formatted.docx --report project/changes.json
python scripts/validate_docx.py project/formatted.docx --plan project/plan.json --report project/validation.json
python scripts/render_office.py project/formatted.docx --plan project/plan.json --backend wps --output-dir project/render
python -m unittest discover -s tests -v
python scripts/package_skill.py --out dist
```

## 特有工作流

先清点全部内容并审阅角色计划，再按机器规则格式化；输入不覆盖、正文不重写。
对照 `references/requirements.md` 与 `references/profile.json`，完成结构检查、渲染后的逐页检查和引文核对后再交付。
目录、页码、字体替代和跨页表格需要排版引擎实际渲染，不能只凭 XML 宣布合格。
源文件原样保存在 `_local/source-audit/`，哈希与段落索引用于追溯。

## 操作记录

- 2026-10-14：备份原 SKILL.md；导入 8 份来源并提取证据；启动基于源规范的规则、保真格式引擎、检查器和开源包改造。
- 确认归档目录为技能实际安装目录；后续验证记录继续追加在此项目。
- 核实中文参考文献字段斜体、显著性星号保留、有效数字建议与样例差异；形成公开规则与来源哈希。
- 新增保真OOXML引擎、角色清点、结构化生成、独立检查器、原生渲染和白名单打包脚本；不依赖私人账号或绝对路径。
- 回归覆盖公式/图片/符号/批注/超链接保真、字符跨度斜体、表单与数据表分流、哈希绑定、失效复核拒绝、坏格式识别。
- 实际WPS验证修复目录复制标题直接格式及物理页序偏移；15页测试文档摘要/目录/正文/参考文献/附录/致谢落在3/5/7/9/11/13/15页，正文编号1/3/5/7保持连续节续编。
- 官方样例独立前向试用结构检查0错误，因内容、封面、引文缺项保留待复核；未把示例当作可提交毕业论文。
- 最终9项公开行为回归通过；技能元数据验证通过；最新15页结构样例与官方样例原生保存后的结构均0错误，内容保真。
- 交付采用白名单生成`dist/BNU-psy-thesis.skill`；开源包排除原始校方文件、原技能备份、测试稿、日志、截图、私人目录和账号。
- 发布包在另一新目录解压后的9项回归全部通过；21个公开文件、60,100字节；8份原始知识来源哈希均未变化。最终核验记录为`_local/final-verification.json`。
- 精简入口：备份和安装测试副本的`SKILL.md`改名为`SKILL.snapshot.txt`，避免递归扫描识别重复技能；安装目录只保留一个正式入口。主入口保留流程及验收约束，命令细节集中至`references/plan-and-review.md`。
- 技能开发时不能仅依赖.gitignore屏蔽入口；安装目录内的备份、解压测试副本都使用非SKILL.md文件名。旧发布包及精简前入口保留在`_local/before-compact-<时间>/`，不删除原记录。
- 2026-11-25：接入校内系统实测规则；增加真实空行、表格居中与宽度缩放、图片等比例缩放、图表题空格及3–5关键词检查。11项回归通过，15页合成论文WPS渲染和最终DOCX结构检查通过；校内复测仅缺官方封面。真实文档、账号、网站报告和模型复测日志均留在工作区或`_local/`，公开包保持白名单。
- 2026-11-25：完成英文原文《The Development of Subitizing in Bilingual Children》及中文译稿的端到端验证。新增封面主标题/题目标签角色和检测用附录格式；33页最终DOCX的网站十个检测部分均无错误，12项回归通过。Claude Code成功生成0结构错误DOCX；Codex、Grok Build、OpenCode因各自运行环境失败而无产物。
- 2026-11-25：新增 GitHub Releases 自动更新、校内检测配置与授权上传、官方样例封面提取脚本。更新先校验并测试，旧版自动备份；账号和邮箱保存到本机密钥目录。校方DOCX仍不进入公开包。
