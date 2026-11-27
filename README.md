# BNU-psy-thesis

将北师大心理学部本科毕业论文格式要求落实为可追溯规则、现有DOCX保真排版与独立验收。保留中文学术写作、开题报告及课程作业场景，按各场景实际要求选择规则。

2.6版随包提供8份校方公开规范原件、来源路径和哈希，并提供旧版更新器的附件恢复逻辑，第三方可直接核查格式依据。2.5.1版把现有封面和前置声明设为自动保真区域，防止封面段落或分节页面设置被正文规则覆盖。2.5版明确物理奇数页规则，保护诚信承诺书与使用授权书同页，自动保存未排版原稿，并支持按GitHub版本号退回。2.3版加入校内检测系统实测规则、GitHub Releases 自动更新、检测入口配置和官方样例封面提取。真实空行、标题对齐、图表题空格、附录格式及超宽对象缩放均有自动检查。检测兼容参数与正式规范差异记录在 `references/requirements.md`。

## 安装

Python 3.9以上，安装 `requirements.txt`。将整个BNU-psy-thesis目录复制到当前 agent 的技能目录；具体位置由该 agent 的配置或文档决定。只复制SKILL.md会缺少规则、脚本、公开规范和封面模板。`assets/sources/`包含全部8份校方公开规范，`references/sources.json`记录路径和SHA-256，`scripts/ensure_sources.py`负责完整性检查与旧版更新后的附件恢复。

每次执行任务前会检查 `references/update.json` 指定的 GitHub Releases。更新包经过哈希校验和回归测试后覆盖本地文件，只保留最新版本；失败恢复使用临时副本，完成后自动删除。运行 `python scripts/update_skill.py --version v2.4.5` 可退回指定历史版本，把版本号换成目标版本即可。网络不可用时使用本地版本。校内检测为可选步骤，首次使用由用户提供账号和邮箱，凭据保存到系统用户配置目录，Windows 使用 `%APPDATA%`，其他系统使用 `XDG_CONFIG_HOME` 或 `~/.config`，不进入项目目录。

机构规范固定为所附来源版本；其他学校或新版规范须提供自己的profile和来源，不能直接声称通用高校格式。

实际最终分页需要本机Word/WPS，或LibreOffice加逐页复核；中文宋体、黑体和Times New Roman须可用。现有论文封面及前置声明整页保留原稿。新写作默认复制 `assets/cover-template.docx` 的官方首页版式并替换封面字段，保留校徽、下划线和信息栏结构；用户有新版官方样例时可作为 `scripts/make_cover.py` 的模板参数替换。缺失研究材料仍需用户提供。当前开发环境实测WPS原生更新目录、保存DOCX及渲染；Word接口与LibreOffice路径需在各自安装环境进一步验证。

## 使用

对代理说：

> 使用BNU-psy-thesis按北师大心理学部规范优化我的本科论文DOCX。保留论文内容、公式和图表，输出新文件与格式检查报告。

代理读取SKILL.md，清点论文、确认角色、运行排版和检查、更新实际目录并渲染逐页核验。自动排版不等同于学术内容审查，最终检查记录会区分已通过项目与待复核项目。

命令行用法和计划格式见 `references/plan-and-review.md`。规则和来源对照见 `references/requirements.md`。脚本拒绝覆盖输入和已有输出；每次迭代使用新文件名。

## 开源边界

原创代码、技能说明与重新表述的规则采用MIT许可证。8份校方公开DOC/DOCX/PDF及 `assets/cover-template.docx` 首页模板随包提供，权利归原作者或发布单位所有，不因本项目MIT许可证而获得重新许可；详见 `assets/sources/README.md`。原技能备份、本机审计、论文原稿、账号和个人文件保留在 `_local/` 或用户项目目录，不进入发布包。公开来源索引保留文件名、相对路径、哈希和证据位置，未包含本机私人路径和账号。

## 验证与打包

```text
python -m unittest discover -s tests -v
python scripts/package_skill.py --out dist
```

产物为 `dist/BNU-psy-thesis.skill`，本质为可解压ZIP。打包直接替换当前发布文件，不保留本地历史副本。打包使用明确白名单，避免把本机论文、备份和测试产物上传。项目不会自动推送代码；论文只有用户明确授权后才上传至检测入口。
