# 公文套红网页工具

本工具是本地运行的第一版公文套红操作台，包含规则维护、Word 上传处理、PDF 在线预览和套红 DOCX 下载。

## 启动

```powershell
cd E:\code\wordedit\redhead-web
node server.js --port 3721
```

然后打开：

```text
http://0.0.0.0:3721
```

服务默认监听 `0.0.0.0`，也可以用 `--host 127.0.0.1` 或环境变量 `HOST` 覆盖。

## 运行要求

- Windows
- Microsoft Word
- 已安装公文字体：方正小标宋简体、仿宋_GB2312、楷体等

后端通过 Word COM 生成套红 DOCX 并导出 PDF，因此预览效果以 Word 渲染结果为准。

## 已支持能力

- 规则在线维护，保存时自动生成历史版本
- 规则页已覆盖当前 `rules.json` 的全部业务规则项；`ruleId/revisionId/updatedAt/note` 等系统元数据由服务自动维护
- `data/rule-coverage.json` 维护“规则路径 -> 应用位置 -> 校验方式”的执行矩阵，并可生成 `docs/rule-execution-matrix.md`
- 规则保存前按 `data/rules.schema.json` 做服务端校验，非法颜色、枚举、数值范围和未知规则项会被拒绝
- 可回滚到历史规则版本
- 上传 `.doc/.docx` 后生成套红 DOCX 和 PDF
- `.doc` 会自动转换为 `.docx` 输出；`.docx` 是否处理前重新转换可通过 `cleanup.convertDocToDocx` 维护
- 套红前自动清理源文尾部多余空白页，只保留规范要求的版记前补空白页
- 版记优先放在正文最后一页版心底部；空间不足或偶数页要求不满足时才另起页或补空白页
- 版记印发机关和印发日期使用 4 号仿宋_GB2312、固定 28pt 行距，排在末条粗线之上一行，左/右各空一字，并使用基线补偿保持上下视觉间隔均衡
- 版记支持按规则增加主送机关行、抄送机关行；主送置于抄送之上，扩展要素之间自动使用中间细分隔线
- 任务串行处理，降低 Word COM 并发冲突
- 单任务默认 300 秒超时，可通过 `REDHEAD_JOB_TIMEOUT_MS` 调整
- 超时后按本任务记录的 Word PID 做清理
- 处理产物默认按“保留 30 天、最多 200 个任务”自动清理，可通过 `REDHEAD_JOB_RETENTION_DAYS`、`REDHEAD_JOB_RETENTION_MAX`、`REDHEAD_CLEANUP_ON_START`、`REDHEAD_CLEANUP_AFTER_JOB` 调整
- 记录页提供产物清理预估和立即清理，清理 `output/<任务ID>` 时会同步清理同名 `uploads/<任务ID>`
- 每个任务生成 `job-status.json`，记录当前阶段、最后更新时间、耗时、PowerShell PID 和 Word PID
- 超时或失败时，服务端会把最后阶段写入错误响应和 `process.log`，用于定位卡在启动 Word、打开文档、导出 PDF、校验等哪一步
- 发文机关标志按版心上边缘下 35mm 定位，红色方正小标宋简体、58pt、居中处理
- 发文字号校验年份全称、六角括号、序号不编虚位、不加“第”、末尾加“号”
- 发文字号与签发人位于发文机关标志下空二行处；发文字号左空一字，签发人右空一字，签发人姓名使用 `楷体_GB2312`
- 版头红色分隔线使用发文字号段落底边框生成，按红线上沿距发文字号下方约 4mm、总长约 442.2pt/156mm、线宽 2.25pt 处理并校验，避免 DOCX/PDF 因浮动线条锚点产生显示差异
- 标题按红色分隔线下空二行编排，默认以正文 3 号字两行 `32pt/11.29mm` 为目标净距，并自动校正
- 主送机关按标题后空行、左顶格、全角冒号、3 号仿宋_GB2312 处理并校验
- 正文基础格式强制黑色，并校验字体、字号、行距、首行左空二字、回行顶格
- 标题按黑色二号方正小标宋、32pt 行距居中处理，并按词义主动分行，避免拆开行政区划、产品名称和年度括号
- 附件说明按正文下空一行、左空二字、3 号仿宋_GB2312 处理，自动规范 `附件：` 全角冒号、多附件阿拉伯数字顺序号、序号左对齐、名称末尾标点和回行悬挂缩进
- 页码按 4 号宋体、`- 1 -` 形式、奇偶页左右空一字、位于版心下边缘下 7mm 处理并校验
- 页码奇偶页对齐、是否仅正文编页码已规则化，可在工作台维护
- 生成后返回格式校验项：页面、页边距、红头、版头红线、文号格式与签发人空字、主送机关、标题字体/颜色/行距/分行、正文基础格式、附件说明、多附件序号对齐、页码、空白页、源文尾部空白页处理、版记要素行数、版记主送/抄送行、中间细分隔线、首末粗分隔线、版记 4 号仿宋与固定 28pt 行距、版记基线补偿、版记文字上下居中、版记末条分隔线位置、印发机关/印发日期空字、PDF 文件
- 标题含 `关于对` 等介词用法时给出提示，不自动改写正式标题内容
- 原文已带落款单位/成文日期时会复用并规范化，不再重复追加
- 默认按加盖公章公文处理，成文日期右空四字编排，发文机关署名以成文日期为准居中
- 署名居中支持 `signature.companyCenterCorrectionPt` 视觉补偿，处理 Word/PDF 字体测量差异
- 公章图片必须使用真实图片配置到 `signature.seal.imagePath`；未配置时会在校验项中提示

## 自动化测试

默认测试不启动 Microsoft Word，适合每次改规则、前端和脚本结构后快速回归：

```powershell
cd E:\code\wordedit\redhead-web
npm test
```

覆盖内容：

- `rules.json`、`default-rules.json` 符合 `data/rules.schema.json`
- Schema 覆盖当前业务规则叶子项
- 工作台 `data-rule` 覆盖全部业务规则项，且不能保留已删除的无效规则项
- PowerShell 脚本读取的规则路径都在 Schema 中
- 规则执行矩阵覆盖全部业务规则项
- 规则执行矩阵中引用的应用函数、代码路径和文档校验标签必须真实存在
- `docs/rule-execution-matrix.md` 必须覆盖当前业务规则叶子项
- 任务产物清理策略 dry-run 和实际删除行为
- `server.js`、`public/app.js`、`lib/rules-validator.js` 语法检查
- `scripts/redhead.ps1` 和 `scripts/modules/*.ps1` PowerShell 语法检查
- 非法规则值拒绝测试

更新规则覆盖矩阵文档：

```powershell
npm run docs:rules
```

需要跑真实 Word COM 端到端回归时使用：

```powershell
npm run test:word
```

该命令会使用 `test-input/tail-blank-test.docx` 生成 DOCX/PDF，因此要求本机 Microsoft Word 和字体环境可用。

## 脚本结构

PowerShell 处理脚本已拆为入口和模块：

- `scripts/redhead.ps1`：稳定入口，保留服务端调用参数，加载模块。
- `scripts/modules/Redhead.Core.ps1`：规则读取、Word COM 工具、版式处理、校验函数。
- `scripts/modules/Redhead.Runner.ps1`：任务执行主流程和阶段状态写入。

## 任务诊断

每次处理会在 `output/<任务ID>/` 下保留：

- `job-status.json`：最后执行阶段和心跳信息。
- `process.log`：任务诊断、PowerShell 标准输出、标准错误。
- `process-pids.json`：本次任务关联的 PowerShell/Word 进程信息。
- `rules.snapshot.json`：本次任务使用的规则快照。
- `meta.json`：本次处理表单参数。
- `result.json`：成功或脚本内失败时写出的处理结果。

如果页面提示超时，优先查看 `job-status.json` 的 `stage/detail` 和 `process.log` 的 `[diagnostics]` 段。
