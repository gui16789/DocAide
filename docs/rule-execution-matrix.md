# 规则执行矩阵

本文件由 `node scripts/generate-rule-coverage-doc.js` 根据 `data/rule-coverage.json` 生成。

- 当前业务规则叶子项：102
- 覆盖条目：20
- 未覆盖规则：0
- 未匹配规则模式：0

## 覆盖分组

| 分组 | 风险 | 规则路径 | 应用位置 | 校验方式 | 备注 |
| --- | --- | --- | --- | --- | --- |
| 规则版本信息 | low | version | stampRules (server.js)<br>saveRules (server.js) | 测试：规则 JSON 能解析并符合 Schema | 用于规则版本展示和历史记录，不直接影响 Word 版式。 |
| 页面尺寸和版心 | high | page.* | Get-PageSizeSpec (scripts/modules/Redhead.Core.ps1)<br>Apply-PageSetup (scripts/modules/Redhead.Core.ps1) | 包含：页面尺寸<br>页边距符合规则 |  |
| 发文机关标志文字样式 | high | redHeader.text<br>redHeader.font<br>redHeader.sizePt<br>redHeader.color<br>redHeader.scalePercent<br>redHeader.characterSpacingPt | Insert-RedHeader (scripts/modules/Redhead.Core.ps1) | 红头文字已插入<br>红头字号、缩放、颜色符合规则 |  |
| 发文机关标志和发文字号位置 | high | redHeader.topFromContentTopMm<br>redHeader.topSpacerPt<br>redHeader.visibleTopInsetPt<br>redHeader.docNoBlankLines<br>redHeader.spaceAfterPt | Insert-RedHeader (scripts/modules/Redhead.Core.ps1)<br>Adjust-RedHeaderLayout (scripts/modules/Redhead.Core.ps1) | 发文机关标志上边缘距版心上边缘35mm<br>发文字号位于发文机关标志下空二行 |  |
| 发文字号和签发人内容 | high | documentNo.defaultText<br>documentNo.signerLabel | Insert-RedHeader (scripts/modules/Redhead.Core.ps1) | 发文字号和签发人已插入<br>发文字号格式符合规范 |  |
| 发文字号和签发人格式 | high | documentNo.font<br>documentNo.sizePt<br>documentNo.signerFont<br>documentNo.leftBlankChars<br>documentNo.signerRightBlankChars | Insert-RedHeader (scripts/modules/Redhead.Core.ps1) | 发文字号左空一字<br>发文字号字体字号符合规范<br>签发人右空一字<br>签发人字体字号符合规范 |  |
| 版头红色分隔线 | high | documentNo.redLineColor<br>documentNo.redLineLengthMm<br>documentNo.redLineWidthPt<br>documentNo.redLineOffsetMm<br>documentNo.redLineLineBoxHeightPt<br>documentNo.spaceAfterPt | Add-DocumentNoRedLine (scripts/modules/Redhead.Core.ps1)<br>Adjust-DocumentNoRedLinePosition (scripts/modules/Redhead.Core.ps1)<br>Adjust-TitlePositionAfterRedLine (scripts/modules/Redhead.Core.ps1) | 红线已设置<br>版头红线颜色线宽符合规则<br>版头红线位于发文字号下4mm<br>版头红线居中且与版心等宽<br>标题编排于红色分隔线下空二行 |  |
| 标题格式和分行 | high | title.* | Normalize-TitleBlock (scripts/modules/Redhead.Core.ps1)<br>Set-TitleParagraph (scripts/modules/Redhead.Core.ps1)<br>Adjust-TitlePositionAfterRedLine (scripts/modules/Redhead.Core.ps1) | 标题编排于红色分隔线下空二行<br>标题字体字号颜色行距居中符合规范<br>标题已合并并按词义主动分行 |  |
| 正文基础格式 | high | body.* | Set-BodyParagraph (scripts/modules/Redhead.Core.ps1)<br>Apply-BodyAndTitleStyle (scripts/modules/Redhead.Core.ps1) | 正文基础字体字号行距颜色符合规范<br>正文自然段左空二字回行顶格 |  |
| 主送机关 | high | mainRecipient.* | Set-MainRecipientParagraph (scripts/modules/Redhead.Core.ps1)<br>Apply-BodyAndTitleStyle (scripts/modules/Redhead.Core.ps1) | 主送机关标题下空一行且顶格<br>主送机关使用全角冒号<br>主送机关字体字号行距符合规范 |  |
| 附件说明 | high | attachment.* | Apply-AttachmentExplanationStyle (scripts/modules/Redhead.Core.ps1)<br>Set-AttachmentExplanationParagraph (scripts/modules/Redhead.Core.ps1) | 附件说明正文下空一行左空二字<br>附件说明字体字号行距符合规范<br>附件说明回行对齐附件名称首字<br>多个附件序号左对齐 |  |
| 署名和成文日期 | high | signature.company<br>signature.date<br>signature.noSeal<br>signature.companyLeftIndentPt<br>signature.dateLeftIndentPt<br>signature.dateRightIndentChars<br>signature.companyCenterCorrectionPt | Apply-Signature (scripts/modules/Redhead.Core.ps1) | 落款单位和成文日期唯一<br>盖章公文模式已启用<br>成文日期右空四字编排<br>发文机关署名以成文日期为准居中 | company/date 可被表单 meta 覆盖；规则值作为默认值和校验基准。 |
| 公章图片 | high | signature.seal.* | Add-SealIfConfigured (scripts/modules/Redhead.Core.ps1) | 真实公章图片已配置<br>公章已插入文档 |  |
| 版记启用和分页位置 | high | imprint.enabled<br>imprint.samePageWhenPossible<br>imprint.requireEvenPage | Add-Imprint (scripts/modules/Redhead.Core.ps1)<br>Align-ImprintTableToContentBottom (scripts/modules/Redhead.Core.ps1) | 总页数与正文/版记结构一致<br>版记位于偶数页<br>版记末条分隔线与版心下边缘重合 |  |
| 版记印发机关和印发日期 | high | imprint.office<br>imprint.date<br>imprint.officeLeftChars<br>imprint.dateRightChars | Set-ImprintIssueCell (scripts/modules/Redhead.Core.ps1) | 印发机关左空一字<br>印发日期右空一字<br>印发机关和印发日期排在末线上方一行 |  |
| 版记文字、行高和上下间隔 | high | imprint.font<br>imprint.sizePt<br>imprint.lineSpacingPt<br>imprint.rowHeightPt<br>imprint.baselineShiftPt<br>imprint.cellPaddingTopPt<br>imprint.cellPaddingBottomPt | Set-ImprintIssueCell (scripts/modules/Redhead.Core.ps1)<br>Set-ImprintElementCell (scripts/modules/Redhead.Core.ps1)<br>Add-ImprintTable (scripts/modules/Redhead.Core.ps1) | 版记文字上下居中且间距一致<br>版记字体字号固定行距符合规则<br>版记文字基线补偿符合规则<br>印发机关和印发日期排在末线上方一行 |  |
| 版记分隔线 | high | imprint.outerLineWidthPt<br>imprint.innerLineWidthPt | Add-ImprintTable (scripts/modules/Redhead.Core.ps1) | 版记要素之间使用中间细分隔线<br>版记首末粗分隔线线宽符合规则<br>版记末条分隔线与版心下边缘重合 |  |
| 版记主送和抄送扩展行 | medium | imprint.mainRecipient.*<br>imprint.cc.* | Get-ImprintRows (scripts/modules/Redhead.Core.ps1)<br>Set-ImprintElementCell (scripts/modules/Redhead.Core.ps1) | 版记要素行数符合规则<br>包含：已编入版记且格式符合规范<br>版记要素之间使用中间细分隔线 |  |
| 页码 | high | pageNumber.* | Apply-BodyPageNumbers (scripts/modules/Redhead.Core.ps1)<br>Add-PageFieldFooter (scripts/modules/Redhead.Core.ps1)<br>Apply-PageSetup (scripts/modules/Redhead.Core.ps1) | 正文页脚已设置页码字段<br>页码字体字号形式符合规范<br>页码位于版心下边缘下7mm<br>页码单双页按规则对齐并空字<br>空白页不编页码<br>空白页后的版记页不编页码<br>版记页按正文连续编页码<br>页码规则已关闭 |  |
| 源文清理和输入转换 | medium | cleanup.* | cleanup.convertDocToDocx (scripts/modules/Redhead.Runner.ps1)<br>cleanup.removeExistingHeadersFooters (scripts/modules/Redhead.Runner.ps1)<br>cleanup.fixDoubleAnmeng (scripts/modules/Redhead.Runner.ps1)<br>Trim-TrailingBlankContent (scripts/modules/Redhead.Core.ps1) | 源文件尾部空白页已处理<br>测试：可选 Word COM 回归<br>代码路径：cleanup.convertDocToDocx<br>代码路径：cleanup.removeExistingHeadersFooters<br>代码路径：cleanup.fixDoubleAnmeng | 输入转换和页眉页脚清理由流程代码执行；尾部空白页有文档级校验。 |

## 规则展开

| 规则路径 | 覆盖分组 | 风险 |
| --- | --- | --- |
| attachment.font | 附件说明 | high |
| attachment.leftBlankChars | 附件说明 | high |
| attachment.lineSpacingPt | 附件说明 | high |
| attachment.sizePt | 附件说明 | high |
| attachment.spaceBeforeLines | 附件说明 | high |
| body.color | 正文基础格式 | high |
| body.firstLineIndentPt | 正文基础格式 | high |
| body.font | 正文基础格式 | high |
| body.latinFont | 正文基础格式 | high |
| body.lineSpacingPt | 正文基础格式 | high |
| body.sizePt | 正文基础格式 | high |
| cleanup.convertDocToDocx | 源文清理和输入转换 | medium |
| cleanup.fixDoubleAnmeng | 源文清理和输入转换 | medium |
| cleanup.removeExistingHeadersFooters | 源文清理和输入转换 | medium |
| cleanup.trimTrailingBlankPages | 源文清理和输入转换 | medium |
| documentNo.defaultText | 发文字号和签发人内容 | high |
| documentNo.font | 发文字号和签发人格式 | high |
| documentNo.leftBlankChars | 发文字号和签发人格式 | high |
| documentNo.redLineColor | 版头红色分隔线 | high |
| documentNo.redLineLengthMm | 版头红色分隔线 | high |
| documentNo.redLineLineBoxHeightPt | 版头红色分隔线 | high |
| documentNo.redLineOffsetMm | 版头红色分隔线 | high |
| documentNo.redLineWidthPt | 版头红色分隔线 | high |
| documentNo.signerFont | 发文字号和签发人格式 | high |
| documentNo.signerLabel | 发文字号和签发人内容 | high |
| documentNo.signerRightBlankChars | 发文字号和签发人格式 | high |
| documentNo.sizePt | 发文字号和签发人格式 | high |
| documentNo.spaceAfterPt | 版头红色分隔线 | high |
| imprint.baselineShiftPt | 版记文字、行高和上下间隔 | high |
| imprint.cc.enabled | 版记主送和抄送扩展行 | medium |
| imprint.cc.label | 版记主送和抄送扩展行 | medium |
| imprint.cc.leftBlankChars | 版记主送和抄送扩展行 | medium |
| imprint.cc.rightBlankChars | 版记主送和抄送扩展行 | medium |
| imprint.cc.text | 版记主送和抄送扩展行 | medium |
| imprint.cellPaddingBottomPt | 版记文字、行高和上下间隔 | high |
| imprint.cellPaddingTopPt | 版记文字、行高和上下间隔 | high |
| imprint.date | 版记印发机关和印发日期 | high |
| imprint.dateRightChars | 版记印发机关和印发日期 | high |
| imprint.enabled | 版记启用和分页位置 | high |
| imprint.font | 版记文字、行高和上下间隔 | high |
| imprint.innerLineWidthPt | 版记分隔线 | high |
| imprint.lineSpacingPt | 版记文字、行高和上下间隔 | high |
| imprint.mainRecipient.enabled | 版记主送和抄送扩展行 | medium |
| imprint.mainRecipient.label | 版记主送和抄送扩展行 | medium |
| imprint.mainRecipient.leftBlankChars | 版记主送和抄送扩展行 | medium |
| imprint.mainRecipient.rightBlankChars | 版记主送和抄送扩展行 | medium |
| imprint.mainRecipient.text | 版记主送和抄送扩展行 | medium |
| imprint.office | 版记印发机关和印发日期 | high |
| imprint.officeLeftChars | 版记印发机关和印发日期 | high |
| imprint.outerLineWidthPt | 版记分隔线 | high |
| imprint.requireEvenPage | 版记启用和分页位置 | high |
| imprint.rowHeightPt | 版记文字、行高和上下间隔 | high |
| imprint.samePageWhenPossible | 版记启用和分页位置 | high |
| imprint.sizePt | 版记文字、行高和上下间隔 | high |
| mainRecipient.font | 主送机关 | high |
| mainRecipient.lineSpacingPt | 主送机关 | high |
| mainRecipient.sizePt | 主送机关 | high |
| mainRecipient.titleSpaceAfterPt | 主送机关 | high |
| page.bottomMarginMm | 页面尺寸和版心 | high |
| page.leftMarginMm | 页面尺寸和版心 | high |
| page.rightMarginMm | 页面尺寸和版心 | high |
| page.size | 页面尺寸和版心 | high |
| page.topMarginMm | 页面尺寸和版心 | high |
| pageNumber.blankChars | 页码 | high |
| pageNumber.distanceBelowContentMm | 页码 | high |
| pageNumber.enabled | 页码 | high |
| pageNumber.evenAlign | 页码 | high |
| pageNumber.font | 页码 | high |
| pageNumber.format | 页码 | high |
| pageNumber.oddAlign | 页码 | high |
| pageNumber.sizePt | 页码 | high |
| redHeader.characterSpacingPt | 发文机关标志文字样式 | high |
| redHeader.color | 发文机关标志文字样式 | high |
| redHeader.docNoBlankLines | 发文机关标志和发文字号位置 | high |
| redHeader.font | 发文机关标志文字样式 | high |
| redHeader.scalePercent | 发文机关标志文字样式 | high |
| redHeader.sizePt | 发文机关标志文字样式 | high |
| redHeader.spaceAfterPt | 发文机关标志和发文字号位置 | high |
| redHeader.text | 发文机关标志文字样式 | high |
| redHeader.topFromContentTopMm | 发文机关标志和发文字号位置 | high |
| redHeader.topSpacerPt | 发文机关标志和发文字号位置 | high |
| redHeader.visibleTopInsetPt | 发文机关标志和发文字号位置 | high |
| signature.company | 署名和成文日期 | high |
| signature.companyCenterCorrectionPt | 署名和成文日期 | high |
| signature.companyLeftIndentPt | 署名和成文日期 | high |
| signature.date | 署名和成文日期 | high |
| signature.dateLeftIndentPt | 署名和成文日期 | high |
| signature.dateRightIndentChars | 署名和成文日期 | high |
| signature.noSeal | 署名和成文日期 | high |
| signature.seal.enabled | 公章图片 | high |
| signature.seal.heightMm | 公章图片 | high |
| signature.seal.imagePath | 公章图片 | high |
| signature.seal.verticalOffsetPt | 公章图片 | high |
| signature.seal.widthMm | 公章图片 | high |
| title.color | 标题格式和分行 | high |
| title.font | 标题格式和分行 | high |
| title.lineSpacingPt | 标题格式和分行 | high |
| title.maxLineChars | 标题格式和分行 | high |
| title.sizePt | 标题格式和分行 | high |
| title.spaceAfterPt | 标题格式和分行 | high |
| title.topBlankLines | 标题格式和分行 | high |
| version | 规则版本信息 | low |

