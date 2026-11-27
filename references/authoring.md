# Word、数据输入与渲染

## 使用默认模板

随附 `assets/reading-template.docx` 来自本次验证过的六栏样例，已将原论文正文替换为字段标记，保留样式和页面。`scripts/build_review.py` 只用于这个默认模板；它会检查29个段落及固定标题，防止把未知样例按错位索引填入。

先完成论文阅读与框架分析，再写 UTF-8 JSON：

```json
{
  "title": "目标论文的中文主题｜文献整理",
  "metadata": [
    "英文题目：完整原题",
    "中文题目：完整中文翻译",
    "DOI：原文信息",
    "期刊信息：期刊、年、可核对的卷期页码及日期",
    "作者：按原文顺序列全",
    "作者单位（原文）：所有单位及对应",
    "作者单位：逐项中文翻译",
    "通讯作者：姓名和邮箱"
  ],
  "abstract_en": "完整英文摘要",
  "abstract_zh": "完整中文摘要译文",
  "introduction": "450字以内的连续段落",
  "methods": "450字以内的连续段落，适配文章类型",
  "captions": ["图1 完整图注译文", "图2 完整图注译文"],
  "conclusions": ["核心结论及范围", "研究不足与展望"]
}
```

图注数量由输入决定，图号应从1连续。没有编号正文图时 captions 使用空数组。每个字符串对应一个段落，不写 Markdown 加粗或列表符号。元数据缺失时写明原文未列出，不留下模板标记。

```text
python scripts/build_review.py --content <work/review.json> --output <outputs/文献解读.docx> --report <work/document-report.json>
```

Python 需 lxml。脚本直接复制模板 ZIP，修改正文和文档元数据、移除过期页数统计；保留样式、字体表、关系和页面设置。它核验输入、图序及字数，不替代事实核对或视觉检查。默认拒绝覆盖，加 `--overwrite` 可修订本次输出。

## 新参考文档

用户新给的样例优先。若栏目、表格、页眉、图片或段落结构不同，先读取并按可用的 documents skill 蒸馏实际模板；复制其文档部件，按确认的字段修改，不能直接用默认脚本强行覆盖。不要把每个样例都统一改为六栏，也不要为凑样例页数改变字号。

## 渲染和逐页检查

如环境提供 documents/pdf skill，使用其可用的渲染与检查流程；获取应用提供的 Python、Node 和原生工具位置，避免假定旧路径可用。

在 Windows 且标准渲染器不可用时，可使用已安装的 Microsoft Word：

```powershell
& scripts/render_word_native.ps1 -InputDocx <absolute.docx> -OutputPdf <absolute-work.pdf>
pdftoppm -r 150 -png <absolute-work.pdf> <work/page>
```

脚本创建自己的隐藏 Word 实例、禁用宏、只读打开指定文档、导出 PDF，再关闭该文档和自己的实例，不保存源文档。若 COM 受沙箱限制，按当前权限机制处理；不要反复启动失败实例或关闭用户其他 Word 文档。未导出成功时不能声称已经检查排版。

逐页打开最终 PNG。检查英文与中文字体、上下标和希腊字母、段落溢出、孤立标题、末页少量残行及是否残留样例内容。修复后重新导出并检查实际变更页及分页受影响的后续页。字数受限的总结和自写结论可精简；摘要和图注保持完整。

## 输出组织示例

```text
outputs/
  参考框架分析.md
  文献解读_论文主题.docx
  期刊、题目、作者信息图.png
  思维导图.png
  思维导图.svg
  思维导图.md
  思维导图.json
  论文插图/
    图1.png
    图2.png
  论文插图.zip
```

只创建本次要求的文件。原始 PDF 按用户要求保留或另存，不必重复复制；校对 PDF、页面 PNG、源文本、输入 JSON 和运行报告保留在 `work/`。
