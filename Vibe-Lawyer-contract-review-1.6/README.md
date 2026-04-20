# 合同审核技能 - XML 修订痕迹写入

## 概述

本技能用于中文合同审核与修改，修订和批注均通过直接操作 DOCX 底层 XML 写入，且该实现路径固定不变，确保与 WPS 完全兼容。

## 核心特性

### 1. XML 修订痕迹写入

- **直接操作底层 XML**：不依赖高层封装接口，直接编辑 document.xml、comments.xml、settings.xml
- **修订和批注都走 XML**：删除、插入、批注范围、批注正文全部通过 XML 节点写入，不提供非 XML 备选路径
- **WPS 完全兼容**：严格遵循 ECMA-376 Office Open XML 标准
- **标准修订标记**：使用 `<w:del>`、`<w:ins>`、`<w:commentRangeStart>` 等标准元素
- **完整修订信息**：包含作者、时间、唯一 ID 等必要属性

### 2. 修订类型支持

- **删除标记** (`<w:del>`)：标记被删除的文本
- **插入标记** (`<w:ins>`)：标记新插入的文本
- **批注标记** (`<w:comment>`)：添加批注说明

### 3. WPS 兼容性保证

- 使用标准 OOXML 命名空间
- ISO 8601 时间格式
- 唯一修订 ID 管理
- 自动启用修订跟踪设置

## 使用方法

### 对外唯一入口

- 对外只暴露 `scripts/run_generate_review_docx.ps1`
- `scripts/internal_generate_review_docx.ps1`、`scripts/internal_stage_review_inputs.ps1` 与 `scripts/internal_write_revisions_xml.py` 仅作为内部实现，不作为用户入口
- 包装器内置中文路径兼容、XML 直改、打包与自检，用户无需理解底层实现

### Windows / macOS 调用方式

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\run_generate_review_docx.ps1 `
  -Source "原合同.docx" `
  -Instructions ".\references\review-instructions.example.json" `
  -Output "原合同-修订批注版.docx"
```

```powershell
pwsh -File ./scripts/run_generate_review_docx.ps1 `
  -Source "source.docx" `
  -Instructions "./references/review-instructions.example.json" `
  -Output "reviewed.docx"
```

### 指令文件格式

```json
{
  "author": "合同审核AI",
  "operations": [
    {
      "mode": "revision_comment",
      "anchor_text": "乙方应在验收后30日内付款。",
      "replacement_text": "乙方应在验收合格且收到合法有效发票后10个工作日内付款。",
      "comment": {
        "问题": "付款条件与发票、验收标准未闭环。",
        "风险": "易引发付款起算点争议。",
        "修改建议": "补齐验收合格与合法有效发票两个触发条件。",
        "建议条款": "乙方应在验收合格且收到合法有效发票后10个工作日内付款。"
      }
    }
  ]
}
```

### 自有模板比对增强版

当需要用“对方合同 + 自有模板”自动生成修订批注版时，仍然使用同一个入口，只是在 `instructions.json` 中增加 `template_compare` 配置：

```json
{
  "author": "合同审核AI",
  "template_compare": {
    "template_path": "./company-template.docx",
    "focus_topics": ["付款", "违约责任", "争议解决"],
    "max_operations": 6,
    "mode": "revision_comment",
    "min_similarity_for_revision": 0.35,
    "min_alignment_confidence": 0.45,
    "allow_missing_clause_insert": true,
    "missing_clause_mode": "revision_comment"
  },
  "operations": []
}
```

调用方式不变：

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\run_generate_review_docx.ps1 `
  -Source "对方合同.docx" `
  -Instructions ".\references\template-compare-review.example.json" `
  -Output "对方合同-按模板修订批注版.docx"
```

MVP 当前支持的对比主题包括：

- `付款`
- `验收`
- `违约责任`
- `解除`
- `责任限制`
- `知识产权`
- `保密`
- `数据合规`
- `争议解决`

增强版当前规则为：

- 以对方合同为主文本，以自有模板为风控参考
- 先按条款标题、正文相似度、主题命中等信号做多因子对齐，再决定生成 `revision_comment` 或 `comment`
- 即使合同顺序重排，也优先按条款内容对齐，而不是简单按段落顺序替换
- 对弱标题或无编号短标题的条款，会尝试按“短标题 + 主题命中 + 后续正文”识别为条款单元
- 对没有明确标题的正文块，会按主题命中向后聚合连续正文，作为条款兜底识别
- 当同主题条款相似度不足时，仅提示风险并给出模板建议条款，不直接改写原文
- 支持显式 `insertion_texts`，可在指定段落后新增完整条款段落块
- `allow_missing_clause_insert` 打开后，会优先按模板相对顺序自动插入缺失条款，支持插在首个已匹配条款前或两个已匹配条款之间
- 自动缺失条款插入未命中时，仍会退回风险提示，不会机械拼接模板全文
- 输出摘要中会附带 `alignment_reasons_summary` 与 `missing_reasons_summary`，便于解释对齐依据与缺失原因
- 输出摘要中还会附带 `review_summary_lines` 与 `review_summary_text`，可直接作为人工审阅摘要
- `template_path` 支持相对路径，相对基准为当前 `instructions.json` 所在目录
- 可直接参考 `references/template-compare-review.example.json`

## XML 结构示例

### 删除标记

```xml
<w:p>
  <w:r>
    <w:del w:author="审核人" w:date="2026-03-19T10:00:00Z" w:id="1">
      <w:t>被删除的文本内容</w:t>
    </w:del>
  </w:r>
</w:p>
```

### 插入标记

```xml
<w:p>
  <w:r>
    <w:ins w:author="审核人" w:date="2026-03-19T10:00:00Z" w:id="2">
      <w:t>新插入的文本内容</w:t>
    </w:ins>
  </w:r>
</w:p>
```

### 批注标记

```xml
<w:p>
  <w:commentRangeStart w:id="101"/>
  <w:r>
    <w:t>需要批注的文本</w:t>
  </w:r>
  <w:commentRangeEnd w:id="101"/>
  <w:r>
    <w:commentReference w:id="101"/>
  </w:r>
</w:p>
```

### 批注内容 (comments.xml)

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="101" w:author="审核人" w:date="2026-03-19T10:00:00Z">
    <w:p>
      <w:r>
        <w:t>问题：该条款责任边界不清</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>风险：可能导致违约责任无法明确划分</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>建议：修改为"违约方应承担全部直接损失"</w:t>
      </w:r>
    </w:p>
  </w:comment>
</w:comments>
```

### 修订跟踪设置 (settings.xml)

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:trackRevisions w:val="1"/>
  <w:showRevisions w:val="1"/>
  <w:revisionView>
    <w:markup>1</w:markup>
  </w:revisionView>
</w:settings>
```

## 实现流程

1. **解压 DOCX**：将 DOCX 文件解压到临时目录
2. **解析 XML**：读取 document.xml、comments.xml、settings.xml
3. **写入修订**：
   - 查找目标文本位置
   - 创建 `<w:del>` 或 `<w:ins>` 标记
   - 添加批注范围标记和引用
4. **更新批注**：在 comments.xml 中添加批注内容
5. **启用跟踪**：在 settings.xml 中启用修订跟踪
6. **重新打包**：将修改后的文件重新打包为 DOCX

## 最小环境要求

- 最终用户仅需 PowerShell
- Windows 推荐 `powershell.exe` 或 `pwsh`
- macOS 推荐 `pwsh`
- 不要求安装 Word / WPS / Office
- 不要求安装 Python 即可生成最终 `docx`

### 开发验证环境

- 仅在开发、测试或调试内部实现时才需要 Python
- Python 侧额外依赖：`lxml`

## 文件结构

```
contract-review-docx-auditor/
├── SKILL.md                          # 技能主文档
├── scripts/
│   ├── run_generate_review_docx.ps1  # 对外唯一入口
│   ├── internal_generate_review_docx.ps1
│   ├── internal_stage_review_inputs.ps1
│   └── internal_write_revisions_xml.py
├── references/
│   └── xml-revision-spec.md          # XML 实现详细规范
└── README.md                         # 本文件
```

## WPS 兼容性测试

生成修订版 DOCX 后，必须在 WPS Office 中验证：

1. 打开生成的 DOCX 文件
2. 点击"审阅"选项卡
3. 确认所有修订标记正确显示（删除线、下划线等）
4. 确认所有批注正确显示在批注框中
5. 确认可以接受/拒绝修订

## 注意事项

### 必须遵守的规则

1. **时间格式**：必须使用 ISO 8601 格式 `YYYY-MM-DDThh:mm:ssZ`
2. **唯一 ID**：每个修订的 ID 必须唯一且为正整数
3. **完整属性**：必须包含 author、date、id
4. **命名空间**：必须使用正确的 OOXML 命名空间
5. **启用跟踪**：必须在 settings.xml 中启用 trackRevisions

### 常见错误

1. **WPS 不显示修订**：
   - 检查 settings.xml 是否启用 trackRevisions
   - 检查修订标记属性是否完整

2. **批注不显示**：
   - 检查 comments.xml 格式是否正确
   - 检查批注 ID 是否匹配

3. **XML 解析错误**：
   - 检查命名空间声明是否正确
   - 检查 XML 是否符合 Well-Formed 要求

## 参考文档

- [xml-revision-spec.md](references/xml-revision-spec.md) - 详细实现规范
- [SKILL.md](SKILL.md) - 技能完整说明
- [ECMA-376 标准](http://www.ecma-international.org/publications/standards/Ecma-376.htm)
- [ISO/IEC 29500](https://www.iso.org/standard/71691.html)

## 技术支持

如遇到问题，请检查：
1. XML 结构是否符合 ECMA-376 标准
2. 所有必要属性是否完整
3. WPS 是否为最新版本
4. 修订 ID 是否唯一
5. 是否通过 `scripts/run_generate_review_docx.ps1` 启动，而不是直接调用内部脚本

## 许可证

本技能仅供学习和研究使用。
