[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$Instructions,
    [Parameter(Mandatory = $true)][string]$Output,
    [string]$CleanOutput,
    [string]$Author = "合同审核AI",
    [ValidateSet("comment", "revision", "revision_comment")][string]$DefaultMode = "revision_comment",
    [string]$Date,
    [string]$InternalScriptRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:EffectiveScriptRoot = if (-not [string]::IsNullOrWhiteSpace($InternalScriptRoot)) {
    $InternalScriptRoot
} elseif (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $PSScriptRoot
} elseif (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
    Split-Path -Parent $PSCommandPath
} else {
    (Get-Location).Path
}

$script:W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
$script:PR = "http://schemas.openxmlformats.org/package/2006/relationships"
$script:CT = "http://schemas.openxmlformats.org/package/2006/content-types"
$script:XML = "http://www.w3.org/XML/1998/namespace"
$script:CommentsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
$script:SettingsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
$script:CommentsType = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
$script:SettingsType = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
$script:CommentFontName = "SimSun"

function FullPath([string]$PathText) { [System.IO.Path]::GetFullPath($PathText) }
function HasNonAscii([AllowNull()][string]$Text) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    foreach ($char in $Text.ToCharArray()) {
        if ([int][char]$char -gt 127) { return $true }
    }
    $false
}
function Norm([AllowNull()][string]$Text) { if ($null -eq $Text) { "" } else { ([regex]::Replace($Text, "\s+", " ")).Trim() } }
function Iso([AllowNull()][string]$Value) { if ([string]::IsNullOrWhiteSpace($Value)) { [DateTimeOffset]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ") } else { $Value } }

function Ensure-DotNetCompression {
    Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
}

function LoadXml([string]$Path) {
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    $doc.Load($Path)
    $doc
}

function SaveXml([System.Xml.XmlDocument]$Doc, [string]$Path) {
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = New-Object System.Text.UTF8Encoding($true)
    $settings.Indent = $false
    $writer = [System.Xml.XmlWriter]::Create($Path, $settings)
    try { $Doc.Save($writer) } finally { $writer.Dispose() }
}

function Ns([System.Xml.XmlDocument]$Doc) {
    $ns = New-Object System.Xml.XmlNamespaceManager($Doc.NameTable)
    $ns.AddNamespace("w", $script:W)
    $ns.AddNamespace("pr", $script:PR)
    $ns.AddNamespace("ct", $script:CT)
    $ns
}

function SelectNodesNs($XmlObject, [string]$XPath) {
    $results = @(Select-Xml -Xml $XmlObject -XPath $XPath -Namespace @{ w = $script:W; pr = $script:PR; ct = $script:CT } | ForEach-Object { $_.Node })
    return ,$results
}

function SelectOneNs($XmlObject, [string]$XPath) {
    $nodes = @(SelectNodesNs -XmlObject $XmlObject -XPath $XPath)
    if ($nodes.Length -gt 0) { $nodes[0] } else { $null }
}

function SharedBytes([string]$Path) {
    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
    try {
        $memory = New-Object System.IO.MemoryStream
        try { $stream.CopyTo($memory); $memory.ToArray() } finally { $memory.Dispose() }
    }
    finally { $stream.Dispose() }
}

function CopyDocx([string]$From, [string]$To) {
    $parent = Split-Path -Parent $To
    if ($parent) { [System.IO.Directory]::CreateDirectory($parent) | Out-Null }
    [System.IO.File]::WriteAllBytes($To, (SharedBytes $From))
}

function Use-Staging([string[]]$Paths) {
    foreach ($path in @($Paths)) {
        if (HasNonAscii $path) { return $true }
    }
    $false
}

function Invoke-Staging([string]$SourcePath, [string]$InstructionsPath) {
    $stagingScript = Join-Path $script:EffectiveScriptRoot "internal_stage_review_inputs.ps1"
    if (-not (Test-Path -LiteralPath $stagingScript)) {
        throw "缺少 staging 脚本: $stagingScript"
    }

    $raw = & $stagingScript -SourcePath $SourcePath -InstructionsPath $InstructionsPath | Out-String
    if ([string]::IsNullOrWhiteSpace($raw)) {
        throw "staging 脚本未返回结果: $stagingScript"
    }

    $stage = $raw | ConvertFrom-Json
    foreach ($field in @("workspace_root", "staged_source", "staged_instructions", "staged_reviewed_docx", "staged_clean_docx")) {
        if ($null -eq $stage.PSObject.Properties[$field] -or [string]::IsNullOrWhiteSpace([string]$stage.$field)) {
            throw "staging 输出缺少字段 ${field}: $stagingScript"
        }
    }
    $stage
}

function Read-JsonFile([string]$Path) {
    Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
}

function Prepare-TemplateCompareInput([object]$Config, [string]$OriginalInstructionsPath, [string]$ProcessingInstructionsPath, [AllowNull()][object]$StagingInfo) {
    if ($null -eq $Config.PSObject.Properties["template_compare"]) {
        return [pscustomobject]@{
            InstructionsPath = $ProcessingInstructionsPath
            CleanupRoot      = $null
            OriginalTemplatePath = $null
        }
    }

    $compareConfig = $Config.template_compare
    $templatePropertyName = if ($null -ne $compareConfig.PSObject.Properties["template_path"]) {
        "template_path"
    } elseif ($null -ne $compareConfig.PSObject.Properties["template"]) {
        "template"
    } else {
        $null
    }
    $templateProperty = if ($null -ne $templatePropertyName) { [string]$compareConfig.$templatePropertyName } else { $null }
    if ([string]::IsNullOrWhiteSpace($templateProperty)) { throw "template_compare 缺少 template_path" }

    $templatePath = ResolvePathFromBaseFile -BaseFilePath $OriginalInstructionsPath -CandidatePath $templateProperty
    if (-not (Test-Path -LiteralPath $templatePath)) { throw "template_compare 模板文件不存在: $templatePath" }
    if ([System.IO.Path]::GetExtension($templatePath).ToLowerInvariant() -ne ".docx") { throw "template_compare 模板必须是 .docx 文件" }

    $needsWorkspace = ($null -ne $StagingInfo) -or (HasNonAscii $templatePath)
    if (-not $needsWorkspace) {
        return [pscustomobject]@{
            InstructionsPath = $ProcessingInstructionsPath
            CleanupRoot      = $null
            OriginalTemplatePath = $templatePath
        }
    }

    $workspaceRoot = if ($null -ne $StagingInfo -and -not [string]::IsNullOrWhiteSpace([string]$StagingInfo.workspace_root)) {
        Join-Path ([string]$StagingInfo.workspace_root) "template_compare"
    } else {
        Join-Path ([System.IO.Path]::GetTempPath()) ("template-compare-stage-" + [guid]::NewGuid().ToString("N"))
    }
    [System.IO.Directory]::CreateDirectory($workspaceRoot) | Out-Null

    $stagedTemplatePath = Join-Path $workspaceRoot "template.docx"
    CopyDocx $templatePath $stagedTemplatePath
    $compareConfig.$templatePropertyName = "./template.docx"

    $preparedInstructionsPath = Join-Path $workspaceRoot "instructions.json"
    $preparedInstructionsJson = $Config | ConvertTo-Json -Depth 20
    [System.IO.File]::WriteAllText($preparedInstructionsPath, $preparedInstructionsJson, (New-Object System.Text.UTF8Encoding($false)))

    [pscustomobject]@{
        InstructionsPath = $preparedInstructionsPath
        CleanupRoot      = $(if ($null -ne $StagingInfo) { $null } else { $workspaceRoot })
        OriginalTemplatePath = $templatePath
    }
}

function NewDoc([string]$LocalName, [string]$NsUri, [string]$Prefix) {
    $doc = New-Object System.Xml.XmlDocument
    [void]$doc.AppendChild($doc.CreateXmlDeclaration("1.0", "UTF-8", $null))
    $root = if ($Prefix) { $doc.CreateElement($Prefix, $LocalName, $NsUri) } else { $doc.CreateElement($LocalName, $NsUri) }
    [void]$doc.AppendChild($root)
    $doc
}

function WNode([System.Xml.XmlDocument]$Doc, [string]$Name) { $Doc.CreateElement("w", $Name, $script:W) }
function SetWAttr([System.Xml.XmlElement]$Node, [string]$Name, [string]$Value) {
    $attr = $Node.OwnerDocument.CreateAttribute("w", $Name, $script:W)
    $attr.Value = $Value
    [void]$Node.Attributes.Append($attr)
}

function AddCommentFontProps([System.Xml.XmlElement]$RunProps) {
    $fonts = WNode $RunProps.OwnerDocument "rFonts"
    SetWAttr $fonts "ascii" $script:CommentFontName
    SetWAttr $fonts "hAnsi" $script:CommentFontName
    SetWAttr $fonts "eastAsia" $script:CommentFontName
    SetWAttr $fonts "cs" $script:CommentFontName
    SetWAttr $fonts "hint" "eastAsia"
    [void]$RunProps.AppendChild($fonts)
}

function NextId([string[]]$Values) {
    $max = -1
    foreach ($value in $Values) { if ($value -match '^\d+$') { $n = [int]$value; if ($n -gt $max) { $max = $n } } }
    $max + 1
}

function RunPropsClone([System.Xml.XmlNode]$Paragraph) {
    $node = SelectOneNs -XmlObject $Paragraph -XPath "./w:r/w:rPr"
    if ($null -eq $node) { $null } else { $node.CloneNode($true) }
}

function SnapshotParagraph([System.Xml.XmlNode]$Paragraph) {
    $ppr = $null
    $content = New-Object System.Collections.Generic.List[System.Xml.XmlNode]
    foreach ($child in @($Paragraph.ChildNodes)) {
        if ($child.LocalName -eq "pPr" -and $child.NamespaceURI -eq $script:W) { $ppr = $child.CloneNode($true) } else { $content.Add($child.CloneNode($true)) }
        [void]$Paragraph.RemoveChild($child)
    }
    [pscustomobject]@{ PPr = $ppr; Content = $content.ToArray() }
}

function AddRunText([System.Xml.XmlElement]$Run, [string]$Text, [bool]$Deleted) {
    $doc = $Run.OwnerDocument
    $tag = if ($Deleted) { "delText" } else { "t" }
    $buffer = New-Object System.Text.StringBuilder
    $flush = {
        if ($buffer.Length -eq 0) { return }
        $node = WNode $doc $tag
        $value = $buffer.ToString()
        if ($value.StartsWith(" ") -or $value.EndsWith(" ") -or $value.Contains("  ")) {
            $space = $doc.CreateAttribute("xml", "space", $script:XML)
            $space.Value = "preserve"
            [void]$node.Attributes.Append($space)
        }
        $node.InnerText = $value
        [void]$Run.AppendChild($node)
        [void]$buffer.Clear()
    }
    foreach ($char in ([string]$Text).ToCharArray()) {
        if ($char -eq "`n") { & $flush; [void]$Run.AppendChild((WNode $doc "br")) }
        elseif ($char -eq "`t") { & $flush; [void]$Run.AppendChild((WNode $doc "tab")) }
        else { [void]$buffer.Append([string]$char) }
    }
    & $flush
    if ($Run.ChildNodes.Count -eq 0) { $empty = WNode $doc $tag; $empty.InnerText = ""; [void]$Run.AppendChild($empty) }
}

function AppendRun([System.Xml.XmlNode]$Parent, [string]$Text, $RunProps, [bool]$Deleted = $false) {
    $doc = $Parent.OwnerDocument
    $run = WNode $doc "r"
    if ($null -ne $RunProps) { [void]$run.AppendChild($doc.ImportNode($RunProps, $true)) }
    AddRunText -Run $run -Text $Text -Deleted $Deleted
    [void]$Parent.AppendChild($run)
}

function RevNode([System.Xml.XmlDocument]$Doc, [string]$Name, [string]$Text, [int]$Id, [string]$AuthorText, [string]$DateText, $RunProps) {
    $node = WNode $Doc $Name
    SetWAttr $node "id" "$Id"
    SetWAttr $node "author" $AuthorText
    SetWAttr $node "date" $DateText
    AppendRun -Parent $node -Text $Text -RunProps $RunProps -Deleted ($Name -eq "del")
    $node
}

function NewDiffSegment([string]$Kind, [string]$Text) {
    [pscustomobject]@{ Kind = $Kind; Text = $Text }
}

function MergeDiffSegments($Segments) {
    $merged = @()
    foreach ($segment in $Segments) {
        if ($null -eq $segment -or [string]::IsNullOrEmpty([string]$segment.Text)) { continue }
        if ($merged.Count -gt 0) {
            $lastIndex = $merged.Count - 1
            $last = $merged[$lastIndex]
        } else {
            $lastIndex = -1
            $last = $null
        }
        if ($lastIndex -ge 0 -and $last.Kind -eq $segment.Kind) {
            $merged[$lastIndex] = (NewDiffSegment $last.Kind ([string]$last.Text + [string]$segment.Text))
        } else {
            $merged += ,(NewDiffSegment $segment.Kind ([string]$segment.Text))
        }
    }
    return $merged
}

function IsMergeableEqualSegment([string]$Text) {
    if ([string]::IsNullOrEmpty($Text) -or $Text -match '\s') { return $false }
    if ($Text -cmatch '^[\p{IsCJKUnifiedIdeographs}]{1,2}$') { return $true }
    if ($Text -cmatch '^[A-Za-z0-9]{1,4}$') { return $true }
    $false
}

function CollapseDiffWindow($WindowSegments) {
    $segments = @($WindowSegments)
    if ($segments.Count -eq 0) { return @() }
    $changeCount = @($segments | Where-Object { $_.Kind -ne "equal" }).Count
    $hasEqual = @($segments | Where-Object { $_.Kind -eq "equal" }).Count -gt 0
    if ($changeCount -le 1 -and -not $hasEqual) { return $segments }
    $oldText = ((@($segments | Where-Object { $_.Kind -in @("equal", "del") } | ForEach-Object { [string]$_.Text })) -join "")
    $newText = ((@($segments | Where-Object { $_.Kind -in @("equal", "ins") } | ForEach-Object { [string]$_.Text })) -join "")
    $collapsed = @()
    if (-not [string]::IsNullOrEmpty($oldText)) { $collapsed += ,(NewDiffSegment "del" $oldText) }
    if (-not [string]::IsNullOrEmpty($newText)) { $collapsed += ,(NewDiffSegment "ins" $newText) }
    return $collapsed
}

function CoalesceChangeWindows($Segments) {
    $segments = @($Segments)
    if ($segments.Count -eq 0) { return @() }
    $coalesced = @()
    $window = @()
    for ($i = 0; $i -lt $segments.Count; $i++) {
        $segment = $segments[$i]
        $prevKind = if ($i -gt 0) { [string]$segments[$i - 1].Kind } else { $null }
        $nextKind = if ($i + 1 -lt $segments.Count) { [string]$segments[$i + 1].Kind } else { $null }
        $isConnector = (
            $segment.Kind -eq "equal" -and
            $null -ne $prevKind -and
            $null -ne $nextKind -and
            $prevKind -ne "equal" -and
            $nextKind -ne "equal" -and
            (IsMergeableEqualSegment ([string]$segment.Text))
        )
        if ($segment.Kind -ne "equal" -or $isConnector) {
            $window += ,$segment
            continue
        }
        if ($window.Count -gt 0) {
            $coalesced += @(CollapseDiffWindow $window)
            $window = @()
        }
        $coalesced += ,$segment
    }
    if ($window.Count -gt 0) {
        $coalesced += @(CollapseDiffWindow $window)
    }
    return (MergeDiffSegments $coalesced)
}

function TokenizeDiffText([AllowNull()][string]$Text) {
    if ([string]::IsNullOrEmpty($Text)) { return @() }
    $pattern = "[0-9]+(?:[.,][0-9]+)*|[A-Za-z]+(?:[-_'][A-Za-z]+)*|[\p{IsCJKUnifiedIdeographs}]|[^\S\r\n]+|\r\n|\n|\r|."
    $tokens = New-Object System.Collections.Generic.List[string]
    foreach ($match in [System.Text.RegularExpressions.Regex]::Matches($Text, $pattern)) { $tokens.Add($match.Value) }
    return ,($tokens.ToArray())
}

function SliceTokenText($Tokens, [int]$Start, [int]$Count) {
    if ($Count -le 0) { return "" }
    if ($Count -eq 1) { return [string]$Tokens[$Start] }
    $end = $Start + $Count - 1
    (@($Tokens[$Start..$end]) -join "")
}

function CommonPrefixCount($LeftTokens, $RightTokens) {
    $left = @($LeftTokens)
    $right = @($RightTokens)
    $max = [Math]::Min($left.Count, $right.Count)
    $count = 0
    while ($count -lt $max -and $left[$count] -ceq $right[$count]) { $count++ }
    $count
}

function CommonSuffixCount($LeftTokens, $RightTokens, [int]$PrefixCount) {
    $left = @($LeftTokens)
    $right = @($RightTokens)
    $leftRemain = $left.Count - $PrefixCount
    $rightRemain = $right.Count - $PrefixCount
    $max = [Math]::Min($leftRemain, $rightRemain)
    $count = 0
    while ($count -lt $max -and $left[($left.Count - 1 - $count)] -ceq $right[($right.Count - 1 - $count)]) { $count++ }
    $count
}

function GetFallbackDiffSegments($OldTokens, $NewTokens) {
    $left = @($OldTokens)
    $right = @($NewTokens)
    $segments = New-Object System.Collections.Generic.List[object]
    $prefixCount = CommonPrefixCount $left $right
    $suffixCount = CommonSuffixCount $left $right $prefixCount
    $oldMiddleCount = $left.Count - $prefixCount - $suffixCount
    $newMiddleCount = $right.Count - $prefixCount - $suffixCount
    if ($prefixCount -gt 0) { $segments.Add((NewDiffSegment "equal" (SliceTokenText $left 0 $prefixCount))) }
    if ($oldMiddleCount -gt 0) { $segments.Add((NewDiffSegment "del" (SliceTokenText $left $prefixCount $oldMiddleCount))) }
    if ($newMiddleCount -gt 0) { $segments.Add((NewDiffSegment "ins" (SliceTokenText $right $prefixCount $newMiddleCount))) }
    if ($suffixCount -gt 0) { $segments.Add((NewDiffSegment "equal" (SliceTokenText $left ($left.Count - $suffixCount) $suffixCount))) }
    return (MergeDiffSegments $segments)
}

function GetMidDiffSegments($OldTokens, $NewTokens) {
    $left = @($OldTokens)
    $right = @($NewTokens)
    if ($left.Count -eq 0 -and $right.Count -eq 0) { return @() }
    if ($left.Count -eq 0) { return (NewDiffSegment "ins" ((@($right) -join ""))) }
    if ($right.Count -eq 0) { return (NewDiffSegment "del" ((@($left) -join ""))) }

    $cellLimit = 4000000
    $cells = [int64]($left.Count + 1) * [int64]($right.Count + 1)
    if ($cells -gt $cellLimit) { return (GetFallbackDiffSegments $left $right) }

    $dp = New-Object 'int[,]' ($left.Count + 1), ($right.Count + 1)
    for ($i = 0; $i -le $left.Count; $i++) { $dp[$i, 0] = $i }
    for ($j = 0; $j -le $right.Count; $j++) { $dp[0, $j] = $j }

    for ($i = 1; $i -le $left.Count; $i++) {
        for ($j = 1; $j -le $right.Count; $j++) {
            if ($left[($i - 1)] -ceq $right[($j - 1)]) {
                $dp[$i, $j] = $dp[($i - 1), ($j - 1)]
            } else {
                $deleteCost = $dp[($i - 1), $j] + 1
                $insertCost = $dp[$i, ($j - 1)] + 1
                $replaceCost = $dp[($i - 1), ($j - 1)] + 2
                $dp[$i, $j] = [Math]::Min($replaceCost, [Math]::Min($deleteCost, $insertCost))
            }
        }
    }

    $units = New-Object System.Collections.Generic.List[object]
    $i = $left.Count
    $j = $right.Count
    while ($i -gt 0 -or $j -gt 0) {
        if ($i -gt 0 -and $j -gt 0 -and $left[($i - 1)] -ceq $right[($j - 1)] -and $dp[$i, $j] -eq $dp[($i - 1), ($j - 1)]) {
            $units.Add((NewDiffSegment "equal" ([string]$left[($i - 1)])))
            $i--
            $j--
        } elseif ($i -gt 0 -and $dp[$i, $j] -eq ($dp[($i - 1), $j] + 1)) {
            $units.Add((NewDiffSegment "del" ([string]$left[($i - 1)])))
            $i--
        } elseif ($j -gt 0 -and $dp[$i, $j] -eq ($dp[$i, ($j - 1)] + 1)) {
            $units.Add((NewDiffSegment "ins" ([string]$right[($j - 1)])))
            $j--
        } else {
            if ($j -gt 0) { $units.Add((NewDiffSegment "ins" ([string]$right[($j - 1)]))); $j-- }
            if ($i -gt 0) { $units.Add((NewDiffSegment "del" ([string]$left[($i - 1)]))); $i-- }
        }
    }

    $ordered = $units.ToArray()
    [array]::Reverse($ordered)
    return (MergeDiffSegments $ordered)
}

function GetMinimalDiffSegments([AllowNull()][string]$OldText, [AllowNull()][string]$NewText) {
    $left = if ($null -eq $OldText) { "" } else { [string]$OldText }
    $right = if ($null -eq $NewText) { "" } else { [string]$NewText }
    $leftTokens = @(TokenizeDiffText $left)
    $rightTokens = @(TokenizeDiffText $right)
    $prefixCount = CommonPrefixCount $leftTokens $rightTokens
    $suffixCount = CommonSuffixCount $leftTokens $rightTokens $prefixCount
    $segments = New-Object System.Collections.Generic.List[object]
    $leftMiddleCount = $leftTokens.Count - $prefixCount - $suffixCount
    $rightMiddleCount = $rightTokens.Count - $prefixCount - $suffixCount
    if ($prefixCount -gt 0) { $segments.Add((NewDiffSegment "equal" (SliceTokenText $leftTokens 0 $prefixCount))) }
    if ($leftMiddleCount -gt 0 -or $rightMiddleCount -gt 0) {
        $leftMiddle = if ($leftMiddleCount -gt 0) { @($leftTokens[$prefixCount..($prefixCount + $leftMiddleCount - 1)]) } else { @() }
        $rightMiddle = if ($rightMiddleCount -gt 0) { @($rightTokens[$prefixCount..($prefixCount + $rightMiddleCount - 1)]) } else { @() }
        foreach ($segment in @(GetMidDiffSegments $leftMiddle $rightMiddle)) { $segments.Add($segment) }
    }
    if ($suffixCount -gt 0) { $segments.Add((NewDiffSegment "equal" (SliceTokenText $leftTokens ($leftTokens.Count - $suffixCount) $suffixCount))) }
    return (CoalesceChangeWindows (MergeDiffSegments $segments))
}

function WriteRevisionSegments([System.Xml.XmlNode]$Paragraph, $Segments, [ref]$NextRevision, [string]$AuthorText, [string]$DateText, $RunProps) {
    $changed = $false
    foreach ($segment in @($Segments)) {
        if ($null -eq $segment -or [string]::IsNullOrEmpty([string]$segment.Text)) { continue }
        switch ($segment.Kind) {
            "equal" {
                AppendRun -Parent $Paragraph -Text ([string]$segment.Text) -RunProps $RunProps -Deleted $false
            }
            "del" {
                [void]$Paragraph.AppendChild((RevNode $Paragraph.OwnerDocument "del" ([string]$segment.Text) $NextRevision.Value $AuthorText $DateText $RunProps))
                $NextRevision.Value++
                $changed = $true
            }
            "ins" {
                [void]$Paragraph.AppendChild((RevNode $Paragraph.OwnerDocument "ins" ([string]$segment.Text) $NextRevision.Value $AuthorText $DateText $RunProps))
                $NextRevision.Value++
                $changed = $true
            }
            default {
                throw "不支持的 diff 片段类型: $($segment.Kind)"
            }
        }
    }
    $changed
}

function ParagraphText([System.Xml.XmlNode]$Paragraph) {
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($node in (SelectNodesNs -XmlObject $Paragraph -XPath ".//w:t | .//w:delText | .//w:instrText | .//w:tab | .//w:br | .//w:cr")) {
        switch ($node.LocalName) {
            "t" { $parts.Add($node.InnerText) }
            "delText" { $parts.Add($node.InnerText) }
            "instrText" { $parts.Add($node.InnerText) }
            "tab" { $parts.Add("`t") }
            default { $parts.Add("`n") }
        }
    }
    $parts -join ""
}

function ParagraphRefs([System.Xml.XmlDocument]$Doc) {
    $refs = New-Object System.Collections.Generic.List[object]
    $occurrences = @{}
    $paragraphNodes = @(SelectNodesNs -XmlObject $Doc -XPath "//w:p")
    if ($paragraphNodes.Count -eq 1 -and $paragraphNodes[0] -is [System.Array]) { $paragraphNodes = @($paragraphNodes[0]) }
    for ($i = 0; $i -lt $paragraphNodes.Count; $i++) {
        $p = $paragraphNodes[$i]
        $text = ParagraphText $p
        $normalized = Norm $text
        if (-not $occurrences.ContainsKey($normalized)) { $occurrences[$normalized] = 0 }
        $occurrences[$normalized] = [int]$occurrences[$normalized] + 1
        $styleNode = SelectOneNs -XmlObject $p -XPath "./w:pPr/w:pStyle"
        if ($styleNode -is [System.Array]) { $styleNode = $styleNode[0] }
        $styleId = if ($null -eq $styleNode) { "" } else { [string]$styleNode.GetAttribute("val", $script:W) }
        $refs.Add([pscustomobject]@{
            Index = $i
            Node = $p
            Text = $text
            Normalized = $normalized
            Occurrence = [int]$occurrences[$normalized]
            StyleId = $styleId
        })
    }
    return @($refs.ToArray())
}

function ResolvePathFromBaseFile([string]$BaseFilePath, [string]$CandidatePath) {
    if ([string]::IsNullOrWhiteSpace($CandidatePath)) { throw "路径不能为空" }
    if ([System.IO.Path]::IsPathRooted($CandidatePath)) { return FullPath $CandidatePath }
    $baseDirectory = Split-Path -Parent $BaseFilePath
    FullPath (Join-Path $baseDirectory $CandidatePath)
}

function TemplateCompareCatalog() {
    @(
        [pscustomobject]@{
            Name = "付款"
            Keywords = @("付款", "支付", "价款", "结算", "款项", "发票")
            Risk = "付款条件不清时，容易引发价款争议。"
            Suggestion = "参考自有模板补强付款触发条件和付款期限。"
        },
        [pscustomobject]@{
            Name = "验收"
            Keywords = @("验收", "交付", "测试", "交付标准", "交付成果")
            Risk = "验收标准不清时，容易引发交付争议。"
            Suggestion = "参考自有模板明确交付标准和验收机制。"
        },
        [pscustomobject]@{
            Name = "违约责任"
            Keywords = @("违约", "赔偿", "违约责任", "损失", "违约金")
            Risk = "违约责任偏弱时，守约方索赔空间会受限。"
            Suggestion = "参考自有模板补强违约责任和赔偿范围。"
        },
        [pscustomobject]@{
            Name = "解除"
            Keywords = @("解除", "终止", "解约")
            Risk = "解除条件不清时，退出机制容易失衡。"
            Suggestion = "参考自有模板明确解除条件和解除后责任。"
        },
        [pscustomobject]@{
            Name = "责任限制"
            Keywords = @("责任限制", "责任上限", "最高责任", "间接损失", "损失上限")
            Risk = "责任限制失衡时，风险承担可能明显偏离。"
            Suggestion = "参考自有模板校准责任上限和例外责任。"
        },
        [pscustomobject]@{
            Name = "知识产权"
            Keywords = @("知识产权", "著作权", "专利", "技术成果", "成果归属")
            Risk = "知识产权归属不清时，容易影响成果控制权。"
            Suggestion = "参考自有模板明确成果归属和侵权责任。"
        },
        [pscustomobject]@{
            Name = "保密"
            Keywords = @("保密", "秘密信息", "商业秘密", "披露")
            Risk = "保密义务不足时，容易造成信息泄露风险。"
            Suggestion = "参考自有模板补强保密范围和期限。"
        },
        [pscustomobject]@{
            Name = "数据合规"
            Keywords = @("数据", "个人信息", "隐私", "网络安全", "信息安全")
            Risk = "数据义务缺失时，容易产生合规风险。"
            Suggestion = "参考自有模板补强数据处理和安全责任。"
        },
        [pscustomobject]@{
            Name = "争议解决"
            Keywords = @("争议解决", "仲裁", "管辖", "法院", "法律适用")
            Risk = "争议解决约定偏弱时，争议处理成本会提高。"
            Suggestion = "参考自有模板明确争议解决路径和管辖。"
        }
    )
}

function CompareTextKey([AllowNull()][string]$Text) {
    $normalized = Norm $Text
    if ([string]::IsNullOrWhiteSpace($normalized)) { return "" }
    ([regex]::Replace($normalized.ToLowerInvariant(), "[^\p{L}\p{Nd}]+", ""))
}

function CompareTextBigrams([AllowNull()][string]$Text) {
    $clean = CompareTextKey $Text
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    if ([string]::IsNullOrWhiteSpace($clean)) { return ,$set }
    if ($clean.Length -lt 2) {
        [void]$set.Add($clean)
        return ,$set
    }
    for ($i = 0; $i -lt ($clean.Length - 1); $i++) {
        [void]$set.Add($clean.Substring($i, 2))
    }
    return ,$set
}

function CompareTextSimilarity([AllowNull()][string]$Left, [AllowNull()][string]$Right) {
    $leftNorm = Norm $Left
    $rightNorm = Norm $Right
    if ([string]::IsNullOrWhiteSpace($leftNorm) -or [string]::IsNullOrWhiteSpace($rightNorm)) { return 0.0 }
    if ($leftNorm -eq $rightNorm) { return 1.0 }
    $leftSet = CompareTextBigrams $leftNorm
    $rightSet = CompareTextBigrams $rightNorm
    if ($leftSet.Count -eq 0 -or $rightSet.Count -eq 0) { return 0.0 }
    $intersection = 0
    foreach ($item in $leftSet) {
        if ($rightSet.Contains($item)) { $intersection++ }
    }
    $union = $leftSet.Count + $rightSet.Count - $intersection
    if ($union -le 0) { return 0.0 }
    [double]$intersection / [double]$union
}

function CompareTopicScore([string]$Text, $Topic) {
    $score = 0
    foreach ($keyword in @($Topic.Keywords)) {
        if (-not [string]::IsNullOrWhiteSpace($keyword) -and $Text.Contains($keyword)) {
            $score++
        }
    }
    $score
}

function ResolveCompareTopic([string]$Text, $Catalog) {
    $bestTopic = $null
    $bestScore = 0
    foreach ($topic in $Catalog) {
        $score = CompareTopicScore -Text $Text -Topic $topic
        if ($score -gt $bestScore) {
            $bestScore = $score
            $bestTopic = $topic
        }
    }
    if ($bestScore -le 0) { return $null }
    $bestTopic
}

function ParagraphStyleId([System.Xml.XmlNode]$Paragraph) {
    $styleNode = SelectOneNs -XmlObject $Paragraph -XPath "./w:pPr/w:pStyle"
    if ($null -eq $styleNode) { return "" }
    $value = $styleNode.GetAttribute("val", $script:W)
    if ([string]::IsNullOrWhiteSpace($value)) { "" } else { $value }
}

function CompareParagraphRefsFromDocx([string]$DocxPath, $Catalog) {
    $workspace = Join-Path ([System.IO.Path]::GetTempPath()) ("compare-docx-" + [guid]::NewGuid().ToString("N"))
    try {
        ExpandDocx $DocxPath $workspace
        $documentPath = Join-Path $workspace "word\document.xml"
        if (-not (Test-Path -LiteralPath $documentPath)) { throw "文档缺少 word/document.xml: $DocxPath" }
        $doc = LoadXml $documentPath
        $paragraphs = ParagraphRefs $doc
        $occurrences = @{}
        $results = New-Object System.Collections.Generic.List[object]
        for ($i = 0; $i -lt $paragraphs.Count; $i++) {
            $ref = $paragraphs[$i]
            if ([string]::IsNullOrWhiteSpace($ref.Normalized)) { continue }
            $topic = ResolveCompareTopic -Text $ref.Text -Catalog $Catalog
            $normalizedText = [string]$ref.Normalized
            if (-not $occurrences.ContainsKey($normalizedText)) { $occurrences[$normalizedText] = 0 }
            $occurrences[$normalizedText] = [int]$occurrences[$normalizedText] + 1
            $results.Add([pscustomobject]@{
                Index = $i
                Text = [string]$ref.Text
                Normalized = $normalizedText
                Occurrence = [int]$occurrences[$normalizedText]
                TopicName = $(if ($null -ne $topic) { [string]$topic.Name } else { "" })
                StyleId = ParagraphStyleId $ref.Node
                SimilarityKey = CompareTextKey $ref.Text
            })
        }
        return @($results.ToArray())
    }
    finally {
        if (Test-Path -LiteralPath $workspace) { Remove-Item -LiteralPath $workspace -Recurse -Force }
    }
}

function IsClauseHeadingParagraph([string]$Text, [string]$StyleId) {
    $normalized = Norm $Text
    if ([string]::IsNullOrWhiteSpace($normalized)) { return $false }
    if (-not [string]::IsNullOrWhiteSpace($StyleId) -and $StyleId -match '(?i)heading|title') { return $true }
    if ($normalized -match '^(第[一二三四五六七八九十百零〇\d]+条)\b') { return $true }
    if ($normalized -match '^[一二三四五六七八九十百零〇]+、') { return $true }
    if ($normalized -match '^\(?\d+(?:\.\d+){0,3}\)?[\.\s、]') { return $true }
    if ($normalized -match '^（[一二三四五六七八九十百零〇]+）') { return $true }
    $false
}

function ClauseNumberKey([string]$Text) {
    $normalized = Norm $Text
    if ([string]::IsNullOrWhiteSpace($normalized)) { return "" }
    if ($normalized -match '^(第[一二三四五六七八九十百零〇\d]+条)') { return $matches[1] }
    if ($normalized -match '^([一二三四五六七八九十百零〇]+、)') { return $matches[1] }
    if ($normalized -match '^(\(?\d+(?:\.\d+){0,3}\)?)[\.\s、]') { return $matches[1] }
    if ($normalized -match '^(（[一二三四五六七八九十百零〇]+）)') { return $matches[1] }
    ""
}

function ClauseTitleText([string]$Text) {
    $normalized = Norm $Text
    if ([string]::IsNullOrWhiteSpace($normalized)) { return "" }
    $title = $normalized
    foreach ($pattern in @(
        '^(第[一二三四五六七八九十百零〇\d]+条)\s*',
        '^[一二三四五六七八九十百零〇]+、\s*',
        '^\(?\d+(?:\.\d+){0,3}\)?[\.\s、]*',
        '^（[一二三四五六七八九十百零〇]+）\s*'
    )) {
        $title = [regex]::Replace($title, $pattern, "")
    }
    $title.Trim()
}

function BuildClauseRefs($Paragraphs, $Catalog) {
    $Paragraphs = @($Paragraphs)
    $clauses = New-Object System.Collections.Generic.List[object]
    $currentClause = $null
    for ($i = 0; $i -lt $Paragraphs.Count; $i++) {
        $paragraph = $Paragraphs[$i]
        $styleId = if ($null -ne $paragraph.PSObject.Properties["StyleId"]) { [string]$paragraph.StyleId } else { "" }
        $isHeading = IsClauseHeadingParagraph -Text $paragraph.Text -StyleId $styleId
        if ($isHeading -or $null -eq $currentClause) {
            if ($null -ne $currentClause) {
                $bodyTexts = @($currentClause.Paragraphs | Where-Object { -not $_.IsHeading } | ForEach-Object { [string]$_.Text })
                $currentClause | Add-Member -NotePropertyName BodyText -NotePropertyValue ((@($bodyTexts) -join "`n").Trim()) -Force
                $fullTexts = @($currentClause.Paragraphs | ForEach-Object { [string]$_.Text })
                $fullText = ((@($fullTexts) -join "`n").Trim())
                $currentClause | Add-Member -NotePropertyName FullText -NotePropertyValue $fullText -Force
                $topic = ResolveCompareTopic -Text $fullText -Catalog $Catalog
                $currentClause | Add-Member -NotePropertyName TopicName -NotePropertyValue $(if ($null -ne $topic) { [string]$topic.Name } else { "" }) -Force
                $clauses.Add($currentClause)
            }
            $paragraphEntries = New-Object System.Collections.Generic.List[object]
            $paragraphEntries.Add([pscustomobject]@{
                Index = $paragraph.Index
                Text = [string]$paragraph.Text
                Normalized = [string]$paragraph.Normalized
                Occurrence = [int]$paragraph.Occurrence
                IsHeading = $isHeading
            }) | Out-Null
            $currentClause = [pscustomobject]@{
                StartIndex = [int]$paragraph.Index
                EndIndex = [int]$paragraph.Index
                HeadingText = [string]$paragraph.Text
                TitleText = (ClauseTitleText $paragraph.Text)
                NumberKey = (ClauseNumberKey $paragraph.Text)
                Paragraphs = $paragraphEntries
            }
        } else {
            $currentClause.EndIndex = [int]$paragraph.Index
            $currentClause.Paragraphs.Add([pscustomobject]@{
                Index = $paragraph.Index
                Text = [string]$paragraph.Text
                Normalized = [string]$paragraph.Normalized
                Occurrence = [int]$paragraph.Occurrence
                IsHeading = $false
            }) | Out-Null
        }
    }
    if ($null -ne $currentClause) {
        $bodyTexts = @($currentClause.Paragraphs | Where-Object { -not $_.IsHeading } | ForEach-Object { [string]$_.Text })
        $currentClause | Add-Member -NotePropertyName BodyText -NotePropertyValue ((@($bodyTexts) -join "`n").Trim()) -Force
        $fullTexts = @($currentClause.Paragraphs | ForEach-Object { [string]$_.Text })
        $fullText = ((@($fullTexts) -join "`n").Trim())
        $currentClause | Add-Member -NotePropertyName FullText -NotePropertyValue $fullText -Force
        $topic = ResolveCompareTopic -Text $fullText -Catalog $Catalog
        $currentClause | Add-Member -NotePropertyName TopicName -NotePropertyValue $(if ($null -ne $topic) { [string]$topic.Name } else { "" }) -Force
        $clauses.Add($currentClause)
    }
    return @($clauses.ToArray())
}

function RepresentativeClauseParagraph($Clause) {
    $bodyParagraphs = New-Object System.Collections.Generic.List[object]
    foreach ($entry in @($Clause.Paragraphs)) {
        if (-not $entry.IsHeading) { $bodyParagraphs.Add($entry) | Out-Null }
    }
    if ($bodyParagraphs.Count -gt 0) {
        return (@($bodyParagraphs.ToArray() | Sort-Object { ([string]$_.Normalized).Length } -Descending))[0]
    }
    (@($Clause.Paragraphs))[0]
}

function CompareClauseAlignment($SourceClause, $TemplateClause, [string]$ExpectedTopic) {
    $titleScore = CompareTextSimilarity -Left ([string]$SourceClause.TitleText) -Right ([string]$TemplateClause.TitleText)
    $bodyScore = CompareTextSimilarity -Left ([string]$SourceClause.BodyText) -Right ([string]$TemplateClause.BodyText)
    $fullScore = CompareTextSimilarity -Left ([string]$SourceClause.FullText) -Right ([string]$TemplateClause.FullText)
    $topicScore = if (-not [string]::IsNullOrWhiteSpace($ExpectedTopic) -and [string]$SourceClause.TopicName -eq $ExpectedTopic -and [string]$TemplateClause.TopicName -eq $ExpectedTopic) { 1.0 } else { 0.0 }
    $numberScore = if (-not [string]::IsNullOrWhiteSpace([string]$SourceClause.NumberKey) -and [string]$SourceClause.NumberKey -eq [string]$TemplateClause.NumberKey) { 1.0 } else { 0.0 }
    $confidence = ($titleScore * 0.25) + ($bodyScore * 0.35) + ($fullScore * 0.20) + ($topicScore * 0.10) + ($numberScore * 0.10)
    $reasons = New-Object System.Collections.Generic.List[string]
    if ($numberScore -gt 0) { $reasons.Add("条号一致") | Out-Null }
    if ($titleScore -ge 0.6) { $reasons.Add("标题接近") | Out-Null }
    if ($bodyScore -ge 0.45) { $reasons.Add("正文接近") | Out-Null }
    if ($topicScore -gt 0) { $reasons.Add("主题命中") | Out-Null }
    [pscustomobject]@{
        Confidence = [math]::Round($confidence, 4)
        TitleScore = [math]::Round($titleScore, 4)
        BodyScore = [math]::Round($bodyScore, 4)
        FullScore = [math]::Round($fullScore, 4)
        NumberScore = [math]::Round($numberScore, 4)
        TopicScore = [math]::Round($topicScore, 4)
        Reason = if ($reasons.Count -gt 0) { (@($reasons) -join "、") } else { "主题回退匹配" }
    }
}

function MissingClauseInsertionIndex($SourceClauses, $MatchedPairs, [string]$TopicName) {
    $pairs = @($MatchedPairs | Sort-Object SourceEndIndex)
    if ($pairs.Count -gt 0) {
        return [int]$pairs[($pairs.Count - 1)].SourceEndIndex
    }
    if (@($SourceClauses).Count -gt 0) {
        return [int](@($SourceClauses | Sort-Object EndIndex)[-1]).EndIndex
    }
    0
}

function ResolveCompareFocusTopics($CompareConfig, $Catalog) {
    $requested = @()
    if ($null -ne $CompareConfig.PSObject.Properties["focus_topics"]) {
        $requested = @($CompareConfig.focus_topics | ForEach-Object { ([string]$_).Trim() })
    } elseif ($null -ne $CompareConfig.PSObject.Properties["topics"]) {
        $requested = @($CompareConfig.topics | ForEach-Object { ([string]$_).Trim() })
    }
    if ($requested.Count -eq 0) {
        return @($Catalog | ForEach-Object { [string]$_.Name })
    }
    $valid = New-Object System.Collections.Generic.List[string]
    foreach ($name in $requested) {
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if (-not $valid.Contains($name)) { $valid.Add($name) }
    }
    if ($valid.Count -eq 0) { throw "template_compare.focus_topics 不能为空" }
    return @($valid.ToArray())
}

function NewTemplateCompareComment($Topic, [string]$Mode, [string]$TemplateText) {
    $suggestion = if ($Mode -eq "comment") {
        "该条款与自有模板的差异较大，建议人工确认交易口径后按自有模板补强。"
    } else {
        [string]$Topic.Suggestion
    }
    [ordered]@{
        "问题" = "对方合同在$($Topic.Name)条款上的表述与自有模板不一致，当前文本未采用自有模板的风控口径。"
        "风险" = [string]$Topic.Risk
        "修改建议" = $suggestion
        "建议条款" = $TemplateText
        "修改依据" = "依据自有模板同主题条款比对结果。"
    }
}

function OperationIdentity($Operation) {
    if ($null -ne $Operation.Location) {
        $hasParagraphIndex = $null -ne $Operation.Location.PSObject.Properties["paragraph_index"]
        $hasParagraph = $null -ne $Operation.Location.PSObject.Properties["paragraph"]
        $hasInsertionTexts = $null -ne $Operation.PSObject.Properties["InsertionTexts"] -and @($Operation.InsertionTexts).Count -gt 0
        if ($hasInsertionTexts) {
            $placement = if ($null -ne $Operation.PSObject.Properties["InsertionPlacement"] -and -not [string]::IsNullOrWhiteSpace([string]$Operation.InsertionPlacement)) { ([string]$Operation.InsertionPlacement).ToLowerInvariant() } else { "after" }
            $topic = if ($null -ne $Operation.PSObject.Properties["TopicName"] -and -not [string]::IsNullOrWhiteSpace([string]$Operation.TopicName)) { [string]$Operation.TopicName } else { "" }
            $joinedInsertion = [string]::Join("`n", @($Operation.InsertionTexts | ForEach-Object { Norm ([string]$_) }))
            $textHash = if ([string]::IsNullOrWhiteSpace($joinedInsertion)) {
                ""
            } else {
                $sha = [System.Security.Cryptography.SHA256]::Create()
                try {
                    [System.BitConverter]::ToString($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($joinedInsertion))).Replace("-", "").ToLowerInvariant()
                }
                finally {
                    $sha.Dispose()
                }
            }
            if ($hasParagraphIndex) { return "loc:index:$([int]$Operation.Location.paragraph_index):insert:${placement}:${topic}:${textHash}" }
            if ($hasParagraph) { return "loc:paragraph:$([int]$Operation.Location.paragraph):insert:${placement}:${topic}:${textHash}" }
        }
        if ($hasParagraphIndex) { return "loc:index:$([int]$Operation.Location.paragraph_index)" }
        if ($hasParagraph) { return "loc:paragraph:$([int]$Operation.Location.paragraph)" }
    }
    $anchor = Norm ([string]$Operation.AnchorText)
    $matchType = if ([string]::IsNullOrWhiteSpace([string]$Operation.MatchType)) { "exact" } else { ([string]$Operation.MatchType).ToLowerInvariant() }
    $occurrence = if ($null -eq $Operation.Occurrence) { 1 } else { [int]$Operation.Occurrence }
    "anchor:${matchType}:${occurrence}:${anchor}"
}

function GenerateTemplateCompareOperations([string]$SourceDocx, [string]$InstructionsPathForResolution, $CompareConfig, [string]$EffectiveAuthor, [string]$EffectiveDate, [string]$DefaultMode) {
    $templateProperty = if ($null -ne $CompareConfig.PSObject.Properties["template_path"]) {
        [string]$CompareConfig.template_path
    } elseif ($null -ne $CompareConfig.PSObject.Properties["template"]) {
        [string]$CompareConfig.template
    } else {
        $null
    }
    if ([string]::IsNullOrWhiteSpace($templateProperty)) { throw "template_compare 缺少 template_path" }
    $templatePath = ResolvePathFromBaseFile -BaseFilePath $InstructionsPathForResolution -CandidatePath $templateProperty
    if (-not (Test-Path -LiteralPath $templatePath)) { throw "template_compare 模板文件不存在: $templatePath" }
    if ([System.IO.Path]::GetExtension($templatePath).ToLowerInvariant() -ne ".docx") { throw "template_compare 模板必须是 .docx 文件" }

    $catalog = @(TemplateCompareCatalog)
    $focusTopics = @(ResolveCompareFocusTopics -CompareConfig $CompareConfig -Catalog $catalog)
    $sourceRefs = @(CompareParagraphRefsFromDocx -DocxPath $SourceDocx -Catalog $catalog)
    $templateRefs = @(CompareParagraphRefsFromDocx -DocxPath $templatePath -Catalog $catalog)
    $sourceClauses = @(BuildClauseRefs -Paragraphs $sourceRefs -Catalog $catalog)
    $templateClauses = @(BuildClauseRefs -Paragraphs $templateRefs -Catalog $catalog)
    $maxOperations = if ($null -ne $CompareConfig.PSObject.Properties["max_operations"]) { [int]$CompareConfig.max_operations } else { 8 }
    if ($maxOperations -le 0) { throw "template_compare.max_operations 必须为正整数" }
    $revisionMode = if ($null -ne $CompareConfig.PSObject.Properties["mode"] -and -not [string]::IsNullOrWhiteSpace([string]$CompareConfig.mode)) {
        ([string]$CompareConfig.mode).ToLowerInvariant()
    } else {
        $DefaultMode.ToLowerInvariant()
    }
    if ($revisionMode -notin @("comment", "revision", "revision_comment")) { throw "template_compare.mode 不合法: $revisionMode" }
    $revisionThreshold = if ($null -ne $CompareConfig.PSObject.Properties["min_similarity_for_revision"]) { [double]$CompareConfig.min_similarity_for_revision } else { 0.35 }
    $minAlignmentConfidence = if ($null -ne $CompareConfig.PSObject.Properties["min_alignment_confidence"]) { [double]$CompareConfig.min_alignment_confidence } else { 0.45 }
    $allowMissingClauseInsert = $false
    if ($null -ne $CompareConfig.PSObject.Properties["allow_missing_clause_insert"]) { $allowMissingClauseInsert = [bool]$CompareConfig.allow_missing_clause_insert }
    $missingClauseMode = if ($null -ne $CompareConfig.PSObject.Properties["missing_clause_mode"] -and -not [string]::IsNullOrWhiteSpace([string]$CompareConfig.missing_clause_mode)) {
        ([string]$CompareConfig.missing_clause_mode).ToLowerInvariant()
    } elseif ($allowMissingClauseInsert) {
        "revision_comment"
    } else {
        "comment"
    }
    if ($missingClauseMode -notin @("comment", "revision", "revision_comment")) { throw "template_compare.missing_clause_mode 不合法: $missingClauseMode" }
    $operations = New-Object System.Collections.Generic.List[object]
    $matchedTopics = New-Object System.Collections.Generic.List[string]
    $missingTopics = New-Object System.Collections.Generic.List[string]
    $usedSource = New-Object 'System.Collections.Generic.HashSet[int]'
    $matchedPairs = New-Object System.Collections.Generic.List[object]

    foreach ($topicName in $focusTopics) {
        if ($operations.Count -ge $maxOperations) { break }
        $topic = $null
        foreach ($topicItem in $catalog) {
            if ([string]$topicItem.Name -eq [string]$topicName) {
                $topic = $topicItem
                break
            }
        }
        if ($null -eq $topic) { continue }
        $sourceCandidates = @($sourceClauses | Where-Object { $_.TopicName -eq $topicName } | Sort-Object StartIndex)
        $templateCandidates = @($templateClauses | Where-Object { $_.TopicName -eq $topicName } | Sort-Object StartIndex)
        if ($templateCandidates.Count -eq 0) { continue }
        if ($sourceCandidates.Count -eq 0) {
            $missingTopics.Add($topicName)
            continue
        }
        $usedTemplate = New-Object 'System.Collections.Generic.HashSet[int]'
        foreach ($sourceCandidate in $sourceCandidates) {
            if ($operations.Count -ge $maxOperations) { break }
            if (-not $usedSource.Add([int]$sourceCandidate.StartIndex)) { continue }
            $bestTemplate = $null
            $bestTemplateSlot = -1
            $bestAlignment = $null
            for ($slot = 0; $slot -lt $templateCandidates.Count; $slot++) {
                if ($usedTemplate.Contains($slot)) { continue }
                $templateCandidate = $templateCandidates[$slot]
                $alignment = CompareClauseAlignment -SourceClause $sourceCandidate -TemplateClause $templateCandidate -ExpectedTopic $topicName
                if ($null -eq $bestAlignment -or $alignment.Confidence -gt $bestAlignment.Confidence) {
                    $bestAlignment = $alignment
                    $bestTemplate = $templateCandidate
                    $bestTemplateSlot = $slot
                }
            }
            if ($null -eq $bestTemplate) { continue }
            [void]$usedTemplate.Add($bestTemplateSlot)
            if ((Norm $sourceCandidate.FullText) -eq (Norm $bestTemplate.FullText)) { continue }
            $sourceParagraph = RepresentativeClauseParagraph $sourceCandidate
            $templateParagraph = RepresentativeClauseParagraph $bestTemplate
            if ((Norm $sourceParagraph.Text) -eq (Norm $templateParagraph.Text) -and (Norm $sourceCandidate.FullText) -ne (Norm $bestTemplate.FullText)) {
                continue
            }
            $mode = if ($revisionMode -eq "comment") {
                "comment"
            } elseif ($bestAlignment.Confidence -lt $minAlignmentConfidence -or $bestAlignment.BodyScore -lt $revisionThreshold) {
                "comment"
            } else {
                $revisionMode
            }
            $comment = NewTemplateCompareComment -Topic $topic -Mode $mode -TemplateText ([string]$bestTemplate.FullText)
            $comment["修改依据"] = "依据自有模板同主题条款比对结果；对齐依据：$($bestAlignment.Reason)；置信度：$($bestAlignment.Confidence)。"
            $operations.Add([pscustomobject]@{
                AnchorText = [string]$sourceParagraph.Text
                Location = $null
                Mode = $mode
                ReplacementText = $(if ($mode -in @("revision", "revision_comment")) { [string]$templateParagraph.Text } else { $null })
                Comment = $comment
                CommentLines = @(
                    "问题：$($comment["问题"])",
                    "风险：$($comment["风险"])",
                    "修改建议：$($comment["修改建议"])",
                    "建议条款：$($comment["建议条款"])",
                    "修改依据：$($comment["修改依据"])"
                )
                MatchType = "exact"
                Occurrence = [int]$sourceParagraph.Occurrence
            })
            $matchedPairs.Add([pscustomobject]@{
                TopicName = $topicName
                SourceEndIndex = [int]$sourceCandidate.EndIndex
            }) | Out-Null
            if (-not $matchedTopics.Contains($topicName)) { $matchedTopics.Add($topicName) }
        }
    }

    foreach ($topicName in $missingTopics) {
        if ($operations.Count -ge $maxOperations) { break }
        $templateClause = $null
        foreach ($candidate in @($templateClauses | Where-Object { $_.TopicName -eq $topicName } | Sort-Object StartIndex)) {
            $templateClause = $candidate
            break
        }
        if ($null -eq $templateClause) { continue }
        $topic = $null
        foreach ($candidate in $catalog) {
            if ([string]$candidate.Name -eq $topicName) {
                $topic = $candidate
                break
            }
        }
        if ($null -eq $topic) { continue }
        $comment = [ordered]@{
            "问题" = "对方合同缺少自有模板中的$topicName条款。"
            "风险" = [string]$topic.Risk
            "修改建议" = $(if ($allowMissingClauseInsert) { "按自有模板新增该条款，并保留修订痕迹。"} else { "建议结合交易背景补充该条款，当前先提示风险。"})
            "建议条款" = [string]$templateClause.FullText
            "修改依据" = "依据自有模板同主题条款比对结果。"
        }
        $insertIndex = MissingClauseInsertionIndex -SourceClauses $sourceClauses -MatchedPairs @($matchedPairs.ToArray()) -TopicName $topicName
        $location = [pscustomobject]@{ paragraph_index = $insertIndex }
        if ($allowMissingClauseInsert -and $missingClauseMode -in @("revision", "revision_comment")) {
            $operations.Add([pscustomobject]@{
                AnchorText = $null
                Location = $location
                Mode = $missingClauseMode
                ReplacementText = $null
                Comment = $comment
                CommentLines = @(
                    "问题：$($comment["问题"])",
                    "风险：$($comment["风险"])",
                    "修改建议：$($comment["修改建议"])",
                    "建议条款：$($comment["建议条款"])",
                    "修改依据：$($comment["修改依据"])"
                )
                MatchType = "exact"
                Occurrence = 1
                InsertionTexts = @($templateClause.Paragraphs | ForEach-Object { [string]$_.Text })
                InsertionPlacement = "after"
            })
        } else {
            $anchorParagraph = RepresentativeClauseParagraph (@($sourceClauses | Sort-Object EndIndex)[-1])
            if ($null -eq $anchorParagraph) { continue }
            $operations.Add([pscustomobject]@{
                AnchorText = [string]$anchorParagraph.Text
                Location = $null
                Mode = "comment"
                ReplacementText = $null
                Comment = $comment
                CommentLines = @(
                    "问题：$($comment["问题"])",
                    "风险：$($comment["风险"])",
                    "修改建议：$($comment["修改建议"])",
                    "建议条款：$($comment["建议条款"])",
                    "修改依据：$($comment["修改依据"])"
                )
                MatchType = "exact"
                Occurrence = [int]$anchorParagraph.Occurrence
            })
        }
    }

    if ($operations.Count -eq 0) {
        $pairCount = [Math]::Min(@($sourceRefs).Count, @($templateRefs).Count)
        for ($i = 0; $i -lt $pairCount; $i++) {
            if ($operations.Count -ge $maxOperations) { break }
            $sourceCandidate = $sourceRefs[$i]
            $templateCandidate = $templateRefs[$i]
            if ((Norm $sourceCandidate.Text) -eq (Norm $templateCandidate.Text)) { continue }
            $topic = ResolveCompareTopic -Text ("$($sourceCandidate.Text) $($templateCandidate.Text)") -Catalog $catalog
            if ($null -eq $topic) {
                $topic = [pscustomobject]@{
                    Name = "对应条款"
                    Risk = "该条款与自有模板存在差异，可能导致风险控制口径不一致。"
                    Suggestion = "参考自有模板对对应条款进行补强。"
                }
            }
            $comment = NewTemplateCompareComment -Topic $topic -Mode $revisionMode -TemplateText ([string]$templateCandidate.Text)
            $operations.Add([pscustomobject]@{
                AnchorText = [string]$sourceCandidate.Text
                Location = $null
                Mode = $revisionMode
                ReplacementText = $(if ($revisionMode -in @("revision", "revision_comment")) { [string]$templateCandidate.Text } else { $null })
                Comment = $comment
                CommentLines = @(
                    "问题：$($comment["问题"])",
                    "风险：$($comment["风险"])",
                    "修改建议：$($comment["修改建议"])",
                    "建议条款：$($comment["建议条款"])",
                    "修改依据：$($comment["修改依据"])"
                )
                MatchType = "exact"
                Occurrence = [int]$sourceCandidate.Occurrence
            })
        }
    }

    [pscustomobject]@{
        TemplatePath = $templatePath
        Operations = @($operations.ToArray())
        FocusTopics = @($focusTopics)
        MatchedTopics = @($matchedTopics.ToArray())
        MissingTopics = @($missingTopics.ToArray())
        AlignmentConfidence = $minAlignmentConfidence
        MissingClauseInsertionEnabled = $allowMissingClauseInsert
    }
}

function ResolveParagraphTargetIndex($Paragraphs, $Location) {
    $Paragraphs = @($Paragraphs)
    if ($null -eq $Location) { return $null }
    $hasParagraphIndex = $null -ne $Location.PSObject.Properties["paragraph_index"]
    $hasParagraph = $null -ne $Location.PSObject.Properties["paragraph"]
    if (-not $hasParagraphIndex -and -not $hasParagraph) { throw "location 缺少 paragraph 或 paragraph_index" }

    if ($hasParagraphIndex) {
        $index = [int]$Location.paragraph_index
    } else {
        $paragraphNumber = [int]$Location.paragraph
        if ($paragraphNumber -le 0) { throw "location.paragraph 必须为正整数: $paragraphNumber" }
        $index = $paragraphNumber - 1
    }

    if ($index -lt 0 -or $index -ge $Paragraphs.Count) { throw "段落索引超出范围: $index" }
    return $index
}

function FindTargets($Paragraphs, $Operations) {
    $Paragraphs = @($Paragraphs)
    $Operations = @($Operations)
    $targets = New-Object System.Collections.Generic.List[object]
    $used = New-Object System.Collections.Generic.HashSet[int]
    foreach ($op in $Operations) {
        $idx = ResolveParagraphTargetIndex -Paragraphs $Paragraphs -Location $op.Location
        if ($null -eq $idx) {
            $matched = New-Object System.Collections.Generic.List[int]
            for ($i = 0; $i -lt $Paragraphs.Count; $i++) {
                $ok = if ($op.MatchType -eq "contains") { $Paragraphs[$i].Normalized.Contains((Norm $op.AnchorText)) } else { $Paragraphs[$i].Normalized -eq (Norm $op.AnchorText) }
                if ($ok) { $matched.Add($i) }
            }
            if ($matched.Count -lt $op.Occurrence) { throw "未找到第 $($op.Occurrence) 个匹配段落: $($op.AnchorText)" }
            $idx = $matched[$op.Occurrence - 1]
        }
        $allowsDuplicateTarget = $null -ne $op.PSObject.Properties["InsertionTexts"] -and @($op.InsertionTexts).Count -gt 0
        if (-not $allowsDuplicateTarget -and -not $used.Add($idx)) {
            $label = if ($null -ne $op.Location) { "段落索引 $idx" } else { [string]$op.AnchorText }
            throw "同一段落被重复命中: $label"
        }
        $targets.Add($Paragraphs[$idx])
    }
    return ,($targets.ToArray())
}

function AddCommentNode([System.Xml.XmlDocument]$Doc, [int]$Id, [string]$AuthorText, [string]$DateText, [string[]]$Lines) {
    $comment = WNode $Doc "comment"
    SetWAttr $comment "id" "$Id"
    SetWAttr $comment "author" $AuthorText
    SetWAttr $comment "date" $DateText
    foreach ($line in $Lines) {
        $p = WNode $Doc "p"
        $r = WNode $Doc "r"
        $rPr = WNode $Doc "rPr"
        $style = WNode $Doc "rStyle"
        SetWAttr $style "val" "CommentText"
        [void]$rPr.AppendChild($style)
        AddCommentFontProps -RunProps $rPr
        [void]$r.AppendChild($rPr)
        AddRunText -Run $r -Text $line -Deleted $false
        [void]$p.AppendChild($r)
        [void]$comment.AppendChild($p)
    }
    [void]$Doc.DocumentElement.AppendChild($comment)
}

function CRStart([System.Xml.XmlDocument]$Doc, [int]$Id) { $n = WNode $Doc "commentRangeStart"; SetWAttr $n "id" "$Id"; $n }
function CREnd([System.Xml.XmlDocument]$Doc, [int]$Id) { $n = WNode $Doc "commentRangeEnd"; SetWAttr $n "id" "$Id"; $n }
function CRefRun([System.Xml.XmlDocument]$Doc, [int]$Id) {
    $r = WNode $Doc "r"; $rPr = WNode $Doc "rPr"; $style = WNode $Doc "rStyle"; SetWAttr $style "val" "CommentReference"; [void]$rPr.AppendChild($style); [void]$r.AppendChild($rPr)
    $ref = WNode $Doc "commentReference"; SetWAttr $ref "id" "$Id"; [void]$r.AppendChild($ref); $r
}

function EnsureRelationship([System.Xml.XmlDocument]$Doc, [string]$Type, [string]$Target) {
    if ($null -ne (SelectOneNs -XmlObject $Doc -XPath "/pr:Relationships/pr:Relationship[@Type='$Type' and @Target='$Target']")) { return }
    $ids = @()
    foreach ($rel in (SelectNodesNs -XmlObject $Doc -XPath "/pr:Relationships/pr:Relationship")) { $ids += ([string]$rel.Attributes["Id"].Value).Replace("rId", "") }
    $node = $Doc.CreateElement("Relationship", $script:PR)
    [void]$node.SetAttribute("Id", "rId$(NextId $ids)")
    [void]$node.SetAttribute("Type", $Type)
    [void]$node.SetAttribute("Target", $Target)
    [void]$Doc.DocumentElement.AppendChild($node)
}

function EnsureOverride([System.Xml.XmlDocument]$Doc, [string]$PartName, [string]$ContentType) {
    $existing = SelectOneNs -XmlObject $Doc -XPath "/ct:Types/ct:Override[@PartName='$PartName']"
    if ($null -ne $existing) { [void]$existing.SetAttribute("ContentType", $ContentType); return }
    $node = $Doc.CreateElement("Override", $script:CT)
    [void]$node.SetAttribute("PartName", $PartName)
    [void]$node.SetAttribute("ContentType", $ContentType)
    [void]$Doc.DocumentElement.AppendChild($node)
}

function EnsureComments([string]$Workspace) {
    $commentsPath = Join-Path $Workspace "word\comments.xml"
    $relsPath = Join-Path $Workspace "word\_rels\document.xml.rels"
    $typesPath = Join-Path $Workspace "[Content_Types].xml"
    $doc = if (Test-Path $commentsPath) { LoadXml $commentsPath } else { NewDoc "comments" $script:W "w" }
    [System.IO.Directory]::CreateDirectory((Split-Path -Parent $relsPath)) | Out-Null
    $rels = if (Test-Path $relsPath) { LoadXml $relsPath } else { NewDoc "Relationships" $script:PR "" }
    EnsureRelationship $rels $script:CommentsRel "comments.xml"
    SaveXml $rels $relsPath
    $types = LoadXml $typesPath
    EnsureOverride $types "/word/comments.xml" $script:CommentsType
    SaveXml $types $typesPath
    SaveXml $doc $commentsPath
    $doc
}

function EnsureSettings([string]$Workspace) {
    $settingsPath = Join-Path $Workspace "word\settings.xml"
    $relsPath = Join-Path $Workspace "word\_rels\document.xml.rels"
    $typesPath = Join-Path $Workspace "[Content_Types].xml"
    $doc = if (Test-Path $settingsPath) { LoadXml $settingsPath } else { NewDoc "settings" $script:W "w" }
    if ($null -eq (SelectOneNs -XmlObject $doc -XPath "/w:settings/w:trackRevisions")) { [void]$doc.DocumentElement.AppendChild((WNode $doc "trackRevisions")) }
    if ($null -eq (SelectOneNs -XmlObject $doc -XPath "/w:settings/w:showRevisions")) { [void]$doc.DocumentElement.AppendChild((WNode $doc "showRevisions")) }
    [System.IO.Directory]::CreateDirectory((Split-Path -Parent $relsPath)) | Out-Null
    $rels = if (Test-Path $relsPath) { LoadXml $relsPath } else { NewDoc "Relationships" $script:PR "" }
    EnsureRelationship $rels $script:SettingsRel "settings.xml"
    SaveXml $rels $relsPath
    $types = LoadXml $typesPath
    EnsureOverride $types "/word/settings.xml" $script:SettingsType
    SaveXml $types $typesPath
    SaveXml $doc $settingsPath
}

function CurrentCommentId([System.Xml.XmlDocument]$Doc) {
    $ids = @()
    foreach ($node in (SelectNodesNs -XmlObject $Doc -XPath "/w:comments/w:comment")) { $ids += [string]$node.Attributes.GetNamedItem("id", $script:W).Value }
    NextId $ids
}

function CurrentRevisionId([System.Xml.XmlDocument]$Doc) {
    $ids = @()
    foreach ($path in @("//w:ins", "//w:del")) {
        foreach ($node in (SelectNodesNs -XmlObject $Doc -XPath $path)) {
            $attr = $node.Attributes.GetNamedItem("id", $script:W)
            if ($null -ne $attr) { $ids += [string]$attr.Value }
        }
    }
    NextId $ids
}

function InsertParagraphAfter([System.Xml.XmlNode]$AnchorParagraph, [System.Xml.XmlNode]$NewParagraph) {
    $parent = $AnchorParagraph.ParentNode
    $nextSibling = $AnchorParagraph.NextSibling
    if ($null -eq $nextSibling) {
        [void]$parent.AppendChild($NewParagraph)
    } else {
        [void]$parent.InsertBefore($NewParagraph, $nextSibling)
    }
}

function InsertParagraphBefore([System.Xml.XmlNode]$AnchorParagraph, [System.Xml.XmlNode]$NewParagraph) {
    $parent = $AnchorParagraph.ParentNode
    [void]$parent.InsertBefore($NewParagraph, $AnchorParagraph)
}

function NewInsertedParagraph([System.Xml.XmlDocument]$Doc, [string]$Text, [string]$Mode, [Nullable[int]]$CommentId, [int]$NextRevisionId, [string]$AuthorText, [string]$DateText, $RunProps, $ParagraphProps) {
    $p = WNode $Doc "p"
    $commentValue = if ($null -ne $CommentId) { [int]$CommentId } else { $null }
    if ($null -ne $ParagraphProps) { [void]$p.AppendChild($Doc.ImportNode($ParagraphProps, $true)) }
    if ($null -ne $commentValue) { [void]$p.AppendChild((CRStart $Doc $commentValue)) }
    if ($Mode -in @("revision", "revision_comment")) {
        [void]$p.AppendChild((RevNode $Doc "ins" $Text $NextRevisionId $AuthorText $DateText $RunProps))
    } else {
        AppendRun -Parent $p -Text $Text -RunProps $RunProps -Deleted $false
    }
    if ($null -ne $commentValue) {
        [void]$p.AppendChild((CREnd $Doc $commentValue))
        [void]$p.AppendChild((CRefRun $Doc $commentValue))
    }
    $p
}

function Reset-Workspace([string]$Workspace) {
    if (Test-Path -LiteralPath $Workspace) { Remove-Item -LiteralPath $Workspace -Recurse -Force }
    [System.IO.Directory]::CreateDirectory($Workspace) | Out-Null
}

function Supports-ShellZipFallback {
    if ($null -ne (Get-Variable -Name IsWindows -ErrorAction SilentlyContinue)) {
        return [bool]$IsWindows
    }
    [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Windows)
}

function Resolve-PythonExecutable {
    foreach ($candidate in @("python", "python3")) {
        $command = Get-Command -Name $candidate -ErrorAction SilentlyContinue | Where-Object { $_.CommandType -in @("Application", "ExternalScript") } | Select-Object -First 1
        if ($null -ne $command -and -not [string]::IsNullOrWhiteSpace([string]$command.Definition)) {
            return [string]$command.Definition
        }
    }
    throw "未找到可用的 Python 可执行文件（python / python3）"
}

function ExpandDocxWithDotNet([string]$Docx, [string]$Workspace) {
    Ensure-DotNetCompression
    [System.IO.Compression.ZipFile]::ExtractToDirectory($Docx, $Workspace)
}

function ExpandDocxWithShell([string]$Docx, [string]$Workspace) {
    $shell = New-Object -ComObject Shell.Application
    try {
        $sourceNs = $shell.NameSpace($Docx)
        $targetNs = $shell.NameSpace($Workspace)
        if ($null -eq $sourceNs -or $null -eq $targetNs) {
            throw "Shell ZIP 后端无法打开源或目标路径。"
        }
        $targetNs.CopyHere($sourceNs.Items(), 16)
        $deadline = (Get-Date).AddSeconds(15)
        while ((Get-Date) -lt $deadline) {
            if (Test-Path (Join-Path $Workspace "word\document.xml")) { break }
            Start-Sleep -Milliseconds 200
        }
        if (-not (Test-Path (Join-Path $Workspace "word\document.xml"))) {
            throw "Shell ZIP 后端解压超时。"
        }
    }
    finally {
        if ($shell -is [System.__ComObject]) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
        }
    }
}

function ExpandDocx([string]$Docx, [string]$Workspace) {
    $dotNetError = $null
    Reset-Workspace $Workspace
    try {
        ExpandDocxWithDotNet $Docx $Workspace
        return
    }
    catch {
        $dotNetError = $_
    }

    if (-not (Supports-ShellZipFallback)) {
        throw "DOCX 解压失败。.NET ZipArchive: $($dotNetError.Exception.Message)。当前平台不支持 Shell.Application 回退，请在 macOS/Linux 上优先使用 pwsh + .NET ZipArchive。"
    }

    Reset-Workspace $Workspace
    try {
        ExpandDocxWithShell $Docx $Workspace
        return
    }
    catch {
        throw "DOCX 解压失败。.NET ZipArchive: $($dotNetError.Exception.Message)；Shell.Application: $($_.Exception.Message)"
    }
}

function PackDocxWithDotNet([string]$Workspace, [string]$Docx) {
    Ensure-DotNetCompression
    $workspaceRoot = [System.IO.Path]::GetFullPath($Workspace)
    if (-not $workspaceRoot.EndsWith([System.IO.Path]::DirectorySeparatorChar)) {
        $workspaceRoot += [System.IO.Path]::DirectorySeparatorChar
    }
    $stream = [System.IO.File]::Open($Docx, [System.IO.FileMode]::Create)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create, $false)
        try {
            Get-ChildItem -LiteralPath $Workspace -Recurse -File | Sort-Object FullName | ForEach-Object {
                $rel = $_.FullName.Substring($workspaceRoot.Length).Replace("\", "/")
                $entry = $zip.CreateEntry($rel, [System.IO.Compression.CompressionLevel]::Optimal)
                $target = $entry.Open()
                try { $file = [System.IO.File]::OpenRead($_.FullName); try { $file.CopyTo($target) } finally { $file.Dispose() } } finally { $target.Dispose() }
            }
        } finally { $zip.Dispose() }
    } finally { $stream.Dispose() }
}

function PackDocxWithShell([string]$Workspace, [string]$Docx) {
    [System.IO.File]::WriteAllBytes($Docx, [byte[]](80,75,5,6,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0))
    $shell = New-Object -ComObject Shell.Application
    try {
        $zipNs = $shell.NameSpace($Docx)
        $srcNs = $shell.NameSpace($Workspace)
        if ($null -eq $zipNs -or $null -eq $srcNs) {
            throw "Shell ZIP 后端无法创建 ZIP 容器。"
        }
        $zipNs.CopyHere($srcNs.Items(), 16)
        $deadline = (Get-Date).AddSeconds(15)
        while ((Get-Date) -lt $deadline) {
            $probe = $shell.NameSpace($Docx)
            if ($null -ne $probe -and $null -ne $probe.ParseName("[Content_Types].xml") -and $null -ne $probe.ParseName("word")) {
                break
            }
            Start-Sleep -Milliseconds 250
        }
    }
    finally {
        if ($shell -is [System.__ComObject]) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
        }
    }
}

function PackDocx([string]$Workspace, [string]$Docx) {
    if (Test-Path $Docx) { Remove-Item -LiteralPath $Docx -Force }
    $parent = Split-Path -Parent $Docx
    if ($parent) { [System.IO.Directory]::CreateDirectory($parent) | Out-Null }
    $dotNetError = $null
    try {
        PackDocxWithDotNet $Workspace $Docx
        return
    }
    catch {
        $dotNetError = $_
    }

    if (-not (Supports-ShellZipFallback)) {
        throw "DOCX 回包失败。.NET ZipArchive: $($dotNetError.Exception.Message)。当前平台不支持 Shell.Application 回退，请在 macOS/Linux 上优先使用 pwsh + .NET ZipArchive。"
    }

    if (Test-Path -LiteralPath $Docx) { Remove-Item -LiteralPath $Docx -Force }
    try {
        PackDocxWithShell $Workspace $Docx
        return
    }
    catch {
        throw "DOCX 回包失败。.NET ZipArchive: $($dotNetError.Exception.Message)；Shell.Application: $($_.Exception.Message)"
    }
}

function ApplyReviewed([string]$Workspace, [object[]]$Operations, [string]$AuthorText, [string]$DateText) {
    $documentPath = Join-Path $Workspace "word\document.xml"
    $doc = LoadXml $documentPath
    $paragraphs = ParagraphRefs $doc
    $targets = FindTargets -Paragraphs $paragraphs -Operations $Operations
    $commentsDoc = $null
    $nextComment = 0
    if (@($Operations | Where-Object { $_.Mode -in @("comment", "revision_comment") }).Count -gt 0) { $commentsDoc = EnsureComments $Workspace; $nextComment = CurrentCommentId $commentsDoc }
    if (@($Operations | Where-Object { $_.Mode -in @("revision", "revision_comment") }).Count -gt 0) { EnsureSettings $Workspace }
    $nextRevision = CurrentRevisionId $doc
    $commentCount = 0
    $revisionCount = 0
    for ($i = 0; $i -lt $Operations.Count; $i++) {
        $op = $Operations[$i]
        $target = $targets[$i]
        $p = $target.Node
        $insertionTexts = if ($null -ne $op.PSObject.Properties["InsertionTexts"]) { @($op.InsertionTexts | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) } else { @() }
        if (@($insertionTexts).Count -gt 0) {
            $insertionPlacement = if ($null -ne $op.PSObject.Properties["InsertionPlacement"] -and -not [string]::IsNullOrWhiteSpace([string]$op.InsertionPlacement)) { ([string]$op.InsertionPlacement).ToLowerInvariant() } else { "after" }
            $anchorRunProps = RunPropsClone $p
            $anchorSnapshot = SnapshotParagraph $p
            if ($null -ne $anchorSnapshot.PPr) { [void]$p.AppendChild($p.OwnerDocument.ImportNode($anchorSnapshot.PPr, $true)) }
            if ($anchorSnapshot.Content.Count -gt 0) {
                foreach ($node in $anchorSnapshot.Content) { [void]$p.AppendChild($p.OwnerDocument.ImportNode($node, $true)) }
            } else {
                AppendRun $p $target.Text $anchorRunProps $false
            }
            $commentId = $null
            if ($op.Mode -in @("comment", "revision_comment")) {
                $commentId = $nextComment; $nextComment++
                AddCommentNode -Doc $commentsDoc -Id $commentId -AuthorText $AuthorText -DateText $DateText -Lines $op.CommentLines
                $commentCount++
            }
            $insertAnchor = $p
            $paragraphProps = if ($null -ne $anchorSnapshot.PPr) { $p.OwnerDocument.ImportNode($anchorSnapshot.PPr, $true) } else { $null }
            for ($insertIndex = 0; $insertIndex -lt $insertionTexts.Count; $insertIndex++) {
                $text = $insertionTexts[$insertIndex]
                $newParagraph = NewInsertedParagraph -Doc $p.OwnerDocument -Text $text -Mode $op.Mode -CommentId $commentId -NextRevisionId $nextRevision -AuthorText $AuthorText -DateText $DateText -RunProps $anchorRunProps -ParagraphProps $paragraphProps
                if ($insertionPlacement -eq "before" -and $insertIndex -eq 0) {
                    InsertParagraphBefore -AnchorParagraph $insertAnchor -NewParagraph $newParagraph
                } else {
                    InsertParagraphAfter -AnchorParagraph $insertAnchor -NewParagraph $newParagraph
                }
                $insertAnchor = $newParagraph
                if ($op.Mode -in @("revision", "revision_comment")) { $nextRevision++; $revisionCount++ }
            }
            continue
        }
        $runProps = RunPropsClone $p
        $snapshot = SnapshotParagraph $p
        if ($null -ne $snapshot.PPr) { [void]$p.AppendChild($p.OwnerDocument.ImportNode($snapshot.PPr, $true)) }
        $commentId = $null
        if ($op.Mode -in @("comment", "revision_comment")) {
            $commentId = $nextComment; $nextComment++
            [void]$p.AppendChild((CRStart $p.OwnerDocument $commentId))
            AddCommentNode -Doc $commentsDoc -Id $commentId -AuthorText $AuthorText -DateText $DateText -Lines $op.CommentLines
            $commentCount++
        }
        if ($op.Mode -eq "comment") {
            if ($snapshot.Content.Count -gt 0) { foreach ($node in $snapshot.Content) { [void]$p.AppendChild($p.OwnerDocument.ImportNode($node, $true)) } } else { AppendRun $p $target.Text $runProps $false }
        } else {
            $segments = GetMinimalDiffSegments $target.Text ([string]$op.ReplacementText)
            $changed = WriteRevisionSegments -Paragraph $p -Segments $segments -NextRevision ([ref]$nextRevision) -AuthorText $AuthorText -DateText $DateText -RunProps $runProps
            if (-not $changed) { AppendRun $p $target.Text $runProps $false } else { $revisionCount++ }
        }
        if ($null -ne $commentId) { [void]$p.AppendChild((CREnd $p.OwnerDocument $commentId)); [void]$p.AppendChild((CRefRun $p.OwnerDocument $commentId)) }
    }
    SaveXml $doc $documentPath
    if ($null -ne $commentsDoc) { SaveXml $commentsDoc (Join-Path $Workspace "word\comments.xml") }
    [pscustomobject]@{ Comments = $commentCount; Revisions = $revisionCount }
}

function ReplaceParagraphText([System.Xml.XmlNode]$Paragraph, [string]$Text) {
    $runProps = RunPropsClone $Paragraph
    $snapshot = SnapshotParagraph $Paragraph
    if ($null -ne $snapshot.PPr) { [void]$Paragraph.AppendChild($Paragraph.OwnerDocument.ImportNode($snapshot.PPr, $true)) }
    AppendRun $Paragraph $Text $runProps $false
}

function ApplyClean([string]$Workspace, [object[]]$Operations) {
    $documentPath = Join-Path $Workspace "word\document.xml"
    $doc = LoadXml $documentPath
    $paragraphs = ParagraphRefs $doc
    $targets = FindTargets -Paragraphs $paragraphs -Operations $Operations
    $count = 0
    for ($i = 0; $i -lt $Operations.Count; $i++) {
        $operation = $Operations[$i]
        $target = $targets[$i].Node
        $insertionTexts = if ($null -ne $operation.PSObject.Properties["InsertionTexts"]) { @($operation.InsertionTexts | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) } else { @() }
        if (@($insertionTexts).Count -gt 0 -and $operation.Mode -in @("revision", "revision_comment")) {
            $insertionPlacement = if ($null -ne $operation.PSObject.Properties["InsertionPlacement"] -and -not [string]::IsNullOrWhiteSpace([string]$operation.InsertionPlacement)) { ([string]$operation.InsertionPlacement).ToLowerInvariant() } else { "after" }
            $runProps = RunPropsClone $target
            $paragraphProps = $null
            $snapshot = SnapshotParagraph $target
            if ($null -ne $snapshot.PPr) {
                $paragraphProps = $target.OwnerDocument.ImportNode($snapshot.PPr, $true)
                [void]$target.AppendChild($target.OwnerDocument.ImportNode($snapshot.PPr, $true))
            }
            if ($snapshot.Content.Count -gt 0) {
                foreach ($node in $snapshot.Content) { [void]$target.AppendChild($target.OwnerDocument.ImportNode($node, $true)) }
            } else {
                AppendRun $target $targets[$i].Text $runProps $false
            }
            $insertAnchor = $target
            for ($insertIndex = 0; $insertIndex -lt $insertionTexts.Count; $insertIndex++) {
                $text = $insertionTexts[$insertIndex]
                $p = WNode $target.OwnerDocument "p"
                if ($null -ne $paragraphProps) { [void]$p.AppendChild($target.OwnerDocument.ImportNode($paragraphProps, $true)) }
                AppendRun $p $text $runProps $false
                if ($insertionPlacement -eq "before" -and $insertIndex -eq 0) {
                    InsertParagraphBefore -AnchorParagraph $insertAnchor -NewParagraph $p
                } else {
                    InsertParagraphAfter -AnchorParagraph $insertAnchor -NewParagraph $p
                }
                $insertAnchor = $p
                $count++
            }
            continue
        }
        if ($operation.Mode -in @("revision", "revision_comment")) { ReplaceParagraphText $target ([string]$operation.ReplacementText); $count++ }
    }
    SaveXml $doc $documentPath
    $count
}

function ValidateDocx([string]$Docx, [bool]$NeedComments, [bool]$NeedRevisions) {
    $workspace = Join-Path ([System.IO.Path]::GetTempPath()) ("review-validate-" + [guid]::NewGuid().ToString("N"))
    try {
        ExpandDocx $Docx $workspace
        $docPath = Join-Path $workspace "word\document.xml"
        if (-not (Test-Path $docPath)) { throw "输出文件缺少 word/document.xml: $Docx" }
        $doc = LoadXml $docPath
        if ($NeedComments) {
            $commentsPath = Join-Path $workspace "word\comments.xml"
            if (-not (Test-Path $commentsPath)) { throw "输出文件缺少 word/comments.xml: $Docx" }
            if ($null -eq (SelectOneNs -XmlObject $doc -XPath "//w:commentRangeStart")) { throw "输出文件未写入批注标记: $Docx" }
            $comments = LoadXml $commentsPath
            if ($null -eq (SelectOneNs -XmlObject $comments -XPath "//w:comment//w:rPr/w:rFonts[@w:eastAsia='$($script:CommentFontName)']")) {
                throw "输出文件批注字体未锁定为宋体(SimSun): $Docx"
            }
        }
        if ($NeedRevisions) {
            $settingsPath = Join-Path $workspace "word\settings.xml"
            if (-not (Test-Path $settingsPath)) { throw "输出文件缺少 word/settings.xml: $Docx" }
            $settings = LoadXml $settingsPath
            if ($null -eq (SelectOneNs -XmlObject $settings -XPath "/w:settings/w:trackRevisions")) { throw "输出文件未开启修订模式: $Docx" }
            if ($null -eq (SelectOneNs -XmlObject $settings -XPath "/w:settings/w:showRevisions")) { throw "输出文件未开启修订显示: $Docx" }
            $hasInsertion = $null -ne (SelectOneNs -XmlObject $doc -XPath "//w:ins")
            $hasDeletion = $null -ne (SelectOneNs -XmlObject $doc -XPath "//w:del")
            if (-not $hasInsertion -and -not $hasDeletion) { throw "输出文件未写入修订痕迹: $Docx" }
        }
    } finally { if (Test-Path $workspace) { Remove-Item -LiteralPath $workspace -Recurse -Force } }
}

$sourcePath = FullPath $Source
$instructionsPath = FullPath $Instructions
$outputPath = FullPath $Output
$cleanOutputPath = if ([string]::IsNullOrWhiteSpace($CleanOutput)) { $null } else { FullPath $CleanOutput }

if (-not (Test-Path -LiteralPath $sourcePath)) { throw "源文件不存在: $sourcePath" }
if (-not (Test-Path -LiteralPath $instructionsPath)) { throw "指令文件不存在: $instructionsPath" }
if ([System.IO.Path]::GetExtension($sourcePath).ToLowerInvariant() -ne ".docx") { throw "source 必须是 .docx 文件" }
if ([System.IO.Path]::GetExtension($instructionsPath).ToLowerInvariant() -ne ".json") { throw "instructions 必须是 .json 文件" }

$useStaging = Use-Staging @($sourcePath, $instructionsPath, $outputPath, $cleanOutputPath)
$stagingInfo = $null
$processingSourcePath = $sourcePath
$processingInstructionsPath = $instructionsPath
$processingOutputPath = $outputPath
$processingCleanOutputPath = $cleanOutputPath
$templateCompareInput = [pscustomobject]@{
    InstructionsPath = $null
    CleanupRoot      = $null
    OriginalTemplatePath = $null
}

if ($useStaging) {
    $stagingInfo = Invoke-Staging -SourcePath $sourcePath -InstructionsPath $instructionsPath
    $processingSourcePath = FullPath ([string]$stagingInfo.staged_source)
    $processingInstructionsPath = FullPath ([string]$stagingInfo.staged_instructions)
    $processingOutputPath = FullPath ([string]$stagingInfo.staged_reviewed_docx)
    if ($cleanOutputPath) {
        $processingCleanOutputPath = FullPath ([string]$stagingInfo.staged_clean_docx)
    } else {
        $processingCleanOutputPath = $null
    }
}

$summary = $null
try {
    $config = Read-JsonFile $processingInstructionsPath
    $templateCompareInput = Prepare-TemplateCompareInput -Config $config -OriginalInstructionsPath $instructionsPath -ProcessingInstructionsPath $processingInstructionsPath -StagingInfo $stagingInfo
    $effectiveAuthor = if ($null -ne $config.PSObject.Properties["author"] -and -not [string]::IsNullOrWhiteSpace([string]$config.author)) { [string]$config.author } else { $Author }
    $effectiveDate = Iso ($(if ($null -ne $config.PSObject.Properties["date"]) { [string]$config.date } else { $Date }))
    $explicitOperations = @()
    foreach ($raw in @($config.operations)) {
        $comment = if ($null -ne $raw.PSObject.Properties["comment"]) {
            if ($raw.comment -is [string]) { [ordered]@{ "问题" = [string]$raw.comment } } else {
                $h = [ordered]@{}
                foreach ($key in @("问题", "风险", "修改建议", "建议条款", "修改依据")) {
                    $prop = $raw.comment.PSObject.Properties[$key]
                    if ($null -ne $prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
                        $h[$key] = [string]$prop.Value
                    }
                }
                if ($h.Count -eq 0) { $null } else { $h }
            }
        } else { $null }
        $commentLines = @()
        if ($null -ne $comment) {
            foreach ($key in @("问题", "风险", "修改建议", "建议条款", "修改依据")) {
                if ($comment.Contains($key) -and -not [string]::IsNullOrWhiteSpace([string]$comment[$key])) {
                    $commentLines += "${key}：$($comment[$key])"
                }
            }
        }
        if ($commentLines.Count -eq 0) { $commentLines = @("问题：未提供批注内容") }
        $mode = [string]$raw.mode; if ([string]::IsNullOrWhiteSpace($mode)) { $mode = $DefaultMode }; $mode = $mode.ToLowerInvariant()
        if ($mode -notin @("comment", "revision", "revision_comment")) { throw "operation mode 不合法: $mode" }
        $occurrence = if ($null -ne $raw.PSObject.Properties["occurrence"]) { [int]$raw.occurrence } else { 1 }
        $matchType = if ($null -ne $raw.PSObject.Properties["match_type"] -and -not [string]::IsNullOrWhiteSpace([string]$raw.match_type)) { ([string]$raw.match_type).ToLowerInvariant() } else { "exact" }
        $location = if ($null -ne $raw.PSObject.Properties["location"]) { $raw.location } else { $null }
        $anchorText = if ($null -ne $raw.PSObject.Properties["anchor_text"]) { [string]$raw.anchor_text } else { $null }
        $insertionTexts = if ($null -ne $raw.PSObject.Properties["insertion_texts"]) { @($raw.insertion_texts | ForEach-Object { [string]$_ }) } else { @() }
        $insertionPlacement = if ($null -ne $raw.PSObject.Properties["insertion_placement"] -and -not [string]::IsNullOrWhiteSpace([string]$raw.insertion_placement)) { ([string]$raw.insertion_placement).ToLowerInvariant() } else { "after" }
        if ($null -eq $location -and [string]::IsNullOrWhiteSpace($anchorText)) { throw "operation 缺少 anchor_text 或 location" }
        $hasReplacementText = $null -ne $raw.PSObject.Properties["replacement_text"]
        if ($mode -in @("revision", "revision_comment") -and -not $hasReplacementText -and $insertionTexts.Count -eq 0) { throw "修订 operation 缺少 replacement_text 或 insertion_texts" }
        if ($mode -in @("comment", "revision_comment") -and $null -eq $comment) { throw "批注 operation 缺少 comment" }
        if ($insertionPlacement -notin @("after", "before")) { throw "operation insertion_placement 不合法: $insertionPlacement" }
        $explicitOperations += [pscustomobject]@{
            AnchorText = $anchorText
            Location = $location
            Mode = $mode
            ReplacementText = $(if ($hasReplacementText) { [string]$raw.replacement_text } else { $null })
            Comment = $comment
            CommentLines = $commentLines
            MatchType = $matchType
            Occurrence = $occurrence
            InsertionTexts = $insertionTexts
            InsertionPlacement = $insertionPlacement
        }
    }
    $templateCompareResult = $null
    $generatedOperations = @()
    if ($null -ne $config.PSObject.Properties["template_compare"]) {
        try {
            $compareHelperPath = Join-Path $script:EffectiveScriptRoot "generate_template_compare_ops.py"
            if (-not (Test-Path -LiteralPath $compareHelperPath)) { throw "缺少模板比对辅助脚本: $compareHelperPath" }
            $compareOutputPath = Join-Path ([System.IO.Path]::GetTempPath()) ("template-compare-" + [guid]::NewGuid().ToString("N") + ".json")
            $pythonExecutable = Resolve-PythonExecutable
            try {
                & $pythonExecutable $compareHelperPath --source $processingSourcePath --instructions ([string]$templateCompareInput.InstructionsPath) --default-mode $DefaultMode --output-json $compareOutputPath | Out-Null
                if (-not (Test-Path -LiteralPath $compareOutputPath)) { throw "模板比对辅助脚本未生成结果文件" }
                $templateCompareResult = Read-JsonFile $compareOutputPath
                if ($null -ne $templateCompareResult -and -not [string]::IsNullOrWhiteSpace([string]$templateCompareInput.OriginalTemplatePath)) {
                    $templateCompareResult.template_path = [string]$templateCompareInput.OriginalTemplatePath
                }
            }
            finally {
                if (Test-Path -LiteralPath $compareOutputPath) { Remove-Item -LiteralPath $compareOutputPath -Force }
            }
        }
        catch {
            $stack = if ([string]::IsNullOrWhiteSpace([string]$_.ScriptStackTrace)) { "" } else { "`n$($_.ScriptStackTrace)" }
            throw "template_compare 生成失败: $($_.Exception.Message)$stack"
        }
        foreach ($rawGenerated in @($templateCompareResult.operations)) {
            $comment = if ($null -ne $rawGenerated.PSObject.Properties["comment"]) { $rawGenerated.comment } else { $null }
            $commentLines = @()
            if ($null -ne $comment) {
                foreach ($key in @("问题", "风险", "修改建议", "建议条款", "修改依据")) {
                    if ($null -ne $comment.PSObject.Properties[$key] -and -not [string]::IsNullOrWhiteSpace([string]$comment.$key)) {
                        $commentLines += ("{0}：{1}" -f $key, [string]$comment.$key)
                    }
                }
            }
            $generatedOperations += [pscustomobject]@{
                AnchorText = $(if ($null -ne $rawGenerated.PSObject.Properties["anchor_text"]) { [string]$rawGenerated.anchor_text } else { $null })
                Location = $(if ($null -ne $rawGenerated.PSObject.Properties["location"]) { $rawGenerated.location } else { $null })
                Mode = ([string]$rawGenerated.mode).ToLowerInvariant()
                ReplacementText = $(if ($null -ne $rawGenerated.PSObject.Properties["replacement_text"]) { [string]$rawGenerated.replacement_text } else { $null })
                Comment = $comment
                CommentLines = $commentLines
                MatchType = $(if ($null -ne $rawGenerated.PSObject.Properties["match_type"]) { [string]$rawGenerated.match_type } else { "exact" })
                Occurrence = $(if ($null -ne $rawGenerated.PSObject.Properties["occurrence"]) { [int]$rawGenerated.occurrence } else { 1 })
                InsertionTexts = $(if ($null -ne $rawGenerated.PSObject.Properties["insertion_texts"]) { @($rawGenerated.insertion_texts | ForEach-Object { [string]$_ }) } else { @() })
                InsertionPlacement = $(if ($null -ne $rawGenerated.PSObject.Properties["insertion_placement"]) { ([string]$rawGenerated.insertion_placement).ToLowerInvariant() } else { "after" })
                TopicName = $(if ($null -ne $rawGenerated.PSObject.Properties["topic_name"]) { [string]$rawGenerated.topic_name } else { $null })
                AlignmentReason = $(if ($null -ne $rawGenerated.PSObject.Properties["alignment_reason"]) { [string]$rawGenerated.alignment_reason } else { $null })
                AlignmentConfidence = $(if ($null -ne $rawGenerated.PSObject.Properties["alignment_confidence"]) { [double]$rawGenerated.alignment_confidence } else { $null })
                MissingReason = $(if ($null -ne $rawGenerated.PSObject.Properties["missing_reason"]) { [string]$rawGenerated.missing_reason } else { $null })
                SourceClauseDetection = $(if ($null -ne $rawGenerated.PSObject.Properties["source_clause_detection"]) { [string]$rawGenerated.source_clause_detection } else { $null })
                TemplateClauseDetection = $(if ($null -ne $rawGenerated.PSObject.Properties["template_clause_detection"]) { [string]$rawGenerated.template_clause_detection } else { $null })
            }
        }
    }
    $operations = @()
    $seenOperationIds = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($operation in @($explicitOperations)) {
        [void]$seenOperationIds.Add((OperationIdentity $operation))
        $operations += $operation
    }
    foreach ($operation in @($generatedOperations)) {
        $identity = OperationIdentity $operation
        if ($seenOperationIds.Add($identity)) {
            $operations += $operation
        }
    }
    if ($operations.Count -eq 0) { throw "instructions.json 必须包含非空 operations 数组，或提供可生成比对结果的 template_compare 配置" }

    CopyDocx $processingSourcePath $processingOutputPath
    $reviewWorkspace = Join-Path ([System.IO.Path]::GetTempPath()) ("review-docx-" + [guid]::NewGuid().ToString("N"))
    try { ExpandDocx $processingOutputPath $reviewWorkspace; $review = ApplyReviewed $reviewWorkspace $operations $effectiveAuthor $effectiveDate; PackDocx $reviewWorkspace $processingOutputPath } finally { if (Test-Path $reviewWorkspace) { Remove-Item -LiteralPath $reviewWorkspace -Recurse -Force } }
    $needComments = @($operations | Where-Object { $_.Mode -in @("comment", "revision_comment") }).Count -gt 0
    $needRevisions = @($operations | Where-Object { $_.Mode -in @("revision", "revision_comment") }).Count -gt 0
    ValidateDocx $processingOutputPath $needComments $needRevisions

    if ($useStaging) { CopyDocx $processingOutputPath $outputPath }

    $summary = [ordered]@{
        source               = $sourcePath
        instructions         = $instructionsPath
        output               = $outputPath
        author               = $effectiveAuthor
        date                 = $effectiveDate
        operations_applied   = $operations.Count
        comments_applied     = $review.Comments
        revisions_applied    = $review.Revisions
        staging_applied      = $useStaging
        processing_source    = $processingSourcePath
        processing_instructions = $processingInstructionsPath
        processing_output    = $processingOutputPath
        review_comment_impl  = "direct_xml"
        xml_editing_locked   = $true
        template_compare_applied = ($null -ne $templateCompareResult)
        template_generated_operations = @($generatedOperations).Count
    }
    if ($null -ne $templateCompareResult) {
        $summary["template_compare_template"] = [string]$templateCompareResult.template_path
        $summary["template_compare_focus_topics"] = @($templateCompareResult.focus_topics)
        $summary["template_compare_matched_topics"] = @($templateCompareResult.matched_topics)
        $summary["template_compare_missing_topics"] = @($templateCompareResult.missing_topics)
        $summary["alignment_reasons_summary"] = @($templateCompareResult.alignment_reasons_summary)
        $summary["missing_reasons_summary"] = @($templateCompareResult.missing_reasons_summary)
        $summary["review_summary_lines"] = @($templateCompareResult.review_summary_lines)
        $summary["review_summary_text"] = [string]$templateCompareResult.review_summary_text
    }
    if ($useStaging) {
        $summary["staging_workspace"] = [string]$stagingInfo.workspace_root
        $summary["staging_reason"] = "non_ascii_path"
    }
    if ($cleanOutputPath) {
        CopyDocx $processingSourcePath $processingCleanOutputPath
        $cleanWorkspace = Join-Path ([System.IO.Path]::GetTempPath()) ("clean-docx-" + [guid]::NewGuid().ToString("N"))
        try { ExpandDocx $processingCleanOutputPath $cleanWorkspace; $cleanChanges = ApplyClean $cleanWorkspace $operations; PackDocx $cleanWorkspace $processingCleanOutputPath } finally { if (Test-Path $cleanWorkspace) { Remove-Item -LiteralPath $cleanWorkspace -Recurse -Force } }
        ValidateDocx $processingCleanOutputPath $false $false
        if ($useStaging) { CopyDocx $processingCleanOutputPath $cleanOutputPath }
        $summary["clean_output"] = $cleanOutputPath
        $summary["clean_changes"] = $cleanChanges
        $summary["processing_clean_output"] = $processingCleanOutputPath
    }

    $summary | ConvertTo-Json -Depth 6
}
finally {
    if ($null -ne $templateCompareInput -and -not [string]::IsNullOrWhiteSpace([string]$templateCompareInput.CleanupRoot) -and (Test-Path -LiteralPath ([string]$templateCompareInput.CleanupRoot))) {
        Remove-Item -LiteralPath ([string]$templateCompareInput.CleanupRoot) -Recurse -Force
    }
    if ($useStaging -and $null -ne $stagingInfo -and -not [string]::IsNullOrWhiteSpace([string]$stagingInfo.workspace_root) -and (Test-Path -LiteralPath ([string]$stagingInfo.workspace_root))) {
        Remove-Item -LiteralPath ([string]$stagingInfo.workspace_root) -Recurse -Force
    }
}
