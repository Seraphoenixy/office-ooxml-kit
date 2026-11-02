param(
    [Parameter(Mandatory)][ValidateSet('extract','inspect','fill')][string]$Action,
    [Parameter(Mandatory)][string]$Path,
    [string]$Output,
    [string]$Data
)
$ErrorActionPreference = 'Stop'
# Console code pages differ between hosts; errors are also written as UTF-8 so callers never see mojibake.
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
Add-Type -AssemblyName System.IO.Compression.FileSystem
$W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
function Read-Part($zip, $name) {
    $entry=$zip.GetEntry($name); if(-not $entry){return $null}
    $reader=[IO.StreamReader]::new($entry.Open())
    try { $xml=[xml]::new(); $xml.PreserveWhitespace=$true; $xml.LoadXml($reader.ReadToEnd()); return ,$xml } finally {$reader.Dispose()}
}
function Ns($xml) { $n=[Xml.XmlNamespaceManager]::new($xml.NameTable); $n.AddNamespace('w',$W); return ,$n }
function Attr($node,$key) { if($node){return $node.GetAttribute($key,$W)}; return '' }
function Text($node,$ns) {
    if($node.LocalName -ne 'p') { return (@($node.SelectNodes('.//w:p',$ns) | ForEach-Object {Text $_ $ns}) -join "`n").Trim() }
    $parts=[Collections.Generic.List[string]]::new()
    foreach($t in $node.SelectNodes('.//w:t | .//w:tab | .//w:br',$ns)) {
        if($t.LocalName -eq 't'){$parts.Add($t.InnerText)} elseif($t.LocalName -eq 'tab'){$parts.Add("`t")} else {$parts.Add("`n")}
    }
    return ($parts -join '').Trim()
}
function Save-Json($value,$file) { [IO.File]::WriteAllText($file,($value|ConvertTo-Json -Depth 70),[Text.UTF8Encoding]::new($false)) }function File-Hash($file) { $stream=[IO.File]::Open($file,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite);$sha=[Security.Cryptography.SHA256]::Create();try{return [Convert]::ToHexString($sha.ComputeHash($stream))}finally{$sha.Dispose();$stream.Dispose()} }
$zip=$null
try {
$packageStream=[IO.File]::Open((Resolve-Path -LiteralPath $Path).Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
$zip=[IO.Compression.ZipArchive]::new($packageStream,[IO.Compression.ZipArchiveMode]::Read,$false)
    $xml=Read-Part $zip 'word/document.xml'; $ns=Ns $xml
    if($Action -eq 'inspect') {
        $parts=@(foreach($e in $zip.Entries){$st=$e.Open();try{$sha=[Security.Cryptography.SHA256]::Create();$hash=[Convert]::ToHexString($sha.ComputeHash($st));$sha.Dispose()}finally{$st.Dispose()};@{name=$e.FullName;size=$e.Length;sha256=$hash}})
        $result=@{sha256=(File-Hash $Path); parts=$parts; sections=@($xml.SelectNodes('//w:sectPr',$ns)|ForEach-Object {$_.OuterXml}); tables=@($xml.SelectNodes('/w:document/w:body/w:tbl',$ns)|ForEach-Object {$_.OuterXml})}
        Save-Json $result $Output
    } elseif($Action -eq 'extract') {
        $styles=Read-Part $zip 'word/styles.xml'; $numbering=Read-Part $zip 'word/numbering.xml'
        $styleMap=@{}; if($styles){$sn=Ns $styles;foreach($s in $styles.SelectNodes('//w:style',$sn)){$styleMap[(Attr $s 'styleId')]=$s}}
        function Style-Prop($id,$prop) { $seen=@{};while($id -and $styleMap.ContainsKey($id) -and -not $seen.ContainsKey($id)){$seen[$id]=$true;$s=$styleMap[$id];$v=$s.SelectSingleNode("./w:pPr/$prop",$ns);if($v){return $v};$id=Attr ($s.SelectSingleNode('./w:basedOn',$ns)) 'val'};return $null }
        $numMap=@{};$abstractMap=@{};$counters=@{};$lastNumber=''
        if($numbering){$nn=Ns $numbering;foreach($a in $numbering.SelectNodes('//w:abstractNum',$nn)){$abstractMap[(Attr $a 'abstractNumId')]=$a};foreach($n in $numbering.SelectNodes('//w:num',$nn)){$numMap[(Attr $n 'numId')]=$n}}
        function Format-Number($value,$fmt){switch($fmt){'upperLetter'{return [string][char](64+$value)}'lowerLetter'{return [string][char](96+$value)}default{return [string]$value}}}
        function Paragraph-Number($p,$style) {
            $np=$p.SelectSingleNode('./w:pPr/w:numPr',$ns);if(-not $np){$np=Style-Prop $style 'w:numPr'};if(-not $np){return ''}
            $id=Attr ($np.SelectSingleNode('./w:numId',$ns)) 'val';$il=Attr ($np.SelectSingleNode('./w:ilvl',$ns)) 'val'
            if(-not $id){$id=Attr (Style-Prop $style 'w:numPr/w:numId') 'val'}
            if(-not $il){$il=Attr (Style-Prop $style 'w:numPr/w:ilvl') 'val'}
            if(-not $il){$il='0'}
            if(-not $numMap.ContainsKey($id)){return ''};$num=$numMap[$id];$aid=Attr ($num.SelectSingleNode('./w:abstractNumId',$ns)) 'val';$abs=$abstractMap[$aid];if(-not $abs){return ''}
            $level=$abs.SelectSingleNode("./w:lvl[@w:ilvl='$il']",$ns);$over=$num.SelectSingleNode("./w:lvlOverride[@w:ilvl='$il']",$ns);if($over -and $over.SelectSingleNode('./w:lvl',$ns)){$level=$over.SelectSingleNode('./w:lvl',$ns)};if(-not $level){return ''}
            $key="$id/$il";$start=Attr ($level.SelectSingleNode('./w:start',$ns)) 'val';if(-not $start){$start='1'};if($over -and $over.SelectSingleNode('./w:startOverride',$ns)){$start=Attr ($over.SelectSingleNode('./w:startOverride',$ns)) 'val'}
            if(-not $counters.ContainsKey($key)){$counters[$key]=[int]$start}else{$counters[$key]++}
            foreach($k in @($counters.Keys)){if($k.StartsWith("$id/") -and [int]($k.Split('/')[1]) -gt [int]$il){$counters.Remove($k)}}
            $label=Attr ($level.SelectSingleNode('./w:lvlText',$ns)) 'val'
            for($j=0;$j -le [int]$il;$j++){$ref=$abs.SelectSingleNode("./w:lvl[@w:ilvl='$j']",$ns);$rk="$id/$j";$v=if($counters.ContainsKey($rk)){$counters[$rk]}else{1};$fmt=Attr ($ref.SelectSingleNode('./w:numFmt',$ns)) 'val';$label=$label.Replace('%'+($j+1),(Format-Number $v $fmt))}
            return $label
        }
        # 表格单元格里的「序号」常是样式绑定的自动编号（正文为空，Word 才显示 1、2、3…）。
        # 只取 w:t 会得到空串，CR 里就成了 --。编号按 abstractNum 计数：共用同一 abstract 的
        # 连续单元格（如 1 号之后的 26 号）继续累加；startOverride 只在该 numId *首次*出现时
        # 重置计数，之后沿用累加（SWTP-013 的三行共用 numId=54，Word 显示 1、2、3 而不是 1、1、1）。
        $cellCounters=@{}
        $cellStarted=@{}
        function Style-NumId($style) {
            $id=$style
            for($guard=0;$guard -lt 40 -and $id -and $styleMap.ContainsKey($id);$guard++){
                $s=$styleMap[$id];$np=$s.SelectSingleNode('./w:pPr/w:numPr',$ns)
                if($np){$v=Attr ($np.SelectSingleNode('./w:numId',$ns)) 'val';if($v){return $v}}
                $id=Attr ($s.SelectSingleNode('./w:basedOn',$ns)) 'val'
            }
            return ''
        }
        function Cell-Text($cell) {
            $text=(Text $cell $ns)
            if($text){return $text}
            $p=$cell.SelectSingleNode('./w:p',$ns);if(-not $p){return $text}
            $style=Attr ($p.SelectSingleNode('./w:pPr/w:pStyle',$ns)) 'val'
            $np=$p.SelectSingleNode('./w:pPr/w:numPr',$ns);$id=''
            if($np){$id=Attr ($np.SelectSingleNode('./w:numId',$ns)) 'val'}
            if(-not $id){$id=Style-NumId $style}
            if(-not $id -or -not $numMap.ContainsKey($id)){return $text}
            $il='0'
            if($np){$il=Attr ($np.SelectSingleNode('./w:ilvl',$ns)) 'val';if(-not $il){$il='0'}}
            $num=$numMap[$id];$aid=Attr ($num.SelectSingleNode('./w:abstractNumId',$ns)) 'val'
            $abs=$abstractMap[$aid];if(-not $abs){return $text}
            $level=$abs.SelectSingleNode("./w:lvl[@w:ilvl='$il']",$ns);$over=$num.SelectSingleNode("./w:lvlOverride[@w:ilvl='$il']",$ns)
            if($over -and $over.SelectSingleNode('./w:lvl',$ns)){$level=$over.SelectSingleNode('./w:lvl',$ns)}
            if(-not $level){return $text}
            $start='1';$startNode=$null
            if($over){$startNode=$over.SelectSingleNode('./w:startOverride',$ns)}
            # startOverride 只在这个 numId 的第一次出现时生效，之后沿用累加：
            # 同一条规程里 1、2、3 共用同一个带 override 的 numId，若每次都重置就全成了 1。
            if($startNode -and -not $cellStarted.ContainsKey($id)){$start=Attr $startNode 'val';$cellCounters[$aid]=[int]$start;$cellStarted[$id]=$true}
            elseif($cellCounters.ContainsKey($aid)){$cellCounters[$aid]++}
            else{$initial=Attr ($level.SelectSingleNode('./w:start',$ns)) 'val';if(-not $initial){$initial='1'};$cellCounters[$aid]=[int]$initial}
            $value=$cellCounters[$aid]
            $label=Attr ($level.SelectSingleNode('./w:lvlText',$ns)) 'val'
            $label=$label.Replace('%'+([int]$il+1),(Format-Number $value (Attr ($level.SelectSingleNode('./w:numFmt',$ns)) 'val')))
            return $label
        }
        $blocks=[Collections.Generic.List[object]]::new();$heads=@('','','','','','','','','');$paragraphIndex=0;$tableIndex=0;$bodyIndex=0;$section='前置内容';$sectionTitle='前置内容'
        $warnings=[Collections.Generic.List[string]]::new();$body=$xml.SelectSingleNode('/w:document/w:body',$ns)
        # 图片按内容哈希比较：升版时 Word 会重排媒体文件名（image7.emf 变成 image9.emf），
        # 按文件名比较会把整篇图片都报成变化。哈希相同即同一张图，位置另外记录。
        # 图表里的文字读不出来，所以这里只提供“哪张图、多大、在哪个单元格”，
        # 图片内容是否讲对了必须由人看。
        $RNS='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        $relTargets=@{}
        $docRels=Read-Part $zip 'word/_rels/document.xml.rels'
        if($docRels){foreach($r in $docRels.SelectNodes('//*[local-name()="Relationship"]')){$relTargets[$r.GetAttribute('Id')]=$r}}
        $mediaCache=@{};$imgStats=@{count=0;ole=0}
        function Rel-Info($id) {
            $rel=$relTargets[$id];if(-not $rel){return $null}
            $target=$rel.GetAttribute('Target')
            if($rel.GetAttribute('TargetMode') -eq 'External'){return @{part=$target;external=$true;sha256='';ext='';bytes=0;token='⟦图:外链⟧'}}
            if($target.StartsWith('/')){$target=$target.Substring(1)}else{$target='word/'+$target}
            while($target.Contains('/../')){$target=[regex]::Replace($target,'[^/]+/\.\./','')}
            if($mediaCache.ContainsKey($target)){return $mediaCache[$target]}
            $entry=$zip.GetEntry($target);if(-not $entry){return $null}
            $stream=$entry.Open()
            try{$sha=[Security.Cryptography.SHA256]::Create();$hash=[Convert]::ToHexString($sha.ComputeHash($stream));$sha.Dispose()}finally{$stream.Dispose()}
            $info=@{part=$target;external=$false;sha256=$hash;ext=[IO.Path]::GetExtension($target).TrimStart('.').ToLower();bytes=$entry.Length;token='⟦图:'+$hash.Substring(0,8)+'⟧'}
            $mediaCache[$target]=$info;return $info
        }
        function Shape-Emu($style,$axis) {
            if(-not $style){return 0}
            if($style -match "$axis`:\s*([\d.]+)pt"){return [long]([double]$Matches[1]*12700)}
            if($style -match "$axis`:\s*([\d.]+)in"){return [long]([double]$Matches[1]*914400)}
            if($style -match "$axis`:\s*([\d.]+)cm"){return [long]([double]$Matches[1]*360000)}
            return 0
        }
        # 返回段落里的图片，按文档顺序。a:blip 是普通图片；v:imagedata 是 OLE 对象
        # （本规程里的 Visio 测试系统图）的预览图。预览可渲染，内嵌对象本身不复制。
        function Paragraph-Images($p) {
            $list=[Collections.Generic.List[object]]::new()
            foreach($node in $p.SelectNodes('.//*[local-name()="drawing" or local-name()="object"]')){
                if($node.SelectSingleNode('ancestor::*[local-name()="drawing" or local-name()="object"]')){continue}
                if($node.LocalName -eq 'drawing'){
                    $blip=$node.SelectSingleNode('.//*[local-name()="blip"]');if(-not $blip){continue}
                    $info=Rel-Info ($blip.GetAttribute('embed',$RNS));if(-not $info){continue}
                    $extent=$node.SelectSingleNode('.//*[local-name()="extent"]')
                    $cx=0;$cy=0
                    if($extent){$cx=[long]$extent.GetAttribute('cx');$cy=[long]$extent.GetAttribute('cy')}
                    $docPr=$node.SelectSingleNode('.//*[local-name()="docPr"]')
                    $name='';$alt=''
                    if($docPr){$name=$docPr.GetAttribute('name');$alt=$docPr.GetAttribute('descr')}
                    $imgStats.count++
                    $list.Add(@{role='picture';part=$info.part;sha256=$info.sha256;ext=$info.ext;bytes=$info.bytes;token=$info.token;cx=$cx;cy=$cy;name=$name;alt=$alt})
                } else {
                    $preview=$node.SelectSingleNode('.//*[local-name()="imagedata"]');if(-not $preview){continue}
                    $info=Rel-Info ($preview.GetAttribute('id',$RNS));if(-not $info){continue}
                    $shape=$node.SelectSingleNode('.//*[local-name()="shape"]')
                    $style=if($shape){$shape.GetAttribute('style')}else{''}
                    $ole=$node.SelectSingleNode('.//*[local-name()="OLEObject"]')
                    $oleInfo=if($ole){Rel-Info ($ole.GetAttribute('id',$RNS))}else{$null}
                    $imgStats.count++;$imgStats.ole++
                    $list.Add(@{role='olePreview';part=$info.part;sha256=$info.sha256;ext=$info.ext;bytes=$info.bytes;token=$info.token;cx=(Shape-Emu $style 'width');cy=(Shape-Emu $style 'height');name='';alt='';olePart=if($oleInfo){$oleInfo.part}else{''}})
                }
            }
            return $list.ToArray()
        }
        # 图片在文字里以 token 表示，位置留在原文顺序中：单元格文字带上 token 后，
        # 只有图片换了而文字没变时，整行仍然会被判为差异，不会静默漏掉。
        function Paragraph-Rich($p) {
            $t=Text $p $ns
            $imgs=@(Paragraph-Images $p)
            if($imgs.Count){$t=$t+(($imgs|ForEach-Object{$_.token}) -join '')}
            return @{text=$t;images=$imgs}
        }
        function Cell-Rich($cell) {
            $lines=[Collections.Generic.List[string]]::new();$imgs=[Collections.Generic.List[object]]::new()
            foreach($p in @($cell.SelectNodes('.//w:p',$ns))){
                $rich=Paragraph-Rich $p
                if($rich.text){$lines.Add($rich.text)}
                foreach($i in $rich.images){$imgs.Add($i)}
            }
            $text=($lines -join "`n").Trim()
            if(-not $text -and $imgs.Count -eq 0){$text=(Cell-Text $cell)}
            return @{text=$text;images=$imgs.ToArray()}
        }
        foreach($el in $body.ChildNodes){
            $bodyIndex++;if($el.LocalName -eq 'p'){
                $paragraphIndex++;$t=Text $el $ns;$style=Attr ($el.SelectSingleNode('./w:pPr/w:pStyle',$ns)) 'val';$outline=$el.SelectSingleNode('./w:pPr/w:outlineLvl',$ns);if(-not $outline){$outline=Style-Prop $style 'w:outlineLvl'};$lev=Attr $outline 'val';$label=Paragraph-Number $el $style
                $category='body';if($style -match '(?i)toc|目录'){$category='toc'}
                $isHeading=$lev -match '^[0-8]$' -and $category -ne 'toc'
                if($isHeading -and $t){$lv=[int]$lev;$heads[$lv]=$t;for($j=$lv+1;$j -lt 9;$j++){$heads[$j]=''};$sectionTitle=$t;$section=(($heads|Where-Object {$_}) -join ' / ')}
                if($t -eq '修订历史'){$section='修订历史';$sectionTitle=$t}
                if($section -eq '修订历史'){$category='revision'}
                $images=@(Paragraph-Images $el)
                if($images.Count){$t=$t+(($images|ForEach-Object{$_.token}) -join '')}
                if(-not $t){continue}
                $blocks.Add(@{kind=if($isHeading){'heading'}else{'paragraph'};text=$t;number=$label;style=$style;section=$section;sectionTitle=$sectionTitle;category=$category;locator="word/document.xml:/w:document/w:body/w:p[$paragraphIndex]";bodyIndex=$bodyIndex;images=$images})
            } elseif($el.LocalName -eq 'tbl') {
                $tableIndex++;$ri=0
                # 一张表的行宽由 gridCol 决定，行内单元格可能跨列（gridSpan）。只数单元格个数会把
                # 「标签 + 跨两列的值」和「序号 + 步骤 + 判据」当成不同宽度的表，重建时就被拆散。
                $grid=@($el.SelectNodes('./w:tblGrid/w:gridCol',$ns)|ForEach-Object {[int](Attr $_ 'w')})
                # 表的排版方式决定列宽来源：fixed 按 gridCol 固定，autofit（多数原文档）由 Word 按内容
                # 重算。把 autofit 表强行按 gridCol 固定，会把「序号」这类窄列压得换行。
                $tblPr=$el.SelectSingleNode('./w:tblPr',$ns)
                $layoutMode=Attr ($tblPr.SelectSingleNode('./w:tblLayout',$ns)) 'type'
                if($layoutMode -ne 'fixed'){$layoutMode='autofit'}
                foreach($row in $el.SelectNodes('./w:tr',$ns)){$ri++;$tcNodes=@($row.SelectNodes('./w:tc | ./w:sdt/w:sdtContent/w:tc',$ns));$cells=@();$cellImages=[Collections.Generic.List[object]]::new();$spans=@();foreach($tc in $tcNodes){$rich=Cell-Rich $tc;$cells+=$rich.text;$cellImages.Add(@($rich.images));$g=$tc.SelectSingleNode('./w:tcPr/w:gridSpan',$ns);if($g){$spans+=[int](Attr $g 'val')}else{$spans+=1}};$t=$cells -join ' | ';if(-not $t.Trim(' ','|')){continue};$blocks.Add(@{kind='tableRow';text=$t;cells=$cells;cellImages=$cellImages.ToArray();spans=$spans;grid=$grid;layout=$layoutMode;number='';section=$section;sectionTitle=$sectionTitle;category=if($section -eq '修订历史'){'revision'}elseif($section -eq '前置内容'){'front'}else{'body'};tableIndex=$tableIndex;rowIndex=$ri;locator="word/document.xml:/w:document/w:body/w:tbl[$tableIndex]/w:tr[$ri]";bodyIndex=$bodyIndex})}
            } elseif($el.LocalName -notin @('sectPr')) {
                $t=Text $el $ns;if($t){$blocks.Add(@{kind='opaque';text=$t;number='';section=$section;sectionTitle=$sectionTitle;category='front';locator="word/document.xml:/w:document/w:body/*[$bodyIndex]";bodyIndex=$bodyIndex})}
            }
        }
        foreach($pair in @(@('修订痕迹','//w:ins | //w:del | //w:moveFrom | //w:moveTo'),@('图片或绘图','//w:drawing | //w:pict'),@('文本框','//w:txbxContent'),@('公式','//*[local-name()="oMath"]'),@('嵌套表格','//w:tbl//w:tbl'),@('内容控件','//w:sdt'))){$count=$xml.SelectNodes($pair[1],$ns).Count;if($count){$warnings.Add("$($pair[0])：$count 处；需要人工核对，文字比较不能验证其完整语义或外观。")}}
        if($imgStats.count){$unique=@($mediaCache.Values|Where-Object {-not $_.external -and $_.part -like 'word/media/*'}|ForEach-Object {$_.part}|Select-Object -Unique).Count;$warnings.Add("图片：$($imgStats.count) 处（$unique 个媒体文件，其中 OLE 对象预览 $($imgStats.ole) 处）；按内容哈希比较，图中数值与界面文字须人工核对。")}
        $features=@(foreach($e in $zip.Entries){if($e.FullName -match '^word/(header|footer|footnotes|endnotes|comments).*\.xml$'){ $part=Read-Part $zip $e.FullName;$pn=Ns $part;@{part=$e.FullName;text=(Text $part $pn)}}})
        foreach($f in $features){$warnings.Add("附属内容 $($f.part) 已提取，需核对其变化。")}
        $all=($blocks|ForEach-Object {$_.text}) -join "`n";$id=[regex]::Match($all,'BDS\d+[A-Z]+\d+').Value;$ver=[regex]::Match($all,'A/\d+').Value
        $result=@{schemaVersion=1;file=(Resolve-Path -LiteralPath $Path).Path;sha256=(File-Hash $Path);metadata=@{documentId=$id;version=$ver;title=''};blocks=$blocks.ToArray();warnings=$warnings.ToArray();ancillary=$features;numberingCaveat='自动编号按 OOXML 样式与编号定义重建；复杂重启、章节关联编号需结合原文人工核对。'}
        Save-Json $result $Output
    } else {
        $payload=Get-Content -LiteralPath $Data -Raw | ConvertFrom-Json -Depth 70
        # 图片来源是两份原文：按部分名从对应版本复制媒体字节，不做转码，
        # 显示效果与原文一致，图片内容哈希也与提取时一致。
        $RNS='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        $WPNS='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
        $ANS='http://schemas.openxmlformats.org/drawingml/2006/main'
        $PICNS='http://schemas.openxmlformats.org/drawingml/2006/picture'
        $sources=@{}
        foreach($side in @('old','new')){
            $sourcePath=if($payload.sources){$payload.sources.$side}else{$null}
            if($sourcePath -and (Test-Path -LiteralPath $sourcePath)){$sources[$side]=[IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $sourcePath).Path)}
        }
        $newParts=[Collections.Generic.List[object]]::new()
        $newRels=[Collections.Generic.List[object]]::new()
        $partBySha=@{};$imageSeq=0;$drawingSeq=0
        $relsXml=Read-Part $zip 'word/_rels/document.xml.rels'
        $takenRelIds=@{}
        if($relsXml){foreach($r in $relsXml.SelectNodes('//*[local-name()="Relationship"]')){$takenRelIds[$r.GetAttribute('Id')]=$true}}
        function Add-ImagePart($image) {
            $src=$sources[[string]$image.source]
            if(-not $src){throw "缺少图片来源：$($image.source)"}
            $sha=[string]$image.sha256
            if($partBySha.ContainsKey($sha)){return $partBySha[$sha]}
            $entry=$src.GetEntry([string]$image.part)
            if(-not $entry){throw "图片来源 $($image.source) 中找不到 $($image.part)"}
            $stream=$entry.Open()
            try{$buffer=[IO.MemoryStream]::new();$stream.CopyTo($buffer);$bytes=$buffer.ToArray()}finally{$stream.Dispose()}
            $extension=if($image.ext){[string]$image.ext}else{'png'}
            $script:imageSeq++
            $partName="word/media/cr$($script:imageSeq).$extension"
            $relId="rIdCr$($script:imageSeq)"
            while($takenRelIds.ContainsKey($relId)){$script:imageSeq++;$partName="word/media/cr$($script:imageSeq).$extension";$relId="rIdCr$($script:imageSeq)"}
            $takenRelIds[$relId]=$true
            $newParts.Add(@{name=$partName;bytes=$bytes})
            $newRels.Add(@{id=$relId;target=('media/cr'+$script:imageSeq+'.'+$extension);extension=$extension})
            $info=@{part=$partName;rid=$relId;name=[IO.Path]::GetFileName($partName)}
            $partBySha[$sha]=$info
            return $info
        }
        # 图片按原尺寸等比缩放到可用宽度内，最多高 9cm；只缩放，不裁切。
        function New-ImageParagraph($image,$maxWidthEmu) {
            $info=Add-ImagePart $image
            $cx=[long]$image.cx;$cy=[long]$image.cy
            if($cx -le 0 -or $cy -le 0){$cx=3600000;$cy=2700000}
            $maxHeightEmu=[long]3240000
            $scale=[Math]::Min(1.0,[Math]::Min([double]$maxWidthEmu/[double]$cx,[double]$maxHeightEmu/[double]$cy))
            if($scale -le 0){$scale=1.0}
            # 变量名不能叫 $w：PowerShell 不区分大小写，会覆盖上面 Word 命名空间的 $W。
            $drawWidth=[long]($cx*$scale);$drawHeight=[long]($cy*$scale)
            $script:drawingSeq++
            $id=$script:drawingSeq
            $xmlText="<w:p xmlns:w=`"$W`" xmlns:r=`"$RNS`" xmlns:wp=`"$WPNS`" xmlns:a=`"$ANS`" xmlns:pic=`"$PICNS`"><w:pPr><w:jc w:val=`"center`"/></w:pPr><w:r><w:drawing><wp:inline distT=`"0`" distB=`"0`" distL=`"0`" distR=`"0`"><wp:extent cx=`"$drawWidth`" cy=`"$drawHeight`"/><wp:effectExtent l=`"0`" t=`"0`" r=`"0`" b=`"0`"/><wp:docPr id=`"$id`" name=`"图片 $id`"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect=`"1`"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri=`"http://schemas.openxmlformats.org/drawingml/2006/picture`"><pic:pic><pic:nvPicPr><pic:cNvPr id=`"$id`" name=`"$($info.name)`"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed=`"$($info.rid)`"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x=`"0`" y=`"0`"/><a:ext cx=`"$drawWidth`" cy=`"$drawHeight`"/></a:xfrm><a:prstGeom prst=`"rect`"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
            $fragment=[Xml.XmlDocument]::new();$fragment.PreserveWhitespace=$true;$fragment.LoadXml($xmlText)
            return $xml.ImportNode($fragment.DocumentElement,$true)
        }
        $table=$xml.SelectSingleNode('/w:document/w:body/w:tbl',$ns);$rows=$table.SelectNodes('./w:tr',$ns)
        if($rows.Count -ne 21 -or (Text $rows[0] $ns) -ne '更改申请单'){throw '模板结构不匹配，拒绝按未知位置填充。'}
        function Set-Cell($cell,[string]$value) {
            $p=$cell.SelectSingleNode('./w:p',$ns);if(-not $p){$p=$xml.CreateElement('w','p',$W)}
            $new=$p.CloneNode($true);foreach($child in @($new.ChildNodes)){if($child.LocalName -ne 'pPr'){[void]$new.RemoveChild($child)}}
            $firstRun=$cell.SelectSingleNode('.//w:r[w:t]',$ns)
            $rp=if($firstRun){$firstRun.SelectSingleNode('./w:rPr',$ns)}else{$null}
            $lines=$value -split "`n"
            foreach($c in @($cell.ChildNodes)){if($c.LocalName -ne 'tcPr'){[void]$cell.RemoveChild($c)}}
            foreach($line in $lines){$dest=$new.CloneNode($true);$run=$xml.CreateElement('w','r',$W);if($rp){[void]$run.AppendChild($rp.CloneNode($true))};$text=$xml.CreateElement('w','t',$W);$space=$xml.CreateAttribute('xml','space','http://www.w3.org/XML/1998/namespace');$space.Value='preserve';[void]$text.Attributes.Append($space);$text.InnerText=$line;[void]$run.AppendChild($text);[void]$dest.AppendChild($run);foreach($pr in $dest.SelectNodes('./w:pPr/w:keepNext | ./w:pPr/w:pageBreakBefore',$ns)){[void]$pr.ParentNode.RemoveChild($pr)};[void]$cell.AppendChild($dest)}
        }
        function Set-RichCell($cell,$parts,$baseCell) {
            $base=if($baseCell){$baseCell}else{$cell.CloneNode($true)}
            $outerWidth=[int](Attr ($cell.SelectSingleNode('./w:tcPr/w:tcW',$ns)) 'w')
            if($outerWidth -le 0){$outerWidth=6000}
            $cellMaxEmu=([long]($outerWidth-240))*635
            foreach($child in @($cell.ChildNodes)){if($child.LocalName -ne 'tcPr'){[void]$cell.RemoveChild($child)}}
            foreach($part in $parts){
                if($part.type -eq 'paragraph'){
                    $scratch=$base.CloneNode($true);Set-Cell $scratch ([string]$part.text)
                    foreach($p in $scratch.SelectNodes('./w:p',$ns)){[void]$cell.AppendChild($p.CloneNode($true))}
                } elseif($part.type -eq 'image') {
                    # 原文图片直接嵌入左右状态列，缩放后按原顺序排在文字之后。
                    [void]$cell.AppendChild((New-ImageParagraph $part $cellMaxEmu))
                } else {
                    $allRows=[Collections.Generic.List[object]]::new()
                    if($part.header){$allRows.Add(@($part.header))}
                    foreach($values in $part.rows){$allRows.Add(@($values))}
                    # 表格的列数来自 gridCol；行内单元格按 gridSpan 跨列。用「一行里的单元格个数」
                    # 当列数，会把「标签 + 跨列值」压成两列并把步骤行挤掉，规程表格就是这样被拆散的。
                    $columns=[int]$part.grid.Count
                    if($columns -le 0){$columns=($allRows|ForEach-Object {$_.Count}|Measure-Object -Maximum).Maximum}
                    $tbl=$xml.CreateElement('w','tbl',$W);$pr=$xml.CreateElement('w','tblPr',$W)
                    $width=$xml.CreateElement('w','tblW',$W);[void]$width.SetAttribute('w',$W,'5000');[void]$width.SetAttribute('type',$W,'pct');[void]$pr.AppendChild($width)
                    $layout=$xml.CreateElement('w','tblLayout',$W)
                    # 原表是 autofit 时保持 autofit：交给 Word 按内容分配列宽，与原文观感一致；
                    # 原表是 fixed 时才按 gridCol 固定列宽。
                    if($part.layout -eq 'fixed'){[void]$layout.SetAttribute('type',$W,'fixed')}else{[void]$layout.SetAttribute('type',$W,'autofit')}
                    [void]$pr.AppendChild($layout)
                    $borders=$xml.CreateElement('w','tblBorders',$W)
                    foreach($edge in @('top','left','bottom','right','insideH','insideV')){$border=$xml.CreateElement('w',$edge,$W);[void]$border.SetAttribute('val',$W,'single');[void]$border.SetAttribute('sz',$W,'4');[void]$border.SetAttribute('color',$W,'000000');[void]$borders.AppendChild($border)}
                    [void]$pr.AppendChild($borders);[void]$tbl.AppendChild($pr)
                    $available=[int](Attr ($cell.SelectSingleNode('./w:tcPr/w:tcW',$ns)) 'w')-180
                    # 列宽沿用原文 gridCol 的比例，这样原文的列宽关系得以保留；没有 gridCol 时退回固定权重。
                    if([int]$part.grid.Count -gt 0){
                        $total=($part.grid|Measure-Object -Sum).Sum
                        $tblGridPositions=@();foreach($g in $part.grid){$tblGridPositions+=[int]($available*([int]$g/[double]$total))}
                    } else {
                        $weights=switch([int]$columns){2 {@(0.40,0.60)} 3 {@(0.12,0.25,0.63)} 5 {@(0.08,0.24,0.15,0.12,0.41)} default {@(1..$columns|ForEach-Object {1.0/$columns})}}
                        $tblGridPositions=@();for($j=0;$j -lt $columns;$j++){$tblGridPositions+=[int]($available*$weights[$j])}
                    }
                    $grid=$xml.CreateElement('w','tblGrid',$W)
                    foreach($pos in $tblGridPositions){$col=$xml.CreateElement('w','gridCol',$W);[void]$col.SetAttribute('w',$W,[string]$pos);[void]$grid.AppendChild($col)}
                    [void]$tbl.AppendChild($grid)
                    $rowIndex=0
                    foreach($values in $allRows){
                        $tr=$xml.CreateElement('w','tr',$W)
                        if($part.header -and $rowIndex -eq 0){$trpr=$xml.CreateElement('w','trPr',$W);[void]$trpr.AppendChild($xml.CreateElement('w','tblHeader',$W));[void]$tr.AppendChild($trpr)}
                        # 该行的跨列信息；表头行与数据行分开保存，这里按 allRows 的序号取。
                        $rowSpans=@()
                        $headerOffset=0;if($part.header){$headerOffset=1}
                        $spanIndex=$rowIndex-$headerOffset
                        if($headerOffset -eq 1 -and $rowIndex -eq 0){if($part.headerSpans){$rowSpans=@($part.headerSpans)}}
                        elseif($part.rowSpans -and $spanIndex -ge 0 -and $spanIndex -lt $part.rowSpans.Count){$rowSpans=@($part.rowSpans[$spanIndex])}
                        $col=0
                        for($j=0;$j -lt $values.Count;$j++){
                            $span=if($j -lt $rowSpans.Count){[int]$rowSpans[$j]}else{1};if($span -lt 1){$span=1}
                            if($col+$span -gt $columns){$span=[Math]::Max(1,$columns-$col)}
                            $cellWidth=0;for($k=$col;$k -lt $col+$span;$k++){$cellWidth+=$tblGridPositions[$k]}
                            $tc=$xml.CreateElement('w','tc',$W);$tcpr=$xml.CreateElement('w','tcPr',$W);$cw=$xml.CreateElement('w','tcW',$W);[void]$cw.SetAttribute('type',$W,'dxa');[void]$cw.SetAttribute('w',$W,[string]$cellWidth);[void]$tcpr.AppendChild($cw)
                            if($span -gt 1){$gs=$xml.CreateElement('w','gridSpan',$W);[void]$gs.SetAttribute('val',$W,[string]$span);[void]$tcpr.AppendChild($gs)}
                            [void]$tc.AppendChild($tcpr)
                            if($values[$j] -is [string] -or -not $values[$j].parts){
                                $scratch=$base.CloneNode($true);Set-Cell $scratch ([string]$values[$j])
                                foreach($p in $scratch.SelectNodes('./w:p',$ns)){
                                    $copy=$p.CloneNode($true)
                                    if($part.header -and $rowIndex -eq 0){foreach($run in $copy.SelectNodes('./w:r',$ns)){$rp=$run.SelectSingleNode('./w:rPr',$ns);if(-not $rp){$rp=$xml.CreateElement('w','rPr',$W);[void]$run.PrependChild($rp)};[void]$rp.AppendChild($xml.CreateElement('w','b',$W))}}
                                    [void]$tc.AppendChild($copy)
                                }
                            } else {
                                # 富单元格：文字与图片按原顺序渲染，沿用外层单元格的字体设置。
                                Set-RichCell $tc $values[$j].parts $base
                            }
                            [void]$tr.AppendChild($tc);$col+=$span
                        }
                        # 原文行短于网格时，用跨列单元格补齐，避免 Word 把缺列的行拉宽。
                        if($col -lt $columns){$tc=$xml.CreateElement('w','tc',$W);$tcpr=$xml.CreateElement('w','tcPr',$W);$cw=$xml.CreateElement('w','tcW',$W);[void]$cw.SetAttribute('type',$W,'dxa');[void]$cw.SetAttribute('w',$W,[string]($tblGridPositions[$col..($columns-1)]|Measure-Object -Sum|ForEach-Object {$_.Sum}));[void]$tcpr.AppendChild($cw);$gs=$xml.CreateElement('w','gridSpan',$W);[void]$gs.SetAttribute('val',$W,[string]($columns-$col));[void]$tcpr.AppendChild($gs);[void]$tc.AppendChild($tcpr);[void]$tc.AppendChild($xml.CreateElement('w','p',$W));[void]$tr.AppendChild($tc)}
                        [void]$tbl.AppendChild($tr);$rowIndex++
                    }
                    [void]$cell.AppendChild($tbl)
                }
            }
            # Word requires a closing paragraph after a nested table in a cell.
            if($cell.LastChild.LocalName -ne 'p'){[void]$cell.AppendChild($xml.CreateElement('w','p',$W))}
        }
        function Mark-Uncertain($element) {
            foreach($paragraph in $element.SelectNodes('.//w:p',$ns)){
                if((Text $paragraph $ns) -notmatch '待确认|待审核|待核实|待补充|需.*确认'){continue}
                foreach($run in $paragraph.SelectNodes('./w:r[w:t]',$ns)){
                    $properties=$run.SelectSingleNode('./w:rPr',$ns)
                    if(-not $properties){$properties=$xml.CreateElement('w','rPr',$W);[void]$run.PrependChild($properties)}
                    foreach($color in @($properties.SelectNodes('./w:color',$ns))){[void]$properties.RemoveChild($color)}
                    $color=$xml.CreateElement('w','color',$W);[void]$color.SetAttribute('val',$W,'FF0000');[void]$properties.AppendChild($color)
                }
            }
        }
        # All editable slots are explicit physical cells, including vertically merged cells.
        foreach($slot in $payload.slots){$cells=$rows[[int]$slot.row].SelectNodes('./w:tc',$ns);Set-Cell $cells[[int]$slot.cell] ([string]$slot.text)}
        $prototype=$rows[20].CloneNode($true);[void]$table.RemoveChild($rows[20])
        # Keep the template header and all change rows in one outer table.
        $details=$table
        foreach($change in $payload.changes){$row=$prototype.CloneNode($true);foreach($h in @($row.SelectNodes('./w:trPr/w:trHeight | ./w:trPr/w:cantSplit',$ns))){[void]$h.ParentNode.RemoveChild($h)};$cells=$row.SelectNodes('./w:tc',$ns);if($change.oldParts){Set-RichCell $cells[0] $change.oldParts}else{Set-Cell $cells[0] ([string]$change.old)};if($change.newParts){Set-RichCell $cells[1] $change.newParts}else{Set-Cell $cells[1] ([string]$change.new)};Set-Cell $cells[2] ([string]$change.remark);[void]$details.AppendChild($row)}
        # Fill visible empty cells; vertical-merge continuation cells are not independent cells.
        # 只含图片的单元格文字为空，不能补成 --，否则图片被覆盖、页面只剩占位符。
        foreach($cell in $table.SelectNodes('.//w:tc',$ns)){
            $merge=$cell.SelectSingleNode('./w:tcPr/w:vMerge',$ns)
            if($merge -and (Attr $merge 'val') -ne 'restart'){continue}
            if($cell.SelectSingleNode('.//*[local-name()="drawing"]')){continue}
            if(-not $cell.SelectSingleNode('./w:tbl',$ns) -and -not (Text $cell $ns).Trim()){Set-Cell $cell '--'}
        }
        # Highlight only form values and editorial remarks; source quotations retain their formatting.
        $finalRows=$table.SelectNodes('./w:tr',$ns)
        for($i=0;$i -lt $finalRows.Count;$i++){
            if($i -lt 20){Mark-Uncertain $finalRows[$i]}
            else{Mark-Uncertain ($finalRows[$i].SelectNodes('./w:tc',$ns)[2])}
        }
        $settings=Read-Part $zip 'word/settings.xml';if($settings){$sns=Ns $settings;$upd=$settings.SelectSingleNode('//w:updateFields',$sns);if(-not $upd){$upd=$settings.CreateElement('w','updateFields',$W);[void]$settings.DocumentElement.AppendChild($upd)};[void]$upd.SetAttribute('val',$W,'false')}
        # Remove case-specific metadata and cached header/footer prose, preserving their layout.
        $replacements=@{'BDSNASA-3511'=$payload.crId;'BDSNASA-3507'='待确认';'BDS2230AJT014'=$payload.documentId;'中央处理软件加载规范'=$payload.title;'谷陆鹏'='待确认';'2026/10/16'='待确认'}
        $edits=@{'word/document.xml'=$xml.OuterXml};if($settings){$edits['word/settings.xml']=$settings.OuterXml}
        foreach($e in $zip.Entries){if($e.FullName -match '^word/(header|footer).*\.xml$'){$part=Read-Part $zip $e.FullName;$pn=Ns $part;foreach($t in $part.SelectNodes('//w:t',$pn)){foreach($key in $replacements.Keys){$t.InnerText=$t.InnerText.Replace($key,[string]$replacements[$key])}};$edits[$e.FullName]=$part.OuterXml}}
        if($zip.GetEntry('docProps/core.xml')){$core=Read-Part $zip 'docProps/core.xml';foreach($n in $core.DocumentElement.ChildNodes){if($n.LocalName -in @('creator','lastModifiedBy','title','subject','description','keywords')){$n.InnerText=''}};$edits['docProps/core.xml']=$core.OuterXml}
        # 新增图片要在包内补关系项和内容类型声明，否则 Word 会报文件损坏。
        if($newRels.Count){
            if($relsXml){
                foreach($r in $newRels){
                    $node=$relsXml.CreateElement('Relationship','http://schemas.openxmlformats.org/package/2006/relationships')
                    $node.SetAttribute('Id',$r.id)
                    $node.SetAttribute('Type','http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
                    $node.SetAttribute('Target',$r.target)
                    [void]$relsXml.DocumentElement.AppendChild($node)
                }
                $edits['word/_rels/document.xml.rels']=$relsXml.OuterXml
            }
            $contentTypes=Read-Part $zip '[Content_Types].xml'
            if($contentTypes){
                $typeByExt=@{png='image/png';jpg='image/jpeg';jpeg='image/jpeg';gif='image/gif';bmp='image/bmp';emf='image/x-emf';wmf='image/x-wmf';tif='image/tiff';tiff='image/tiff'}
                foreach($extension in @($newRels|ForEach-Object {$_.extension}|Select-Object -Unique)){
                    if($contentTypes.SelectSingleNode("//*[local-name()='Default' and @Extension='$extension']")){continue}
                    $default=$contentTypes.CreateElement('Default','http://schemas.openxmlformats.org/package/2006/content-types')
                    $default.SetAttribute('Extension',$extension)
                    if($typeByExt.ContainsKey($extension)){$default.SetAttribute('ContentType',$typeByExt[$extension])}else{$default.SetAttribute('ContentType','application/octet-stream')}
                    [void]$contentTypes.DocumentElement.AppendChild($default)
                }
                $edits['[Content_Types].xml']=$contentTypes.OuterXml
            }
        }
        if(Test-Path -LiteralPath $Output){throw '输出已存在，禁止覆盖。'}
        $outZip=[IO.Compression.ZipFile]::Open($Output,[IO.Compression.ZipArchiveMode]::Create)
        try{foreach($e in $zip.Entries){$entry=$outZip.CreateEntry($e.FullName);$dest=$entry.Open();try{if($edits.ContainsKey($e.FullName)){$bytes=[Text.UTF8Encoding]::new($false).GetBytes($edits[$e.FullName]);$dest.Write($bytes,0,$bytes.Length)}else{$src=$e.Open();try{$src.CopyTo($dest)}finally{$src.Dispose()}}}finally{$dest.Dispose()}};foreach($part in $newParts){$entry=$outZip.CreateEntry($part.name);$dest=$entry.Open();try{$dest.Write($part.bytes,0,$part.bytes.Length)}finally{$dest.Dispose()}}}finally{$outZip.Dispose()}
    }
} finally {if($zip){$zip.Dispose()};if($sources){foreach($source in $sources.Values){$source.Dispose()}}}



