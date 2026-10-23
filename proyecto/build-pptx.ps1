$ErrorActionPreference = 'Stop'
$base   = Split-Path -Parent $MyInvocation.MyCommand.Path
$json   = Join-Path $base 'contenido-pptx.json'
$outFile= Join-Path $base 'presentacion-ecosistema.pptx'

$raw  = [IO.File]::ReadAllText($json, [Text.Encoding]::UTF8)
$deck = $raw | ConvertFrom-Json
$slides = @($deck.slides)
$total  = $slides.Count

# ---------- constants ----------
$MID   = [string][char]0x00B7
$INK   = 'F6EFE6'
$MUTED = 'B3A196'
$DIM   = '6B5B50'
$ACC   = 'EC8839'
$ACC2  = 'E6562C'
$WARN  = 'E8BD76'
$OKC   = '8FD6A8'
$CW    = 12.093
$M     = 0.62
$TOTALW= 13.333
$TOTALH= 7.5

function E([double]$inches) { return [string][int][math]::Round($inches * 914400) }

function X([string]$s) {
  if ($null -eq $s) { return '' }
  return ($s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;')
}

function Runs {
  param([string]$text, [double]$sz, [string]$color, [bool]$bold = $false)
  $segs = [regex]::Split([string]$text, '(\*\*)')
  $isBold = $bold
  $out = ''
  foreach ($s in $segs) {
    if ($s -eq '**') { $isBold = -not $isBold; continue }
    if ($s -eq '') { continue }
    $b = if ($isBold) { '1' } else { '0' }
    $out += "<a:r><a:rPr lang=""es-MX"" sz=""$([int]($sz*100))"" b=""$b"" dirty=""0""><a:solidFill><a:srgbClr val=""$color""/></a:solidFill><a:latin typeface=""Segoe UI""/><a:cs typeface=""Segoe UI""/></a:rPr><a:t>$(X $s)</a:t></a:r>"
  }
  if ($out -eq '') { $out = "<a:r><a:rPr lang=""es-MX"" sz=""$([int]($sz*100))"" b=""0""><a:solidFill><a:srgbClr val=""$color""/></a:solidFill></a:rPr><a:t></a:t></a:r>" }
  return $out
}

function P {
  param([string]$text = '', [double]$sz = 11, [string]$color = $INK, [bool]$bold = $false, [double]$before = 0, [double]$after = 0, [double]$ln = 100, [bool]$bullet = $false, [string]$align = 'l')
  $pr = "<a:pPr algn=""$align"""
  if ($bullet) { $pr += " marL=""$(E 0.22)"" indent=""$(E -0.22)""" }
  $pr += "><a:lnSpc><a:spcPct val=""$([int]($ln*1000))""/></a:lnSpc>"
  if ($before -gt 0) { $pr += "<a:spcBef><a:spcPts val=""$([int]($before*100))""/></a:spcBef>" }
  if ($after -gt 0) { $pr += "<a:spcAft><a:spcPts val=""$([int]($after*100))""/></a:spcAft>" }
  if ($bullet) { $pr += '<a:buFont typeface="Arial"/><a:buChar char="&#8226;"/>' } else { $pr += '<a:buNone/>' }
  $pr += '</a:pPr>'
  if ($text -eq '') { return "<a:p>$pr</a:p>" }
  return "<a:p>$pr$(Runs $text $sz $color $bold)</a:p>"
}

$script:sid = 2
function New-Sp {
  param(
    [double]$x, [double]$y, [double]$w, [double]$h,
    [string]$geom = 'rect', [double]$adj = -1,
    [string]$fill = 'none', [int]$fa = 100000,
    [string]$line = 'none', [int]$la = 100000, [int]$lw = 9525,
    [string]$paras = '', [string]$anchor = 't',
    [double]$l = 0, [double]$t = 0, [double]$r = 0, [double]$b = 0,
    [string]$name = 'shape'
  )
  $id = $script:sid; $script:sid++
  if ($geom -eq 'rect' -and $adj -lt 0) {
    $geomXml = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
  } else {
    $adjXml = ''
    if ($adj -ge 0) { $adjXml = "<a:avLst><a:gd name=""adj"" fmla=""val $([int]$adj)""/></a:avLst>" }
    else { $adjXml = '<a:avLst/>' }
    $geomXml = "<a:prstGeom prst=""$geom"">$adjXml</a:prstGeom>"
  }
  if ($fill -eq 'none') { $fillXml = '<a:noFill/>' }
  else {
    if ($fa -lt 100000) { $fillXml = "<a:solidFill><a:srgbClr val=""$fill""><a:alpha val=""$fa""/></a:srgbClr></a:solidFill>" }
    else { $fillXml = "<a:solidFill><a:srgbClr val=""$fill""/></a:solidFill>" }
  }
  if ($line -eq 'none') { $lineXml = '<a:ln><a:noFill/></a:ln>' }
  else {
    if ($la -lt 100000) { $lineXml = "<a:ln w=""$lw""><a:solidFill><a:srgbClr val=""$line""><a:alpha val=""$la""/></a:srgbClr></a:solidFill></a:ln>" }
    else { $lineXml = "<a:ln w=""$lw""><a:solidFill><a:srgbClr val=""$line""/></a:solidFill></a:ln>" }
  }
  $tx = ''
  if ($paras -ne '') {
    $tx = "<p:txBody><a:bodyPr wrap=""square"" lIns=""$(E $l)"" tIns=""$(E $t)"" rIns=""$(E $r)"" bIns=""$(E $b)"" anchor=""$anchor""/><a:lstStyle/>$paras</p:txBody>"
  }
  return "<p:sp><p:nvSpPr><p:cNvPr id=""$id"" name=""$name $id""/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=""$(E $x)"" y=""$(E $y)""/><a:ext cx=""$(E $w)"" cy=""$(E $h)""/></a:xfrm>$geomXml$fillXml$lineXml</p:spPr>$tx</p:sp>"
}

function New-Text {
  param([double]$x, [double]$y, [double]$w, [double]$h, [string]$paras, [string]$anchor = 't', [string]$name = 'text', [string]$align = 'l')
  return (New-Sp -x $x -y $y -w $w -h $h -paras $paras -anchor $anchor -name $name)
}

function New-Bg {
  $id = $script:sid; $script:sid++
  return "<p:sp><p:nvSpPr><p:cNvPr id=""$id"" name=""bg $id""/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""$(E $TOTALW)"" cy=""$(E $TOTALH)""/></a:xfrm><a:prstGeom prst=""rect""><a:avLst/></a:prstGeom><a:gradFill rotWithShape=""1""><a:gsLst><a:gs pos=""0""><a:srgbClr val=""1C1008""/></a:gs><a:gs pos=""48000""><a:srgbClr val=""100906""/></a:gs><a:gs pos=""100000""><a:srgbClr val=""070403""/></a:gs></a:gsLst><a:lin ang=""2700000"" scaled=""1""/></a:gradFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>"
}

function New-Chrome {
  param([int]$index, [string]$section)
  $s = ''
  $pct = [double]$index / [double]$total
  if ($pct -gt 1) { $pct = 1 }
  $s += New-Sp -x 0 -y 0 -w ($TOTALW * $pct) -h 0.035 -fill $ACC -fa 90000 -name 'progress'
  $s += New-Sp -x $M -y 0.30 -w 0.15 -h 0.15 -geom 'diamond' -fill $ACC -name 'mark'
  $s += New-Text -x ($M + 0.24) -y 0.27 -w 2.2 -h 0.22 -paras (P 'LUXIMIA' 8 '8D7A6D' $true 0 0 100) -name 'brand'
  $s += New-Text -x (9.7) -y 0.29 -w 3.0 -h 0.2 -paras (P ($section.ToUpper()) 7.5 $DIM $true 0 0 100 $false 'r') -name 'section'
  $s += New-Sp -x $M -y 6.965 -w $CW -h 0.008 -fill 'FFFFFF' -fa 9000 -name 'footerline'
  $s += New-Text -x $M -y 7.02 -w 6.5 -h 0.18 -paras (P ("LUXIMIA  $MID  " + $section.ToUpper()) 6.8 $DIM $false 0 0 100) -name 'footer'
  $s += New-Text -x (11.2) -y 7.02 -w 1.51 -h 0.18 -paras (P ("$('{0:D2}' -f $index) / $total") 6.8 $DIM $false 0 0 100 $false 'r') -name 'pagenum'
  return $s
}

function LinesNeeded([string]$text, [int]$cpl) {
  if ($cpl -lt 10) { $cpl = 10 }
  $plain = ([string]$text) -replace '\*\*',''
  if ($plain.Length -eq 0) { return 1 }
  $n = 0
  foreach ($seg in ($plain -split "`n")) {
    $n += [math]::Max(1, [math]::Ceiling($seg.Length / [double]$cpl))
  }
  if ($n -lt 1) { $n = 1 }
  return $n
}

function BlockHeight($b) {
  switch ([string]$b.type) {
    'stats' { return 1.06 }
    'cards' {
      $cols = 3; if ($null -ne $b.cols) { $cols = [int]$b.cols }
      $w = ($CW - ($cols - 1) * 0.16) / $cols
      $cpl = [math]::Floor(($w - 0.38) / 0.070)
      $maxLines = 1
      foreach ($it in $b.items) {
        $ln = LinesNeeded ([string]$it.p) $cpl
        if ($ln -gt $maxLines) { $maxLines = $ln }
      }
      $rows = [math]::Ceiling($b.items.Count / [double]$cols)
      $rowH = 0.44 + $maxLines * 0.190 + 0.18
      return $rows * $rowH + ($rows - 1) * 0.14
    }
    'bullets' {
      $n = $b.items.Count
      $cols = 1; if ($null -ne $b.cols) { $cols = [int]$b.cols }
      if ($cols -eq 2) {
        $half = [math]::Ceiling($n / 2.0)
        $hL = 0; $hR = 0
        for ($i = 0; $i -lt $half; $i++) { $hL += (LinesNeeded ([string]$b.items[$i]) 85) * 0.205 + 0.02 }
        for ($i = $half; $i -lt $n; $i++) { $hR += (LinesNeeded ([string]$b.items[$i]) 85) * 0.205 + 0.02 }
        return [math]::Max($hL, $hR)
      } else {
        $h = 0
        foreach ($it in $b.items) { $h += (LinesNeeded ([string]$it) 150) * 0.205 + 0.02 }
        return $h
      }
    }
    'note' { return 0.22 + (LinesNeeded ([string]$b.text) 96) * 0.200 }
    'table' {
      $rows = $b.rows.Count + 1
      if ($rows -gt 7) { return $rows * 0.345 + 0.02 }
      return $rows * 0.50 + 0.02
    }
    'quote' { return 1.25 }
    default { return 0.8 }
  }
}

function Render-Cards($b, [double]$x, [double]$y, [double]$w) {
  $cols = 3; if ($null -ne $b.cols) { $cols = [int]$b.cols }
  $gap = 0.16
  $cw = ($w - ($cols - 1) * $gap) / $cols
  $cpl = [math]::Floor(($cw - 0.42) / 0.070)
  $maxLines = 1
  foreach ($it in $b.items) {
    $ln = LinesNeeded ([string]$it.p) $cpl
    if ($ln -gt $maxLines) { $maxLines = $ln }
  }
  $rowH = 0.44 + $maxLines * 0.190 + 0.18
  $out = ''
  $i = 0
  foreach ($it in $b.items) {
    $row = [math]::Floor($i / $cols)
    $col = $i % $cols
    $cx = $x + $col * ($cw + $gap)
    $cy = $y + $row * ($rowH + 0.14)
    $paras = (P ([string]$it.h) 12 $INK $true 0 4 100)
    $first = $true
    foreach ($line in (([string]$it.p) -split "`n")) {
      $bef = 0
      if (-not $first) { $bef = 1.5 }
      $paras += (P $line 10 $MUTED $false $bef 0 104)
      $first = $false
    }
    $out += New-Sp -x $cx -y $cy -w $cw -h $rowH -geom 'roundRect' -adj 7000 -fill 'FFFFFF' -fa 6500 -line 'FFFFFF' -la 11000 -paras $paras -l 0.19 -t 0.16 -r 0.19 -b 0.12 -anchor 't' -name 'card'
    $i++
  }
  return $out
}

function Render-Stats($b, [double]$x, [double]$y, [double]$w) {
  $items = @($b.items)
  $n = $items.Count
  $gap = 0.14
  $cwd = ($w - ($n - 1) * $gap) / $n
  $out = ''
  for ($i = 0; $i -lt $n; $i++) {
    $it = $items[$i]
    $cx = $x + $i * ($cwd + $gap)
    $paras = (P ([string]$it.n) 24 $ACC $true 0 5 100) + (P ([string]$it.l) 9 $MUTED $false 0 0 106)
    $out += New-Sp -x $cx -y $y -w $cwd -h 1.06 -geom 'roundRect' -adj 7000 -fill 'FFFFFF' -fa 6000 -line 'FFFFFF' -la 10000 -paras $paras -l 0.15 -t 0.13 -r 0.14 -b 0.10 -anchor 't' -name 'stat'
  }
  return $out
}

function Render-Bullets($b, [double]$x, [double]$y, [double]$w) {
  $items = @($b.items)
  $n = $items.Count
  $cols = 1; if ($null -ne $b.cols) { $cols = [int]$b.cols }
  $out = ''
  if ($cols -eq 2) {
    $half = [int][math]::Ceiling($n / 2.0)
    $cw = ($w - 0.32) / 2
    $leftParas = ''
    for ($i = 0; $i -lt $half; $i++) { $leftParas += (P ([string]$items[$i]) 11.5 $MUTED $false 0 0 108 $true) }
    $rightParas = ''
    for ($i = $half; $i -lt $n; $i++) { $rightParas += (P ([string]$items[$i]) 11.5 $MUTED $false 0 0 108 $true) }
    $out += New-Text -x $x -y $y -w $cw -h 4.2 -paras $leftParas -name 'bulletsL'
    $out += New-Text -x ($x + $cw + 0.32) -y $y -w $cw -h 4.2 -paras $rightParas -name 'bulletsR'
  } else {
    $paras = ''
    foreach ($it in $items) { $paras += (P ([string]$it) 11.5 $MUTED $false 0 0 108 $true) }
    $out += New-Text -x $x -y $y -w $w -h 4.2 -paras $paras -name 'bullets'
  }
  return $out
}

function Render-Note($b, [double]$x, [double]$y, [double]$w) {
  $tone = 'plain'; if ($null -ne $b.tone) { $tone = [string]$b.tone }
  $fill = 'FFFFFF'; $fa = 4500; $line = 'FFFFFF'; $la = 9000; $bar = $ACC
  if ($tone -eq 'accent') { $fill = $ACC; $fa = 9000; $line = $ACC; $la = 28000; $bar = $ACC }
  if ($tone -eq 'warn')   { $fill = $WARN; $fa = 8000; $line = $WARN; $la = 28000; $bar = $WARN }
  $h = 0.22 + (LinesNeeded ([string]$b.text) 96) * 0.200
  $out = New-Sp -x $x -y $y -w $w -h $h -geom 'roundRect' -adj 9000 -fill $fill -fa $fa -line $line -la $la -name 'note'
  $out += New-Sp -x ($x + 0.02) -y ($y + 0.10) -w 0.035 -h ($h - 0.20) -fill $bar -name 'notebar'
  $out += New-Text -x ($x + 0.20) -y ($y + 0.10) -w ($w - 0.36) -h ($h - 0.16) -paras (P ([string]$b.text) 11 $INK $false 0 0 104) -anchor 't' -name 'notetext'
  return $out
}

function Render-Table($b, [double]$x, [double]$y, [double]$w) {
  $rows = @($b.rows)
  $head = @($b.head)
  $dense = ($rows.Count + 1) -gt 7
  $rowH = if ($dense) { 0.345 } else { 0.50 }
  $fs = if ($dense) { 9 } else { 10.5 }
  $hfs = 7.0
  if ($null -ne $b.widths) { $widths = @($b.widths) } else { $widths = @(1.0 / $head.Count) * $head.Count }
  $ids = $script:sid; $script:sid++
  $grid = ''
  foreach ($wd in $widths) { $grid += "<a:gridCol w=""$(E ($w * [double]$wd))""/>" }
  $trs = ''
  # header
  $trs += "<a:tr h=""$(E $rowH)"">"
  for ($c = 0; $c -lt $head.Count; $c++) {
    $cellParas = P (([string]$head[$c]).ToUpper()) $hfs $DIM $true 0 0 100
    $trs += "<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>$cellParas</a:txBody><a:tcPr marL=""$(E 0.10)"" marR=""$(E 0.08)"" marT=""$(E 0.04)"" marB=""$(E 0.04)"" anchor=""ctr""><a:lnB w=""9525""><a:solidFill><a:srgbClr val=""FFFFFF""><a:alpha val=""14000""/></a:srgbClr></a:solidFill></a:lnB><a:solidFill><a:srgbClr val=""FFFFFF""><a:alpha val=""4000""/></a:srgbClr></a:solidFill></a:tcPr></a:tc>"
  }
  $trs += '</a:tr>'
  $r = 0
  foreach ($row in $rows) {
    $fillA = if ($r % 2 -eq 0) { '3000' } else { '0' }
    $fillXml = if ($fillA -eq '0') { '<a:noFill/>' } else { "<a:solidFill><a:srgbClr val=""FFFFFF""><a:alpha val=""$fillA""/></a:srgbClr></a:solidFill>" }
    $trs += "<a:tr h=""$(E $rowH)"">"
    for ($c = 0; $c -lt $head.Count; $c++) {
      $val = ''
      if ($c -lt $row.Count) { $val = [string]$row[$c] }
      $col = if ($c -eq 0) { $INK } else { $MUTED }
      $bold = ($c -eq 0)
      $celLn = 102
      if ($dense) { $celLn = 98 }
      $cellParas = P $val $fs $col $bold 0 0 $celLn
      $trs += "<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>$cellParas</a:txBody><a:tcPr marL=""$(E 0.10)"" marR=""$(E 0.08)"" marT=""$(E 0.05)"" marB=""$(E 0.05)"" anchor=""ctr""><a:lnB w=""9525""><a:solidFill><a:srgbClr val=""FFFFFF""><a:alpha val=""9000""/></a:srgbClr></a:solidFill></a:lnB>$fillXml</a:tcPr></a:tc>"
    }
    $trs += '</a:tr>'
    $r++
  }
  $h = ($rows.Count + 1) * $rowH
  return "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=""$ids"" name=""table $ids""/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp=""1""/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x=""$(E $x)"" y=""$(E $y)""/><a:ext cx=""$(E $w)"" cy=""$(E $h)""/></p:xfrm><a:graphic><a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/table""><a:tbl><a:tblPr firstRow=""1"" bandRow=""0""/><a:tblGrid>$grid</a:tblGrid>$trs</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"
}

# ---------- slide builders ----------
function New-ContentSlide($s, [int]$index) {
  $script:sid = 2
  $sp = New-Bg
  $sp += New-Chrome -index $index -section ([string]$s.section)
  $y = 0.55
  $sp += New-Sp -x $M -y ($y + 0.09) -w 0.34 -h 0.012 -fill $ACC -name 'kline'
  $sp += New-Text -x ($M + 0.46) -y $y -w 7.5 -h 0.26 -paras (P (([string]$s.kicker).ToUpper()) 10.5 $ACC $true 0 0 100) -name 'kicker'
  $y += 0.40

  $title = [string]$s.title
  $tl = LinesNeeded $title 46
  if ($tl -gt 2) { $tl = 2 }
  $tsz = 30; if ($tl -eq 1) { $tsz = 32 }
  $th = 0.54 * $tl + 0.08
  $sp += New-Text -x $M -y $y -w $CW -h $th -paras (P $title $tsz $INK $true 0 0 94) -name 'title'
  $y += $th + 0.04

  if ($null -ne $s.lead -and [string]$s.lead -ne '') {
    $ll = LinesNeeded ([string]$s.lead) 98
    $lh = $ll * 0.225 + 0.03
    $sp += New-Text -x $M -y $y -w $CW -h $lh -paras (P ([string]$s.lead) 12.5 $MUTED $false 0 0 112) -name 'lead'
    $y += $lh + 0.20
  }

  $avail = 6.80 - $y
  $blocks = @($s.blocks)
  $heights = @()
  $sum = 0
  foreach ($b in $blocks) { $h = BlockHeight $b; $heights += $h; $sum += $h }
  $gaps = [math]::Max(0, $blocks.Count - 1) * 0.16
  $totalH = $sum + $gaps
  $gap = 0.16
  if ($totalH -gt $avail -and $blocks.Count -gt 1) {
    $gap = [math]::Max(0.06, ($avail - $sum) / ($blocks.Count - 1))
  }
  # Si sobra espacio, repartirlo entre separacion y un desplazamiento moderado.
  $extra = $avail - $totalH
  if ($extra -gt 0.2 -and $blocks.Count -gt 1) {
    $gap = [math]::Min(0.40, $gap + $extra * 0.10)
  }
  if ($extra -gt 0.2) {
    $y += [math]::Min(0.85, $extra * 0.40)
  }

  $bi = 0
  foreach ($b in $blocks) {
    $bh = $heights[$bi]
    switch ([string]$b.type) {
      'stats'   { $sp += Render-Stats $b $M $y $CW }
      'cards'   { $sp += Render-Cards $b $M $y $CW }
      'bullets' { $sp += Render-Bullets $b $M $y $CW }
      'note'    { $sp += Render-Note $b $M $y $CW }
      'table'   { $sp += Render-Table $b $M $y $CW }
      'quote'   { $sp += Render-Note @{text=$b.text; tone='accent'} $M $y $CW }
    }
    $y += $bh + $gap
    $bi++
  }
  if ($totalH -gt $avail + 0.05) {
    Write-Host ("  ! slide {0}: contenido {1:N2}in > disponible {2:N2}in" -f $index, $totalH, $avail) -ForegroundColor Yellow
  }
  return $sp
}

function New-CoverSlide($s) {
  $script:sid = 2
  $sp = New-Bg
  $sp += New-Sp -x 0 -y 0 -w ($TOTALW * 0.011) -h 0.035 -fill $ACC -name 'progress'
  $sp += New-Sp -x $M -y 0.30 -w 0.15 -h 0.15 -geom 'diamond' -fill $ACC -name 'mark'
  $sp += New-Text -x ($M + 0.24) -y 0.27 -w 2.2 -h 0.22 -paras (P 'LUXIMIA' 8 '8D7A6D' $true 0 0 100) -name 'brand'

  $sp += New-Text -x $M -y 1.98 -w 8.6 -h 0.26 -paras (P (([string]$s.kicker).ToUpper()) 11 $ACC $true 0 0 120) -name 'kicker'
  $sp += New-Text -x $M -y 2.32 -w 8.7 -h 2.05 -paras (P ([string]$s.title) 44 $INK $true 0 0 92) -name 'title'
  $sp += New-Text -x $M -y 4.52 -w 7.5 -h 0.62 -paras (P ([string]$s.lead) 13 $MUTED $false 0 0 116) -name 'lead'

  $chipX = $M
  foreach ($c in @($s.chips)) {
    $cwid = 0.46 + ([string]$c).Length * 0.082
    $paras = (P ([string]$c) 10 $INK $false 0 0 100)
    $sp += New-Sp -x $chipX -y 5.55 -w $cwid -h 0.34 -geom 'roundRect' -adj 50000 -fill $ACC -fa 10000 -line $ACC -la 35000 -paras $paras -l 0.13 -t 0.06 -r 0.10 -b 0.02 -anchor 'ctr' -name 'chip'
    $chipX += $cwid + 0.12
  }
  if ($null -ne $s.note -and [string]$s.note -ne '') {
    $sp += New-Text -x $M -y 6.18 -w 8 -h 0.22 -paras (P ([string]$s.note) 9 $DIM $false 0 0 100) -name 'note'
  }

  # orbital mark
  $cx = 11.05; $cy = 3.10
  $sp += New-Sp -x ($cx - 1.25) -y ($cy - 1.25) -w 2.5 -h 2.5 -geom 'ellipse' -fill 'none' -line $ACC -la 16000 -lw 12700 -name 'ring'
  $sp += New-Sp -x ($cx - 0.95) -y ($cy - 0.95) -w 1.9 -h 1.9 -geom 'ellipse' -fill 'none' -line 'FFFFFF' -la 7000 -lw 9525 -name 'ring2'
  $sp += New-Sp -x ($cx - 0.48) -y ($cy - 0.48) -w 0.96 -h 0.96 -geom 'diamond' -fill $ACC -fa 90000 -line $ACC2 -la 90000 -lw 19050 -name 'gem'
  return $sp
}

function New-ClosingSlide($s) {
  $script:sid = 2
  $sp = New-Bg
  $sp += New-Sp -x 0 -y 0 -w $TOTALW -h 0.035 -fill $ACC -name 'progress'
  $sp += New-Sp -x $M -y 0.30 -w 0.15 -h 0.15 -geom 'diamond' -fill $ACC -name 'mark'
  $sp += New-Text -x ($M + 0.24) -y 0.27 -w 2.2 -h 0.22 -paras (P 'LUXIMIA' 8 '8D7A6D' $true 0 0 100) -name 'brand'

  if ($null -ne $s.kicker -and [string]$s.kicker -ne '') {
    $sp += New-Text -x $M -y 1.28 -w 8.4 -h 0.26 -paras (P (([string]$s.kicker).ToUpper()) 11 $ACC $true 0 0 120) -name 'kicker'
  }
  $sp += New-Text -x $M -y 1.72 -w 8.6 -h 2.05 -paras (P ([string]$s.title) 40 $INK $true 0 0 94) -name 'title'
  $sp += New-Text -x $M -y 3.86 -w 7.9 -h 0.72 -paras (P ([string]$s.lead) 12.5 $MUTED $false 0 0 116) -name 'lead'

  $cx = 11.0; $cy = 2.55
  $sp += New-Sp -x ($cx - 0.85) -y ($cy - 0.85) -w 1.7 -h 1.7 -geom 'ellipse' -fill 'none' -line $ACC -la 18000 -lw 12700 -name 'ring'
  $sp += New-Sp -x ($cx - 0.34) -y ($cy - 0.34) -w 0.68 -h 0.68 -geom 'diamond' -fill $ACC -fa 90000 -line $ACC2 -la 90000 -lw 19050 -name 'gem'
  $sp += New-Text -x 9.45 -y 3.62 -w 3.1 -h 0.24 -paras (P 'LUXIMIA' 12 $INK $true 0 0 300 $false 'ctr') -name 'wordmark'
  $sp += New-Text -x 9.45 -y 3.90 -w 3.1 -h 0.20 -paras (P 'PROJECTS WITH SOUL' 7.5 $DIM $false 0 0 300 $false 'ctr') -name 'tagline'
  $sp += New-Text -x 9.2 -y 4.45 -w 3.6 -h 0.5 -paras (P ([string]$s.note) 13 $MUTED $false 0 0 120 $false 'ctr') -name 'thanks'
  return $sp
}

# ---------- package ----------
function SlideXml($shapes) {
  return "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><p:sld xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:p=""http://schemas.openxmlformats.org/presentationml/2006/main""><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=""1"" name=""""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""0"" cy=""0""/><a:chOff x=""0"" y=""0""/><a:chExt cx=""0"" cy=""0""/></a:xfrm></p:grpSpPr>$shapes</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr><p:transition spd=""med""><p:fade/></p:transition></p:sld>"
}

$slideXmlList = @()
for ($i = 0; $i -lt $total; $i++) {
  $s = $slides[$i]
  Write-Host ("Slide {0}/{1}: {2}" -f ($i + 1), $total, $s.title)
  if ($s.type -eq 'cover') { $shapes = New-CoverSlide $s }
  elseif ($s.type -eq 'closing') { $shapes = New-ClosingSlide $s }
  else { $shapes = New-ContentSlide $s ($i + 1) }
  $slideXmlList += (SlideXml $shapes)
}

# ---------- static parts ----------
$ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
for ($i = 0; $i -lt $total; $i++) {
  $ct += "<Override PartName=""/ppt/slides/slide$($i+1).xml"" ContentType=""application/vnd.openxmlformats-officedocument.presentationml.slide+xml""/>"
}
$ct += '</Types>'

$rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'

$pres = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>'
for ($i = 0; $i -lt $total; $i++) { $pres += "<p:sldId id=""$($i+256)"" r:id=""rId$($i+2)""/>" }
$pres += '</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="es-MX"/></a:defPPr></p:defaultTextStyle></p:presentation>'

$presRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
for ($i = 0; $i -lt $total; $i++) { $presRels += "<Relationship Id=""rId$($i+2)"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"" Target=""slides/slide$($i+1).xml""/>" }
$presRels += '</Relationships>'

$master = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="0A0705"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="2800" b="1"><a:solidFill><a:srgbClr val="F6EFE6"/></a:solidFill><a:latin typeface="Segoe UI"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz="1400"><a:solidFill><a:srgbClr val="B3A196"/></a:solidFill><a:latin typeface="Segoe UI"/></a:defRPr></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr sz="1200"><a:solidFill><a:srgbClr val="F6EFE6"/></a:solidFill><a:latin typeface="Segoe UI"/></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>'

$masterRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>'

$layout = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'

$layoutRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>'

$theme = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Luximia"><a:themeElements><a:clrScheme name="Luximia"><a:dk1><a:srgbClr val="F6EFE6"/></a:dk1><a:lt1><a:srgbClr val="0A0705"/></a:lt1><a:dk2><a:srgbClr val="EC8839"/></a:dk2><a:lt2><a:srgbClr val="B3A196"/></a:lt2><a:accent1><a:srgbClr val="EC8839"/></a:accent1><a:accent2><a:srgbClr val="E6562C"/></a:accent2><a:accent3><a:srgbClr val="7FD8E8"/></a:accent3><a:accent4><a:srgbClr val="8FD6A8"/></a:accent4><a:accent5><a:srgbClr val="E8BD76"/></a:accent5><a:accent6><a:srgbClr val="C4B5FD"/></a:accent6><a:hlink><a:srgbClr val="7FD8E8"/></a:hlink><a:folHlink><a:srgbClr val="EC8839"/></a:folHlink></a:clrScheme><a:fontScheme name="Luximia"><a:majorFont><a:latin typeface="Segoe UI"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Segoe UI"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Luximia"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="28575"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'

$now = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
$core = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties"" xmlns:dc=""http://purl.org/dc/elements/1.1/"" xmlns:dcterms=""http://purl.org/dc/terms/"" xmlns:dcmitype=""http://purl.org/dc/dcmitype/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""><dc:title>Del trabajo manual al trabajo sistematizado</dc:title><dc:subject>Launcher, mapa del ecosistema y gobierno de accesos</dc:subject><dc:creator>Luximia</dc:creator><cp:lastModifiedBy>Luximia</cp:lastModifiedBy><cp:revision>1</cp:revision><dcterms:created xsi:type=""dcterms:W3CDTF"">$now</dcterms:created><dcterms:modified xsi:type=""dcterms:W3CDTF"">$now</dcterms:modified></cp:coreProperties>"

$app = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Properties xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"" xmlns:vt=""http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes""><Application>Microsoft Office PowerPoint</Application><PresentationFormat>Widescreen</PresentationFormat><Slides>$total</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><MMClips>0</MMClips><ScaleCrop>false</ScaleCrop><Company>Luximia</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion></Properties>"

$parts = [ordered]@{}
$parts['[Content_Types].xml'] = $ct
$parts['_rels/.rels'] = $rootRels
$parts['docProps/core.xml'] = $core
$parts['docProps/app.xml'] = $app
$parts['ppt/presentation.xml'] = $pres
$parts['ppt/_rels/presentation.xml.rels'] = $presRels
$parts['ppt/slideMasters/slideMaster1.xml'] = $master
$parts['ppt/slideMasters/_rels/slideMaster1.xml.rels'] = $masterRels
$parts['ppt/slideLayouts/slideLayout1.xml'] = $layout
$parts['ppt/slideLayouts/_rels/slideLayout1.xml.rels'] = $layoutRels
$parts['ppt/theme/theme1.xml'] = $theme
for ($i = 0; $i -lt $total; $i++) {
  $parts["ppt/slides/slide$($i+1).xml"] = $slideXmlList[$i]
  $parts["ppt/slides/_rels/slide$($i+1).xml.rels"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>'
}

# validate XML before zipping
foreach ($k in $parts.Keys) {
  if ($k -like '*.xml' -or $k -like '*.rels') {
    $null = [xml]$parts[$k]
  }
}

if (Test-Path $outFile) { Remove-Item $outFile -Force }
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$fs = [IO.File]::Open($outFile, [IO.FileMode]::Create)
$zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Create, $false)
$enc = New-Object Text.UTF8Encoding($false)
foreach ($k in $parts.Keys) {
  $entry = $zip.CreateEntry($k, [IO.Compression.CompressionLevel]::Optimal)
  $st = $entry.Open()
  $bytes = $enc.GetBytes($parts[$k])
  $st.Write($bytes, 0, $bytes.Length)
  $st.Close()
}
$zip.Dispose()
$fs.Close()

$fi = Get-Item $outFile
Write-Host ("OK: {0} ({1:N0} bytes, {2} slides)" -f $fi.Name, $fi.Length, $total) -ForegroundColor Green
