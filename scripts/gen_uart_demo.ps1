# 生成示例图: STM32 UART 与串口助手连接示意
Add-Type -AssemblyName System.Drawing
$W = 900; $H = 420
$bmp = New-Object System.Drawing.Bitmap $W, $H
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
$g.Clear([System.Drawing.Color]::White)

$fTitle = New-Object System.Drawing.Font("Microsoft YaHei", 16, [System.Drawing.FontStyle]::Bold)
$fBox   = New-Object System.Drawing.Font("Microsoft YaHei", 12, [System.Drawing.FontStyle]::Bold)
$fNote  = New-Object System.Drawing.Font("Microsoft YaHei", 10)

function Draw-Box($x,$y,$w,$h,$text,$fill,$fnt){
  $sb = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromName($fill))
  $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(70,80,90)), 2
  $g.FillRectangle($sb,$x,$y,$w,$h)
  $g.DrawRectangle($pen,$x,$y,$w,$h)
  $fmt = New-Object System.Drawing.StringFormat
  $fmt.Alignment = [System.Drawing.StringAlignment]::Center
  $fmt.LineAlignment = [System.Drawing.StringAlignment]::Center
  $g.DrawString($text,$fnt,[System.Drawing.Brushes]::Black,[System.Drawing.RectangleF]::new($x,$y,$w,$h),$fmt)
  $sb.Dispose(); $pen.Dispose()
}

$g.DrawString("STM32 UART 与电脑串口助手连接示意图", $fTitle, [System.Drawing.Brushes]::Black, 20, 12)

# 左侧: STM32 单片机
Draw-Box 60  150 220 120 "STM32 单片机`nUSART1" "219,238,244" $fBox
# 右侧: 电脑串口助手
Draw-Box 620 150 220 120 "电脑`n串口助手" "255,243,219" $fBox
# 中间: USB 转串口
Draw-Box 360 150 200 120 "USB 转串口`n(如 CH340)" "236,236,236" $fBox

# 连线 TX
$p = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(200,60,60)), 3
$g.DrawLine($p, 280, 175, 360, 175)
$g.DrawString("TX", $fNote, [System.Drawing.Brushes]::Red, 300, 148)
$g.DrawLine($p, 560, 175, 620, 175)
$g.DrawString("TX", $fNote, [System.Drawing.Brushes]::Red, 585, 148)
$p.Dispose()

# 连线 RX
$p = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(60,120,200)), 3
$g.DrawLine($p, 280, 245, 360, 245)
$g.DrawString("RX", $fNote, [System.Drawing.Brushes]::Blue, 300, 255)
$g.DrawLine($p, 560, 245, 620, 245)
$g.DrawString("RX", $fNote, [System.Drawing.Brushes]::Blue, 585, 255)
$p.Dispose()

# GND
$p = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(90,90,90)), 3
$g.DrawLine($p, 280, 315, 360, 315)
$g.DrawLine($p, 560, 315, 620, 315)
$g.DrawString("GND(共地)", $fNote, [System.Drawing.Brushes]::Gray, 400, 320)
$p.Dispose()

$g.DrawString("TX 接 RX、RX 接 TX（交叉），并共地；波特率与 8-N-1 两端一致。", $fNote, [System.Drawing.Brushes]::Gray, 20, 385)

$out = Join-Path (Split-Path $PSScriptRoot -Parent) "examples\img\uart_demo.png"
$bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "saved: $out"
