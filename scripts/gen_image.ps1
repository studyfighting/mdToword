# 用 Windows GDI+ 生成一张带中文标注的示意图(总线层次架构)
Add-Type -AssemblyName System.Drawing
$W = 860; $H = 520
$bmp = New-Object System.Drawing.Bitmap $W, $H
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
$g.Clear([System.Drawing.Color]::White)

$fTitle = New-Object System.Drawing.Font("Microsoft YaHei", 15, [System.Drawing.FontStyle]::Bold)
$fBox   = New-Object System.Drawing.Font("Microsoft YaHei", 11, [System.Drawing.FontStyle]::Bold)
$fNote  = New-Object System.Drawing.Font("Microsoft YaHei", 10)

function Draw-Box($x,$y,$w,$h,$text,$fill,[System.Drawing.Font]$fnt){
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

# 标题
$g.DrawString("SoC 总线层次结构示意（AXI / AHB / APB）", $fTitle, [System.Drawing.Brushes]::Black, 20, 12)

# 高速层
Draw-Box 40  70 180 60 "CPU 核心" "255,223,160" $fBox
Draw-Box 240 70 180 60 "DMA / 高性能外设" "255,223,160" $fBox
Draw-Box 440 70 180 60 "DDR 控制器" "255,223,160" $fBox

# 高速总线条
$gb = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(60,120,190)), 6
$g.DrawLine($gb, 30, 150, 630, 150)
$g.DrawString("AXI / AHB 高速总线（多主设备 + 仲裁器）", $fNote, [System.Drawing.Brushes]::Blue, 660, 140)
$gb.Dispose()

# 桥
Draw-Box 250 175 180 60 "AXI/AHB → APB 桥" "200,230,180" $fBox

# APB 总线条
$gp = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(140,90,40)), 5
$g.DrawLine($gp, 30, 260, 630, 260)
$g.DrawString("APB 低速外设总线（单主设备）", $fNote, [System.Drawing.Brushes]::Brown, 660, 250)
$gp.Dispose()

# 外设层
$per = @("UART","I2C/SMB","SPI/SSI","GPIO","PWM","TCU 定时器","SAR ADC","WDT / OST","RTC")
$x = 40; $y = 290; $pw = 175; $ph = 42
for($i=0; $i -lt $per.Count; $i++){
  $px = $x + ($i % 3) * ($pw + 20)
  $py = $y + [math]::Floor($i/3) * ($ph + 20)
  Draw-Box $px $py $pw $ph $per[$i] "222,235,247" $fBox
}

# 连接线(主->桥, 桥->APB)
$pl = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(60,120,190)), 2
$g.DrawLine($pl, 130, 130, 340, 170)
$g.DrawLine($pl, 340, 235, 330, 256)
$pl.Dispose()

# 底部说明
$g.DrawString("寄存器配置走 APB；大数据搬运（图像/码流）走 AXI——控制通路与数据通路分离。", $fNote, [System.Drawing.Brushes]::Gray, 20, 480)

$out = Join-Path $PSScriptRoot "_sample_bus.png"
$bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "saved: $out"
