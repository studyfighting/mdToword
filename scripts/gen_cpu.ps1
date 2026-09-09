# 生成 CPU 内核内部结构示意图（含寄存器组：PC/SP/LR/PSR 等）
Add-Type -AssemblyName System.Drawing

$W = 1080; $H = 740
$bmp = New-Object System.Drawing.Bitmap $W, $H
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
$g.Clear([System.Drawing.Color]::White)

$fTitle = New-Object System.Drawing.Font("Microsoft YaHei", 16, [System.Drawing.FontStyle]::Bold)
$fBox   = New-Object System.Drawing.Font("Microsoft YaHei", 11, [System.Drawing.FontStyle]::Bold)
$fText  = New-Object System.Drawing.Font("Microsoft YaHei", 10)
$fReg   = New-Object System.Drawing.Font("Microsoft YaHei", 10.5)

function Draw-Box($x,$y,$w,$h,$text,$fill){
  $sb = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, $fill[0], $fill[1], $fill[2]))
  $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(90,100,110)), 2
  $g.FillRectangle($sb,$x,$y,$w,$h)
  $g.DrawRectangle($pen,$x,$y,$w,$h)
  $fmt = New-Object System.Drawing.StringFormat
  $fmt.Alignment = [System.Drawing.StringAlignment]::Center
  $fmt.LineAlignment = [System.Drawing.StringAlignment]::Center
  $g.DrawString($text,$fBox,[System.Drawing.Brushes]::Black,[System.Drawing.RectangleF]::new($x,$y,$w,$h),$fmt)
  $sb.Dispose(); $pen.Dispose()
}

function Draw-Arrow($x1,$y1,$x2,$y2,$color){
  $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb($color[0],$color[1],$color[2])), 2
  $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::ArrowAnchor
  $g.DrawLine($pen,$x1,$y1,$x2,$y2)
  $pen.Dispose()
}

# 标题
$g.DrawString("CPU 内核内部结构示意（以 ARM Cortex-M 为例）", $fTitle, [System.Drawing.Brushes]::Black, 20, 12)

# CPU 外框
$outer = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(120,120,140)), 3
$g.DrawRectangle($outer, 40, 60, 1000, 620)
$g.DrawString("CPU 内核 Core", $fBox, [System.Drawing.Brushes]::Black, 50, 66)
$outer.Dispose()

# 1. 控制单元（左上）
Draw-Box 70 105 320 190 "控制单元`nControl Unit`n`n取指 Fetch → 译码 Decode → 执行 Execute" @(222,235,247)

# 2. 寄存器组（右上）
$reg = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,255,240,200))
$regPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(90,100,110)), 2
$g.FillRectangle($reg, 430, 105, 520, 320)
$g.DrawRectangle($regPen, 430, 105, 520, 320)
$g.DrawString("寄存器组 Register File", $fBox, [System.Drawing.Brushes]::Black, 440, 112)
$regText = "通用寄存器 R0 ~ R12`n`n栈指针 SP (R13)`n`n链接寄存器 LR (R14)`n`n程序计数器 PC (R15)`n`n状态寄存器 PSR / FLAGS"
$g.DrawString($regText, $fReg, [System.Drawing.Brushes]::Black, 460, 150)
$reg.Dispose(); $regPen.Dispose()

# 3. ALU（左下）
Draw-Box 70 330 320 160 "ALU 算术逻辑单元`n`n加减乘除 · 逻辑运算 · 比较" @(200,235,210)

# 4. 总线接口（右下）
Draw-Box 430 460 520 150 "总线接口 Bus Interface`n`n经 AHB / APB 总线连接 Flash / SRAM / 外设`n（取指、读写数据）" @(215,210,240)

# 数据流箭头
Draw-Arrow 390 150 428 150 @(60,120,190)     # 控制单元 -> 寄存器组
Draw-Arrow 430 330 392 370 @(140,90,40)       # 寄存器组 -> ALU
Draw-Arrow 392 400 428 360 @(60,140,90)       # ALU -> 寄存器组
Draw-Arrow 200 295 200 328 @(60,120,190)      # 控制单元 -> ALU（控制）
Draw-Arrow 230 490 230 428 @(140,90,40)       # 寄存器组 <-> 总线接口（双向之一）
Draw-Arrow 690 428 690 458 @(140,90,40)       # 寄存器组 <-> 总线接口（双向之二）

# 数据流标注
$g.DrawString("控制信号", $fText, [System.Drawing.Brushes]::Gray, 395, 128)
$g.DrawString("操作数", $fText, [System.Drawing.Brushes]::Gray, 400, 316)
$g.DrawString("结果写回", $fText, [System.Drawing.Brushes]::Gray, 356, 400)
$g.DrawString("控制", $fText, [System.Drawing.Brushes]::Gray, 205, 300)
$g.DrawString("读写数据", $fText, [System.Drawing.Brushes]::Gray, 640, 420)
$g.DrawString("取指/写回", $fText, [System.Drawing.Brushes]::Gray, 240, 495)

# 底部说明
$g.DrawString("PC 指向下一条要执行的指令地址；SP 指向栈顶；指令经总线接口从 Flash 取回，经译码后由 ALU 执行，结果写回寄存器或内存。", $fText, [System.Drawing.Brushes]::Gray, 40, 692)

$out = Join-Path $PSScriptRoot 'output.png'
$bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "saved: $out"
