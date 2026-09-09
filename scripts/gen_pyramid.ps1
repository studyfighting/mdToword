# 生成存储金字塔示意图（保留表格信息、分层图形，最底层与 md 保持一致）
Add-Type -AssemblyName System.Drawing

$W = 1400; $H = 760
$bmp = New-Object System.Drawing.Bitmap $W, $H
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
$g.Clear([System.Drawing.Color]::White)

$fTitle = New-Object System.Drawing.Font("Microsoft YaHei", 16, [System.Drawing.FontStyle]::Bold)
$fName  = New-Object System.Drawing.Font("Microsoft YaHei", 12, [System.Drawing.FontStyle]::Bold)
$fNote  = New-Object System.Drawing.Font("Microsoft YaHei", 9)

$g.DrawString("嵌入式存储金字塔（按 CPU 访问速度从快到慢）", $fTitle, [System.Drawing.Brushes]::Black, 20, 12)

$layers = @(
  @{ name = "① 寄存器 Registers";            note = "✅ 易失 · 1个时钟周期 · CPU 直接操作临时变量";                       w = 260; rgb = @(255,200,200) },
  @{ name = "② Cache (SRAM)";                note = "✅ 易失 · 2-10个时钟周期 · 硬件自动缓存热门数据";                   w = 380; rgb = @(255,220,170) },
  @{ name = "③ 内部 SRAM（主内存）";          note = "✅ 易失 · 几十个时钟周期 · 变量 / 堆栈 / DMA 缓冲区";                w = 500; rgb = @(255,240,150) },
  @{ name = "④ 外部 DRAM (SDRAM/DDR)";       note = "✅ 易失 · 上百个时钟周期 · 大容量扩展内存(如 Linux 系统)";          w = 620; rgb = @(220,240,170) },
  @{ name = "⑤ 内部 FLASH (NOR Flash)";      note = "❌ 非易失 · 读慢·写/擦极慢 · 存程序代码(.text)，支持 XIP";          w = 760; rgb = @(200,235,210) },
  @{ name = "⑥ 外部 FLASH (SPI NOR/NAND)";   note = "❌ 非易失 · 读较慢·写/擦极慢 · 存文件系统/资源/OTA；NAND 须载入 RAM"; w = 910; rgb = @(190,225,240) },
  @{ name = "⑦ ROM / EEPROM / eFuse";        note = "❌ 非易失 · 最慢 · 存 BootROM / UID / 密钥(eFuse) / 配置参数(EEPROM)"; w = 1060; rgb = @(215,210,240) }
)

$centerX = 700
$y = 62
$layerH = 84
$gap = 8

foreach($L in $layers){
  $x = $centerX - [int]($L.w / 2)
  $sb = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, $L.rgb[0], $L.rgb[1], $L.rgb[2]))
  $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(90,100,110)), 2
  $g.FillRectangle($sb, $x, $y, $L.w, $layerH)
  $g.DrawRectangle($pen, $x, $y, $L.w, $layerH)

  $fmtC = New-Object System.Drawing.StringFormat
  $fmtC.Alignment = [System.Drawing.StringAlignment]::Center
  $fmtC.LineAlignment = [System.Drawing.StringAlignment]::Center

  # 名称（上半部）
  $g.DrawString($L.name, $fName, [System.Drawing.Brushes]::Black, [System.Drawing.RectangleF]::new($x + 10, $y + 6, $L.w - 20, 34), $fmtC)
  # 说明（下半部）
  $g.DrawString($L.note, $fNote, [System.Drawing.Brushes]::Black, [System.Drawing.RectangleF]::new($x + 10, $y + 42, $L.w - 20, 34), $fmtC)

  $sb.Dispose(); $pen.Dispose()
  $y += $layerH + $gap
}

$g.DrawString("越靠上：越快、越贵、容量越小（易失）；越靠下：越慢、越便宜、容量越大（非易失）。", $fNote, [System.Drawing.Brushes]::Gray, 20, ($y + 6))

$out = Join-Path $PSScriptRoot 'output.png'
$bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "saved: $out"
