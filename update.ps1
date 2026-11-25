#requires -Version 5.1
<#
  WS Stock Analyzer - جمع البيانات والتحليل (الأسهم الحلال في البورصة المصرية)
  1) الأسهم وقطاعاتها + الفلترة الشرعية : stocks.templatesnippet.com
  2) القيمة العادلة                      : arabicstock.com
  3) التحليل الفني (كل المؤشرات + فريمات متعددة) : TradingView scanner
  ثم يحسب درجة لكل سهم، يرتبهم، يكتب data.js ويفتح الواجهة.
#>
param(
    [switch]$NoOpen,          # لا تفتح الواجهة بعد التحديث
    [switch]$ForceRefresh,    # تجاهل الكاش وأعد تحميل القيم العادلة
    [switch]$PricesOnly,      # v2.2: تحديث الأسعار الحالية بس (تحديث 12:00) من غير إعادة التحليل والترتيب
    [switch]$Auto             # v3.1: المهمة المجدولة — بيقرر لوحده: تحديث كامل لو فيه جلسة خلصت ومتحدثتش، أو أسعار بس بعد 12
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ======================= الإعدادات =======================
$Version           = 'v3.1'
$IncludeDoubtful   = $false   # true = اقبل الأسهم "المشكوك فيها" شرعياً
$DailyDays         = 5        # عدد الجلسات اللي يتعرض تغيرها اليومي (أسبوع تداول = 5)
$FullAfter         = '14:40'  # التحديث الكامل بيتعمل بعد الميعاد ده في أيام التداول (المهمة 2:45)
$SessionClose      = '14:30'  # نهاية جلسة البورصة: سعر إغلاق النهارده بيتسجل في الكاش بعدها بس (قبلها التغير لحظي)
$RequestDelayMs    = 250      # مهلة بين طلبات arabicstock
# أوزان الدرجة النهائية (المجموع = 1)
$W_Tech = 0.40; $W_Value = 0.35; $W_Quality = 0.15; $W_Liquidity = 0.10
# =========================================================

$Root    = Split-Path -Parent $MyInvocation.MyCommand.Path
$Cache   = Join-Path $Root 'cache'
$LogFile = Join-Path $Root 'update.log'
New-Item -ItemType Directory -Force $Cache | Out-Null
$UA  = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36'
$Inv = [Globalization.CultureInfo]::InvariantCulture
$Utf8 = New-Object System.Text.UTF8Encoding($false)
$Errors = New-Object System.Collections.ArrayList

function Write-Log([string]$m) {
    $line = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $m
    Write-Host $line
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

function Get-Text([string]$url) {
    for ($i = 1; $i -le 3; $i++) {
        try {
            $r = Invoke-WebRequest -Uri $url -UserAgent $UA -UseBasicParsing -TimeoutSec 45
            return [Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray())
        } catch {
            if ($i -eq 3) { throw }
            Start-Sleep -Seconds (2 * $i)
        }
    }
}

function Get-Num($s) {
    if ($null -eq $s) { return $null }
    $m = [regex]::Match(([string]$s -replace ',', ''), '-?\d+(\.\d+)?')
    if ($m.Success) { return [double]::Parse($m.Value, $Inv) } else { return $null }
}
function Get-Clamp([double]$v, [double]$lo, [double]$hi) { [Math]::Max($lo, [Math]::Min($hi, $v)) }
function Get-Round($v, $d = 2) { if ($null -eq $v) { $null } else { [Math]::Round([double]$v, $d) } }

function Read-JsonFile($path) {
    if (Test-Path $path) { try { return ([IO.File]::ReadAllText($path, $Utf8) | ConvertFrom-Json) } catch { } }
    return $null
}
function Save-JsonFile($obj, $path) { [IO.File]::WriteAllText($path, ($obj | ConvertTo-Json -Depth 8 -Compress), $Utf8) }

# ------------------------------------------------------------------
# زخم الشراء (v3.1): رقم من 0 لـ 100 بيقيس قوة الشراء دلوقتي، من غير ما يأثر على الترتيب
#   الحجم (20) + تدفق الفلوس CMF و MFI (20) + قوة الترند ADX/DMI (15) + MACD (15)
#   + RSI (10) + الاتجاه (المتوسطات) (10) + اختراق قمة الشهر بحجم (10)
# ------------------------------------------------------------------
$MomCols = @('relative_volume_10d_calc', 'MoneyFlow', 'ChaikinMoneyFlow', 'ADX', 'ADX+DI', 'ADX-DI',
    'MACD.macd', 'MACD.signal', 'MACD.macd[1]', 'MACD.signal[1]', 'RSI', 'RSI[1]', 'EMA20', 'EMA50', 'High.1M')
function Get-Momentum($t) {
    $n = { param($k) $v = $t[$k]; if ($null -eq $v) { $null } else { [double]$v } }
    $close = & $n 'close'; $chg = & $n 'change'; if ($null -eq $close) { return $null }
    $parts = New-Object System.Collections.ArrayList
    $add = { param($name, $pts, $max, $detail) [void]$parts.Add([pscustomobject]@{ name = $name; pts = [Math]::Round($pts, 1); max = $max; detail = $detail }) }

    # 1) الحجم: حجم أعلى من المتوسط والسهم طالع = شراء حقيقي. حجم عالي والسهم نازل = بيع
    $rv = & $n 'relative_volume_10d_calc'
    $p = if ($null -ne $rv -and $null -ne $chg -and $chg -gt 0) { (Get-Clamp (($rv - 0.5) / 1.5) 0 1) * 20 } else { 0 }
    & $add 'حجم التداول مع الصعود' $p 20 $(if ($null -ne $rv) { '{0:N1}x المتوسط، السهم {1}' -f $rv, $(if ($chg -gt 0) { 'طالع' } elseif ($chg -lt 0) { 'نازل' } else { 'ثابت' }) } else { '—' })

    # 2) تدفق الفلوس: CMF (فلوس داخلة/خارجة بالحجم) + MFI (RSI بالحجم)
    $cmf = & $n 'ChaikinMoneyFlow'; $mfi = & $n 'MoneyFlow'
    $pc = if ($null -ne $cmf) { (Get-Clamp (($cmf + 0.05) / 0.3) 0 1) * 10 } else { 0 }
    $pm = if ($null -eq $mfi) { 0 } elseif ($mfi -gt 85) { 5 } elseif ($mfi -ge 50) { (0.6 + 0.4 * (Get-Clamp (($mfi - 50) / 25) 0 1)) * 10 } else { (Get-Clamp (($mfi - 30) / 20) 0 1) * 4 }
    & $add 'تدفق الفلوس (CMF + MFI)' ($pc + $pm) 20 ('CMF {0} · MFI {1}' -f $(if ($null -ne $cmf) { $cmf.ToString('0.00', $Inv) } else { '—' }), $(if ($null -ne $mfi) { $mfi.ToString('0', $Inv) } else { '—' }))

    # 3) قوة الترند: ADX عالي و +DI فوق −DI
    $adx = & $n 'ADX'; $pdi = & $n 'ADX+DI'; $mdi = & $n 'ADX-DI'
    $p = if ($null -ne $adx -and $null -ne $pdi -and $null -ne $mdi -and $pdi -gt $mdi) { (Get-Clamp (($adx - 15) / 20) 0 1) * 15 } else { 0 }
    & $add 'قوة الترند (ADX/DMI)' $p 15 $(if ($null -ne $adx) { 'ADX {0:N0} · +DI {1:N0} / −DI {2:N0}' -f $adx, $pdi, $mdi } else { '—' })

    # 4) MACD: فوق خط الإشارة، والفرق بيكبر
    $m = & $n 'MACD.macd'; $s = & $n 'MACD.signal'; $m1 = & $n 'MACD.macd[1]'; $s1 = & $n 'MACD.signal[1]'
    $p = 0; $det = '—'
    if ($null -ne $m -and $null -ne $s) {
        $grow = ($null -ne $m1 -and $null -ne $s1) -and (($m - $s) -gt ($m1 - $s1))
        $p = if ($m -gt $s) { 9 + $(if ($grow) { 6 } else { 0 }) } elseif ($grow) { 4 } else { 0 }
        $det = '{0} خط الإشارة{1}' -f $(if ($m -gt $s) { 'فوق' } else { 'تحت' }), $(if ($grow) { '، والزخم بيزيد' } else { '، والزخم بيقل' })
    }
    & $add 'MACD' $p 15 $det

    # 5) RSI: المنطقة الصحية 55-70، فوق 75 تشبع شرائي
    $rsi = & $n 'RSI'; $rsi1 = & $n 'RSI[1]'
    $p = if ($null -eq $rsi) { 0 } elseif ($rsi -gt 75) { 2 } elseif ($rsi -gt 70) { 6 } elseif ($rsi -ge 55) { 10 } elseif ($rsi -ge 50) { 6 } elseif ($null -ne $rsi1 -and $rsi -gt $rsi1) { 2 } else { 0 }
    & $add 'RSI' $p 10 $(if ($null -ne $rsi) { '{0:N0}{1}' -f $rsi, $(if ($rsi -gt 75) { ' (تشبع شرائي)' } elseif ($null -ne $rsi1) { $(if ($rsi -gt $rsi1) { ' وطالع' } else { ' ونازل' }) } else { '' }) } else { '—' })

    # 6) الاتجاه: السعر فوق EMA20، و EMA20 فوق EMA50
    $e20 = & $n 'EMA20'; $e50 = & $n 'EMA50'
    $p = $(if ($e20 -and $close -gt $e20) { 5 } else { 0 }) + $(if ($e20 -and $e50 -and $e20 -gt $e50) { 5 } else { 0 })
    & $add 'الاتجاه (EMA20 / EMA50)' $p 10 $(if ($e20) { 'السعر {0} EMA20، و EMA20 {1} EMA50' -f $(if ($close -gt $e20) { 'فوق' } else { 'تحت' }), $(if ($e50 -and $e20 -gt $e50) { 'فوق' } else { 'تحت' }) } else { '—' })

    # 7) اختراق: السعر عند أو فوق أعلى سعر في الشهر، بحجم أعلى من المتوسط
    $hi = & $n 'High.1M'
    $near = $hi -and $close -ge $hi * 0.99
    $breakout = $near -and $null -ne $rv -and $rv -ge 1.5 -and $chg -gt 0
    $p = if ($breakout) { 10 } elseif ($near) { 5 } elseif ($hi -and $close -ge $hi * 0.95) { 2 } else { 0 }
    & $add 'اختراق قمة الشهر' $p 10 $(if ($hi) { 'قمة الشهر {0} · السعر {1:N1}% منها' -f $hi, (($close / $hi - 1) * 100) } else { '—' })

    $score = 0.0; foreach ($x in $parts) { $score += $x.pts }
    $label = if ($score -ge 70) { 'strong' } elseif ($score -ge 55) { 'pos' } elseif ($score -ge 40) { 'neutral' } else { 'weak' }
    [pscustomobject]@{ score = [Math]::Round($score, 0); label = $label; breakout = [bool]$breakout; rv = Get-Round $rv 2; parts = @($parts) }
}

$DataFile = Join-Path $Root 'data.js'

# ------------------------------------------------------------------
# حالة التحديث (v3.1): status.js بيتكتب طول التحديث، والواجهة بتعرض منه شريط التقدم
# ------------------------------------------------------------------
$StatusFile = Join-Path $Root 'status.js'
$script:statusStart = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss'); $script:statusMode = 'full'
function Set-Status([double]$pct, [string]$msg, [switch]$Done, [string]$Err) {
    $o = [ordered]@{ running = -not $Done; mode = $script:statusMode; pct = [Math]::Round([Math]::Min(100, [Math]::Max(0, $pct))); msg = $msg
        startedAt = $script:statusStart; at = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss'); error = $(if ($Err) { $Err } else { $null }) }
    try { [IO.File]::WriteAllText($StatusFile, 'window.EGX_STATUS = ' + ($o | ConvertTo-Json -Compress) + ';', $Utf8) } catch { }
}
# لو التحديث وقع في أي خطوة، الشريط يقول كده بدل ما يفضل "شغال" على طول
trap { Set-Status 100 'التحديث ما كملش' -Done -Err $_.Exception.Message; Write-Log "  ! خطأ: $($_.Exception.Message)"; break }

# ------------------------------------------------------------------
# التشغيل التلقائي (v3.1): المهمة بتشتغل 12:00 و 2:45 وكمان لما الجهاز يصحى أو تسجل دخول،
# فلو الجهاز كان نايم وقت الميعاد، التحديث اللي فات بيتعمل أول ما يصحى
# ------------------------------------------------------------------
if ($Auto) {
    $now = Get-Date
    $isTD = { param($d) @('Friday', 'Saturday') -notcontains $d.DayOfWeek.ToString() }
    $meta = if (Test-Path $DataFile) { [IO.File]::ReadAllText($DataFile, $Utf8) } else { '' }
    $gm = [regex]::Match($meta, '"generatedAt":"([^"]+)"'); $pm = [regex]::Match($meta, '"pricesAt":"([^"]+)"')
    $lastFull = if ($gm.Success) { [datetime]::Parse($gm.Groups[1].Value, $Inv) } else { [datetime]::MinValue }
    $lastPx = if ($pm.Success) { [datetime]::Parse($pm.Groups[1].Value, $Inv) } else { $lastFull }
    if ($lastPx -lt $lastFull) { $lastPx = $lastFull }
    # آخر جلسة خلصت: أقرب يوم تداول الساعة $FullAfter (النهارده لو عدّينا الميعاد)
    $lastClose = $null
    for ($k = 0; $k -lt 10 -and -not $lastClose; $k++) {
        $d = $now.Date.AddDays(-$k)
        if (& $isTD $d) { $c = [datetime]::ParseExact($d.ToString('yyyy-MM-dd') + " $FullAfter", 'yyyy-MM-dd HH:mm', $Inv); if ($c -le $now) { $lastClose = $c } }
    }
    $noon = [datetime]::ParseExact($now.ToString('yyyy-MM-dd') + ' 12:00', 'yyyy-MM-dd HH:mm', $Inv)
    if ($lastClose -and $lastFull -lt $lastClose) {
        Write-Log "تشغيل تلقائي: آخر تحديث كامل $($lastFull.ToString('yyyy-MM-dd HH:mm')) قبل جلسة $($lastClose.ToString('yyyy-MM-dd')) - تحديث كامل"
    } elseif ((& $isTD $now) -and $now -ge $noon -and $lastPx -lt $noon) {
        Write-Log 'تشغيل تلقائي: تحديث الأسعار (بعد 12)'; $PricesOnly = $true
    } else { return }   # كله متحدث، مفيش حاجة تتعمل
}
if ($PricesOnly -and -not (Test-Path $DataFile)) { $PricesOnly = $false }   # مفيش تحليل سابق: نعمل تحديث كامل
if ($PricesOnly) { $script:statusMode = 'prices' }
Set-Status 1 'بدء التحديث...'
# الواجهة بتفتح مع بداية التحديث، فتشوف شريط التقدم وهو شغال
if (-not $NoOpen) { Start-Process (Join-Path $Root 'dashboard.html') }

# ------------------------------------------------------------------
# 0) تحديث الأسعار فقط (v2.2): السعر وتغير اليوم من TradingView، والترتيب والتحليل زي ما هم
# ------------------------------------------------------------------
if ($PricesOnly) {
    Write-Log '===== تحديث الأسعار فقط ====='
    Set-Status 10 'تحميل الأسعار الحالية من TradingView...'
    $raw = [IO.File]::ReadAllText($DataFile, $Utf8).Trim()
    $data = ($raw -replace '^window\.EGX_DATA\s*=\s*', '' -replace ';\s*$', '') | ConvertFrom-Json
    $pxCols = @('name', 'close', 'change', 'time') + $MomCols
    $body = @{ filter = @(@{ left = 'type'; operation = 'equal'; right = 'stock' }); columns = $pxCols; range = @(0, 1000) } | ConvertTo-Json -Depth 5
    $tvp = Invoke-RestMethod -Uri 'https://scanner.tradingview.com/egypt/scan' -Method Post -Body $body -ContentType 'application/json' -UserAgent $UA -TimeoutSec 60
    $px = @{}; $pxT = @{}
    foreach ($row in $tvp.data) {
        $k = ([string]$row.d[0]).ToUpper(); $px[$k] = $row.d
        $h = @{}; for ($i = 0; $i -lt $pxCols.Count; $i++) { $h[$pxCols[$i]] = $row.d[$i] }; $pxT[$k] = $h
    }
    $dOf = { param($tm) if ($tm) { [DateTimeOffset]::FromUnixTimeSeconds([long]$tm).UtcDateTime.ToString('yyyy-MM-dd') } }
    # آخر جلسة على TradingView: لو جلسة جديدة بدأت، تتضاف أول عمود في التغير اليومي وأقدم عمود يتشال
    $tvDate = @($px.Values | ForEach-Object { & $dOf $_[3] } | Where-Object { $_ } | Group-Object | Sort-Object Count -Descending | Select-Object -First 1).Name
    $days = @($data.days)
    $newDay = $tvDate -and ($days.Count -eq 0 -or $tvDate -gt $days[0])
    $sameDay = $tvDate -and $days.Count -and $tvDate -eq $days[0]
    if ($newDay) { $data.days = @(@($tvDate) + @($days | Select-Object -First ($DailyDays - 1))) }
    $n = 0
    foreach ($s in @($data.stocks)) {
        $d = $px[$s.symbol]; $todayPct = $null
        if ($d -and $null -ne $d[1]) {
            $s.price = Get-Round $d[1] 3; $s.change = Get-Round $d[2] 2; $n++
            if ((& $dOf $d[3]) -eq $tvDate) { $todayPct = Get-Round $d[2] 2 }   # سهم ما اتداولش النهارده = —
            $s | Add-Member momentum (Get-Momentum $pxT[$s.symbol]) -Force    # الزخم بيتحدث مع الأسعار
        }
        # بالـ index مش بالـ pipeline، عشان الأيام الفاضية (null) ما تتشالش وتلخبط الأعمدة
        $old = @($s.daily); $nd = New-Object System.Collections.Generic.List[object]
        if ($newDay) {
            $nd.Add($todayPct)
            for ($i = 0; $i -lt $DailyDays - 1; $i++) { $nd.Add($(if ($i -lt $old.Count) { $old[$i] } else { $null })) }
            $s.daily = $nd.ToArray()
        } elseif ($sameDay) {
            for ($i = 0; $i -lt $DailyDays; $i++) { $nd.Add($(if ($i -eq 0) { $todayPct } elseif ($i -lt $old.Count) { $old[$i] } else { $null })) }
            $s.daily = $nd.ToArray()
        }
    }
    if ($newDay) { Write-Log "  جلسة جديدة ($tvDate) اتضافت للتغير اليومي" }
    Set-Status 80 'حفظ الأسعار والزخم...'
    if ($data.purif) { foreach ($p in $data.purif.PSObject.Properties) { $d = $px[$p.Name]; if ($d -and $null -ne $d[1]) { $p.Value.price = Get-Round $d[1] 3 } } }
    $data | Add-Member pricesAt ((Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')) -Force
    [IO.File]::WriteAllText($DataFile, 'window.EGX_DATA = ' + ($data | ConvertTo-Json -Depth 10 -Compress) + ';', $Utf8)
    Write-Log ("تم تحديث أسعار {0} سهم (الترتيب والتحليل من آخر تحديث كامل {1})" -f $n, $data.generatedAt)
    Set-Status 100 ("تم تحديث أسعار {0} سهم" -f $n) -Done
    return
}

Write-Log '===== بدء التحديث ====='

# ------------------------------------------------------------------
# 1) الفحص الشرعي + القطاعات
# ------------------------------------------------------------------
Write-Log 'تحميل بيانات الفحص الشرعي...'
Set-Status 2 'تحميل الفحص الشرعي...'
$shariahCacheFile = Join-Path $Cache 'shariah.json'
try {
    $shRaw = Get-Text 'https://stocks.templatesnippet.com/data/stocks_merged.json'
    [IO.File]::WriteAllText($shariahCacheFile, $shRaw, $Utf8)
} catch {
    [void]$Errors.Add("الفحص الشرعي: $($_.Exception.Message) - تم استخدام آخر نسخة محفوظة")
    if (-not (Test-Path $shariahCacheFile)) { throw 'تعذر تحميل بيانات الفحص الشرعي ولا توجد نسخة محفوظة' }
    $shRaw = [IO.File]::ReadAllText($shariahCacheFile, $Utf8)
}
$shariah = $shRaw | ConvertFrom-Json

# نفس منطق الحكم الإجمالي المستخدم في موقع "اسهم حلال"
$SourceNames = @{
    halal_bourse = 'بورصة حلال'; musaffa = 'مصفّى'; kashif = 'كاشف'; halal_invest = 'حلال إنفست'; faisal_bank = 'بنك فيصل'
    osoul = 'أسطول'; thndr = 'Islamicly / Thndr'; mohamed_ramadan_sharia = 'جروب الأسهم الشرعية فقط'
    egx33_sharia = 'مؤشر الشريعة EGX33'; nourborsa = 'نور البورصة'
}
function Get-ShariahVerdict($e) {
    $sources = @(); $compliantCount = 0
    foreach ($p in $e.PSObject.Properties) {
        $v = $p.Value
        if ($v -is [System.Management.Automation.PSCustomObject] -and $v.PSObject.Properties['status']) {
            $nm = if ($SourceNames.ContainsKey($p.Name)) { $SourceNames[$p.Name] } else { $p.Name }
            $sources += [pscustomobject]@{ key = $p.Name; name = $nm; status = $(if ($null -ne $v.status) { $v.status } else { 'not_checked' }); note = $v.note }
            if ($null -eq $v.status) { continue }
            if ($v.status -eq 'compliant') { $compliantCount++ }
        }
    }
    $t = $e.haram_earnings_percentage; $a = $e.loans_percentage
    $verdict =
        if ($e.core_activity_compliant -eq $false) { 'non_compliant' }
        elseif (($null -ne $a -and $a -gt 50) -or $e.haram_investments_compliant -eq $false) { 'non_compliant' }
        elseif ($compliantCount -gt 2) { 'compliant' }
        elseif ($null -eq $t -and $null -eq $a) { 'doubtful' }
        elseif (($null -eq $t -or $t -le 0) -and ($null -eq $a -or $a -le 0)) { 'no_loans' }
        elseif (($null -eq $t -or $t -le 5) -and ($null -eq $a -or $a -le 33)) { 'compliant' }
        else { 'doubtful' }
    return [pscustomobject]@{
        verdict = $verdict; compliantSources = $compliantCount
        totalSources = @($sources | Where-Object { $_.status -ne 'not_checked' }).Count
        sources = $sources
    }
}

$shMap = @{}
foreach ($e in $shariah) {
    $v = Get-ShariahVerdict $e
    $shMap[$e.symbol.ToUpper()] = [pscustomobject]@{
        symbol = $e.symbol.ToUpper(); nameAr = $e.name_ar; sector = $e.category_ar; sectorEn = $e.category
        verdict = $v.verdict; compliantSources = $v.compliantSources; totalSources = $v.totalSources
        haramPct = $e.haram_earnings_percentage; loansPct = $e.loans_percentage; sources = $v.sources
    }
}
Write-Log ("الفحص الشرعي: {0} سهم" -f $shMap.Count)

# ------------------------------------------------------------------
# 2) التحليل الفني من TradingView (كل أسهم البورصة المصرية)
# ------------------------------------------------------------------
Write-Log 'تحميل التحليل الفني من TradingView...'
Set-Status 5 'تحميل التحليل الفني من TradingView...'
$tvCols = @('name','description','sector','close','change','volume',
    'Recommend.All','Recommend.MA','Recommend.Other',
    'Recommend.All|60','Recommend.All|240','Recommend.All|1W','Recommend.All|1M',
    'RSI','MACD.macd','MACD.signal','Stoch.K','Stoch.D','ADX','ADX+DI','ADX-DI','CCI20','W.R','Mom','AO',
    'BB.upper','BB.lower','EMA20','EMA50','EMA200','SMA50','SMA200','Ichimoku.BLine','P.SAR','UO','Stoch.RSI.K',
    'Pivot.M.Classic.S1','Pivot.M.Classic.R1','Perf.W','Perf.1M','Perf.3M','relative_volume_10d_calc','ATR',
    'market_cap_basic','price_earnings_ttm','average_volume_30d_calc','Perf.6M','Perf.Y','Perf.YTD','change[1]',
    'time','close[1]','close[2]','MoneyFlow','ChaikinMoneyFlow','MACD.macd[1]','MACD.signal[1]','RSI[1]','High.1M')
$tvBody = @{ filter = @(@{ left = 'type'; operation = 'equal'; right = 'stock' }); columns = $tvCols; range = @(0, 1000) } | ConvertTo-Json -Depth 5
$tv = Invoke-RestMethod -Uri 'https://scanner.tradingview.com/egypt/scan' -Method Post -Body $tvBody -ContentType 'application/json' -UserAgent $UA -TimeoutSec 60
$tvMap = @{}
foreach ($row in $tv.data) {
    $o = @{}
    for ($i = 0; $i -lt $tvCols.Count; $i++) { $o[$tvCols[$i]] = $row.d[$i] }
    $tvMap[([string]$o['name']).ToUpper()] = $o
}
Write-Log ("TradingView: {0} سهم مدرج" -f $tvMap.Count)

# ------------------------------------------------------------------
# 3) الفلترة الشرعية
# ------------------------------------------------------------------
$allowed = @('compliant', 'no_loans'); if ($IncludeDoubtful) { $allowed += 'doubtful' }
$kept = @(); $excluded = @()
foreach ($sym in ($tvMap.Keys | Sort-Object)) {
    $t = $tvMap[$sym]; $s = $shMap[$sym]
    if ($null -eq $s) {
        $excluded += [pscustomobject]@{ symbol = $sym; name = $t['description']; sector = $t['sector']; reason = 'not_screened' }
    } elseif ($allowed -notcontains $s.verdict) {
        $excluded += [pscustomobject]@{ symbol = $sym; name = $s.nameAr; sector = $s.sector; reason = $s.verdict }
    } else { $kept += $sym }
}
Write-Log ("بعد الفلترة الشرعية: {0} سهم متوافق، {1} مستبعد" -f $kept.Count, $excluded.Count)

# ------------------------------------------------------------------
# 4) القيمة العادلة من arabicstock.com (كاش يومي لأن الموقع يحدّث بعد إغلاق الجلسة)
# ------------------------------------------------------------------
$today = (Get-Date).ToString('yyyy-MM-dd')
$urlMapFile = Join-Path $Cache 'arabicstock_urls.json'
$fvCacheFile = Join-Path $Cache 'fairvalue.json'

$urlMap = @{}
$urlCache = Read-JsonFile $urlMapFile
$urlFresh = $urlCache -and $urlCache.date -and ((Get-Date) - [datetime]::Parse($urlCache.date, $Inv)).TotalDays -lt 7
if ($urlFresh -and -not $ForceRefresh) {
    foreach ($p in $urlCache.map.PSObject.Properties) { $urlMap[$p.Name] = $p.Value }
} else {
    Write-Log 'بناء فهرس صفحات arabicstock...'
    $catBase = 'https://arabicstock.com/category/' + [uri]::EscapeDataString('أسهم-البورصة-المصرية')
    $page = 1; $maxPage = 1
    do {
        try {
            $html = Get-Text ($catBase + $(if ($page -gt 1) { "?page=$page" } else { '' }))
            foreach ($m in [regex]::Matches($html, '\?page=(\d+)')) { $maxPage = [Math]::Max($maxPage, [int]$m.Groups[1].Value) }
            foreach ($m in [regex]::Matches($html, 'href="(https://arabicstock\.com/\d+/[^"]+)"')) {
                $u = $m.Groups[1].Value
                $sym = ($u.Split('/')[-1].Split('-')[-1]).ToUpper()
                if ($sym -match '^[A-Z][A-Z0-9]{1,7}$' -and -not $urlMap.ContainsKey($sym)) { $urlMap[$sym] = $u }
            }
        } catch { [void]$Errors.Add("arabicstock صفحة $page : $($_.Exception.Message)") }
        $page++; Start-Sleep -Milliseconds $RequestDelayMs
    } while ($page -le $maxPage -and $page -le 40)
    if ($urlMap.Count -gt 0) { Save-JsonFile ([pscustomobject]@{ date = $today; map = $urlMap }) $urlMapFile }
    elseif ($urlCache) { foreach ($p in $urlCache.map.PSObject.Properties) { $urlMap[$p.Name] = $p.Value } }
    Write-Log ("فهرس arabicstock: {0} سهم" -f $urlMap.Count)
}

function Get-FairValue([string]$url, [string]$html) {
    $cells = @{}; $cls = @{}
    foreach ($m in [regex]::Matches($html, '(?s)data-label="([^"]+)">\s*<div class="content_divtxt([^"]*)">\s*(.*?)\s*</div>')) {
        $k = $m.Groups[1].Value
        if (-not $cells.ContainsKey($k)) {
            $cells[$k] = [Net.WebUtility]::HtmlDecode(($m.Groups[3].Value -replace '<[^>]+>', '').Trim())
            $cls[$k] = $m.Groups[2].Value
        }
    }
    $txt = [Net.WebUtility]::HtmlDecode((($html -replace '(?s)<script.*?</script>', '' -replace '(?s)<style.*?</style>', '' -replace '<[^>]+>', ' ') -replace '\s+', ' '))
    $healthCls = $cls['الأداء العام (الصحة المالية)']
    $health = if ($healthCls -match 'green') { 1 } elseif ($healthCls -match 'red') { -1 } elseif ($cells['الأداء العام (الصحة المالية)'] -and $cells['الأداء العام (الصحة المالية)'] -ne '—') { 0 } else { $null }
    $conf = $null
    $cm = [regex]::Match($txt, 'درجة الثقة في القيمة العادلة\s*!\s*(\S+)')
    if ($cm.Success) { $conf = $cm.Groups[1].Value }
    $model = {
        param($lbl)
        $mm = [regex]::Match($txt, [regex]::Escape($lbl) + '\s*!\s*(-?[\d,.]+)\s*جنيه')
        if ($mm.Success) { Get-Num $mm.Groups[1].Value } else { $null }
    }
    $pio = [regex]::Match($txt, 'Piotroski Score\s*!\s*(\d+)')
    $alt = [regex]::Match($txt, 'Altman Z-Score\s*!\s*(-?[\d.]+)')
    [pscustomobject]@{
        fetched      = $today
        url          = $url
        symbol       = $cells['رمز السهم']
        fairValue    = Get-Num $cells['القيمة العادلة للسهم']
        fairPct      = Get-Num $cells['القيمة العادلة (%)']
        fairClass    = $cells['تصنيف القيمة العادلة']
        analystFV    = Get-Num $cells['القيمة العادلة (هدف المحللين)']
        analystPct   = Get-Num $cells['القيمة العادلة (هدف المحللين) %']
        analystClass = $cells['تصنيف القيمة العادلة (هدف المحللين)']
        healthText   = $cells['الأداء العام (الصحة المالية)']
        health       = $health
        confidence   = $conf
        dy           = Get-Num $cells['نسبة توزيعات الأرباح السنوية DY']
        pe           = Get-Num $cells['معدل P/E']
        pb           = Get-Num $cells['مضاعف القيمة الدفترية Price/Book']
        debtCap      = Get-Num $cells['الديون/إجمالي رأس المال TD/TC']
        nextEarnings = $cells['تاريخ إعلان الأرباح القادم']
        piotroski    = if ($pio.Success) { [int]$pio.Groups[1].Value } else { $null }
        altmanZ      = if ($alt.Success) { Get-Num $alt.Groups[1].Value } else { $null }
        dcf          = & $model 'القيمة العادلة باستخدام DCF'
        graham       = & $model 'Ben Graham Formula'
        ddm          = & $model 'القيمة العادلة باستخدام DDM'
        epv          = & $model 'القيمة العادلة باستخدام EPV'
        comps        = & $model 'القيمة العادلة باستخدام Comps'
    }
}

$fvCache = @{}
$fvc = Read-JsonFile $fvCacheFile
if ($fvc) { foreach ($p in $fvc.PSObject.Properties) { $fvCache[$p.Name] = $p.Value } }
$todo = @($kept | Where-Object { $urlMap[$_] -and -not ($fvCache[$_] -and $fvCache[$_].fetched -eq $today -and -not $ForceRefresh) })
Write-Log ("القيمة العادلة: تحميل {0} صفحة..." -f $todo.Count)
$fetchedNow = 0; $fvi = 0
Set-Status 8 ("القيمة العادلة: 0 / {0}" -f $todo.Count)
foreach ($sym in $todo) {
    $fvi++; Set-Status (8 + 52 * $fvi / [Math]::Max(1, $todo.Count)) ("القيمة العادلة من السهم العربي: {0} / {1}" -f $fvi, $todo.Count)   # أبطأ خطوة (8% → 60%)
    try {
        $fvCache[$sym] = Get-FairValue $urlMap[$sym] (Get-Text $urlMap[$sym])
        $fetchedNow++
        if ($fetchedNow % 20 -eq 0) { Write-Log "  القيمة العادلة: $fetchedNow / $($todo.Count)"; Save-JsonFile $fvCache $fvCacheFile }
    } catch { [void]$Errors.Add("القيمة العادلة $sym : $($_.Exception.Message)") }
    Start-Sleep -Milliseconds $RequestDelayMs
}
Save-JsonFile $fvCache $fvCacheFile
Write-Log ("القيمة العادلة: تم تحديث {0} سهم (الباقي من كاش اليوم)" -f $fetchedNow)

# ------------------------------------------------------------------
# 4b) الأسعار اليومية لآخر أسبوع (Yahoo Finance: SYMBOL.CA)
#     لحساب نسبة التغير في كل جلسة على حدة
# ------------------------------------------------------------------
# v2.2: Yahoo ما بيديش سعر إغلاق أيام الأحد، فبقينا نسجل إغلاق كل جلسة من TradingView
#       في cache\closes.json مع كل تحديث كامل، و Yahoo بيكمّل الأيام القديمة بس
Write-Log 'تحميل الأسعار اليومية...'
$todayStr = (Get-Date).ToString('yyyy-MM-dd')
$closesFile = Join-Path $Cache 'closes.json'
$closeStore = @{}   # التاريخ -> (السهم -> سعر الإغلاق)
$cs = Read-JsonFile $closesFile
if ($cs) { foreach ($p in $cs.PSObject.Properties) { $m = @{}; foreach ($q in $p.Value.PSObject.Properties) { $m[$q.Name] = [double]$q.Value }; $closeStore[$p.Name] = $m } }

$yh = @{}; $dayCount = @{}; $histFail = 0; $yi = 0
foreach ($sym in $kept) {
    $yi++; if ($yi % 10 -eq 1) { Set-Status (60 + 15 * $yi / [Math]::Max(1, $kept.Count)) ("الأسعار اليومية: {0} / {1}" -f $yi, $kept.Count) }
    $h = @{}
    try {
        $r = Invoke-RestMethod -Uri "https://query1.finance.yahoo.com/v8/finance/chart/$sym.CA?range=1mo&interval=1d" -UserAgent $UA -TimeoutSec 20
        $res = $r.chart.result[0]; $ts = $res.timestamp; $cl = $res.indicators.quote[0].close
        for ($i = 0; $i -lt $ts.Count; $i++) {
            $d = [DateTimeOffset]::FromUnixTimeSeconds([long]$ts[$i]).UtcDateTime.ToString('yyyy-MM-dd')
            $dayCount[$d] = 1 + [int]$dayCount[$d]          # يوم التداول بيتحسب حتى لو السعر ناقص (زي الأحد)
            if ($null -ne $cl[$i]) { $h[$d] = [double]$cl[$i] }
        }
    } catch { $histFail++ }
    $yh[$sym] = $h
    Start-Sleep -Milliseconds 80
}
# آخر جلسة على TradingView (أغلب الأسهم)
$tvDate = @($kept | ForEach-Object { $tm = $tvMap[$_]['time']; if ($tm) { [DateTimeOffset]::FromUnixTimeSeconds([long]$tm).UtcDateTime.ToString('yyyy-MM-dd') } } |
    Group-Object | Sort-Object Count -Descending | Select-Object -First 1).Name
$sessionToday = $tvDate -eq $todayStr
# جلسة النهارده شغالة: بتتعرض في التغير اليومي على طول (تغير لحظي)، بس سعرها ما بيتسجلش غير بعد الإغلاق
$liveToday = $sessionToday -and ((Get-Date) -lt [datetime]::ParseExact("$todayStr $SessionClose", 'yyyy-MM-dd HH:mm', $Inv))

# تقويم الجلسات = أيام Yahoo (لمعظم الأسهم) + الأيام المسجلة عندنا + آخر جلسة على TradingView
$minCount = [Math]::Max(1, [int](($kept.Count - $histFail) * 0.3))
$calSet = @{}
foreach ($d in $dayCount.Keys) { if ($dayCount[$d] -ge $minCount) { $calSet[$d] = 1 } }
foreach ($d in $closeStore.Keys) { $calSet[$d] = 1 }
if ($tvDate) { $calSet[$tvDate] = 1 }
$lastDay = if ($tvDate) { $tvDate } else { $todayStr }   # ما نعرضش يوم جلسته لسه ما بدأتش
$calendar = @($calSet.Keys | Where-Object { $_ -le $lastDay } | Sort-Object -Descending)

# تسجيل أسعار TradingView: إغلاق آخر جلسة (لو خلصت) + الجلستين اللي قبلها
if ($tvDate) {
    $ix = [array]::IndexOf($calendar, $tvDate)
    $slots = @(@{ col = 'close'; d = $(if (-not $liveToday) { $tvDate }) })
    if ($ix -ge 0 -and $ix + 1 -lt $calendar.Count) { $slots += @{ col = 'close[1]'; d = $calendar[$ix + 1] } }
    if ($ix -ge 0 -and $ix + 2 -lt $calendar.Count) { $slots += @{ col = 'close[2]'; d = $calendar[$ix + 2] } }
    foreach ($sym in $kept) {
        $t = $tvMap[$sym]; if (-not $t['time']) { continue }
        if ([DateTimeOffset]::FromUnixTimeSeconds([long]$t['time']).UtcDateTime.ToString('yyyy-MM-dd') -ne $tvDate) { continue }   # سهم ما اتداولش في آخر جلسة
        foreach ($s in $slots) {
            if (-not $s.d -or $null -eq $t[$s.col]) { continue }
            if (-not $closeStore.ContainsKey($s.d)) { $closeStore[$s.d] = @{} }
            $closeStore[$s.d][$sym] = [double]$t[$s.col]
        }
    }
}
# حفظ آخر 40 جلسة
$csOut = [ordered]@{}
foreach ($d in @($closeStore.Keys | Sort-Object -Descending | Select-Object -First 40 | Sort-Object)) { $csOut[$d] = $closeStore[$d] }
Save-JsonFile ([pscustomobject]$csOut) $closesFile

# أسعار الإغلاق النهائية لكل سهم: المسجل عندنا من TradingView أولاً، وبعده Yahoo
$closeMap = @{}; $hist = @{}
foreach ($sym in $kept) {
    $m = @{}
    foreach ($d in $yh[$sym].Keys) { $m[$d] = $yh[$sym][$d] }
    foreach ($d in $closeStore.Keys) { if ($closeStore[$d].ContainsKey($sym)) { $m[$d] = $closeStore[$d][$sym] } }
    $t = $tvMap[$sym]
    if ($liveToday -and $t['time'] -and $null -ne $t['close'] -and [DateTimeOffset]::FromUnixTimeSeconds([long]$t['time']).UtcDateTime.ToString('yyyy-MM-dd') -eq $todayStr) {
        $m[$todayStr] = [double]$t['close']   # السعر اللحظي لجلسة النهارده
    }
    $closeMap[$sym] = $m
    $pts = New-Object System.Collections.ArrayList
    foreach ($d in ($m.Keys | Sort-Object)) { [void]$pts.Add([pscustomobject]@{ d = $d; c = $m[$d] }) }
    $hist[$sym] = $pts   # بتستخدم في رد فعل السعر على الأخبار
}
# الجلسات المعروضة: آخر 5 جلسات، وأول ما جلسة جديدة تبدأ بتتضاف وأقدم جلسة تتشال
$tradingDays = @($calendar | Select-Object -First $DailyDays)
if ($liveToday) { Write-Log "  جلسة النهارده لسه شغالة - تغيرها في الجدول لحظي لحد الإغلاق" }
if ($histFail -gt 0) { Write-Log "  الأسعار اليومية: $histFail سهم غير متاح على Yahoo" }
Write-Log ("الأسعار اليومية: آخر جلسة على TradingView {0}، الجلسات المعروضة: {1}" -f $tvDate, ($tradingDays -join ', '))

function Get-DailyChanges([string]$sym) {
    $m = $closeMap[$sym]
    foreach ($d in $tradingDays) {
        $pct = $null
        $i = [array]::IndexOf($calendar, $d)
        if ($m -and $i -ge 0 -and $i + 1 -lt $calendar.Count) {
            $pd = $calendar[$i + 1]   # الجلسة اللي قبلها مباشرة (لو سعرها ناقص بيظهر — بدل رقم غلط)
            if ($m.ContainsKey($d) -and $m.ContainsKey($pd) -and $m[$pd] -gt 0) { $pct = [Math]::Round(($m[$d] / $m[$pd] - 1) * 100, 2) }
        }
        $pct
    }
}

# ------------------------------------------------------------------
# 5) التحليل والتقييم
# ------------------------------------------------------------------
function Get-TechAnalysis($t) {
    $close = [double]$t['close']; $checks = @(); $sum = 0.0; $max = 0.0
    function Add-Check([string]$name, $cond, [double]$w, [string]$detail) {
        if ($null -eq $cond) { return }
        $script:max_ += $w
        $val = if ($cond -eq 1) { $w } elseif ($cond -eq -1) { -$w } else { 0 }
        $script:sum_ += $val
        $script:checks_ += [pscustomobject]@{ name = $name; signal = $cond; detail = $detail }
    }
    $script:sum_ = 0.0; $script:max_ = 0.0; $script:checks_ = @()
    $n = { param($k) $v = $t[$k]; if ($null -eq $v) { $null } else { [double]$v } }

    $e20 = & $n 'EMA20'; $e50 = & $n 'EMA50'; $e200 = & $n 'EMA200'
    if ($e20)  { Add-Check 'السعر مقابل EMA20'  ([int][Math]::Sign($close - $e20)) 1 ('EMA20 = {0:N2}' -f $e20) }
    if ($e20 -and $e50) { Add-Check 'EMA20 مقابل EMA50 (الاتجاه القصير)' ([int][Math]::Sign($e20 - $e50)) 1 ('{0:N2} / {1:N2}' -f $e20, $e50) }
    if ($e200) { Add-Check 'السعر مقابل EMA200 (الاتجاه الطويل)' ([int][Math]::Sign($close - $e200)) 1.5 ('EMA200 = {0:N2}' -f $e200) }
    $macd = & $n 'MACD.macd'; $sig = & $n 'MACD.signal'
    if ($null -ne $macd -and $null -ne $sig) { Add-Check 'MACD مقابل خط الإشارة' ([int][Math]::Sign($macd - $sig)) 1 ('{0:N3} / {1:N3}' -f $macd, $sig) }
    $rsi = & $n 'RSI'
    if ($null -ne $rsi) {
        $r = if ($rsi -gt 75) { -1 } elseif ($rsi -ge 50) { 1 } elseif ($rsi -lt 30) { 1 } elseif ($rsi -lt 40) { -1 } else { 0 }
        $d = if ($rsi -gt 75) { 'تشبع شرائي' } elseif ($rsi -lt 30) { 'تشبع بيعي (فرصة ارتداد)' } else { '' }
        Add-Check 'RSI (14)' $r 1 ('{0:N1} {1}' -f $rsi, $d)
    }
    $adx = & $n 'ADX'; $pdi = & $n 'ADX+DI'; $mdi = & $n 'ADX-DI'
    if ($null -ne $adx -and $null -ne $pdi -and $null -ne $mdi) {
        $r = if ($adx -lt 20) { 0 } else { [int][Math]::Sign($pdi - $mdi) }
        Add-Check 'ADX / DMI (قوة واتجاه الترند)' $r 1 ('ADX {0:N1}  +DI {1:N1}  -DI {2:N1}' -f $adx, $pdi, $mdi)
    }
    $sk = & $n 'Stoch.K'; $sd = & $n 'Stoch.D'
    if ($null -ne $sk -and $null -ne $sd) {
        $r = if ($sk -gt 80 -and $sk -lt $sd) { -1 } elseif ($sk -lt 20 -and $sk -gt $sd) { 1 } else { [int][Math]::Sign($sk - $sd) }
        Add-Check 'Stochastic %K/%D' $r 0.5 ('{0:N1} / {1:N1}' -f $sk, $sd)
    }
    $cci = & $n 'CCI20'
    if ($null -ne $cci) { $r = if ($cci -lt -100) { 1 } elseif ($cci -gt 200) { -1 } else { [int][Math]::Sign($cci) }; Add-Check 'CCI (20)' $r 0.5 ('{0:N0}' -f $cci) }
    $wr = & $n 'W.R'
    if ($null -ne $wr) { $r = if ($wr -lt -80) { 1 } elseif ($wr -gt -20) { -1 } else { 0 }; Add-Check 'Williams %R' $r 0.5 ('{0:N1}' -f $wr) }
    $ao = & $n 'AO';  if ($null -ne $ao)  { Add-Check 'Awesome Oscillator' ([int][Math]::Sign($ao)) 0.5 ('{0:N3}' -f $ao) }
    $mom = & $n 'Mom'; if ($null -ne $mom) { Add-Check 'Momentum (10)' ([int][Math]::Sign($mom)) 0.5 ('{0:N2}' -f $mom) }
    $ich = & $n 'Ichimoku.BLine'; if ($ich) { Add-Check 'Ichimoku (Base Line)' ([int][Math]::Sign($close - $ich)) 1 ('{0:N2}' -f $ich) }
    $psar = & $n 'P.SAR'; if ($psar) { Add-Check 'Parabolic SAR' ([int][Math]::Sign($close - $psar)) 1 ('{0:N2}' -f $psar) }
    $bbu = & $n 'BB.upper'; $bbl = & $n 'BB.lower'
    if ($bbu -and $bbl -and $bbu -gt $bbl) {
        $pos = ($close - $bbl) / ($bbu - $bbl)
        $r = if ($pos -gt 1) { -1 } elseif ($pos -lt 0.2) { 1 } else { 0 }
        Add-Check 'Bollinger Bands (موقع السعر)' $r 0.5 ('{0:P0} من النطاق' -f $pos)
    }
    $uo = & $n 'UO'; if ($null -ne $uo) { $r = if ($uo -lt 30) { 1 } elseif ($uo -gt 70) { -1 } else { [int][Math]::Sign($uo - 50) }; Add-Check 'Ultimate Oscillator' $r 0.5 ('{0:N1}' -f $uo) }
    $p1m = & $n 'Perf.1M'; if ($null -ne $p1m) { Add-Check 'أداء شهر' ([int][Math]::Sign($p1m)) 0.5 ('{0:N1}%' -f $p1m) }
    $p3m = & $n 'Perf.3M'; if ($null -ne $p3m) { Add-Check 'أداء 3 شهور' ([int][Math]::Sign($p3m)) 0.5 ('{0:N1}%' -f $p3m) }
    $rv = & $n 'relative_volume_10d_calc'; $chg = & $n 'change'
    if ($null -ne $rv -and $null -ne $chg -and $rv -gt 1.5) { Add-Check 'حجم تداول مرتفع مع اتجاه اليوم' ([int][Math]::Sign($chg)) 0.5 ('{0:N1}x متوسط الحجم' -f $rv) }

    $own = if ($script:max_ -gt 0) { $script:sum_ / $script:max_ } else { 0 }

    # تقييمات TradingView متعددة الفريمات (-1 .. +1)
    $tfs = [ordered]@{ 'Recommend.All|60' = 0.15; 'Recommend.All|240' = 0.15; 'Recommend.All' = 0.40; 'Recommend.All|1W' = 0.20; 'Recommend.All|1M' = 0.10 }
    $tvs = 0.0; $tvw = 0.0
    foreach ($k in $tfs.Keys) { $v = & $n $k; if ($null -ne $v) { $tvs += $v * $tfs[$k]; $tvw += $tfs[$k] } }
    $tvScore = if ($tvw -gt 0) { $tvs / $tvw } else { 0 }

    $score = 50 + 50 * (0.6 * $tvScore + 0.4 * $own)
    [pscustomobject]@{
        score = Get-Round $score 1; own = Get-Round $own 3; tvComposite = Get-Round $tvScore 3
        tf = [pscustomobject]@{
            h1 = Get-Round (& $n 'Recommend.All|60') 3; h4 = Get-Round (& $n 'Recommend.All|240') 3
            d1 = Get-Round (& $n 'Recommend.All') 3; w1 = Get-Round (& $n 'Recommend.All|1W') 3; m1 = Get-Round (& $n 'Recommend.All|1M') 3
            ma = Get-Round (& $n 'Recommend.MA') 3; osc = Get-Round (& $n 'Recommend.Other') 3
        }
        checks = $script:checks_
        support = Get-Round (& $n 'Pivot.M.Classic.S1'); resistance = Get-Round (& $n 'Pivot.M.Classic.R1')
        rsi = Get-Round $rsi 1; atr = Get-Round (& $n 'ATR') 3
    }
}

$prevRanksFile = Join-Path $Cache 'last_ranks.json'
$prevRanksDate = if (Test-Path $prevRanksFile) { (Get-Item $prevRanksFile).LastWriteTime.ToString('yyyy-MM-dd') } else { $null }
$prev = Read-JsonFile $prevRanksFile
$prevRank = @{}; if ($prev) { foreach ($p in $prev.PSObject.Properties) { $prevRank[$p.Name] = [int]$p.Value } }

Set-Status 76 'حساب الدرجات والترتيب...'
$results = @()
foreach ($sym in $kept) {
    $t = $tvMap[$sym]; $s = $shMap[$sym]; $fv = $fvCache[$sym]
    $close = [double]$t['close']
    $tech = Get-TechAnalysis $t

    # --- التقييم (القيمة العادلة) ---
    $upside = $null; $valueScore = $null; $analystUp = $null
    # لو سعر السهم على arabicstock مختلف جداً عن TradingView (عملة مختلفة مثل الأسهم المقومة بالدولار، أو تجزئة)
    # نستخدم نسبة الصعود المنشورة على الموقع بدلاً من حسابها على السعر الحالي
    $sameUnit = $false
    if ($fv -and $fv.fairValue -and $fv.fairValue -gt 0 -and $null -ne $fv.fairPct -and $close -gt 0) {
        $implied = $fv.fairValue / (1 + $fv.fairPct / 100)
        $ratio = $close / $implied
        $sameUnit = ($ratio -gt 0.65 -and $ratio -lt 1.5)
    }
    if ($sameUnit) {
        $upside = ($fv.fairValue - $close) / $close * 100
        if ($fv.analystFV -and $fv.analystFV -gt 0) { $analystUp = ($fv.analystFV - $close) / $close * 100 }
    } elseif ($fv -and $null -ne $fv.fairPct) {
        $upside = $fv.fairPct
        if ($null -ne $fv.analystPct) { $analystUp = $fv.analystPct }
    }
    if ($null -ne $upside) {
        $u = if ($null -ne $analystUp) { 0.7 * $upside + 0.3 * $analystUp } else { $upside }
        $confW = switch -Regex ([string]$fv.confidence) { 'عالي' { 1.0 } 'متوسط' { 0.8 } 'منخفض' { 0.6 } default { 0.75 } }
        $valueScore = 50 + (Get-Clamp $u -50 50) * $confW
    }

    # --- الجودة ---
    $q = @()
    if ($fv) {
        if ($null -ne $fv.health) { $q += $(switch ([int]$fv.health) { 1 { 85 } 0 { 55 } default { 20 } }) }
        if ($null -ne $fv.piotroski) { $q += $fv.piotroski / 9 * 100 }
        if ($null -ne $fv.altmanZ) { $q += (Get-Clamp (($fv.altmanZ - 1.1) / (3 - 1.1)) 0 1) * 100 }
    }
    $qualityScore = if ($q.Count) { ($q | Measure-Object -Average).Average } else { $null }

    # --- السيولة (قيمة التداول اليومي) ---
    $avgVol = if ($t['average_volume_30d_calc']) { [double]$t['average_volume_30d_calc'] } else { [double]$t['volume'] }
    $traded = $avgVol * $close
    $liqScore = if ($traded -gt 0) { (Get-Clamp ((([Math]::Log10($traded)) - 5) / 3) 0 1) * 100 } else { 0 }

    $vs = if ($null -ne $valueScore) { $valueScore } else { 40 }   # بدون قيمة عادلة = خصم
    $qs = if ($null -ne $qualityScore) { $qualityScore } else { 45 }
    $final = $W_Tech * $tech.score + $W_Value * $vs + $W_Quality * $qs + $W_Liquidity * $liqScore

    $results += [pscustomobject]@{
        symbol = $sym; name = $s.nameAr; nameEn = $t['description']
        sector = $(if ($s.sector) { $s.sector } else { 'أخرى' }); tvSector = $t['sector']
        price = Get-Round $close 3; change = Get-Round $t['change'] 2
        perfW = Get-Round $t['Perf.W'] 2; perf1M = Get-Round $t['Perf.1M'] 2; perf3M = Get-Round $t['Perf.3M'] 2
        perf6M = Get-Round $t['Perf.6M'] 2; perfY = Get-Round $t['Perf.Y'] 2; perfYTD = Get-Round $t['Perf.YTD'] 2
        daily = @(Get-DailyChanges $sym)
        volume = $t['volume']; tradedValue = [Math]::Round($traded); marketCap = $t['market_cap_basic']; pe = Get-Round $t['price_earnings_ttm'] 1
        shariah = [pscustomobject]@{ verdict = $s.verdict; compliantSources = $s.compliantSources; totalSources = $s.totalSources; haramPct = $s.haramPct; loansPct = $s.loansPct; sources = $s.sources }
        fair = $(if ($fv) { [pscustomobject]@{
                value = $fv.fairValue; upside = Get-Round $upside 1; cls = $fv.fairClass; confidence = $fv.confidence
                analyst = $fv.analystFV; analystUpside = Get-Round $analystUp 1
                dcf = $fv.dcf; graham = $fv.graham; ddm = $fv.ddm; epv = $fv.epv; comps = $fv.comps
                health = $fv.healthText; piotroski = $fv.piotroski; altmanZ = $fv.altmanZ; dy = $fv.dy; pb = $fv.pb
                debtCap = $fv.debtCap; nextEarnings = $fv.nextEarnings; url = $fv.url; date = $fv.fetched } } else { $null })
        tech = $tech
        momentum = Get-Momentum $t
        scores = [pscustomobject]@{ tech = $tech.score; value = Get-Round $valueScore 1; quality = Get-Round $qualityScore 1; liquidity = Get-Round $liqScore 1 }
        score = Get-Round $final 1
    }
}

$results = @($results | Sort-Object -Property score -Descending)
$newRanks = @{}
for ($i = 0; $i -lt $results.Count; $i++) {
    $r = $results[$i]; $rank = $i + 1
    $r | Add-Member rank $rank
    $r | Add-Member prevRank $(if ($prevRank.ContainsKey($r.symbol)) { $prevRank[$r.symbol] } else { $null })
    $newRanks[$r.symbol] = $rank
}
foreach ($g in ($results | Group-Object sector)) {
    $j = 1; foreach ($r in ($g.Group | Sort-Object score -Descending)) { $r | Add-Member sectorRank $j; $j++ }
}
Save-JsonFile $newRanks $prevRanksFile

# ------------------------------------------------------------------
# 5a) التوقع مقابل الفعلي (v3.1): ترتيب البرنامج في آخر يوم قبل الجلسة
#     مقابل الترتيب الفعلي لمكاسب الأسهم في الجلسة دي (من الأعلى للأقل)
#     بيتحسب لكل جلسة خلصت، ويتحفظ تاريخه في cache\rank_eval.json
# ------------------------------------------------------------------
Set-Status 78 'تحديث المقارنات...'
$snapFile = Join-Path $Cache 'rank_snapshots.json'; $evalFile = Join-Path $Cache 'rank_eval.json'
$snaps = @{}
$sf = Read-JsonFile $snapFile
if ($sf) {
    foreach ($p in $sf.PSObject.Properties) { $m = @{}; foreach ($q in $p.Value.PSObject.Properties) { $m[$q.Name] = [int]$q.Value }; $snaps[$p.Name] = $m }
} else {
    # أول مرة: نبني تاريخ الترتيب من النسخ المحفوظة في "old copies" (آخر نسخة في كل يوم)
    $oc = Join-Path $Root 'old copies'
    if (Test-Path $oc) {
        foreach ($dir in (Get-ChildItem $oc -Directory | Sort-Object Name)) {
            $f = Join-Path $dir.FullName 'data.js'; if (-not (Test-Path $f)) { continue }
            try {
                $od = ([IO.File]::ReadAllText($f, $Utf8).Trim() -replace '^window\.EGX_DATA\s*=\s*', '' -replace ';\s*$', '') | ConvertFrom-Json
                $dt = ([datetime]::Parse($od.generatedAt, $Inv)).ToString('yyyy-MM-dd')
                $m = @{}; foreach ($s in @($od.stocks)) { if ($s.rank) { $m[[string]$s.symbol] = [int]$s.rank } }
                if ($m.Count) { $snaps[$dt] = $m }
            } catch { }
        }
        Write-Log ("التوقع مقابل الفعلي: اتبنى تاريخ الترتيب من النسخ القديمة ({0} يوم)" -f $snaps.Count)
    }
}
$snaps[$today] = $newRanks   # آخر تحديث كامل في اليوم هو ترتيب اليوم ده

# مقارنة ترتيب (من آخر يوم قبل الجلسة) بالترتيب الفعلي لمكاسب الجلسة — بتتحسب بنفس الطريقة للترتيب العام وللزخم
#   $snaps: التاريخ -> (السهم -> قيمة)، $higherFirst: true = القيمة الأعلى تيجي الأول (الزخم)، false = الأقل الأول (المركز)
function Update-Eval($snaps, [string]$evalFile, [bool]$higherFirst, [string]$tag) {
    $evals = @{}
    $ef = Read-JsonFile $evalFile
    if ($ef) { foreach ($p in $ef.PSObject.Properties) { $evals[$p.Name] = $p.Value } }
    foreach ($d in $tradingDays) {
        if ($liveToday -and $d -eq $todayStr) { continue }            # الجلسة لسه ما خلصتش
        $base = $snaps.Keys | Where-Object { $_ -lt $d } | Sort-Object -Descending | Select-Object -First 1
        $i = [array]::IndexOf($calendar, $d)
        if (-not $base -or $i -lt 0 -or $i + 1 -ge $calendar.Count) { continue }
        $pd = $calendar[$i + 1]; $pr = $snaps[$base]
        $rows = @()
        foreach ($sym in $pr.Keys) {
            $m = $closeMap[$sym]
            if ($m -and $m.ContainsKey($d) -and $m.ContainsKey($pd) -and $m[$pd] -gt 0) {
                $rows += [pscustomobject]@{ s = $sym; v = $pr[$sym]; g = [Math]::Round(($m[$d] / $m[$pd] - 1) * 100, 2) }
            }
        }
        if ($rows.Count -lt 10) { continue }
        # ترتيبي = ترتيب السهم بين الأسهم اللي اتقارنت، والفعلي = ترتيبه بالمكسب من الأعلى للأقل
        $byPred = @($rows | Sort-Object @{ e = 'v'; Descending = $higherFirst }, @{ e = 's'; Descending = $false })
        for ($k = 0; $k -lt $byPred.Count; $k++) { $byPred[$k] | Add-Member p ($k + 1) }
        $byAct = @($rows | Sort-Object @{ e = 'g'; Descending = $true }, @{ e = 's'; Descending = $false })
        for ($k = 0; $k -lt $byAct.Count; $k++) { $byAct[$k] | Add-Member a ($k + 1) }
        # التعادل (زي مكسب 0% أو نفس درجة الزخم) بياخد متوسط الترتيب في حساب سبيرمان، عشان ما يتحسبش لصالحي
        $avgA = @{}; foreach ($g in ($byAct | Group-Object g)) { $avgA[$g.Name] = ($g.Group | Measure-Object a -Average).Average }
        $avgP = @{}; foreach ($g in ($byPred | Group-Object v)) { $avgP[$g.Name] = ($g.Group | Measure-Object p -Average).Average }
        $n = $rows.Count; $sd = 0.0; foreach ($x in $rows) { $sd += [Math]::Pow($avgP[[string]$x.v] - $avgA[[string]$x.g], 2) }
        $avgG = { param($xs) if (@($xs).Count) { [Math]::Round((@($xs) | Measure-Object g -Average).Average, 2) } else { $null } }
        $top10 = @($byPred | Select-Object -First 10); $bot10 = @($byPred | Select-Object -Last 10)
        $evals[$d] = [pscustomobject]@{
            date = $d; basedOn = $base; prevDay = $pd; n = $n
            rho = [Math]::Round(1 - 6 * $sd / ($n * ($n * $n - 1)), 3)                # معامل سبيرمان: 1 = نفس الترتيب بالظبط
            marketAvg = & $avgG $rows; top10Avg = & $avgG $top10; bottom10Avg = & $avgG $bot10
            top10InTop10 = @($top10 | Where-Object { $_.a -le 10 }).Count
            top10InTopHalf = @($top10 | Where-Object { $_.a -le [Math]::Ceiling($n / 2) }).Count
            rows = @($byAct | ForEach-Object { [pscustomobject]@{ s = $_.s; g = $_.g; a = $_.a; p = $_.p; v = $_.v } })
        }
        Write-Log ("{0} {1} (من يوم {2}): سبيرمان {3}، أفضل 10 {4}% مقابل السوق {5}%" -f $tag, $d, $base, $evals[$d].rho, $evals[$d].top10Avg, $evals[$d].marketAvg)
    }
    $evOut = [ordered]@{}; foreach ($d in @($evals.Keys | Sort-Object -Descending | Select-Object -First 60 | Sort-Object)) { $evOut[$d] = $evals[$d] }
    Save-JsonFile ([pscustomobject]$evOut) $evalFile
    return @($evals.Keys | Sort-Object -Descending | Select-Object -First 30 | ForEach-Object { $evals[$_] })
}
function Save-Snaps($snaps, [string]$file) {
    $o = [ordered]@{}; foreach ($d in @($snaps.Keys | Sort-Object -Descending | Select-Object -First 60 | Sort-Object)) { $o[$d] = $snaps[$d] }
    Save-JsonFile ([pscustomobject]$o) $file
}

$rankEval = @(Update-Eval $snaps $evalFile $false 'مقارنة الترتيب')   # @() عشان يوم واحد يفضل قائمة
Save-Snaps $snaps $snapFile

# ------------------------------------------------------------------
# 5b) مقارنة الزخم (v3.1): درجة زخم الشراء في آخر يوم قبل الجلسة مقابل الترتيب الفعلي لمكاسب الجلسة
# ------------------------------------------------------------------
$momSnapFile = Join-Path $Cache 'momentum_snapshots.json'; $momEvalFile = Join-Path $Cache 'momentum_eval.json'
$momSnaps = @{}
$msf = Read-JsonFile $momSnapFile
if ($msf) {
    foreach ($p in $msf.PSObject.Properties) { $m = @{}; foreach ($q in $p.Value.PSObject.Properties) { $m[$q.Name] = [double]$q.Value }; $momSnaps[$p.Name] = $m }
} else {
    # أول مرة: الزخم من النسخ المحفوظة اللي فيها زخم (من v3.1 وبعدها)
    $oc = Join-Path $Root 'old copies'
    if (Test-Path $oc) {
        foreach ($dir in (Get-ChildItem $oc -Directory | Sort-Object Name)) {
            $f = Join-Path $dir.FullName 'data.js'; if (-not (Test-Path $f)) { continue }
            try {
                $raw = [IO.File]::ReadAllText($f, $Utf8); if (-not $raw.Contains('"momentum"')) { continue }
                $od = ($raw.Trim() -replace '^window\.EGX_DATA\s*=\s*', '' -replace ';\s*$', '') | ConvertFrom-Json
                $dt = ([datetime]::Parse($od.generatedAt, $Inv)).ToString('yyyy-MM-dd')
                $m = @{}; foreach ($s in @($od.stocks)) { if ($s.momentum) { $m[[string]$s.symbol] = [double]$s.momentum.score } }
                if ($m.Count) { $momSnaps[$dt] = $m }
            } catch { }
        }
    }
}
$mToday = @{}; foreach ($r in $results) { if ($r.momentum) { $mToday[$r.symbol] = [double]$r.momentum.score } }
if ($mToday.Count) { $momSnaps[$today] = $mToday }
$momEval = @(Update-Eval $momSnaps $momEvalFile $true 'مقارنة الزخم')
Save-Snaps $momSnaps $momSnapFile

# ------------------------------------------------------------------
# 5b) الأخبار: TradingView (إنجليزي) + Google News (عربي) لكل سهم متوافق
#     تحليل آلي بالكلمات المفتاحية + رد فعل السعر، والاحتفاظ بالأخبار 14 يوم
# ------------------------------------------------------------------
$NewsKeepDays = 14; $NewsRefreshHours = 3
$newsFile = Join-Path $Cache 'news.json'
$newsStore = @{}; $lastNewsFetch = $null
$nf = Read-JsonFile $newsFile
if ($nf) { $lastNewsFetch = $nf.fetchedAt; foreach ($it in @($nf.items)) { if ($it) { $newsStore[$it.id] = $it } } }

function Get-NormAr([string]$s) {
    if (-not $s) { return '' }
    $s = $s.ToLower() -replace '[أإآ]', 'ا' -replace 'ى', 'ي' -replace 'ة', 'ه' -replace 'ؤ', 'و' -replace 'ئ', 'ي' -replace '[ً-ْـ]', ''
    return $s
}
$md5 = [Security.Cryptography.MD5]::Create()
function Get-Hash([string]$s) { ([BitConverter]::ToString($md5.ComputeHash([Text.Encoding]::UTF8.GetBytes($s))) -replace '-', '').Substring(0, 16) }

$GenericWords = @('شركه','للصناعات','الصناعات','للاستثمار','الاستثمار','والاستثمار','للاستثمارات','القابضه','للتنميه','والتنميه','التنميه',
    'المصريه','مصر','العربيه','الدوليه','للتجاره','والتجاره','التجاره','العقاري','العقاريه','الوطنيه','المتحده','العامه','للخدمات','الخدمات',
    'والصناعه','للتصنيع','والتصنيع','الماليه','للاوراق','الاوراق','ش.م.م','و','في','من','الجديده','للمقاولات','المقاولات','ال')
function Get-CoreWords([string]$name) {
    @((Get-NormAr $name) -split '[\s\-\(\)"«»،,\.]+' | Where-Object { $_.Length -ge 3 -and $GenericWords -notcontains $_ })
}
# الخبر يخص الشركة لو: فيه رمز السهم، أو فيه كل الكلمات المميزة لاسمها (أول كلمتين)
# ولو الاسم كلمة واحدة بس لازم كمان يكون في العنوان كلمة سوق (سهم/البورصة/أرباح...)
$MarketTerms = 'سهم|اسهم|البورصه|بورصه|ارباح|خسائر|توزيع|ايرادات|مبيعات|الشركه|شركه|راس المال|مساهمين|الجمعيه|صفقه|استحواذ|نتائج'
function Test-NewsRelevant([string]$nt, [string]$sym, $core) {
    if ($nt -match ('(?<![a-z])' + $sym.ToLower() + '(?![a-z])')) { return $true }
    $need = @($core | Select-Object -First 2)
    if (-not $need.Count) { return $false }
    foreach ($w in $need) { if (-not $nt.Contains($w)) { return $false } }
    if ($need.Count -eq 1) { return ($nt -match $MarketTerms) }
    return $true
}

# قاموس التحليل: [نمط, وزن, تصنيف]. العبارات المركبة أولاً (تُستهلك قبل الكلمات المفردة)
$LexPhrases = @(
    @('(تراجع|انخفاض|تقلص|تقليص|انكماش)\s+(صافي\s+)?(ال)?خسائر', 2.5, 'results'), @('تحول\S*\s+(\S+\s+){0,5}(لل|الي\s+ال)ربح', 3, 'results'),
    @('(تراجع|انخفاض|هبوط|تقلص)\s+(صافي\s+)?(ال)?(ارباح|ربح|ايرادات|مبيعات)', -2.5, 'results'),
    @('(ارتفاع|نمو|زياده|قفزه)\s+(صافي\s+)?(ال)?(ارباح|ربح|ايرادات|مبيعات)', 2.5, 'results'),
    @('(ارباح|ربح|ايرادات|مبيعات)\S*\s+(\S+\s+){0,3}(ترتفع|يرتفع|تقفز|يقفز|تنمو|ينمو|تتضاعف|يتضاعف)', 2.5, 'results'),
    @('(ارباح|ربح|ايرادات|مبيعات)\S*\s+(\S+\s+){0,3}(تتراجع|يتراجع|تنخفض|ينخفض|تهبط|يهبط)', -2.5, 'results'),
    @('(صافي\s+)?خساره|خسائر', -2.5, 'results'), @('صافي\s+(ال)?ربح', 1, 'results'),
    @('(شراء|تشتري|يشتري|تشترى|اعدام)\s+(\S+\s+){0,3}(اسهم|سهم)\s+(ال)?خزينه', 1.5, 'capital'), @('اسهم\s+مجانيه', 2, 'dividend'), @('زياده\s+راس\s+(ال)?مال', 0.5, 'capital'),
    @('(رفع|ترقيه)\s+(ال)?تصنيف', 2, 'rating'), @('(خفض|تخفيض)\s+(ال)?تصنيف', -2, 'rating'), @('سعر\s+(ال)?مستهدف', 0.5, 'rating'),
    @('وقف\s+(ال)?تداول|ايقاف\s+(ال)?تداول', -2, 'risk'), @('مصنع\s+جديد|خط\s+انتاج\s+جديد', 1.5, 'deal'), @('مذكره\s+تفاهم', 1, 'deal')
)
$LexWords = @(
    @('قياسي',2,'results'), @('قفز',2,''), @('قفزه',2,''), @('يقفز',2,''), @('تقفز',2,''), @('تضاعف',2,''), @('ارتفاع',1,''), @('يرتفع',1,''), @('ترتفع',1,''),
    @('نمو',1,''), @('زياده',1,''), @('صعود',1,''), @('مرتفعه',1,''), @('ارباح',1,'results'), @('ايرادات',0.5,'results'),
    @('توزيع',2,'dividend'), @('توزيعات',2,'dividend'), @('كوبون',2,'dividend'),
    @('صفقه',1,'deal'), @('استحواذ',1.5,'deal'), @('تعاقد',1.5,'deal'), @('عقد',1,'deal'), @('اتفاقيه',1,'deal'), @('شراكه',1,'deal'),
    @('توسع',1,'deal'), @('توسعات',1,'deal'), @('اطلاق',1,'deal'), @('تصدير',1,'deal'), @('ترقيه',1,'rating'),
    @('تراجع',-1,''), @('يتراجع',-1,''), @('تتراجع',-1,''), @('انخفاض',-1,''), @('هبوط',-1,''), @('ضغوط',-1,''),
    @('غرامه',-2,'risk'), @('عقوبه',-2,'risk'), @('شطب',-3,'risk'), @('دعوي',-1.5,'risk'), @('نزاع',-1.5,'risk'), @('تعثر',-2,'risk'),
    @('حريق',-2,'risk'), @('استقاله',-1,'risk'), @('تحقيقات',-1.5,'risk'),
    # English
    @('record',2,'results'), @('surge[sd]?',2,''), @('jump(s|ed)?',2,''), @('soar(s|ed)?',2,''), @('ris(e|es|ing)',1,''), @('rose',1,''), @('growth',1,''),
    @('strong',1,''), @('beat(s)?',1.5,'results'), @('profit(s)?',1,'results'), @('dividend(s)?',2,'dividend'), @('acqui(re|res|red|sition)',1.5,'deal'),
    @('contract(s)?',1.5,'deal'), @('agreement',1,'deal'), @('deal',1,'deal'), @('expansion',1,'deal'), @('buyback',1.5,'capital'), @('upgrade[sd]?',2,'rating'),
    @('loss(es)?',-2,'results'), @('declin(e|es|ed)',-1,''), @('fall(s)?',-1,''), @('fell',-1,''), @('drop(s|ped)?',-1,''), @('plung(e|es|ed)',-2,''),
    @('slump(s|ed)?',-2,''), @('miss(es|ed)?',-1.5,'results'), @('pressure(s)?',-0.5,''), @('fined',-1.5,'risk'), @('lawsuit',-1.5,'risk'),
    @('downgrade[sd]?',-2,'rating'), @('suspend(s|ed)?',-2,'risk'), @('weak(er)?',-1,'')
)

function Get-NewsAnalysis([string]$title) {
    $txt = ' ' + (Get-NormAr $title) + ' '
    $score = 0.0; $pos = @(); $neg = @(); $cats = @{}
    $apply = {
        param($pattern, $w, $cat)
        $m = [regex]::Match($txt, $pattern)
        if ($m.Success) {
            $script:score_ += $w
            if ($w -gt 0) { $script:pos_ += $m.Value.Trim() } else { $script:neg_ += $m.Value.Trim() }
            if ($cat) { $script:cats_[$cat] = [double]$script:cats_[$cat] + [Math]::Abs($w) }
            $script:txt_ = [regex]::Replace($txt, $pattern, ' ')
        }
    }
    $script:score_ = 0.0; $script:pos_ = @(); $script:neg_ = @(); $script:cats_ = @{}
    foreach ($p in $LexPhrases) { $script:txt_ = $txt; & $apply $p[0] $p[1] $p[2]; $txt = $script:txt_ }
    foreach ($p in $LexWords) {
        $pat = '(?<!\p{L})(?:و|ف|ب|ل|ال|وال|بال|لل)?' + $p[0] + '\p{L}{0,3}(?!\p{L})'
        $script:txt_ = $txt; & $apply $pat $p[1] $p[2]; $txt = $script:txt_
    }
    $score = $script:score_
    if ($score -ne 0 -and $title -match '\d+(\.\d+)?\s*%') { $score *= 1.25 }
    $label = if ($score -ge 3) { 'strong_pos' } elseif ($score -ge 1) { 'pos' } elseif ($score -le -3) { 'strong_neg' } elseif ($score -le -1) { 'neg' } else { 'neutral' }
    $cat = if ($script:cats_.Count) { ($script:cats_.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key } else { 'general' }
    [pscustomobject]@{ score = [Math]::Round($score, 2); label = $label; category = $cat; pos = @($script:pos_ | Select-Object -Unique); neg = @($script:neg_ | Select-Object -Unique) }
}

# رد فعل السعر: تغير السهم في أول جلسة بعد نشر الخبر (الخبر بعد 2 ظهراً بتوقيت القاهرة يُحسب على الجلسة التالية)
function Get-NewsReaction([string]$sym, [datetime]$pubUtc) {
    $pts = $hist[$sym]; if (-not $pts) { return $null }
    $cairo = $pubUtc.AddHours(2)
    $target = if ($cairo.Hour -ge 14) { $cairo.Date.AddDays(1) } else { $cairo.Date }
    $ts = $target.ToString('yyyy-MM-dd')
    for ($i = 1; $i -lt $pts.Count; $i++) {
        if ($pts[$i].d -ge $ts -and $pts[$i - 1].c -gt 0) {
            return [pscustomobject]@{ date = $pts[$i].d; pct = [Math]::Round(($pts[$i].c / $pts[$i - 1].c - 1) * 100, 2) }
        }
    }
    return $null
}

$needNews = $ForceRefresh -or -not $lastNewsFetch -or ((Get-Date) - [datetime]::Parse($lastNewsFetch, $Inv)).TotalHours -ge $NewsRefreshHours
$cutUtc = (Get-Date).ToUniversalTime().AddDays(-$NewsKeepDays)
if ($needNews) {
    Write-Log 'تحميل الأخبار (TradingView + Google News)...'
    $added = 0; $nfail = 0; $ni = 0
    foreach ($r in $results) {
        $sym = $r.symbol
        $ni++; if ($ni % 5 -eq 1) { Set-Status (80 + 16 * $ni / [Math]::Max(1, $results.Count)) ("الأخبار: {0} / {1}" -f $ni, $results.Count) }
        try {
            $tvn = Invoke-RestMethod "https://news-headlines.tradingview.com/v2/view/headlines/symbol?client=web&lang=en&symbol=EGX:$sym" -UserAgent $UA -TimeoutSec 20
            foreach ($it in @($tvn.items)) {
                if (-not $it) { continue }
                $pub = [DateTimeOffset]::FromUnixTimeSeconds([long]$it.published).UtcDateTime
                $id = 'tv:' + $it.id
                if ($pub -lt $cutUtc -or $newsStore.ContainsKey($id)) { continue }
                $newsStore[$id] = [pscustomobject]@{ id = $id; symbol = $sym; title = [string]$it.title; source = [string]$it.source
                    url = 'https://www.tradingview.com' + $it.storyPath; published = $pub.ToString('yyyy-MM-ddTHH:mm:ssZ'); lang = 'en'; firstSeen = $today }
                $added++
            }
        } catch { $nfail++ }
        try {
            $core = Get-CoreWords $r.name
            $q = $(if ($core.Count) { ($core | Select-Object -First 3) -join ' ' } else { $r.name }) + ' البورصة'
            $xml = [xml](Get-Text ('https://news.google.com/rss/search?q=' + [uri]::EscapeDataString($q) + '&hl=ar&gl=EG&ceid=EG:ar'))
            foreach ($it in @($xml.rss.channel.item)) {
                if (-not $it) { continue }
                $pub = [DateTimeOffset]::Parse([string]$it.pubDate, $Inv).UtcDateTime
                if ($pub -lt $cutUtc) { continue }
                $src = if ($it.source -is [Xml.XmlElement]) { $it.source.'#text' } else { [string]$it.source }
                $title = ([string]$it.title) -replace ('\s+-\s+' + [regex]::Escape([string]$src) + '\s*$'), ''
                $nt = Get-NormAr $title
                if (-not (Test-NewsRelevant $nt $sym $core)) { continue }
                $id = 'gn:' + $sym + ':' + (Get-Hash $nt)
                if ($newsStore.ContainsKey($id)) { continue }
                $newsStore[$id] = [pscustomobject]@{ id = $id; symbol = $sym; title = $title; source = $src; url = [string]$it.link
                    published = $pub.ToString('yyyy-MM-ddTHH:mm:ssZ'); lang = 'ar'; firstSeen = $today }
                $added++
            }
        } catch { $nfail++ }
        Start-Sleep -Milliseconds 150
    }
    $lastNewsFetch = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
    Write-Log ("الأخبار: {0} خبر جديد" -f $added)
    if ($nfail -gt 20) { [void]$Errors.Add("الأخبار: تعذر تحميل $nfail طلب") }
}
# حذف الأخبار الأقدم من 14 يوم فقط
foreach ($k in @($newsStore.Keys)) {
    if ([datetime]::Parse($newsStore[$k].published, $Inv).ToUniversalTime() -lt $cutUtc) { $newsStore.Remove($k) }
}
Save-JsonFile ([pscustomobject]@{ fetchedAt = $lastNewsFetch; items = @($newsStore.Values) }) $newsFile

$keptSet = @{}; foreach ($r in $results) { $keptSet[$r.symbol] = $r }
$seenTitles = @{}
$newsOut = foreach ($it in ($newsStore.Values | Sort-Object published -Descending)) {
    if (-not $keptSet.ContainsKey($it.symbol)) { continue }
    $tkey = $it.symbol + '|' + ((Get-NormAr $it.title) -replace '[^\p{L}\p{N}]', '')
    if ($seenTitles.ContainsKey($tkey)) { continue }; $seenTitles[$tkey] = 1
    if ($it.lang -eq 'ar' -and -not (Test-NewsRelevant (Get-NormAr $it.title) $it.symbol (Get-CoreWords $keptSet[$it.symbol].name))) { continue }
    $a = Get-NewsAnalysis $it.title
    $pubUtc = [datetime]::Parse($it.published, $Inv).ToUniversalTime()
    [pscustomobject]@{
        symbol = $it.symbol; name = $keptSet[$it.symbol].name; sector = $keptSet[$it.symbol].sector
        title = $it.title; source = $it.source; url = $it.url; published = $it.published; lang = $it.lang
        score = $a.score; label = $a.label; category = $a.category; pos = $a.pos; neg = $a.neg
        reaction = Get-NewsReaction $it.symbol $pubUtc
    }
}
$newsOut = @($newsOut)
# ملخص الأخبار لكل سهم (آخر 7 أيام)
$cut7 = (Get-Date).ToUniversalTime().AddDays(-7)
foreach ($r in $results) {
    $mine = @($newsOut | Where-Object { $_.symbol -eq $r.symbol })
    $recent = @($mine | Where-Object { [datetime]::Parse($_.published, $Inv).ToUniversalTime() -ge $cut7 })
    $sum = 0.0; foreach ($n in $recent) { $sum += $n.score }
    $r | Add-Member news ([pscustomobject]@{ count = $mine.Count; count7 = $recent.Count; score7 = [Math]::Round($sum, 1) })
}
Write-Log ("الأخبار المعروضة: {0} خبر (آخر {1} يوم)" -f $newsOut.Count, $NewsKeepDays)

# ------------------------------------------------------------------
# 5c) نسب التطهير لكل الأسهم (المصدر: بيانات "اسهم حلال" اللي فيها كمان أرقام كاشف)
#     كاشف بيقفل نسبة التطهير على موقعه للأعضاء بس، فبنستخدم أرقامه المنشورة عند اسهم حلال
# ------------------------------------------------------------------
$purif = [ordered]@{}
foreach ($e in ($shariah | Sort-Object symbol)) {
    $sym = $e.symbol.ToUpper(); $k = $e.kashif
    $price = if ($tvMap.ContainsKey($sym)) { Get-Round $tvMap[$sym]['close'] 3 } else { $null }
    $purif[$sym] = [pscustomobject]@{
        name = $e.name_ar; sector = $e.category_ar; price = $price
        verdict = $shMap[$sym].verdict
        # اسهم حلال
        h = $e.haram_earnings_percentage; aaoifi = $e.aaoifi_haram_earning_per_share; loans = $e.loans_percentage
        # كاشف
        k = $(if ($k) { $k.haram_percentage } else { $null }); kGrade = $(if ($k) { $k.purity } else { $null })
        kStatus = $(if ($k) { $k.status } else { $null }); kNote = $(if ($k) { $k.note } else { $null })
        kDate = $(if ($k) { $k.statement_date } else { $null })
    }
}

# ------------------------------------------------------------------
# 5d) محفظة ثندر: أحدث ملف Excel في فولدر thndr  ->  my-holdings.js
#     (شيت Closed Lots = عمليات اتباعت، شيت Holdings = أسهم لسه معاك)
# ------------------------------------------------------------------
$thndrDir = Join-Path $Root 'thndr'
$OnCloud = [bool]$env:GITHUB_ACTIONS   # نسخة الويب على GitHub: الصفحة عامة، فمحفظة ثندر ما بتتنشرش
if (-not $OnCloud) { New-Item -ItemType Directory -Force $thndrDir | Out-Null }
$xl = if ($OnCloud) { $null } else { Get-ChildItem $thndrDir -Filter *.xlsx -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike '~$*' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1 }
if ($xl) {
    $zip = $null
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [IO.Compression.ZipFile]::OpenRead($xl.FullName)
        $getXml = { param($name) $e = $zip.GetEntry($name); if (-not $e) { return $null }; $sr = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8); $t = $sr.ReadToEnd(); $sr.Close(); [xml]$t }
        $wbx = & $getXml 'xl/workbook.xml'; $rels = & $getXml 'xl/_rels/workbook.xml.rels'
        $shared = @(); $ssx = & $getXml 'xl/sharedStrings.xml'; if ($ssx) { $shared = @($ssx.sst.si | ForEach-Object { $_.InnerText }) }
        $readSheet = {
            param($sheetName)
            $sh = @($wbx.workbook.sheets.sheet) | Where-Object { $_.name -eq $sheetName } | Select-Object -First 1
            if (-not $sh) { return }
            $rid = $sh.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            $target = (@($rels.Relationships.Relationship) | Where-Object { $_.Id -eq $rid }).Target
            $sx = & $getXml $(if ($target.StartsWith('/')) { $target.TrimStart('/') } else { 'xl/' + $target })
            $hdr = $null
            foreach ($r in @($sx.worksheet.sheetData.row)) {
                $vals = @{}
                foreach ($c in @($r.c)) {
                    $col = $c.r -replace '\d', ''
                    $vals[$col] = if ($c.t -eq 'inlineStr') { $c.is.InnerText } elseif ($c.t -eq 's') { $shared[[int]$c.v] } else { [string]$c.v }
                }
                if (-not $hdr) { $hdr = $vals; continue }
                $o = @{}; foreach ($k in $hdr.Keys) { $o[[string]$hdr[$k]] = $vals[$k] }
                $o
            }
        }
        $toDate = { param($v) if (-not $v) { return '' }; $d = 0.0
            if ([double]::TryParse([string]$v, [Globalization.NumberStyles]::Float, $Inv, [ref]$d)) { [DateTime]::FromOADate($d).ToString('yyyy-MM-dd') }
            else { ([datetime]::Parse([string]$v, $Inv)).ToString('yyyy-MM-dd') } }
        $num = { param($v) $d = 0.0; if ([double]::TryParse([string]$v, [Globalization.NumberStyles]::Float, $Inv, [ref]$d)) { $d } else { $null } }
        $lots = New-Object System.Collections.ArrayList; $i = 0
        foreach ($r in @(& $readSheet 'Closed Lots')) {
            if (-not $r['Ticker']) { continue }; $i++
            [void]$lots.Add([pscustomobject]@{ id = "t$i"; src = 'thndr'; sym = ([string]$r['Ticker']).ToUpper(); qty = & $num $r['Units']
                buyDate = & $toDate $r['Buy Date']; buyPrice = & $num $r['Buy Price']; sellDate = & $toDate $r['Sell Date']; sellPrice = & $num $r['Sell Price']; divs = 0 })
        }
        $nClosed = $lots.Count
        foreach ($r in @(& $readSheet 'Holdings')) {
            if (-not $r['Ticker']) { continue }; $i++
            [void]$lots.Add([pscustomobject]@{ id = "t$i"; src = 'thndr'; sym = ([string]$r['Ticker']).ToUpper(); qty = & $num $r['Units Held']
                buyDate = & $toDate $r['Buy Date']; buyPrice = & $num $r['Buy Price (EGP)']; sellDate = ''; sellPrice = ''; divs = 0 })
        }
        if ($lots.Count -eq 0) { throw 'الملف ما فيهوش شيت Closed Lots أو Holdings' }
        $info = [pscustomobject]@{ file = $xl.Name; modified = $xl.LastWriteTime.ToString('yyyy-MM-dd HH:mm'); closed = $nClosed; open = $lots.Count - $nClosed }
        $js = 'window.MY_HOLDINGS = ' + (@($lots) | ConvertTo-Json -Depth 4 -Compress) + ";`nwindow.MY_HOLDINGS_INFO = " + ($info | ConvertTo-Json -Compress) + ';'
        [IO.File]::WriteAllText((Join-Path $Root 'my-holdings.js'), $js, $Utf8)
        Write-Log ("محفظة ثندر: {0} ({1} عملية مقفولة، {2} لسه معاك)" -f $xl.Name, $nClosed, $info.open)
    } catch { [void]$Errors.Add("ملف ثندر $($xl.Name): $($_.Exception.Message)") }
    finally { if ($zip) { $zip.Dispose() } }
}

$out = [pscustomobject]@{
    version = $Version
    generatedAt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
    days = $tradingDays
    weights = [pscustomobject]@{ tech = $W_Tech; value = $W_Value; quality = $W_Quality; liquidity = $W_Liquidity }
    includeDoubtful = $IncludeDoubtful
    counts = [pscustomobject]@{ market = $tvMap.Count; screened = $shMap.Count; kept = $results.Count; excluded = $excluded.Count; withFair = @($results | Where-Object { $_.fair -and $null -ne $_.fair.upside }).Count }
    stocks = $results
    excluded = $excluded
    purif = [pscustomobject]$purif
    news = $newsOut
    rankEval = $rankEval
    momEval = $momEval
    newsFetchedAt = $lastNewsFetch
    errors = @($Errors)
}
Set-Status 97 'حفظ النتايج...'
$json = $out | ConvertTo-Json -Depth 8 -Compress
[IO.File]::WriteAllText($DataFile, "window.EGX_DATA = $json;", $Utf8)
Write-Log ("تم: {0} سهم مرتب. أخطاء: {1}" -f $results.Count, $Errors.Count)
foreach ($e in $Errors) { Write-Log "  ! $e" }

# ------------------------------------------------------------------
# 6) نسخة محلية من نتيجة كل تحديث في "old copies\<التاريخ_الوقت>"
#    (الواجهة + البيانات تفتح لوحدها، + ranking.csv للفتح في Excel)
# ------------------------------------------------------------------
if (-not $OnCloud) { try {   # على GitHub ما بنحفظش نسخ (كانت هتكبّر المستودع 1 ميجا كل تحديث)
    $snap = Join-Path (Join-Path $Root 'old copies') (Get-Date -Format 'yyyy-MM-dd_HH-mm')
    New-Item -ItemType Directory -Force $snap | Out-Null
    Copy-Item (Join-Path $Root 'dashboard.html'), (Join-Path $Root 'data.js') -Destination $snap -Force
    $csv = $results | ForEach-Object {
        [pscustomobject][ordered]@{
            Rank = $_.rank; Symbol = $_.symbol; Name = $_.name; Sector = $_.sector; Price = $_.price; ChangePct = $_.change
            Day1Pct = $_.daily[0]; Day2Pct = $_.daily[1]; Day3Pct = $_.daily[2]; Day4Pct = $_.daily[3]; Day5Pct = $_.daily[4]
            WeekPct = $_.perfW; MonthPct = $_.perf1M; SixMonthPct = $_.perf6M; YearPct = $_.perfY
            FairValue = $_.fair.value; UpsidePct = $_.fair.upside; AnalystTarget = $_.fair.analyst; AnalystUpsidePct = $_.fair.analystUpside; TechScore = $_.scores.tech; ValueScore = $_.scores.value
            QualityScore = $_.scores.quality; LiquidityScore = $_.scores.liquidity; Score = $_.score; SectorRank = $_.sectorRank
        }
    } | ConvertTo-Csv -NoTypeInformation
    [IO.File]::WriteAllLines((Join-Path $snap 'ranking.csv'), $csv, (New-Object System.Text.UTF8Encoding($true)))
    Write-Log "تم حفظ نسخة في: $snap"
} catch { Write-Log "  ! تعذر حفظ النسخة: $($_.Exception.Message)" } }

Set-Status 100 ("تم: {0} سهم مرتب" -f $results.Count) -Done
