# ======================================================================
#  CÀI DRIVER HP LaserJet Pro M402 - TỰ ĐỘNG, ĐÚNG BẢN 32/64 BIT
#  Chạy qua CaiDriverM402.bat (bấm đúp). Có thể thêm tham số:
#     -IP 192.168.1.50     cài máy in mạng (LAN) thay vì cắm USB
#  Driver lấy từ thư mục Drivers\x64 (Windows 64 bit) hoặc Drivers\x86
#  (Windows 32 bit) cạnh file này. Nhật ký lưu ở cai_driver.log.
# ======================================================================
param(
    [string]$IP = ""
)

$Goc = Split-Path -Parent $MyInvocation.MyCommand.Path
$Log = Join-Path $Goc "cai_driver.log"
try { Start-Transcript -Path $Log -Force | Out-Null } catch {}

function Buoc($t) { Write-Host ""; Write-Host ">>> $t" -ForegroundColor Cyan }
function OK($t)   { Write-Host "    [OK]  $t" -ForegroundColor Green }
function Canh($t) { Write-Host "    [!]   $t" -ForegroundColor Yellow }
function Loi($t)  { Write-Host "    [LOI] $t" -ForegroundColor Red }
function Dung($t) { Loi $t; try { Stop-Transcript | Out-Null } catch {}; exit 1 }

# ---------------------------------------------------------------- 1
Buoc "1/6  Kiểm tra quyền Admin và phiên bản Windows"
$admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
         ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $admin) { Dung "Cần chạy bằng quyền Administrator. Hãy bấm đúp file CaiDriverM402.bat." }

$os    = Get-WmiObject Win32_OperatingSystem
$build = [int]$os.BuildNumber
$x64   = [Environment]::Is64BitOperatingSystem
$kieu  = if ($x64) { "x64" } else { "x86" }
OK "$($os.Caption.Trim()) - build $build - Windows $(if ($x64) {'64 bit'} else {'32 bit'}) → dùng driver thư mục Drivers\$kieu"
$coCmdlet = [bool](Get-Command Add-PrinterDriver -ErrorAction SilentlyContinue)   # Windows 8 trở lên
if (-not $coCmdlet) { Canh "Windows cũ (7): dùng cách cài tương thích, có thể cần rút/cắm lại cáp USB." }

# ---------------------------------------------------------------- 2
Buoc "2/6  Khởi động lại dịch vụ in (Print Spooler)"
try {
    Set-Service Spooler -StartupType Automatic -ErrorAction SilentlyContinue
    Restart-Service Spooler -Force -ErrorAction Stop
    OK "Print Spooler đang chạy"
} catch { Canh "Không khởi động lại được Spooler: $($_.Exception.Message)" }

# ---------------------------------------------------------------- 3
Buoc "3/6  Chuẩn bị file driver (Drivers\$kieu)"
$thuMuc = Join-Path $Goc "Drivers\$kieu"
if (-not (Test-Path $thuMuc)) { New-Item -ItemType Directory -Path $thuMuc -Force | Out-Null }

# Tự giải nén file .zip nếu bạn bỏ nguyên file zip vào thư mục
foreach ($z in Get-ChildItem $thuMuc -Filter *.zip -File -ErrorAction SilentlyContinue) {
    $dich = Join-Path $thuMuc ([IO.Path]::GetFileNameWithoutExtension($z.Name))
    if (-not (Test-Path $dich)) {
        try {
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [IO.Compression.ZipFile]::ExtractToDirectory($z.FullName, $dich)
            OK "Đã giải nén $($z.Name)"
        } catch { Canh "Không giải nén được $($z.Name): $($_.Exception.Message)" }
    }
}

$infs = @(Get-ChildItem $thuMuc -Filter *.inf -Recurse -File -ErrorAction SilentlyContinue)
if ($infs.Count -eq 0) {
    Canh "Thư mục Drivers\$kieu chưa có file driver (*.inf)."
    Canh "Sẽ thử lấy driver có sẵn trong Windows / Windows Update."
} else {
    OK "Tìm thấy $($infs.Count) file .inf"
}

# ---------------------------------------------------------------- 4
Buoc "4/6  Nạp driver vào Windows"
foreach ($d in ($infs | ForEach-Object { $_.DirectoryName } | Select-Object -Unique)) {
    if ($build -ge 10240) { $kq = & pnputil.exe /add-driver "$d\*.inf" /install 2>&1 }
    else                  { $kq = & pnputil.exe -i -a "$d\*.inf" 2>&1 }
    $kq | Where-Object { $_ -match "Published|Added|đã|thành công|Total|Tổng" } |
        ForEach-Object { Write-Host "    $_" }
}

# Đọc tên mẫu máy in M402/M403 ghi trong file .inf
function Lay-TenMau([string]$inf) {
    $txt = Get-Content $inf -Raw -ErrorAction SilentlyContinue
    if (-not $txt) { return @() }
    $chuoi = @{}
    foreach ($m in [regex]::Matches($txt, '(?m)^\s*(\w+)\s*=\s*"([^"]*)"')) { $chuoi[$m.Groups[1].Value] = $m.Groups[2].Value }
    $ten = @()
    foreach ($m in [regex]::Matches($txt, '(?m)^\s*(?:%(\w+)%|"([^"]+)")\s*=\s*[^;\r\n]+')) {
        $n = if ($m.Groups[1].Success) { $chuoi[$m.Groups[1].Value] } else { $m.Groups[2].Value }
        if ($n -match 'M40[23]') { $ten += $n }
    }
    $ten
}

$ungVien = @()   # (tên mẫu, file inf)
foreach ($f in $infs) { foreach ($t in (Lay-TenMau $f.FullName)) { $ungVien += ,@($t, $f.FullName) } }
# ưu tiên: PCL 6 > PCL, M402 > M403
$ungVien = @($ungVien | Sort-Object { -([int]($_[0] -match 'PCL ?6') * 2 + [int]($_[0] -match 'M402')) })
# nếu không đọc được từ .inf, thử các tên thường gặp (driver có sẵn trong Windows)
foreach ($t in @("HP LaserJet Pro M402-M403 n-dne PCL 6", "HP LaserJet Pro M402-M403 PCL 6", "HP LaserJet Pro M402 PCL 6")) {
    $ungVien += ,@($t, $null)
}

$tenDriver = $null
foreach ($u in $ungVien) {
    $ten = $u[0]; $inf = $u[1]
    try {
        if ($coCmdlet) {
            Add-PrinterDriver -Name $ten -ErrorAction Stop
        } else {
            if (-not $inf) { continue }
            & rundll32.exe printui.dll,PrintUIEntry /ia /m "$ten" /f "$inf" | Out-Null
        }
        if ($coCmdlet -and -not (Get-PrinterDriver -Name $ten -ErrorAction SilentlyContinue)) { continue }
        $tenDriver = $ten; $infDung = $inf
        break
    } catch { }
}
if (-not $tenDriver) {
    Loi "Không cài được driver M402."
    Write-Host "    → Tải driver 'HP LaserJet Pro M402' (hoặc HP Universal Print Driver PCL6) từ support.hp.com,"
    Write-Host "      đúng bản Windows $(if ($x64) {'64 bit'} else {'32 bit'}), giải nén rồi bỏ vào thư mục: $thuMuc"
    Write-Host "      sau đó chạy lại file này."
    Dung "Chưa có driver phù hợp."
}
OK "Đã cài driver: $tenDriver"

# ---------------------------------------------------------------- 5
Buoc "5/6  Tạo máy in"
function Tim-MayIn { if ($coCmdlet) { Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.DriverName -eq $tenDriver } } }
$may = Tim-MayIn
if ($may) {
    OK "Đã có máy in: $(@($may)[0].Name)"
} elseif ($coCmdlet) {
    $port = $null
    if ($IP) {
        $port = "IP_$IP"
        if (-not (Get-PrinterPort -Name $port -ErrorAction SilentlyContinue)) { Add-PrinterPort -Name $port -PrinterHostAddress $IP }
        OK "Cổng mạng $IP"
    } else {
        Start-Sleep -Seconds 3                       # chờ Windows tự nhận máy in USB
        $may = Tim-MayIn
        if ($may) { OK "Windows đã tự tạo máy in: $(@($may)[0].Name)" }
        else {
            $usb = Get-PrinterPort -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "USB*" } | Select-Object -Last 1
            if ($usb) { $port = $usb.Name; OK "Dùng cổng USB: $port" }
        }
    }
    if (-not $may -and $port) {
        try { Add-Printer -Name "HP LaserJet Pro M402" -DriverName $tenDriver -PortName $port -ErrorAction Stop; OK "Đã tạo máy in HP LaserJet Pro M402" }
        catch { Loi "Không tạo được máy in: $($_.Exception.Message)" }
    } elseif (-not $may) {
        Canh "Chưa thấy máy in nào được cắm. Hãy cắm cáp USB, bật máy in rồi chạy lại file này"
        Canh "(hoặc chạy:  CaiDriverM402.bat -IP 192.168.x.x  nếu máy in dùng mạng LAN)."
    }
} else {
    Canh "Windows 7: rút cáp USB máy in ra, cắm lại; Windows sẽ tự gắn driver vừa nạp."
}

# ---------------------------------------------------------------- 6
Buoc "6/6  Đặt làm máy in mặc định"
if ($coCmdlet) {
    $may = Tim-MayIn
    if ($may) {
        $ten = @($may)[0].Name
        (New-Object -ComObject WScript.Network).SetDefaultPrinter($ten)
        try { Set-ItemProperty 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows' -Name LegacyDefaultPrinterMode -Value 1 } catch {}
        OK "Máy in mặc định: $ten"
    } else { Canh "Chưa có máy in để đặt mặc định." }
}

Write-Host ""
Write-Host "================  XONG  ================" -ForegroundColor Green
Write-Host "Nên in thử 1 trang: Settings → Printers → HP LaserJet Pro M402 → Print test page."
try { Stop-Transcript | Out-Null } catch {}
