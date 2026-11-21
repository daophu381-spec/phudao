# ======================================================================
#  CHẠY TRÊN MÁY ĐANG CẮM MÁY IN (máy chia sẻ, vd 10.5.208.100)
#  Mở khoá những thứ hay chặn máy khác kết nối vào máy in:
#   - đặt mạng đang dùng thành "Private" (mạng Public chặn chia sẻ)
#   - bật nhóm luật firewall "File and Printer Sharing"
#   - bật dịch vụ chia sẻ (Server, Print Spooler)
#   - tắt RPC Privacy cho in ấn (sửa lỗi 0x0000011b sau bản vá Windows)
#   - chia sẻ máy in M402 nếu chưa chia sẻ
#  Nhật ký: may_chu_mo_chia_se.log
# ======================================================================
$Goc = Split-Path -Parent $MyInvocation.MyCommand.Path
try { Start-Transcript -Path (Join-Path $Goc "may_chu_mo_chia_se.log") -Force | Out-Null } catch {}

function Buoc($t) { Write-Host ""; Write-Host ">>> $t" -ForegroundColor Cyan }
function OK($t)   { Write-Host "    [OK]  $t" -ForegroundColor Green }
function Canh($t) { Write-Host "    [!]   $t" -ForegroundColor Yellow }

Buoc "1/5  Đặt mạng đang dùng thành Private"
try {
    Get-NetConnectionProfile -ErrorAction Stop | ForEach-Object {
        if ($_.NetworkCategory -ne 'Private' -and $_.NetworkCategory -ne 'DomainAuthenticated') {
            Set-NetConnectionProfile -InterfaceIndex $_.InterfaceIndex -NetworkCategory Private -ErrorAction Stop
            OK "$($_.Name): đổi thành Private"
        } else { OK "$($_.Name): đã là $($_.NetworkCategory)" }
    }
} catch { Canh "Không tự đổi được ($($_.Exception.Message)). Tự đổi: Settings → Network → Wi-Fi/Ethernet → chọn mạng → Private." }

Buoc "2/5  Bật chia sẻ file và máy in trên firewall"
try {
    # Mã nhóm này đúng trên mọi ngôn ngữ Windows
    $luat = Get-NetFirewallRule -Group '@FirewallAPI.dll,-32752' -ErrorAction Stop
    $luat | Enable-NetFirewallRule
    OK "Đã bật $(@($luat).Count) luật File and Printer Sharing"
} catch {
    & netsh.exe advfirewall firewall set rule group="File and Printer Sharing" new enable=Yes | Out-Null
    Canh "Dùng netsh (nếu Windows không phải tiếng Anh thì bật tay: Control Panel → Windows Firewall → Allow an app → File and Printer Sharing)."
}

Buoc "3/5  Bật dịch vụ chia sẻ"
foreach ($s in "LanmanServer", "Spooler", "FDResPub", "fdPHost") {
    try { Set-Service $s -StartupType Automatic -ErrorAction Stop; Start-Service $s -ErrorAction SilentlyContinue; OK "Dịch vụ $s" }
    catch { Canh "Dịch vụ $s: $($_.Exception.Message)" }
}

Buoc "4/5  Sửa lỗi 0x0000011b cho in ấn (RPC Privacy)"
try {
    New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Print" -Force | Out-Null
    Set-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Print" -Name RpcAuthnLevelPrivacyEnabled -Value 0 -Type DWord
    Restart-Service Spooler -Force
    OK "Đã áp dụng và khởi động lại Print Spooler"
    Canh "Lưu ý: việc này giảm một lớp mã hoá của dịch vụ in. Chỉ nên dùng trong mạng nội bộ cơ quan."
} catch { Canh $_.Exception.Message }

Buoc "5/5  Chia sẻ máy in"
$ten = $null
try {
    $ds = @(Get-Printer -ErrorAction Stop | Where-Object { $_.Name -notmatch 'PDF|XPS|OneNote|Fax|Microsoft Print|Snagit' })
    $daChiaSe = $ds | Where-Object { $_.Shared } | Select-Object -First 1
    if ($daChiaSe) {
        OK "Đã chia sẻ sẵn: $($daChiaSe.Name) → tên chia sẻ '$($daChiaSe.ShareName)'"
        $ten = $daChiaSe.ShareName
    } else {
        $m = $ds | Where-Object { $_.Name -match 'M40[23]' } | Select-Object -First 1
        if (-not $m) { $m = $ds | Select-Object -First 1 }
        if ($m) {
            Set-Printer -Name $m.Name -Shared $true -ShareName "M402"
            OK "Đã chia sẻ máy in '$($m.Name)' với tên 'M402'"
            $ten = "M402"
        } else { Canh "Không thấy máy in thật nào trên máy này (cắm USB + bật máy in, cài driver rồi chạy lại)." }
    }
} catch { Canh $_.Exception.Message }

$ip = @(Get-WmiObject Win32_NetworkAdapterConfiguration -ErrorAction SilentlyContinue |
        Where-Object { $_.IPEnabled } | ForEach-Object { $_.IPAddress } | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' })
Write-Host ""
Write-Host "================  XONG  ================" -ForegroundColor Green
Write-Host "IP máy này   : $($ip -join ', ')"
if ($ten) { foreach ($i in $ip) { Write-Host "Đường dẫn in : \\$i\$ten" } }
Write-Host ""
Write-Host "NẾU MÁY KHÁC VẪN BỊ HỎI TÀI KHOẢN:"
Write-Host " - Máy này phải có tài khoản Windows CÓ MẬT KHẨU. Máy kia nhập đúng tên + mật khẩu đó."
Write-Host " - Đăng nhập bằng Microsoft account thì nhập email + mật khẩu Microsoft đó (không dùng PIN)."
Write-Host " - Nên đặt IP tĩnh cho máy này (hoặc nhờ quản trị mạng cấp IP cố định) để IP không đổi."
try { Stop-Transcript | Out-Null } catch {}
