# ======================================================================
#  KẾT NỐI MÁY IN ĐƯỢC CHIA SẺ TỪ MÁY KHÁC TRONG MẠNG (mặc định 10.5.208.100)
#  Chạy qua KetNoiMayInMang.bat (bấm đúp) trên MÁY CẦN IN.
#  Tự dò lỗi mạng, tự nhập tài khoản, tự sửa lỗi Windows 0x11b / 0x709 / 0xbcb
#  và nếu vẫn không được thì dùng cách dự phòng (driver cài sẵn + cổng \\IP\tên).
#  Nhật ký: ket_noi_may_in.log  (gửi file này nếu cần nhờ xem lỗi)
# ======================================================================
param([string]$IP = "10.5.208.100")

$Goc = Split-Path -Parent $MyInvocation.MyCommand.Path
try { Start-Transcript -Path (Join-Path $Goc "ket_noi_may_in.log") -Force | Out-Null } catch {}

function Buoc($t) { Write-Host ""; Write-Host ">>> $t" -ForegroundColor Cyan }
function OK($t)   { Write-Host "    [OK]  $t" -ForegroundColor Green }
function Canh($t) { Write-Host "    [!]   $t" -ForegroundColor Yellow }
function Loi($t)  { Write-Host "    [LOI] $t" -ForegroundColor Red }
function Thoat($t){ Loi $t; try { Stop-Transcript | Out-Null } catch {}; exit 1 }

function Test-Cong([string]$ip, [int]$port, [int]$ms = 3000) {
    try {
        $c = New-Object Net.Sockets.TcpClient
        $r = $c.BeginConnect($ip, $port, $null, $null)
        $ok = $r.AsyncWaitHandle.WaitOne($ms, $false) -and $c.Connected
        $c.Close()
        return [bool]$ok
    } catch { return $false }
}

$coCmdlet = [bool](Get-Command Add-Printer -ErrorAction SilentlyContinue)

# ---------------------------------------------------------------- 1
Buoc "1/5  Địa chỉ máy đang chia sẻ máy in"
$nhap = Read-Host "Nhập IP máy chia sẻ (Enter = $IP)"
if ($nhap.Trim()) { $IP = $nhap.Trim() }
$ipToi = @(Get-WmiObject Win32_NetworkAdapterConfiguration -ErrorAction SilentlyContinue |
           Where-Object { $_.IPEnabled } | ForEach-Object { $_.IPAddress } |
           Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' })
Write-Host "    Máy này có IP : $($ipToi -join ', ')"
Write-Host "    Máy chia sẻ   : $IP"

# ---------------------------------------------------------------- 2
Buoc "2/5  Dò đường mạng tới $IP"
& ping.exe -n 2 -w 1500 $IP | Out-Null
$ping = ($LASTEXITCODE -eq 0)
$p445 = Test-Cong $IP 445
$p135 = Test-Cong $IP 135
Write-Host "    Ping          : $(if ($ping) {'được'} else {'không trả lời (có thể do firewall chặn ping, chưa chắc là lỗi)'})"
Write-Host "    Cổng 445 (SMB): $(if ($p445) {'MỞ'} else {'ĐÓNG/không tới được'})"
Write-Host "    Cổng 135 (RPC): $(if ($p135) {'MỞ'} else {'ĐÓNG/không tới được'})"

if (-not $p445) {
    if ($ping) {
        Loi "Thấy máy $IP nhưng cổng chia sẻ bị chặn."
        Write-Host "    → Trên máy $IP: chạy file MayChuMoChiaSe.bat (bật File and Printer Sharing, đặt mạng 'Private')."
    } else {
        Loi "Không tới được máy $IP."
        Write-Host "    Kiểm tra lần lượt:"
        Write-Host "     1) IP máy này có cùng dải với $IP không? (cùng 3 số đầu, vd 10.5.208.x). Khác dải = khác mạng/VLAN."
        Write-Host "     2) Wifi cơ quan hay bật 'AP/Client Isolation' (cách ly thiết bị): máy wifi KHÔNG thấy nhau."
        Write-Host "        Cách xử lý: cắm dây LAN cả hai máy, hoặc nhờ người quản trị wifi tắt cách ly."
        Write-Host "     3) Máy $IP đang tắt / ngủ / đổi IP (nên đặt IP tĩnh cho máy chia sẻ)."
    }
    Thoat "Chưa thể kết nối, hãy xử lý theo gợi ý trên rồi chạy lại."
}
OK "Đường mạng thông."

# ---------------------------------------------------------------- 3
Buoc "3/5  Tìm máy in đang chia sẻ"
function Lay-DanhSach([string]$ip) {
    $kq = & net.exe view "\\$ip" 2>&1
    $ok = ($LASTEXITCODE -eq 0)
    $ds = @()
    foreach ($d in $kq) {
        if ($d -match '^\s*(\S.*?)\s{2,}Print\b') { $ds += $Matches[1].Trim() }
    }
    return @{ OK = $ok; DS = $ds; Raw = $kq }
}
$r = Lay-DanhSach $IP
if (-not $r.OK) {
    Canh "Máy $IP yêu cầu đăng nhập (hoặc từ chối). Nhập TÀI KHOẢN + MẬT KHẨU đang dùng để đăng nhập Windows trên máy $IP."
    Canh "(Tài khoản Windows phải có mật khẩu; tài khoản không mật khẩu hoặc chỉ dùng mã PIN sẽ bị từ chối.)"
    $user = Read-Host "Tên đăng nhập (vd TENMAY\ten hoặc .\ten) - Enter để bỏ qua"
    if ($user.Trim()) {
        & net.exe use "\\$IP\IPC$" /delete /y 2>&1 | Out-Null
        & net.exe use "\\$IP\IPC$" * /user:$user
        $r = Lay-DanhSach $IP
    }
}
$tenShare = $null
if ($r.DS.Count -ge 1) {
    OK "Máy in chia sẻ trên ${IP}: $($r.DS -join ' | ')"
    if ($r.DS.Count -eq 1) { $tenShare = $r.DS[0] }
    else {
        for ($i = 0; $i -lt $r.DS.Count; $i++) { Write-Host "      [$($i+1)] $($r.DS[$i])" }
        $c = Read-Host "Chọn số máy in"
        if ($c -match '^\d+$' -and [int]$c -ge 1 -and [int]$c -le $r.DS.Count) { $tenShare = $r.DS[[int]$c - 1] }
    }
} else {
    Canh "Không liệt kê được máy in tự động."
    $r.Raw | ForEach-Object { Write-Host "      $_" }
    Write-Host "    Mở máy $IP: Settings → Printers → chọn máy in → Printer properties → Sharing, xem 'Share name'."
}
if (-not $tenShare) { $tenShare = (Read-Host "Nhập tên chia sẻ (Share name) của máy in").Trim() }
if (-not $tenShare) { Thoat "Chưa có tên máy in chia sẻ." }
$unc = "\\$IP\$tenShare"

# ---------------------------------------------------------------- 4
Buoc "4/5  Kết nối máy in $unc"
function Them-Truc-Tiep([string]$unc) {
    try {
        if ($coCmdlet) { Add-Printer -ConnectionName $unc -ErrorAction Stop }
        else { & rundll32.exe printui.dll,PrintUIEntry /in /n "$unc"; Start-Sleep -Seconds 6 }
        return $null
    } catch { return $_.Exception.Message }
}

function Sua-Windows-May-Nay {
    # Sửa lỗi 0x0000011b / 0x00000709 / 0x00000bcb sau các bản vá bảo mật Windows
    try { New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Print" -Force | Out-Null
          Set-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Print" -Name RpcAuthnLevelPrivacyEnabled -Value 0 -Type DWord } catch {}
    try { $k = "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint"
          New-Item -Path $k -Force | Out-Null
          Set-ItemProperty $k -Name RestrictDriverInstallationToAdministrators -Value 0 -Type DWord
          Set-ItemProperty $k -Name NoWarningNoElevationOnInstall -Value 1 -Type DWord
          Set-ItemProperty $k -Name UpdatePromptSettings -Value 2 -Type DWord } catch {}
    try { Restart-Service Spooler -Force -ErrorAction Stop } catch {}
    Start-Sleep -Seconds 2
}

$loi = Them-Truc-Tiep $unc
if ($loi) {
    Canh "Lần 1 lỗi: $loi"
    Canh "Đang sửa thiết lập Windows trên máy này rồi thử lại..."
    Sua-Windows-May-Nay
    $loi = Them-Truc-Tiep $unc
}

if ($loi) {
    Canh "Lần 2 lỗi: $loi"
    Canh "Chuyển sang cách dự phòng: dùng driver M402 cài sẵn trên máy này + cổng $unc"
    $drv = $null
    if ($coCmdlet) { $drv = Get-PrinterDriver -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'M40[23]' } | Select-Object -First 1 }
    if (-not $drv) {
        Loi "Máy này chưa có driver HP M402. Hãy chạy CaiDriver\CaiDriverM402.bat trước, rồi chạy lại file này."
        Write-Host "    Nếu mã lỗi là 0x0000011b: trên máy $IP chạy MayChuMoChiaSe.bat rồi thử lại."
        Thoat "Chưa kết nối được máy in."
    }
    try {
        if (-not (Get-PrinterPort -Name $unc -ErrorAction SilentlyContinue)) { Add-PrinterPort -Name $unc -ErrorAction Stop }
        Add-Printer -Name "HP M402 ($IP)" -DriverName $drv.Name -PortName $unc -ErrorAction Stop
        $loi = $null
        OK "Đã tạo máy in bằng driver cài sẵn: $($drv.Name)"
    } catch {
        $loi = $_.Exception.Message
        Loi "Cách dự phòng cũng lỗi: $loi"
    }
}

if ($loi) {
    Write-Host ""
    Write-Host "    GỢI Ý THEO MÃ LỖI:"
    Write-Host "     0x0000011b  → trên máy $IP chạy MayChuMoChiaSe.bat (tắt RPC Privacy) rồi chạy lại."
    Write-Host "     0x00000709  → máy này: chạy lại bằng quyền Admin (bat đã tự xin); máy kia: MayChuMoChiaSe.bat."
    Write-Host "     0x00000bcb  → thiếu driver: chạy CaiDriverM402.bat trên máy này rồi chạy lại."
    Write-Host "     0x00000005 / Access denied → sai tài khoản/mật khẩu máy $IP, hoặc tài khoản đó không có mật khẩu."
    Thoat "Chưa kết nối được. Gửi file ket_noi_may_in.log để được xem tiếp."
}

# ---------------------------------------------------------------- 5
Buoc "5/5  Đặt làm máy in mặc định"
try {
    $may = Get-Printer -ErrorAction Stop | Where-Object { $_.PortName -eq $unc -or $_.Name -like "*$tenShare*" -or $_.Name -like "*$IP*" } | Select-Object -First 1
    if ($may) {
        (New-Object -ComObject WScript.Network).SetDefaultPrinter($may.Name)
        OK "Máy in mặc định: $($may.Name)"
    } else { Canh "Không tìm thấy máy in vừa thêm để đặt mặc định (vẫn có thể đã thêm được, xem Settings → Printers)." }
} catch { Canh "Không đặt được mặc định: $($_.Exception.Message)" }

Write-Host ""
Write-Host "================  XONG  ================" -ForegroundColor Green
Write-Host "In thử 1 trang: Settings → Printers → chọn máy in → Print test page."
try { Stop-Transcript | Out-Null } catch {}
