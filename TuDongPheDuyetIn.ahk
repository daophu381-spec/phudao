; ======================================================================
;  TỰ ĐỘNG PHÊ DUYỆT + IN (máy in HP M402) - BẤM 1 LẦN
;  Cần cài AutoHotkey v2: https://www.autohotkey.com  (bản v2.0 trở lên)
;
;  PHÍM TẮT
;    F8        : Ghi bước "BẤM CHUỘT" tại vị trí con trỏ đang đứng
;    F6        : Ghi bước "BẤM PHÍM" (mặc định Enter) vào cửa sổ đang mở
;    F5        : Ghi bước "CHỜ" thêm vài giây (khi phần mềm xử lý chậm)
;    F7        : Xoá hết các bước đã ghi để ghi lại từ đầu
;    F1        : Xem danh sách các bước đã ghi
;    F9        : CHẠY (Phê duyệt + In) - giống bấm nút to trên màn hình
;    Esc       : Dừng khẩn cấp khi đang chạy
;  Cấu hình được lưu trong file cau_hinh.ini cạnh file này.
; ======================================================================
#Requires AutoHotkey v2.0
#SingleInstance Force
SetTitleMatchMode 2          ; tên cửa sổ chỉ cần "chứa" chuỗi đã lưu
SendMode "Input"

global IniFile  := A_ScriptDir "\cau_hinh.ini"
global Running  := false
global StopFlag := false

KhoiTaoIni()
A_IconTip := "Tự động Phê duyệt + In"
TaoNutBam()

; ----------------------------- PHÍM TẮT --------------------------------
F8:: GhiBuocChuot()
F6:: GhiBuocPhim()
F5:: GhiBuocCho()
F7:: XoaHetBuoc()
F1:: XemDanhSach()
F9:: SetTimer(ChayMotLan, -10)
~Esc:: {
    global StopFlag
    if Running
        StopFlag := true
}

; ----------------------------- KHỞI TẠO --------------------------------
KhoiTaoIni() {
    ; Tạo file INI dạng UTF-16 để lưu được tên cửa sổ tiếng Việt có dấu
    if !FileExist(IniFile) {
        FileAppend("[Chung]`n", IniFile, "UTF-16")
        IniWrite(0,    IniFile, "Chung", "SoBuoc")
        IniWrite(700,  IniFile, "Chung", "NghiSauMoiBuoc_ms")
        IniWrite(20,   IniFile, "Chung", "ChoCuaSoToiDa_giay")
    }
}

SoBuoc() => Integer(IniRead(IniFile, "Chung", "SoBuoc", 0))

; ----------------------- NÚT BẤM NỔI TRÊN MÀN HÌNH ----------------------
TaoNutBam() {
    global NutGui
    NutGui := Gui("+AlwaysOnTop +ToolWindow", "Phê duyệt + In")
    NutGui.SetFont("s13 bold")
    btn := NutGui.Add("Button", "w230 h56", "▶  PHÊ DUYỆT + IN")
    btn.OnEvent("Click", (*) => SetTimer(ChayMotLan, -10))
    NutGui.SetFont("s9 norm")
    NutGui.Add("Button", "w110 h26", "Xem các bước").OnEvent("Click", (*) => XemDanhSach())
    NutGui.Add("Button", "x+10 w110 h26", "Hướng dẫn").OnEvent("Click", (*) => HuongDan())
    NutGui.OnEvent("Close", (*) => (LuuViTriNut(), ExitApp()))

    x := IniRead(IniFile, "Chung", "NutX", "")
    y := IniRead(IniFile, "Chung", "NutY", "")
    if (x != "" && y != "")
        NutGui.Show("x" x " y" y " NoActivate")
    else
        NutGui.Show("NoActivate")

    if SoBuoc() = 0
        HuongDan()
}

LuuViTriNut() {
    try {
        WinGetPos(&x, &y, , , "ahk_id " NutGui.Hwnd)
        IniWrite(x, IniFile, "Chung", "NutX")
        IniWrite(y, IniFile, "Chung", "NutY")
    }
}

HuongDan() {
    MsgBox(
    "CÁCH GHI CÁC BƯỚC (chỉ làm 1 lần):`n`n"
    "1. Mở phần mềm giao dịch, chọn sẵn 1 chứng từ cần duyệt.`n"
    "2. Đặt chuột lên nút [Phê duyệt] (KHÔNG bấm) → nhấn F8.`n"
    "3. Bấm tay nút Phê duyệt. Nếu hiện hộp hỏi xác nhận → đặt chuột`n"
    "    lên nút [Đồng ý]/[Có] → F8, rồi bấm tay nút đó.`n"
    "4. Đặt chuột lên nút [In] → F8, rồi bấm tay nút In.`n"
    "5. Hộp thoại in hiện ra (đã chọn HP M402) → nhấn F6 (ghi phím Enter)`n"
    "    hoặc đặt chuột lên nút [In]/[Print] → F8.`n"
    "6. Nếu phần mềm chậm, nhấn F5 để chèn thời gian chờ ở chỗ cần.`n`n"
    "DÙNG HẰNG NGÀY: chọn chứng từ → bấm nút [▶ PHÊ DUYỆT + IN] hoặc F9.`n"
    "Đang chạy muốn dừng: nhấn Esc.   Ghi lại từ đầu: F7.   Xem bước: F1.",
    "Hướng dẫn", "Iconi")
}

; ------------------------------ GHI BƯỚC -------------------------------
GhiBuocChuot() {
    if Running
        return
    CoordMode "Mouse", "Screen"
    MouseGetPos(&sx, &sy, &hwnd)
    if (hwnd = NutGui.Hwnd) {
        MsgBox("Hãy đặt chuột lên nút của PHẦN MỀM GIAO DỊCH, không phải nút của công cụ này.", "Chưa ghi", "Icon!")
        return
    }
    WinGetPos(&wx, &wy, , , "ahk_id " hwnd)
    tieuDe := WinGetTitle("ahk_id " hwnd)
    chuongTrinh := WinGetProcessName("ahk_id " hwnd)

    tieuDe := HoiTenCuaSo(tieuDe, chuongTrinh)
    if (tieuDe = "#HUY")
        return

    n := SoBuoc() + 1
    s := "Buoc" n
    IniWrite("CLICK",        IniFile, s, "Loai")
    IniWrite(tieuDe,         IniFile, s, "TieuDe")
    IniWrite(chuongTrinh,    IniFile, s, "ChuongTrinh")
    IniWrite(sx - wx,        IniFile, s, "X")
    IniWrite(sy - wy,        IniFile, s, "Y")
    IniWrite(n,              IniFile, "Chung", "SoBuoc")
    ThongBao("Đã ghi bước " n ": bấm chuột tại (" (sx - wx) ", " (sy - wy) ") trong cửa sổ «" tieuDe "»")
}

GhiBuocPhim() {
    if Running
        return
    hwnd := WinExist("A")
    if (!hwnd || hwnd = NutGui.Hwnd) {
        MsgBox("Hãy bấm vào cửa sổ của phần mềm / hộp thoại in trước rồi mới nhấn F6.", "Chưa ghi", "Icon!")
        return
    }
    tieuDe := WinGetTitle("ahk_id " hwnd)
    chuongTrinh := WinGetProcessName("ahk_id " hwnd)

    r := InputBox("Phím cần bấm:`n  {Enter} = Enter    {Tab} = Tab    {Space} = Cách`n  !i = Alt+I    ^p = Ctrl+P",
                  "Ghi bước bấm phím", "w360 h170", "{Enter}")
    if (r.Result != "OK" || r.Value = "")
        return
    tieuDe := HoiTenCuaSo(tieuDe, chuongTrinh)
    if (tieuDe = "#HUY")
        return

    n := SoBuoc() + 1
    s := "Buoc" n
    IniWrite("PHIM",      IniFile, s, "Loai")
    IniWrite(tieuDe,      IniFile, s, "TieuDe")
    IniWrite(chuongTrinh, IniFile, s, "ChuongTrinh")
    IniWrite(r.Value,     IniFile, s, "Phim")
    IniWrite(n,           IniFile, "Chung", "SoBuoc")
    ThongBao("Đã ghi bước " n ": bấm phím " r.Value " trong cửa sổ «" tieuDe "»")
}

GhiBuocCho() {
    if Running
        return
    r := InputBox("Chờ thêm bao nhiêu giây? (ví dụ 2 hoặc 1.5)", "Ghi bước chờ", "w300 h120", "2")
    if (r.Result != "OK" || !IsNumber(r.Value))
        return
    n := SoBuoc() + 1
    s := "Buoc" n
    IniWrite("CHO",                       IniFile, s, "Loai")
    IniWrite(Round(r.Value * 1000),       IniFile, s, "ms")
    IniWrite(n,                           IniFile, "Chung", "SoBuoc")
    ThongBao("Đã ghi bước " n ": chờ " r.Value " giây")
}

; Cho phép xoá bớt phần tên cửa sổ hay thay đổi (số chứng từ, tên khách hàng...)
HoiTenCuaSo(tieuDe, chuongTrinh) {
    r := InputBox(
        "Tên cửa sổ dùng để nhận diện (chương trình: " chuongTrinh ").`n"
        "Nếu tên có phần thay đổi theo từng chứng từ (số CT, tên khách...), hãy XOÁ phần đó, "
        "chỉ giữ phần cố định. Để trống = chỉ nhận theo chương trình.",
        "Xác nhận cửa sổ", "w460 h190", tieuDe)
    return (r.Result = "OK") ? Trim(r.Value) : "#HUY"
}

XoaHetBuoc() {
    if Running
        return
    if MsgBox("Xoá hết " SoBuoc() " bước đã ghi để ghi lại từ đầu?", "Xoá các bước", "YesNo Icon?") != "Yes"
        return
    Loop SoBuoc()
        IniDelete(IniFile, "Buoc" A_Index)
    IniWrite(0, IniFile, "Chung", "SoBuoc")
    ThongBao("Đã xoá hết. Bắt đầu ghi lại bằng F8 / F6 / F5.")
}

XemDanhSach() {
    n := SoBuoc()
    if n = 0 {
        MsgBox("Chưa ghi bước nào. Nhấn nút [Hướng dẫn] để xem cách ghi.", "Danh sách bước")
        return
    }
    t := ""
    Loop n
        t .= A_Index ". " MoTaBuoc("Buoc" A_Index) "`n"
    MsgBox(t "`n(Sửa chi tiết trong file cau_hinh.ini)", "Danh sách " n " bước")
}

MoTaBuoc(s) {
    loai := IniRead(IniFile, s, "Loai", "")
    cs := IniRead(IniFile, s, "TieuDe", "") " [" IniRead(IniFile, s, "ChuongTrinh", "") "]"
    switch loai {
        case "CLICK": return "Bấm chuột (" IniRead(IniFile, s, "X", "") ", " IniRead(IniFile, s, "Y", "") ") – " cs
        case "PHIM":  return "Bấm phím " IniRead(IniFile, s, "Phim", "") " – " cs
        case "CHO":   return "Chờ " IniRead(IniFile, s, "ms", 0) / 1000 " giây"
    }
    return "(không rõ)"
}

; ------------------------------- CHẠY ----------------------------------
ChayMotLan() {
    global Running, StopFlag
    if Running
        return
    n := SoBuoc()
    if n = 0 {
        HuongDan()
        return
    }
    Running := true
    StopFlag := false
    nghi    := Integer(IniRead(IniFile, "Chung", "NghiSauMoiBuoc_ms", 700))
    choToiDa := Integer(IniRead(IniFile, "Chung", "ChoCuaSoToiDa_giay", 20))

    try {
        Loop n {
            if StopFlag
                throw Error("Đã dừng theo yêu cầu (Esc).")
            s := "Buoc" A_Index
            loai := IniRead(IniFile, s, "Loai", "")
            ToolTip("Đang chạy bước " A_Index "/" n ": " MoTaBuoc(s) "`n(Esc để dừng)")

            if (loai = "CHO") {
                ChoCoTheDung(Integer(IniRead(IniFile, s, "ms", 1000)))
                continue
            }

            ; Chờ đúng cửa sổ xuất hiện rồi đưa lên trước
            tieuDe := IniRead(IniFile, s, "TieuDe", "")
            ct     := IniRead(IniFile, s, "ChuongTrinh", "")
            win    := (tieuDe != "" ? tieuDe " " : "") (ct != "" ? "ahk_exe " ct : "")
            if (win != "") {
                if !WinWait(win, , choToiDa)
                    throw Error("Bước " A_Index ": không thấy cửa sổ «" tieuDe "» sau " choToiDa " giây.")
                if StopFlag
                    throw Error("Đã dừng theo yêu cầu (Esc).")
                WinActivate(win)
                if !WinWaitActive(win, , 5)
                    throw Error("Bước " A_Index ": không đưa được cửa sổ «" tieuDe "» lên trước.")
            }

            if (loai = "CLICK") {
                CoordMode "Mouse", "Window"
                Click(IniRead(IniFile, s, "X"), IniRead(IniFile, s, "Y"))
            } else if (loai = "PHIM") {
                Send(IniRead(IniFile, s, "Phim", "{Enter}"))
            }
            ChoCoTheDung(nghi)
        }
        ToolTip("✔ Xong – đã phê duyệt và gửi lệnh in.")
        SoundBeep(1200, 120)
    } catch as e {
        ToolTip()
        SoundBeep(400, 300)
        MsgBox(e.Message "`n`nHãy kiểm tra lại chứng từ trên phần mềm trước khi chạy tiếp.", "Đã dừng", "Icon! 262144")
    }
    Running := false
    SetTimer(() => ToolTip(), -2500)
}

; Sleep nhưng vẫn dừng ngay được khi nhấn Esc
ChoCoTheDung(ms) {
    het := A_TickCount + ms
    while (A_TickCount < het) {
        if StopFlag
            throw Error("Đã dừng theo yêu cầu (Esc).")
        Sleep(50)
    }
}

ThongBao(t) {
    ToolTip(t)
    SetTimer(() => ToolTip(), -3000)
}
