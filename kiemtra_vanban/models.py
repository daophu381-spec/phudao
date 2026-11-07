from __future__ import annotations

from dataclasses import dataclass, field, asdict

# Nhóm lỗi (thứ tự hiển thị)
CAT_SPELL = "Chính tả"
CAT_CAP = "Viết hoa"
CAT_PUNCT = "Dấu câu và dấu cách"
CAT_FONT = "Phông chữ, cỡ chữ, trình bày"
CAT_LAYOUT = "Rơi chữ, rơi dòng, rơi trang"
CAT_FORM = "Thể thức NĐ30"
CATEGORIES = [CAT_SPELL, CAT_CAP, CAT_PUNCT, CAT_FONT, CAT_LAYOUT, CAT_FORM]

# Mức độ
SEV_ERROR = "LỖI"
SEV_WARN = "CẢNH BÁO"      # nhiều khả năng sai nhưng cần người xác nhận (tên riêng, mẫu có sẵn...)
SEV_HINT = "GỢI Ý"         # lệch khuyến nghị NĐ30, không tự sửa nếu là biểu mẫu
SEV_ORDER = {SEV_ERROR: 0, SEV_WARN: 1, SEV_HINT: 2}


@dataclass
class Issue:
    category: str
    severity: str
    where: str            # vị trí: "Đoạn 12", "Bảng 1 / ô 2", "Trang 2 / dòng cuối"
    found: str            # đoạn văn bản có lỗi (trích ngắn)
    message: str          # mô tả lỗi
    fix: str = ""         # gợi ý sửa thành
    order: int = 0        # thứ tự xuất hiện trong văn bản

    def to_dict(self):
        return asdict(self)


@dataclass
class Para:
    """Một đoạn văn bản đã làm phẳng (trong thân văn bản hoặc trong ô bảng)."""
    index: int
    text: str
    where: str
    in_table: bool = False
    obj: object = None            # docx Paragraph (nếu có)
    page: int | None = None


def short(text: str, start: int = 0, end: int | None = None, pad: int = 18) -> str:
    """Trích đoạn ngắn quanh vị trí lỗi để người đọc tìm lại."""
    end = start + 1 if end is None else end
    a, b = max(0, start - pad), min(len(text), end + pad)
    s = text[a:b].replace(" ", "·").replace("\t", "→").replace("\n", "⏎")
    return ("…" if a > 0 else "") + s + ("…" if b < len(text) else "")
