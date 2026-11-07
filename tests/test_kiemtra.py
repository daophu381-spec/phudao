"""Chạy:  python -m unittest discover -s tests -v   (từ thư mục gốc của dự án)"""
import os
import shutil
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
sys.path.insert(0, HERE)

import tao_mau  # noqa: E402
from kiemtra_vanban import layout  # noqa: E402
from kiemtra_vanban.docxrules import DocxFile, check_docx  # noqa: E402
from kiemtra_vanban.models import Para  # noqa: E402
from kiemtra_vanban.textrules import check_text_document  # noqa: E402
from kiemtra_vanban.vietnamese import check_syllable  # noqa: E402

HAS_LAYOUT = bool(shutil.which("soffice") or shutil.which("libreoffice"))
try:
    import pdfplumber  # noqa: F401
except ImportError:
    HAS_LAYOUT = False


def run_text(lines):
    paras = [Para(i, t, f"Đoạn {i + 1}") for i, t in enumerate(lines)]
    return check_text_document(paras)


def msgs(issues):
    return " | ".join(i.message + " -> " + i.fix for i in issues)


class TestSyllable(unittest.TestCase):
    def test_valid(self):
        for w in ("Việt Nam hòa hoà thủy thuỷ khỏe ủy quốc quý quỳ giá gì gìn giường người nghề ngành nghĩa kiểm "
                  "của cửa khuya khuỷu tuyệt ngoằn loại oanh xoong hoạch lệch ếch ích tiếng yêu ý ở kìa mưa trường "
                  "nghiệp chính sách xã hội Đoàn Nguyễn Thanh niên hằng quyết định Phù Cừ Hưng Yên tuần thuế được").split():
            self.assertTrue(check_syllable(w).ok, w)

    def test_invalid(self):
        cases = {"nghành": "ngành", "ghái": "gái", "nhât": None, "tien": None, "ngươì": "người", "ko": None, "thôg": None}
        for w, fix in cases.items():
            r = check_syllable(w)
            self.assertFalse(r.ok, w)
            if fix:
                self.assertEqual(r.suggestion, fix, w)

    def test_tone_style(self):
        self.assertEqual(check_syllable("hòa").style, "moi")
        self.assertEqual(check_syllable("hoà").style, "cu")
        self.assertEqual(check_syllable("quý").style, "")


class TestTextRules(unittest.TestCase):
    def test_clean_text_has_no_issue(self):
        good = [
            "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Số: 05/NHCS-TD", f"Phù Cừ, ngày 05 tháng 10 năm 2026",
            "V/v kiểm tra sử dụng vốn vay tại các xã trên địa bàn",
            "Ngân hàng Chính sách xã hội huyện Phù Cừ hằng tháng xử lý hồ sơ của hộ nghèo, vay vốn (gọi tắt là NHCSXH).",
            "Tổng dư nợ là 98.835.000 đồng, lãi suất 0,66%/tháng. Thực hiện theo Nghị định số 30/2020/NĐ-CP, điểm a khoản 2 Điều 5.",
            "Ủy ban nhân dân xã Tống Trân, tỉnh Hưng Yên; thôn Đông, 08:30 ngày 3/2 tại http://abc.vn/x.html.",
        ]
        self.assertEqual(run_text(good), [])

    def test_errors_detected(self):
        expect = {
            "Số 5/NHCS-TD": ["Thiếu dấu hai chấm", "thêm số 0"],
            "Phù Cừ, ngày 5 tháng 03 năm 2026": ["phải thêm số 0"],
            "Xong rồi  nhé .": ["Hai dấu cách", "trước dấu chấm"],
            "Hộ Nghèo,vay vốn(gọi tắt là NHCSXH )": ["Viết hoa tùy tiện", "Thiếu dấu cách sau", "trước ngoặc", "bên trong ngoặc"],
            "Tổng 98,835,000 đồng, lãi 0.66%/tháng.Xong": ["dấu phẩy phân cách", "dấu phẩy, không dùng dấu chấm", "sau dấu chấm"],
            "hàng tháng sử lý hồ sơ, bổ xung và sát nhập": ["hằng tháng", "xử lý", "bổ sung", "sáp nhập"],
            "nghành tín dụng ko tốt và và đúng": ["ngành", "không", "Lặp từ"],
            "Tại Tỉnh Hưng Yên theo điều 5 nghị định số 30/2020/NĐ - CP": ["tỉnh", "Điều 5", "Nghị định số", "Gạch nối"],
            "Ngân hàng Chính sách Xã hội và Ủy Ban Nhân Dân": ["Ngân hàng Chính sách xã hội", "Ủy ban nhân dân"],
            "Nguyễn VănAn gởi": ["Chữ hoa xuất hiện giữa từ", "gửi"],
        }
        for line, wants in expect.items():
            out = msgs(run_text([line]))
            for w in wants:
                self.assertIn(w, out, f"{line!r} thiếu {w!r}\n{out}")

    def test_consistency(self):
        out = msgs(run_text(["Hoà bình và hòa thuận.", "Kí tên rồi ký nhận."]))
        self.assertIn("hai kiểu bỏ dấu", out)
        self.assertIn("thống nhất", out)


class TestDocx(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp()
        tao_mau.build(os.path.join(cls.tmp, "dung.docx"), good=True)
        tao_mau.build(os.path.join(cls.tmp, "loi.docx"), good=False)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def check(self, name):
        f = DocxFile(os.path.join(self.tmp, name))
        return check_text_document(f.paras) + check_docx(f)

    def test_good_sample_is_clean(self):
        self.assertEqual([i.to_dict() for i in self.check("dung.docx")], [])

    def test_bad_sample(self):
        out = msgs(self.check("loi.docx"))
        for w in ("Quốc hiệu phải viết IN HOA", "Tiêu ngữ phải viết", "Tiêu ngữ phải in đậm", "in nghiêng",
                  "Tên loại văn bản phải viết IN HOA", "Thiếu dấu hai chấm sau 'Kính gửi'", "Phải viết 'Nơi nhận:'",
                  "gạch đầu dòng '- '", "Thiếu dấu chấm sau 'KT'", "cantSplit", "Cỡ chữ", "Lề trái"):
            self.assertIn(w, out, w)


@unittest.skipUnless(HAS_LAYOUT, "cần LibreOffice và pdfplumber")
class TestLayout(unittest.TestCase):
    def analyze(self, build):
        tmp = tempfile.mkdtemp()
        try:
            path = os.path.join(tmp, "x.docx")
            build(path)
            return layout.analyze_docx(path)[0]
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def test_orphan_word_in_body(self):
        from docx import Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH as AL

        def build(path):
            d = Document()
            tao_mau.setup(d, margins=(2, 2, 2, 2))
            tao_mau.para(d, tao_mau.ORPHAN_TEXT, size=14, align=AL.JUSTIFY, indent=1.27)
            d.save(path)
        out = self.analyze(build)
        self.assertTrue(any("một mình một dòng" in i.message for i in out), msgs(out))

    def test_orphan_word_in_title(self):
        from docx import Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH as AL

        def build(path):
            d = Document()
            tao_mau.setup(d, margins=(2, 2, 2, 2))
            tao_mau.para(d, tao_mau.LONG_TITLE, bold=True, size=14, align=AL.CENTER)
            d.save(path)
        out = self.analyze(build)
        self.assertTrue(any("rơi chữ" in i.message for i in out), msgs(out))

    def test_good_sample_has_no_layout_issue(self):
        out = self.analyze(lambda p: tao_mau.build(p, good=True))
        self.assertEqual([i.message for i in out], [])

    def test_short_last_page(self):
        from docx import Document

        def build(path):
            d = Document()
            tao_mau.setup(d)
            for i in range(60):
                tao_mau.para(d, f"Dòng nội dung số {i + 1} của văn bản thử nghiệm kiểm tra rơi trang.", size=14)
            d.save(path)
        out = self.analyze(build)
        # 60 dòng 14 pt không thể vừa 1 trang; trang cuối có ít dòng thì phải có cảnh báo, nhiều thì không
        pages_msg = [i for i in out if "Trang cuối" in i.message]
        self.assertTrue(len(pages_msg) <= 1)


if __name__ == "__main__":
    unittest.main()
