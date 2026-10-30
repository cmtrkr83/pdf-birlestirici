# -*- coding: utf-8 -*-
"""
İlçe PDF Birleştirici

Bir ilçe klasörü seçersiniz. Klasör içindeki her alt klasör bir kurumu
(kurum kodu) temsil eder. Her kurum klasörünün içindeki PDF'ler o kurumun
şubeleridir.

Program:
  * Her kurum klasöründeki PDF'leri isim sırasına (doğal sıralama) göre
    birleştirir ve kurum kodu adıyla tek bir PDF üretir.
  * Oluşturulan okul PDF'lerini kurum klasörü sırasını bozmadan tek bir
    ilçe PDF'si olarak da birleştirir.
  * Sayfa sırası, yön ve orijinal çözünürlük korunur (yeniden ölçekleme
    veya döndürme yapılmaz).
  * Her PDF'in ilk 2 sayfasını sınıf listesi kabul eder, 3. sayfadan
    itibaren 2 sayfayı 1 öğrencinin optik kağıdı sayar.
  * Hata veren dosya ve klasörleri günlükte gösterir.
  * İlçe özeti, okul detayı ve şube detayı sayfalarından oluşan
    rapor.xlsx üretir.

Gereksinim: pip install pypdf openpyxl
"""

import math
import os
import queue
import re
import shutil
import stat
import threading
import traceback
from datetime import datetime

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    PdfReader = None
    PdfWriter = None

# Opsiyonel: okul adı metnini çok daha hızlı çıkarmak için PyMuPDF
try:
    import pymupdf as _pymupdf
except ImportError:
    try:
        import fitz as _pymupdf
    except ImportError:
        _pymupdf = None

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    Workbook = None

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Sınıf listesi sayfa sayısı (optik kağıtlar bu sayfadan sonra başlar)
CLASS_LIST_PAGES = 2
# Bir öğrencinin optik kağıt sayfa sayısı (arkalı önlü)
STUDENT_PAGES = 2
# İlçe klasörü içinde üretilecek çıktı klasörünün adı
OUTPUT_DIR_NAME = "Birlestirilmis_PDFler"
DISTRICT_PDF_FILE_NAME = "Ilce_Birlestirilmis.pdf"
EXCEL_FILE_NAME = "rapor.xlsx"
TEXT_FILE_NAME = "rapor.txt"


def natural_key(text):
    """Dosya adlarını doğal sıralar: 1, 2, 10 (9, 10'dan sonra gelir)."""
    return [int(part) if part.isdigit() else part.lower()
            for part in re.split(r"(\d+)", text)]


def list_pdfs(folder):
    """Bir klasördeki PDF dosyalarını doğal sırada döndürür."""
    files = []
    for name in os.listdir(folder):
        full = os.path.join(folder, name)
        if os.path.isfile(full) and name.lower().endswith(".pdf"):
            files.append(full)
    files.sort(key=lambda p: natural_key(os.path.basename(p)))
    return files


def remove_tree(path):
    """Klasörü (salt okunur dosyalar dahil) siler."""
    def on_error(func, p, exc_info):
        try:
            os.chmod(p, stat.S_IWRITE)
            func(p)
        except Exception:
            pass
    shutil.rmtree(path, onerror=on_error)


def students_from_pages(page_count):
    """Sayfa sayısından öğrenci sayısını hesaplar."""
    optic_pages = max(page_count - CLASS_LIST_PAGES, 0)
    if optic_pages <= 0:
        return 0
    return math.ceil(optic_pages / STUDENT_PAGES)


def first_page_text(path, reader):
    """İlk sayfanın metnini döndürür (mümkünse hızlı PyMuPDF ile)."""
    if _pymupdf is not None:
        try:
            doc = _pymupdf.open(path)
            try:
                text = doc[0].get_text()
            finally:
                doc.close()
            if text:
                return text
        except Exception:
            pass
    try:
        return reader.pages[0].extract_text() or ""
    except Exception:
        return ""


def extract_school_info(text):
    """İlk sayfadaki metinden okul adı, il-ilçe ve sınıf/şube bilgisini çıkarır.

    Yerleşim (aynı sistemden üretilen tüm PDF'lerde tutarlı):
        İL - İLÇE
        OKUL ADI
        SINIFI / ŞUBESİ
        :
        :
        :
        <İL - İLÇE değeri>
        <OKUL ADI değeri>
        <SINIF / ŞUBE değeri>
    """
    result = {"okul_adi": "", "il_ilce": "", "sinif": ""}
    if not text:
        return result

    lines = [l.strip() for l in text.splitlines()]
    start = None
    for i, line in enumerate(lines):
        if "OKUL ADI" in line.upper():
            start = i
            break

    values = []
    if start is not None:
        for line in lines[start + 1:]:
            if not line:
                continue
            if line == ":" or line.upper() in ("SINIFI / ŞUBESİ", "SINIFI/ŞUBESİ"):
                continue
            if line.upper() in ("İL - İLÇE", "IL - ILCE"):
                continue
            if line.upper().startswith("SN") and len(line) <= 3:
                break
            values.append(line)
            if len(values) >= 3:
                break

    if len(values) >= 2:
        result["il_ilce"] = values[0]
        result["okul_adi"] = values[1]
        if len(values) >= 3:
            result["sinif"] = values[2]

    # Yedek yöntem: il-ilçe satırından (BÜYÜK HARF, " - " içerir) sonraki satır
    if not result["okul_adi"]:
        for i, line in enumerate(lines):
            if " - " in line and line == line.upper() \
                    and any(ch.isalpha() for ch in line):
                if i + 1 < len(lines):
                    result["il_ilce"] = line
                    result["okul_adi"] = lines[i + 1]
                    break

    return result


def merge_folder_pdfs(pdf_paths, output_path, log):
    """PDF'leri sırayla tek dosyada birleştirir.

    Orijinal sayfa içeriği, yönü ve boyutu değiştirilmez.

    Döndürür: dict(total_pages, ok_branches, students, errors)
      errors: [(dosya_adi, sebep), ...]
    """
    writer = PdfWriter()
    total_pages = 0
    ok_branches = 0
    students = 0
    errors = []
    branches_info = []
    school_info = {"okul_adi": "", "il_ilce": "", "sinif": ""}

    for path in pdf_paths:
        name = os.path.basename(path)
        try:
            reader = PdfReader(path)
            if reader.is_encrypted:
                try:
                    reader.decrypt("")
                except Exception:
                    pass
            page_count = len(reader.pages)
        except Exception as exc:
            errors.append((name, f"okunamadı: {exc}"))
            branches_info.append({"dosya": name, "pages": "",
                                  "students": "", "durum": "Hatalı"})
            log(f"    [HATA] {name} -> okunamadı: {exc}")
            continue

        # Okul adı ilk sayfadan bir kez alınır (bulunana kadar denenir)
        if not school_info["okul_adi"]:
            info = extract_school_info(first_page_text(path, reader))
            if info["okul_adi"]:
                school_info = info

        try:
            for page in reader.pages:
                writer.add_page(page)
        except Exception as exc:
            errors.append((name, f"birleştirilemedi: {exc}"))
            branches_info.append({"dosya": name, "pages": "",
                                  "students": "", "durum": "Hatalı"})
            log(f"    [HATA] {name} -> birleştirilemedi: {exc}")
            continue

        total_pages += page_count
        stu = students_from_pages(page_count)
        students += stu
        ok_branches += 1
        branches_info.append({"dosya": name, "pages": page_count,
                              "students": stu, "durum": "Başarılı"})

    if ok_branches == 0:
        return {"total_pages": 0, "ok_branches": 0, "students": 0,
                "errors": errors, "branches": branches_info,
                "school_info": school_info}

    try:
        with open(output_path, "wb") as fh:
            writer.write(fh)
    except Exception as exc:
        errors.append((os.path.basename(output_path),
                       f"yazma hatası: {exc}"))
        log(f"    [HATA] çıktı yazılamadı ({output_path}): {exc}")
        return {"total_pages": 0, "ok_branches": 0, "students": 0,
                "errors": errors, "branches": branches_info,
                "school_info": school_info}

    return {"total_pages": total_pages, "ok_branches": ok_branches,
            "students": students, "errors": errors,
            "branches": branches_info, "school_info": school_info}


def merge_created_pdfs(pdf_paths, output_path, log):
    """Oluşturulmuş PDF'leri verilen sırayı koruyarak tek dosyada birleştirir."""
    writer = PdfWriter()
    total_pages = 0

    for path in pdf_paths:
        name = os.path.basename(path)
        try:
            reader = PdfReader(path)
            if reader.is_encrypted:
                try:
                    reader.decrypt("")
                except Exception:
                    pass
            for page in reader.pages:
                writer.add_page(page)
            total_pages += len(reader.pages)
        except Exception as exc:
            log(f"    [HATA] {name} -> ilçe PDF'sine eklenemedi: {exc}")
            return False, total_pages

    if total_pages == 0:
        log("    [HATA] İlçe PDF'si için birleştirilecek okul PDF'i bulunamadı.")
        return False, 0

    try:
        with open(output_path, "wb") as fh:
            writer.write(fh)
    except Exception as exc:
        log(f"    [HATA] ilçe PDF'si yazılamadı ({output_path}): {exc}")
        return False, 0

    return True, total_pages


def write_excel(report_path, district_dir, school_rows, branch_rows,
                error_rows, il_ilce=""):
    """Excel raporunu sayfalı olarak yazar."""
    wb = Workbook()

    title_font = Font(bold=True, size=14, color="FFFFFF")
    title_fill = PatternFill("solid", fgColor="1F4E78")
    head_font = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="2E75B6")
    label_font = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    total_school = len(school_rows)
    total_branches = sum(r["branches"] for r in school_rows)
    total_students = sum(r["students"] for r in school_rows)
    total_pages = sum(r["pages"] for r in school_rows)
    total_errors = len(error_rows)

    # ---- Sayfa 1: İlçe Özeti ----
    ws = wb.active
    ws.title = "İlçe Özeti"
    ws.merge_cells("A1:B1")
    ws["A1"] = "İLÇE PDF BİRLEŞTİRME RAPORU"
    ws["A1"].font = title_font
    ws["A1"].fill = title_fill
    ws["A1"].alignment = center

    rows = [
        ("İlçe klasörü", district_dir),
        ("İl / İlçe", il_ilce or "-"),
        ("Oluşturma tarihi", datetime.now().strftime("%d.%m.%Y %H:%M")),
        ("Toplam okul sayısı", total_school),
        ("Toplam şube sayısı", total_branches),
        ("Toplam öğrenci sayısı", total_students),
        ("Toplam sayfa sayısı", total_pages),
        ("Toplam hatalı dosya/klasör", total_errors),
    ]
    r = 3
    for label, value in rows:
        ws.cell(row=r, column=1, value=label).font = label_font
        ws.cell(row=r, column=2, value=value)
        ws.cell(row=r, column=1).alignment = left
        ws.cell(row=r, column=2).alignment = left
        r += 1
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 60

    # ---- Sayfa 2: Okul Detayı ----
    ws2 = wb.create_sheet("Okul Detayı")
    headers = ["Kurum Kodu", "Okul Adı", "Şube Sayısı", "Öğrenci Sayısı",
               "Sayfa Sayısı", "Hatalı Dosya"]
    for c, h in enumerate(headers, start=1):
        cell = ws2.cell(row=1, column=c, value=h)
        cell.font = head_font
        cell.fill = head_fill
        cell.alignment = center
        cell.border = border
    for i, row in enumerate(school_rows, start=2):
        values = [row["kod"], row["okul_adi"] or "-", row["branches"],
                  row["students"], row["pages"], row["errors"]]
        for c, v in enumerate(values, start=1):
            cell = ws2.cell(row=i, column=c, value=v)
            cell.border = border
            cell.alignment = left if c <= 2 else center
    last = len(school_rows) + 1
    tr = last + 1
    ws2.cell(row=tr, column=1, value="TOPLAM").font = label_font
    ws2.cell(row=tr, column=3, value=total_branches).font = label_font
    ws2.cell(row=tr, column=4, value=total_students).font = label_font
    ws2.cell(row=tr, column=5, value=total_pages).font = label_font
    ws2.cell(row=tr, column=6, value=sum(r["errors"] for r in school_rows)
             ).font = label_font
    widths = [18, 46, 14, 16, 14, 14]
    for c, w in enumerate(widths, start=1):
        ws2.column_dimensions[get_column_letter(c)].width = w
    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = f"A1:F{last}"

    # ---- Sayfa 3: Şube Detayı ----
    ws3 = wb.create_sheet("Şube Detayı")
    headers3 = ["Kurum Kodu", "Okul Adı", "Şube (Dosya)", "Sayfa Sayısı",
                "Öğrenci Sayısı", "Durum"]
    for c, h in enumerate(headers3, start=1):
        cell = ws3.cell(row=1, column=c, value=h)
        cell.font = head_font
        cell.fill = head_fill
        cell.alignment = center
        cell.border = border
    for i, row in enumerate(branch_rows, start=2):
        values = [row["kod"], row.get("okul_adi", "") or "-", row["dosya"],
                  row["pages"], row["students"], row["durum"]]
        for c, v in enumerate(values, start=1):
            cell = ws3.cell(row=i, column=c, value=v)
            cell.border = border
            cell.alignment = left if c <= 3 else center
    widths3 = [18, 40, 45, 14, 16, 12]
    for c, w in enumerate(widths3, start=1):
        ws3.column_dimensions[get_column_letter(c)].width = w
    ws3.freeze_panes = "A2"
    if branch_rows:
        ws3.auto_filter.ref = f"A1:F{len(branch_rows) + 1}"

    # ---- Sayfa 4: Hatalar (varsa) ----
    if error_rows:
        ws4 = wb.create_sheet("Hatalar")
        headers4 = ["Kurum Kodu", "Tür", "Dosya / Klasör", "Açıklama"]
        for c, h in enumerate(headers4, start=1):
            cell = ws4.cell(row=1, column=c, value=h)
            cell.font = head_font
            cell.fill = head_fill
            cell.alignment = center
            cell.border = border
        for i, row in enumerate(error_rows, start=2):
            values = [row["kod"], row["tur"], row["ad"], row["aciklama"]]
            for c, v in enumerate(values, start=1):
                cell = ws4.cell(row=i, column=c, value=v)
                cell.border = border
                cell.alignment = center if c == 2 else left
        widths4 = [18, 14, 45, 60]
        for c, w in enumerate(widths4, start=1):
            ws4.column_dimensions[get_column_letter(c)].width = w

    wb.save(report_path)


def process_district(district_dir, output_dir, log, progress, finished):
    """İlçe klasörünü tarar, birleştirir ve raporu üretir."""
    try:
        subfolders = []
        for name in os.listdir(district_dir):
            full = os.path.join(district_dir, name)
            if os.path.isdir(full) and name != OUTPUT_DIR_NAME:
                subfolders.append(full)
        subfolders.sort(key=lambda p: natural_key(os.path.basename(p)))

        if not subfolders:
            log("Uyarı: İlçe klasöründe hiç alt klasör bulunamadı.")
            finished(False, None, None, None)
            return

        # Mevcut çıktı klasörünü sil (üzerine yaz)
        if os.path.isdir(output_dir):
            log(f"Eski çıktı klasörü siliniyor: {output_dir}")
            remove_tree(output_dir)
        os.makedirs(output_dir, exist_ok=True)

        school_rows = []
        branch_rows = []
        error_rows = []
        district_branches = 0
        district_students = 0
        il_ilce = ""
        created_school_pdfs = []

        progress(0, len(subfolders))

        for idx, folder in enumerate(subfolders, start=1):
            kod = os.path.basename(folder)
            pdfs = list_pdfs(folder)

            if not pdfs:
                log(f"[UYARI] {kod}: klasörde PDF bulunamadı (atlandı).")
                error_rows.append({"kod": kod, "tur": "Klasör",
                                   "ad": kod,
                                   "aciklama": "PDF bulunamadı"})
                progress(idx, len(subfolders))
                continue

            log(f"[İŞLENİYOR] {kod}: {len(pdfs)} şube PDF'i...")
            output_path = os.path.join(output_dir, f"{kod}.pdf")

            result = merge_folder_pdfs(pdfs, output_path, log)
            okul_adi = result["school_info"]["okul_adi"]
            if il_ilce == "" and result["school_info"]["il_ilce"]:
                il_ilce = result["school_info"]["il_ilce"]

            for name, reason in result["errors"]:
                error_rows.append({"kod": kod, "tur": "Dosya",
                                   "ad": name, "aciklama": reason})

            if result["ok_branches"] == 0:
                log(f"  [!] {kod}"
                    f"{' - ' + okul_adi if okul_adi else ''}: "
                    "hiçbir dosya birleştirilemedi.")
            else:
                log(f"  -> {os.path.basename(output_path)} | "
                    f"{okul_adi or '(okul adı okunamadı)'}")
                log(f"     {result['ok_branches']}/{len(pdfs)} şube, "
                    f"{result['students']} öğrenci, "
                    f"{result['total_pages']} sayfa")
                created_school_pdfs.append(output_path)

            # Şube detay satırları (orijinal sırada, tekrar okumadan)
            for info in result["branches"]:
                branch_rows.append({
                    "kod": kod, "okul_adi": okul_adi, "dosya": info["dosya"],
                    "pages": info["pages"], "students": info["students"],
                    "durum": info["durum"]})

            district_branches += result["ok_branches"]
            district_students += result["students"]
            school_rows.append({
                "kod": kod,
                "okul_adi": okul_adi,
                "branches": result["ok_branches"],
                "students": result["students"],
                "pages": result["total_pages"],
                "errors": len(result["errors"]),
            })
            progress(idx, len(subfolders))

        # Okul PDF'lerini, okul klasörlerinin doğal sırasını koruyarak birleştir.
        district_pdf_path = os.path.join(output_dir, DISTRICT_PDF_FILE_NAME)
        district_pdf_ok, district_pdf_pages = merge_created_pdfs(
            created_school_pdfs, district_pdf_path, log)
        if district_pdf_ok:
            log(f"İlçe PDF'si: {district_pdf_path}")
            log(f"İlçe PDF'si toplam sayfa: {district_pdf_pages}")
        else:
            district_pdf_path = None

        # Excel raporu
        excel_path = os.path.join(output_dir, EXCEL_FILE_NAME)
        if Workbook is not None:
            write_excel(excel_path, district_dir, school_rows, branch_rows,
                        error_rows, il_ilce)
        else:
            excel_path = None

        # Metin raporu
        lines = []
        lines.append("=" * 78)
        lines.append("İLÇE PDF BİRLEŞTİRME RAPORU")
        lines.append("=" * 78)
        lines.append(f"İlçe klasörü : {district_dir}")
        if il_ilce:
            lines.append(f"İl / İlçe    : {il_ilce}")
        lines.append(f"Oluşturma    : "
                     f"{datetime.now().strftime('%d.%m.%Y %H:%M')}")
        lines.append("")
        lines.append(f"Toplam okul sayısı   : {len(school_rows)}")
        lines.append(f"Toplam şube sayısı   : {district_branches}")
        lines.append(f"Toplam öğrenci sayısı: {district_students}")
        lines.append(f"Toplam hatalı kayıt  : {len(error_rows)}")
        lines.append("")
        lines.append("-" * 78)
        lines.append(f"{'Kurum Kodu':<14}{'Okul Adı':<36}"
                     f"{'Şube':>6}{'Öğrenci':>9}{'Sayfa':>8}")
        lines.append("-" * 78)
        for row in school_rows:
            ad = (row["okul_adi"] or "-")[:34]
            lines.append(f"{row['kod']:<14}{ad:<36}"
                         f"{row['branches']:>6}{row['students']:>9}"
                         f"{row['pages']:>8}")
        lines.append("-" * 78)
        report_text = "\n".join(lines)
        text_path = os.path.join(output_dir, TEXT_FILE_NAME)
        with open(text_path, "w", encoding="utf-8") as fh:
            fh.write(report_text)

        log("")
        log(report_text)
        log("")
        if excel_path:
            log(f"Excel raporu: {excel_path}")
        if district_pdf_path:
            log(f"Tek ilçe PDF'si: {district_pdf_path}")
        log(f"Metin raporu: {text_path}")

        finished(True, report_text, excel_path, district_pdf_path)
    except Exception:
        log("Beklenmeyen hata:\n" + traceback.format_exc())
        finished(False, None, None, None)


class App:
    def __init__(self, root):
        self.root = root
        self.district_dir = tk.StringVar()
        self.msg_queue = queue.Queue()
        self.running = False

        root.title("İlçe PDF Birleştirici")
        root.geometry("860x640")
        root.minsize(720, 520)

        top = ttk.Frame(root, padding=10)
        top.pack(fill="x")

        ttk.Label(top, text="İlçe klasörü:").pack(side="left")
        ttk.Entry(top, textvariable=self.district_dir).pack(
            side="left", fill="x", expand=True, padx=6)
        ttk.Button(top, text="Klasör Seç", command=self.choose_folder).pack(
            side="left")
        self.run_btn = ttk.Button(top, text="Birleştir ve Raporla",
                                  command=self.start)
        self.run_btn.pack(side="left", padx=6)

        self.progress = ttk.Progressbar(root, mode="determinate")
        self.progress.pack(fill="x", padx=10)

        body = ttk.Frame(root, padding=10)
        body.pack(fill="both", expand=True)

        ttk.Label(body, text="İşlem Günlüğü / Rapor").pack(anchor="w")
        text_frame = ttk.Frame(body)
        text_frame.pack(fill="both", expand=True)

        self.text = tk.Text(text_frame, wrap="none", font=("Consolas", 10))
        scroll_y = ttk.Scrollbar(text_frame, orient="vertical",
                                 command=self.text.yview)
        scroll_x = ttk.Scrollbar(text_frame, orient="horizontal",
                                 command=self.text.xview)
        self.text.configure(yscrollcommand=scroll_y.set,
                            xscrollcommand=scroll_x.set)
        self.text.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        self.status = ttk.Label(root, text="Hazır", anchor="w",
                                relief="sunken", padding=4)
        self.status.pack(fill="x", side="bottom")

        # Sürekli çalışan kuyruk dinleyicisi (loglar akışkan gösterilir)
        self.root.after(80, self.poll_queue)

    def choose_folder(self):
        folder = filedialog.askdirectory(title="İlçe klasörünü seçin")
        if folder:
            self.district_dir.set(os.path.normpath(folder))

    def log(self, message):
        self.msg_queue.put(("log", message))

    def progress_cb(self, value, maximum):
        self.msg_queue.put(("progress", value, maximum))

    def finished(self, ok, report_text, excel_path, district_pdf_path):
        self.msg_queue.put(("done", ok, excel_path, district_pdf_path))

    def start(self):
        district = self.district_dir.get().strip()
        if not district or not os.path.isdir(district):
            messagebox.showerror("Hata", "Lütfen geçerli bir ilçe klasörü seçin.")
            return
        if PdfReader is None:
            messagebox.showerror(
                "Eksik paket",
                "pypdf kurulu değil.\n\nKurmak için:\n  pip install pypdf")
            return

        output_dir = os.path.join(district, OUTPUT_DIR_NAME)
        if os.path.isdir(output_dir):
            cevap = messagebox.askyesno(
                "Uyarı",
                f"Seçilen klasörde zaten bir çıktı klasörü var:\n\n"
                f"{output_dir}\n\n"
                "Bu klasör silinip üzerine yeniden yazılacak. Devam edelim mi?")
            if not cevap:
                return

        self.text.delete("1.0", "end")
        self.run_btn.config(state="disabled")
        self.running = True
        self.progress.config(value=0, maximum=1)
        self.status.config(text="İşleniyor...")

        threading.Thread(
            target=process_district,
            args=(district, output_dir, self.log, self.progress_cb,
                  self.finished),
            daemon=True,
        ).start()

    def poll_queue(self):
        drained = False
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                kind = msg[0]
                if kind == "log":
                    self.text.insert("end", msg[1] + "\n")
                    drained = True
                elif kind == "progress":
                    value, maximum = msg[1], msg[2]
                    self.progress.config(maximum=max(maximum, 1), value=value)
                elif kind == "done":
                    ok, excel_path, district_pdf_path = msg[1], msg[2], msg[3]
                    self.running = False
                    self.run_btn.config(state="normal")
                    if ok:
                        self.progress.config(value=self.progress["maximum"])
                        self.status.config(text="Tamamlandı")
                        messagebox.showinfo(
                            "Bitti",
                            "İşlem tamamlandı.\n\nÇıktı klasörü:\n"
                            f"{os.path.join(self.district_dir.get(), OUTPUT_DIR_NAME)}"
                            + (f"\n\nExcel: {os.path.basename(excel_path)}"
                               if excel_path else "")
                            + (f"\nİlçe PDF'si: "
                               f"{os.path.basename(district_pdf_path)}"
                               if district_pdf_path else ""))
                    else:
                        self.status.config(text="Hata ile bitti")
                        messagebox.showerror(
                            "Hata", "İşlem tamamlanamadı. Günlüğe bakın.")
        except queue.Empty:
            pass
        if drained:
            self.text.see("end")
            self.text.update_idletasks()
        self.root.after(80, self.poll_queue)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
