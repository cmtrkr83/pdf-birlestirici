#!/usr/bin/env python3
"""Ilce klasorlerindeki okul PDF'lerini kurum kodu bazinda birlestirir ve rapor uretir.

Yapi:  <ilce>/<kurum_kodu>/<kurum_kodu>-<sinif>-<program>-<sube>.pdf
Cikti: <ilce>/birlestirilmis/<kurum_kodu>.pdf  +  rapor.xlsx

Sayfalar pypdf ile ham olarak eklenir; yeniden sikistirma, olceklendirme veya
gorsuntu isleme yapilmaz. Kaynak dosyalara dokunulmaz.
"""

import argparse
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path

try:
    import pdfplumber
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from pypdf import PdfReader, PdfWriter
except ImportError:
    sys.exit(
        "Eksik bagimlilik. Kurulum:\n"
        "  python3 -m venv .venv && .venv/bin/pip install pypdf pdfplumber openpyxl\n"
        "Sonra betigi .venv/bin/python ile calistirin."
    )

CIKTI_KLASOR = "birlestirilmis"
RAPOR_ADI = "rapor.xlsx"
KURUM_KODU_DESEN = re.compile(r"^\d{4,10}$")
HATA = "HATA"


@dataclass
class Sube:
    dosya: Path
    sinif_adi: str
    sube_adi: str
    ogrenciler: int
    sayfa: int
    okul_adi: str = ""
    not_: str = ""


@dataclass
class Okul:
    kod: str
    klasor: Path
    subeler: list[Sube] = field(default_factory=list)
    birlesik_yol: Path | None = None
    birlesik_sayfa: int = 0
    hata: str = ""


# ---------------------------------------------------------------- dosya adlari

def dosya_adini_ayir(pdf: Path, kod: str):
    """'708752-6. Sinif-Orta-Agir Zihinsel-A Sube.pdf' -> ('6. Sinif', 'A')"""
    ad = pdf.stem
    if ad.lower().endswith(".pdf"):
        ad = ad[:-4]
    if ad.startswith(kod):
        ad = ad[len(kod):]
    parcalar = [p.strip() for p in ad.split("-") if p.strip()]
    if not parcalar:
        return ad or "Bilinmiyor", "?"
    sube = parcalar[-1]
    sinif_adi = "-".join(parcalar[:-1])
    return (sinif_adi or "Bilinmiyor"), (sube or "?")


def sira_anahtari(sube_adi: str):
    """'A Sube' -> 1, 'B Sube' -> 2 ... alfabetik olmayanlar sona."""
    m = re.match(r"^([A-Za-zÇĞİÖŞÜçğıöşü])", sube_adi.strip())
    if m:
        return (0, m.group(1).upper())
    return (1, sube_adi)


# ------------------------------------------------------------------- okuma

YOKLAMA_BASLIK = "Sınıf Yoklama Listesi"
OKUL_ADI_DESEN = re.compile(r"OKUL\s*ADI\s*:?\s*(.+)")
SINIF_DESEN = re.compile(r"SINIFI\s*/\s*ŞUBESİ\s*:?\s*(.+)")


def yoklama_sayfalari(pdf) -> list:
    """Yoklama listesi sayfalarini dondurur (liste birden fazla sayfaya tasabilir)."""
    sayfalar = []
    for sayfa in pdf.pages:
        metin = sayfa.extract_text() or ""
        if YOKLAMA_BASLIK not in metin:
            break
        sayfalar.append(sayfa)
    return sayfalar


def sn_ogrenci(pdf) -> int:
    """Yoklama listesi sayfalarindaki SN kolonundan en buyuk ogrenci sayisi."""
    enbuyuk = 0
    for sayfa in yoklama_sayfalari(pdf):
        limit = sayfa.width * 0.06
        for w in sayfa.extract_words():
            if w["x0"] < limit and w["text"].isdigit() and float(w["top"]) > 100:
                enbuyuk = max(enbuyuk, int(w["text"]))
    return enbuyuk


def alan_oku(pdf, desen):
    for sayfa in yoklama_sayfalari(pdf):
        for satir in (sayfa.extract_text() or "").splitlines():
            m = desen.search(satir)
            if m:
                return m.group(1).strip()
    return ""


def sube_oku(pdf: Path, kod: str) -> Sube:
    sinif_adi, sube_adi = dosya_adini_ayir(pdf, kod)
    with pdfplumber.open(str(pdf)) as d:
        sayfa = len(d.pages)
        sn = sn_ogrenci(d)
        okul_adi = alan_oku(d, OKUL_ADI_DESEN)
        sinif_pdf = alan_oku(d, SINIF_DESEN)

    # PDF'teki sinif/sube, dosya adindan daha guvenilir
    if sinif_pdf and "/" in sinif_pdf:
        parcalar = [p.strip() for p in sinif_pdf.split("/")]
        sinif_adi = parcalar[0] or sinif_adi
        sube_adi = parcalar[1] or sube_adi

    formül = (sayfa - 2) // 2 if sayfa > 2 else 0
    notlar = []
    # Yapi sabit: 1-2. sayfa yoklama listesi, 3. sayfadan itibaren her ogrenciye
    # 2 sayfalik (onlu arka) cevap kagidi. Bu yuzden ogrenci = (sayfa - 2) / 2.
    # SN kolonu bunu dogrulayan bagimsiz bir olcumdur.
    if formül and (sayfa - 2) % 2:
        notlar.append(f"sayfa={sayfa} tek sayida -> yapı tutarsız")
    if sn and formül and sn != formül:
        ogrenciler = sn
        notlar.append(f"SN={sn} sayfa formulu={formül} -> SN kullanildi")
    else:
        ogrenciler = sn or formül
        if not sn:
            notlar.append("SN okunamadi, sayfa formulu kullanildi")

    return Sube(
        dosya=pdf,
        sinif_adi=sinif_adi,
        sube_adi=sube_adi,
        ogrenciler=ogrenciler,
        sayfa=sayfa,
        okul_adi=okul_adi,
        not_="; ".join(notlar),
    )


# ------------------------------------------------------------------- tarama

def okullari_tara(ilce: Path, ilerleme=None) -> list[Okul]:
    """Ilce klasorunu tarar. ilerleme(i, toplam, okul_kodu, sube_no, sube_toplam)."""
    adaylar = [
        k for k in sorted(ilce.iterdir())
        if k.is_dir() and k.name != CIKTI_KLASOR and KURUM_KODU_DESEN.match(k.name)
        and any(k.glob("*.pdf"))
    ]
    okullar = []
    for i, klasor in enumerate(adaylar):
        okul = Okul(kod=klasor.name, klasor=klasor)
        okul.subeler = []
        pdfler = sorted(klasor.glob("*.pdf"))
        for j, pdf in enumerate(pdfler):
            if ilerleme:
                ilerleme(i + 1, len(adaylar), okul.kod, j + 1, len(pdfler))
            try:
                okul.subeler.append(sube_oku(pdf, okul.kod))
            except Exception as e:
                sinif_adi, sube_adi = dosya_adini_ayir(pdf, okul.kod)
                okul.subeler.append(Sube(
                    dosya=pdf,
                    sinif_adi=sinif_adi,
                    sube_adi=sube_adi,
                    ogrenciler=0,
                    sayfa=0,
                    not_=f"{type(e).__name__}: {e}",
                ))
        okul.subeler.sort(key=lambda s: (s.sinif_adi, sira_anahtari(s.sube_adi)))
        okullar.append(okul)
    return okullar


def birlestir(okul: Okul, cikti: Path, dry_run: bool) -> None:
    """Sayfalari ham olarak birlestirir (yeniden kodlama yok)."""
    hedef = cikti / f"{okul.kod}.pdf"
    okul.birlesik_yol = hedef
    if dry_run:
        okul.birlesik_sayfa = sum(s.sayfa for s in okul.subeler)
        return

    writer = PdfWriter()
    toplam = 0
    for sube in okul.subeler:
        if sube.sayfa == 0:
            continue
        reader = PdfReader(str(sube.dosya))
        for page in reader.pages:
            writer.add_page(page)
            toplam += 1
    with hedef.open("wb") as fh:
        writer.write(fh)
    okul.birlesik_sayfa = toplam


# -------------------------------------------------------------------- rapor

BASLIK_FILL = PatternFill("solid", fgColor="1F4E79")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")


def sayfa_yaz(wb, baslik, sutunlar, satirlar, genislikler=None):
    ws = wb.create_sheet(baslik)
    ws.append(sutunlar)
    for hucre in ws[1]:
        hucre.fill = BASLIK_FILL
        hucre.font = BASLIK_YAZI
        hucre.alignment = Alignment(horizontal="center", vertical="center")
    for satir in satirlar:
        ws.append(satir)
    ws.freeze_panes = "A2"
    for i, g in enumerate(genislikler or [], start=1):
        ws.column_dimensions[get_column_letter(i)].width = g
    return ws


def rapor_uret(okullar: list[Okul], ilce: Path, cikti: Path) -> Path:
    wb = Workbook()
    wb.remove(wb.active)

    toplam_okul = len(okullar)
    toplam_sube = sum(len(o.subeler) for o in okullar)
    toplam_ogrenci = sum(s.ogrenciler for o in okullar for s in o.subeler)
    hatali = [o for o in okullar if o.hata]
    sayfa_yaz(
        wb,
        "Ilce Ozeti",
        ["Ilce", "Okul Sayisi", "Sube Sayisi", "Ogrenci Sayisi", "Hatali Okul", "Toplam Sayfa"],
        [[ilce.name, toplam_okul, toplam_sube, toplam_ogrenci, len(hatali),
          sum(o.birlesik_sayfa for o in okullar)]],
        [20, 13, 13, 16, 13, 13],
    )

    sayfa_yaz(
        wb,
        "Okul Detayi",
        ["Kurum Kodu", "Okul Adi", "Sube Sayisi", "Ogrenci Sayisi",
         "Birlestirilen Sayfa", "Dosya", "Durum", "Not"],
        [
            [
                o.kod,
                next((s.okul_adi for s in o.subeler if s.okul_adi), ""),
                len(o.subeler),
                sum(s.ogrenciler for s in o.subeler),
                o.birlesik_sayfa,
                str(o.birlesik_yol) if o.birlesik_yol else "",
                HATA if o.hata else "OK",
                o.hata,
            ]
            for o in okullar
        ],
        [12, 45, 12, 15, 18, 60, 9, 40],
    )

    satirlar = []
    for o in okullar:
        for s in o.subeler:
            satirlar.append([
                o.kod,
                next((x.okul_adi for x in o.subeler if x.okul_adi), ""),
                s.sinif_adi,
                s.sube_adi,
                s.ogrenciler,
                s.sayfa,
                HATA if s.not_ and not s.sayfa else ("UYARI" if s.not_ else "OK"),
                s.not_,
            ])
    sayfa_yaz(
        wb,
        "Sube Detayi",
        ["Kurum Kodu", "Okul Adi", "Sinif", "Sube", "Ogrenci", "Sayfa", "Durum", "Not"],
        satirlar,
        [12, 45, 32, 12, 10, 9, 9, 45],
    )

    if hatali:
        sayfa_yaz(
            wb,
            "Hatalar",
            ["Kurum Kodu", "Dosya", "Hata"],
            [[o.kod, str(o.klasor), o.hata] for o in hatali],
            [12, 45, 80],
        )

    hedef = cikti / RAPOR_ADI
    wb.save(hedef)
    return hedef


def konsol_ozeti(okullar: list[Okul], ilce: Path, dry_run: bool) -> None:
    toplam_sube = sum(len(o.subeler) for o in okullar)
    toplam_ogrenci = sum(s.ogrenciler for o in okullar for s in o.subeler)
    toplam_sayfa = sum(o.birlesik_sayfa for o in okullar)
    fiil = "(kuru calisma)" if dry_run else ""

    print(f"\nIlce        : {ilce.name} {fiil}")
    print(f"Okul sayisi : {len(okullar)}")
    print(f"Sube sayisi : {toplam_sube}")
    print(f"Ogrenci     : {toplam_ogrenci}")
    print(f"Toplam sayfa: {toplam_sayfa}")
    print("-" * 46)

    for o in okullar:
        ad = next((s.okul_adi for s in o.subeler if s.okul_adi), "")
        print(f"{o.kod}  {len(o.subeler):>2} sube  "
              f"{sum(s.ogrenciler for s in o.subeler):>3} ogrenci  {ad}")

    uyarilar = [(o.kod, s) for o in okullar for s in o.subeler if s.not_]
    if uyarilar:
        print("-" * 46)
        print(f"UYARILAR ({len(uyarilar)}):")
        for kod, s in uyarilar:
            print(f"  {kod} / {s.sinif_adi} {s.sube_adi}: {s.not_}")


# ------------------------------------------------------------------- is akisi

def is_yap(ilce: Path, dry_run=False, birlestirilecek=True, ilerleme=None) -> dict:
    """Tum akisi calistirir. CLI ve GUI ortak kullanir.

    ilerleme: cagrilanabilir (asama, yuzde, metin)
    Doner: {'okullar', 'rapor', 'cikti', 'toplam_ogrenci', 'hatali'}
    """
    def bildir(asama, yuzde, metin):
        if ilerleme:
            ilerleme(asama, yuzde, metin)

    ilce = Path(ilce).expanduser().resolve()
    if not ilce.is_dir():
        raise NotADirectoryError(f"Klasor bulunamadi: {ilce}")

    cikti = ilce / CIKTI_KLASOR

    # 1) tarama + ogrenci sayilari (en yavas asama, %0-45)
    def tarama_ilerleme(i, toplam_okul, kod, sube_no, sube_toplam):
        oran = (i - 1) / max(toplam_okul, 1)
        bildir("tarama", int(oran * 45),
               f"Okul {i}/{toplam_okul} · {kod} · sube {sube_no}/{sube_toplam}")

    bildir("tarama", 0, f"Taranıyor: {ilce.name}")
    okullar = okullari_tara(ilce, ilerleme=tarama_ilerleme)
    if not okullar:
        raise ValueError(
            "Bu klasörde kurum kodu içeren PDF klasörü bulunamadı.\n"
            "Beklenen yapı:  ilçe klasörü / okul kodu klasörü / PDF dosyaları"
        )

    cikti.mkdir(exist_ok=True)

    # 2) birlestirme (%45-95)
    for i, okul in enumerate(okullar, 1):
        bildir("birlestirme", 45 + int((i - 1) / len(okullar) * 50),
               f"Birleştiriliyor: {okul.kod} ({len(okul.subeler)} şube)")
        try:
            if birlestirilecek:
                birlestir(okul, cikti, dry_run)
            else:
                okul.birlesik_yol = cikti / f"{okul.kod}.pdf"
                okul.birlesik_sayfa = sum(s.sayfa for s in okul.subeler)
        except Exception as e:
            okul.hata = f"Birlestirme hatasi: {type(e).__name__}: {e}"

    # 3) rapor (%95-100)
    bildir("rapor", 95, "Rapor yazılıyor...")
    rapor = rapor_uret(okullar, ilce, cikti)

    hatali = [o for o in okullar if o.hata]
    bildir("bitti", 100, "Tamamlandı")
    return {
        "okullar": okullar,
        "rapor": rapor,
        "cikti": cikti,
        "toplam_okul": len(okullar),
        "toplam_sube": sum(len(o.subeler) for o in okullar),
        "toplam_ogrenci": sum(s.ogrenciler for o in okullar for s in o.subeler),
        "toplam_sayfa": sum(o.birlesik_sayfa for o in okullar),
        "hatali": hatali,
        "kuru": dry_run or not birlestirilecek,
    }


# ---------------------------------------------------------------------- main

def main() -> int:
    ap = argparse.ArgumentParser(
        description="Ilce klasorundeki okul PDF'lerini birlestirir, rapor uretir."
    )
    ap.add_argument("ilce", type=Path, help="Ilce klasoru (ornek: ~/Desktop/CK_41-8)")
    ap.add_argument("--dry-run", action="store_true",
                    help="Birlestirme yapmadan yalnizca sayilari hesaplar ve rapor yazar")
    ap.add_argument("--rapor-only", action="store_true",
                    help="Yalnizca rapor uretir, birlestirme yapmaz")
    args = ap.parse_args()

    ilce = args.ilce.expanduser().resolve()
    if not ilce.is_dir():
        sys.exit(f"Hata: klasor bulunamadi -> {ilce}")

    def ilerleme(asama, yuzde, metin):
        print(f"\r  {yuzde:>3}%  {metin}", end="", flush=True)

    try:
        sonuc = is_yap(
            ilce,
            dry_run=args.dry_run,
            birlestirilecek=not args.rapor_only,
            ilerleme=ilerleme,
        )
    except (NotADirectoryError, ValueError) as e:
        sys.exit(f"Hata: {e}")
    print(" " * 70, end="\r")

    okullar = sonuc["okullar"]
    konsol_ozeti(okullar, ilce, sonuc["kuru"])

    if sonuc["hatali"]:
        print("-" * 46)
        print(f"HATALI OKULLAR ({len(sonuc['hatali'])}):")
        for o in sonuc["hatali"]:
            print(f"  {o.kod}: {o.hata}")

    print("-" * 46)
    print(f"Rapor: {sonuc['rapor']}")
    if not sonuc["kuru"]:
        print(f"Birlestirilenler: {sonuc['cikti']}")
    return 1 if sonuc["hatali"] else 0


if __name__ == "__main__":
    sys.exit(main())