#!/usr/bin/env python3
"""PDF Birlestirici - Windows masaustu arayuzu.

Ilce klasorunu secer, her okulun PDF'lerini kurum kodu adiyla tek dosyada
birlestirir ve ilce/okul/sube/ogrenci raporu uretir.
"""

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

# PyInstaller tek dosya modunda gomulu modulun yolunu bul
if getattr(sys, "frozen", False):
    sys.path.insert(0, os.path.dirname(sys.executable))
else:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pdf_birlestir import is_yap  # noqa: E402

UYGULAMA = "PDF Birlestirici"
ASAMA_ADI = {"tarama": "PDF'ler taranıyor", "birlestirme": "Birleştiriliyor",
             "rapor": "Rapor yazılıyor", "bitti": "Tamamlandı"}


class Uygulama(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(UYGULAMA)
        self.geometry("720x620")
        self.minsize(640, 560)
        self.configure(bg="#f4f6f8")

        self.klasor_var = tk.StringVar()
        self.dry_run_var = tk.BooleanVar(value=False)
        self.ozet_var = tk.StringVar(value="Henüz işlem yapılmadı.")

        self.kuyruk: queue.Queue = queue.Queue()
        self.calisiyor = False
        self._ac_butonlari: list = []
        self._son_gecen: str | None = None

        self._arayuz_kur()

    # ------------------------------------------------------------- arayuz
    def _arayuz_kur(self):
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass

        baslik = tk.Label(self, text=UYGULAMA, font=("Segoe UI", 20, "bold"),
                          bg="#f4f6f8", fg="#1f2933")
        baslik.pack(pady=(22, 2))
        tk.Label(self, text="Okul PDF'lerini tek dosyada birleştirir ve rapor üretir",
                 font=("Segoe UI", 10), bg="#f4f6f8", fg="#616e7c").pack(pady=(0, 18))

        # klasor secimi
        kutu = tk.LabelFrame(self, text=" Kaynak Klasör ", font=("Segoe UI", 10, "bold"),
                             bg="#f4f6f8", fg="#1f2933", padx=14, pady=14)
        kutu.pack(fill="x", padx=34)

        satir = tk.Frame(kutu, bg="#f4f6f8")
        satir.pack(fill="x")
        self.girdi = tk.Entry(satir, textvariable=self.klasor_var,
                              font=("Segoe UI", 10), relief="solid", bd=1)
        self.girdi.pack(side="left", fill="x", expand=True, ipady=5)
        self.tara_btn = ttk.Button(satir, text="Klasör Se…", command=self.klasor_sec,
                                   width=14)
        self.tara_btn.pack(side="left", padx=(10, 0))

        # secenekler
        self.dry_check = ttk.Checkbutton(
            kutu, text="Sadece say (birleştirme yapma)", variable=self.dry_run_var)
        self.dry_check.pack(anchor="w", pady=(12, 0))

        # buton
        self.basla_btn = ttk.Button(self, text="Birleştirmeyi Başlat",
                                    command=self.baslat, width=26)
        self.basla_btn.pack(pady=(22, 0))
        self.basla_btn.configure(state="disabled")

        # ilerleme
        self.yuzde_var = tk.DoubleVar(value=0)
        self.asama_var = tk.StringVar(value="Hazır")
        self.progress = ttk.Progressbar(self, variable=self.yuzde_var, maximum=100,
                                        mode="determinate")
        self.progress.pack(fill="x", padx=34, pady=(20, 6))
        self.asama_etiket = tk.Label(self, textvariable=self.asama_var,
                                     font=("Segoe UI", 10, "bold"),
                                     bg="#f4f6f8", fg="#1f2933")
        self.asama_etiket.pack()

        # ozet
        self.ozet_etiket = tk.Label(self, textvariable=self.ozet_var,
                                    font=("Segoe UI", 10), bg="#f4f6f8",
                                    fg="#3e4c59", justify="center")
        self.ozet_etiket.pack(pady=(14, 8))

        # gunluk
        gunluk_kutu = tk.LabelFrame(self, text=" Ayrıntı ", font=("Segoe UI", 10, "bold"),
                                   bg="#f4f6f8", fg="#1f2933", padx=8, pady=8)
        gunluk_kutu.pack(fill="both", expand=True, padx=34, pady=(8, 24))

        self.gunluk = tk.Text(gunluk_kutu, font=("Consolas", 9), height=12,
                              wrap="none", bg="white", relief="solid", bd=1,
                              state="disabled")
        kaydirma = ttk.Scrollbar(gunluk_kutu, orient="vertical", command=self.gunluk.yview)
        self.gunluk.configure(yscrollcommand=kaydirma.set)
        self.gunluk.pack(side="left", fill="both", expand=True)
        kaydirma.pack(side="right", fill="y")

    # ------------------------------------------------------------ eylemler
    def klasor_sec(self):
        yol = filedialog.askdirectory(title="İlçe klasörünü seçin", mustexist=True)
        if yol:
            self.klasor_var.set(yol)
            self.girdi.icursor(0)
            self.girdi.xview_moveto(1)
            self.basla_btn.configure(state="normal")
            self.ozet_var.set(f"{Path(yol).name} seçildi. Başlat'a basabilirsiniz.")

    def _gunluk(self, metin):
        self.gunluk.configure(state="normal")
        self.gunluk.insert("end", metin + "\n")
        self.gunluk.see("end")
        self.gunluk.configure(state="disabled")

    def baslat(self):
        if self.calisiyor:
            return
        yol = self.klasor_var.get().strip()
        if not yol or not Path(yol).is_dir():
            messagebox.showerror("Hata", "Geçerli bir klasör seçin.")
            return

        self.calisiyor = True
        self.basla_btn.configure(state="disabled")
        self.tara_btn.configure(state="disabled")
        self.dry_check.configure(state="disabled")
        self.girdi.configure(state="disabled")
        self.gunluk.configure(state="normal")
        self.gunluk.delete("1.0", "end")
        self.gunluk.configure(state="disabled")
        self.yuzde_var.set(0)
        self._ac_butonlari_temizle()
        self._son_gecen = None

        threading.Thread(target=self._calistir, args=(yol,), daemon=True).start()
        self.after(80, self._kuyruk_drenle)

    def _calistir(self, yol):
        try:
            sonuc = is_yap(
                yol,
                dry_run=self.dry_run_var.get(),
                birlestirilecek=True,
                ilerleme=lambda a, y, m: self.kuyruk.put(("ilerleme", a, y, m)),
            )
            self.kuyruk.put(("bitti", sonuc, None))
        except Exception as e:
            self.kuyruk.put(("hata", f"{type(e).__name__}: {e}", None))

    def _kuyruk_drenle(self):
        dolu = True
        while dolu:
            dolu = False
            try:
                tur = self.kuyruk.get_nowait()
            except queue.Empty:
                break
            dolu = True
            tip = tur[0]

            if tip == "ilerleme":
                _, asama, yuzde, metin = tur
                self.yuzde_var.set(yuzde)
                self.asama_var.set(f"{ASAMA_ADI.get(asama, asama)} — %{yuzde}")
                if metin != self._son_gecen:
                    self._son_gecen = metin
                    self._gunluk(metin)
            elif tip == "bitti":
                self._bitti(tur[1])
            elif tip == "hata":
                self._hata(tur[1])

        if self.calisiyor:
            self.after(80, self._kuyruk_drenle)

    def _bitti(self, sonuc):
        self.calisiyor = False
        self.basla_btn.configure(state="normal")
        self.tara_btn.configure(state="normal")
        self.dry_check.configure(state="normal")
        self.girdi.configure(state="normal")
        self.yuzde_var.set(100)
        self.asama_var.set("Tamamlandı")

        for o in sonuc["okullar"]:
            ad = next((s.okul_adi for s in o.subeler if s.okul_adi), "")
            self._gunluk(
                f"{o.kod}  {len(o.subeler):>2} şube  "
                f"{sum(s.ogrenciler for s in o.subeler):>3} öğrenci  {ad}"
            )
        for o in sonuc["hatali"]:
            self._gunluk(f"HATA  {o.kod}: {o.hata}")
        uyarilar = [(o.kod, s) for o in sonuc["okullar"] for s in o.subeler if s.not_]
        for kod, s in uyarilar:
            self._gunluk(f"UYARI {kod} / {s.sinif_adi} {s.sube_adi}: {s.not_}")

        self._ac_butonlari_kur(sonuc)

        efsane = (f"{sonuc['toplam_okul']} okul  ·  {sonuc['toplam_sube']} şube  ·  "
                  f"{sonuc['toplam_ogrenci']} öğrenci")
        if sonuc["hatali"]:
            efsane += f"  ·  {len(sonuc['hatali'])} hatalı okul"
        self.ozet_var.set(efsane)

        if sonuc["kuru"]:
            mesaj = (f"{sonuc['toplam_okul']} okul tarandı.\n"
                     f"{sonuc['toplam_sube']} şube · {sonuc['toplam_ogrenci']} öğrenci\n\n"
                     f"(Kuru çalışma — birleştirme yapılmadı)")
        else:
            mesaj = (f"{efsane}\n\n"
                     f"Birleştirilen PDF'ler:\n{sonuc['cikti']}\n\n"
                     f"Rapor:\n{sonuc['rapor']}")
        messagebox.showinfo("İşlem tamamlandı", mesaj)

    def _hata(self, mesaj):
        self.calisiyor = False
        self.basla_btn.configure(state="normal")
        self.tara_btn.configure(state="normal")
        self.dry_check.configure(state="normal")
        self.girdi.configure(state="normal")
        self.asama_var.set("Hata oluştu")
        self.ozet_var.set("İşlem başarısız — ayrıntılar aşağıda.")
        self._gunluk(f"HATA: {mesaj}")
        messagebox.showerror("Hata", mesaj)

    def _ac_butonlari_temizle(self):
        for btn in self._ac_butonlari:
            btn.destroy()
        self._ac_butonlari = []

    def _ac_butonlari_kur(self, sonuc):
        self._ac_butonlari_temizle()
        secenekler = [("Raporu aç", sonuc["rapor"])]
        if not sonuc["kuru"]:
            secenekler.insert(0, ("Çıktı klasörünü aç", sonuc["cikti"]))
        for ad, yol in secenekler:
            btn = ttk.Button(self, text=ad, command=lambda y=yol: self._ac(y))
            btn.pack(pady=3)
            self._ac_butonlari.append(btn)

    @staticmethod
    def _ac(yol):
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(yol))  # noqa: S606
            elif sys.platform == "darwin":
                subprocess.run(["open", str(yol)], check=False)
            else:
                subprocess.run(["xdg-open", str(yol)], check=False)
        except Exception:
            messagebox.showinfo("Klasör", str(yol))

    def kapat(self):
        if self.calisiyor:
            if not messagebox.askyesno("Çıkış", "İşlem sürüyor. Yine de çıkılsın mı?"):
                return
        self.destroy()


if __name__ == "__main__":
    Uygulama().mainloop()