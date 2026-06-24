import customtkinter as ctk
import threading
import json
import os
import sys
import time
import base64
import tempfile
import sqlite3
import uuid
from datetime import datetime, date
from calendar import monthrange
from tkinter import messagebox, ttk
import tkinter as tk

try:
    from PIL import Image, ImageTk
    PIL_OK = True
except ImportError:
    PIL_OK = False

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException, StaleElementReferenceException, TimeoutException
)

try:
    import win32com.client
    OUTLOOK_OK = True
except ImportError:
    OUTLOOK_OK = False

# ── Tema ────────────────────────────────────────────────────────────────────
RENK_ANA_ARKA = "#071320"
RENK_PANEL    = "#0d1f30"
RENK_KART     = "#112436"
RENK_VURGU    = "#1a6ea8"
RENK_VURGU2   = "#0f4f7a"
RENK_YESIL    = "#1db954"
RENK_KIRMIZI  = "#e74c3c"
RENK_SARI     = "#f0a500"
RENK_YAZI     = "#e8f0f7"
RENK_YAZI2    = "#8faabf"
RENK_SINIR    = "#1e3448"
RENKLER       = ["#1a6ea8", "#1db954", "#9b59b6", "#e67e22"]

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ── Yollar ──────────────────────────────────────────────────────────────────
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE  = os.path.join(BASE_DIR, "config.json")
DB_FILE      = os.path.join(BASE_DIR, "kayitlar.db")
PROFILE_DIR  = os.path.join(BASE_DIR, "chrome_profile")
os.makedirs(PROFILE_DIR, exist_ok=True)

# ── Config ──────────────────────────────────────────────────────────────────
def load_config():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {"wa_group_name": "", "kurallar": []}

def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

# ── Veritabanı ───────────────────────────────────────────────────────────────
def db_init():
    con = sqlite3.connect(DB_FILE)
    con.execute("""
        CREATE TABLE IF NOT EXISTS kayitlar (
            id       TEXT PRIMARY KEY,
            tarih    TEXT,
            gonderen TEXT,
            kural_ad TEXT,
            mesaj    TEXT,
            resim_var INTEGER DEFAULT 0
        )
    """)
    con.commit()
    con.close()

def db_ekle(gonderen, kural_ad, mesaj, resim_var):
    con = sqlite3.connect(DB_FILE)
    con.execute(
        "INSERT INTO kayitlar VALUES (?,?,?,?,?,?)",
        (str(uuid.uuid4()), datetime.now().isoformat(), gonderen,
         kural_ad, mesaj, 1 if resim_var else 0)
    )
    con.commit()
    con.close()

def db_rapor(yil, ay):
    con = sqlite3.connect(DB_FILE)
    prefix = f"{yil:04d}-{ay:02d}"
    rows = con.execute(
        "SELECT tarih,gonderen,kural_ad,mesaj,resim_var FROM kayitlar "
        "WHERE tarih LIKE ? ORDER BY tarih",
        (prefix + "%",)
    ).fetchall()
    gunluk = {}
    for r in rows:
        gun = r[0][:10]
        gunluk[gun] = gunluk.get(gun, 0) + 1
    con.close()
    return rows, gunluk

# ── Outlook ─────────────────────────────────────────────────────────────────
MAIL_SABLON = """\
Sayın İlgili,

Aşağıda bilgileri paylaşılan araç, {tarih} tarihinde kapı girişinde \
kayıt yapılmamış olup sisteme RED olarak işlenmiştir.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Gönderen  : {gonderen}
Grup      : {grup}
Tarih     : {tarih}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Araç Bilgisi / Açıklama:
{mesaj}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Bu araç kapı giriş kaydı yapılmamış olduğundan \
ilgili birimler tarafından gerekli işlemlerin yapılması gerekmektedir.

Saygılarımızla,
Pregate Araç Kontrol Sistemi
"""

def send_mail(mail_list, kural_ad, gonderen, mesaj, img_paths, grup_adi):
    if not OUTLOOK_OK:
        return False, "pywin32 yüklü değil"
    try:
        tarih = datetime.now().strftime("%d.%m.%Y %H:%M")
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.Subject = f"[KAYIT RED] {kural_ad} – {tarih}"
        mail.Body = MAIL_SABLON.format(
            tarih=tarih,
            gonderen=gonderen,
            grup=grup_adi,
            mesaj=mesaj or "(Mesaj yok)"
        )
        mail.To = "; ".join(mail_list)
        for p in img_paths:
            if os.path.exists(p):
                mail.Attachments.Add(p)
        mail.Send()
        return True, "OK"
    except Exception as e:
        return False, str(e)

# ── WhatsApp Bot ─────────────────────────────────────────────────────────────
class WhatsAppBot:
    def __init__(self, on_log, on_status, on_message):
        self.on_log     = on_log
        self.on_status  = on_status
        self.on_message = on_message
        self.driver     = None
        self.running    = False
        self._seen      = set()
        self._tmp       = tempfile.mkdtemp()

    def start(self):
        opts = Options()
        opts.add_argument(f"--user-data-dir={PROFILE_DIR}")
        opts.add_argument("--profile-directory=WABot")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_argument("--disable-notifications")
        try:
            self.driver = webdriver.Chrome(options=opts)
            self.driver.get("https://web.whatsapp.com")
            self.on_log("🌐 WhatsApp Web açılıyor…")
            self.on_status("QR", "● QR Bekleniyor", RENK_SARI)
            self.on_log("📷 Chrome penceresinde QR taratın (maks 3 dk)…")
            self.running = True
            self._bekle_giris()
        except Exception as e:
            self.on_log(f"❌ Chrome hatası: {e}")
            self.running = False
            self.on_status("ERR", "● Hata", RENK_KIRMIZI)

    def _bekle_giris(self):
        try:
            WebDriverWait(self.driver, 180).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, '[data-testid="chat-list"]'))
            )
            self.on_status("OK", "● Bağlandı ✓", RENK_YESIL)
            self.on_log("✅ WhatsApp bağlantısı kuruldu!")
            self._grup_ac()
        except TimeoutException:
            self.on_log("⏱ QR zaman aşımı. Lütfen tekrar başlatın.")
            self.running = False
            self.on_status("ERR", "● Zaman Aşımı", RENK_KIRMIZI)

    def _grup_ac(self):
        cfg = load_config()
        grup = cfg.get("wa_group_name", "").strip()
        if not grup:
            self.on_log("⚠ Grup adı girilmemiş!")
            return
        self.on_log(f"🔍 '{grup}' grubu aranıyor…")
        try:
            # Arama kutusunu bul
            sb = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, '[data-testid="chat-list-search"]'))
            )
            sb.click(); time.sleep(0.5)
            sb.send_keys(grup); time.sleep(2)
            # İlk sonucu tıkla
            first = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, '[data-testid="cell-frame-container"]'))
            )
            first.click(); time.sleep(2)
            self.on_log(f"📌 '{grup}' grubuna girildi. Mesajlar izleniyor…")
            self._dinle()
        except Exception as e:
            self.on_log(f"❌ Grup açılamadı: {e}")

    def _dinle(self):
        while self.running:
            try:
                self._tara()
            except Exception as e:
                self.on_log(f"⚠ Tarama hatası: {e}")
            time.sleep(4)

    def _tara(self):
        msgs = self.driver.find_elements(
            By.CSS_SELECTOR, '[data-testid="msg-container"]')
        for msg in msgs[-15:]:
            try:
                mid = msg.get_attribute("data-id") or msg.id
                if mid in self._seen:
                    continue
                self._seen.add(mid)

                # Gönderen adı
                gonderen = ""
                try:
                    gonderen = msg.find_element(
                        By.CSS_SELECTOR, 'span[aria-label]'
                    ).get_attribute("aria-label") or ""
                except:
                    pass
                try:
                    gonderen = gonderen or msg.find_element(
                        By.CSS_SELECTOR, '[data-testid="msg-meta"] span'
                    ).text
                except:
                    pass

                # Metin
                text = ""
                try:
                    text = msg.find_element(
                        By.CSS_SELECTOR, '[data-testid="msg-text"]').text
                except:
                    pass

                # Resim
                img_paths = []
                try:
                    imgs = msg.find_elements(By.CSS_SELECTOR, 'img[src^="blob:"]')
                    for img_el in imgs:
                        src = img_el.get_attribute("src")
                        if src:
                            data = self.driver.execute_script("""
                                const url=arguments[0];
                                return new Promise(res=>{
                                    fetch(url).then(r=>r.blob()).then(b=>{
                                        const fr=new FileReader();
                                        fr.onload=()=>res(fr.result);
                                        fr.readAsDataURL(b);
                                    });
                                });
                            """, src)
                            if data and "," in data:
                                raw = base64.b64decode(data.split(",")[1])
                                p = os.path.join(self._tmp, f"img_{int(time.time()*1000)}.jpg")
                                with open(p, "wb") as f:
                                    f.write(raw)
                                img_paths.append(p)
                except:
                    pass

                if text or img_paths:
                    self.on_message(gonderen, text, img_paths)

            except StaleElementReferenceException:
                continue
            except:
                continue

    def stop(self):
        self.running = False
        try:
            if self.driver:
                self.driver.quit()
        except:
            pass
        self.driver = None


# ══════════════════════════════════════════════════════════════════════════════
# AYARLAR PENCERESİ
# ══════════════════════════════════════════════════════════════════════════════
class AyarlarPencere(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Ayarlar")
        self.geometry("720x620")
        self.configure(fg_color=RENK_ANA_ARKA)
        self.grab_set()
        self._kart_list = []
        self._build()
        self._render()

    def _build(self):
        # Başlık
        hdr = ctk.CTkFrame(self, fg_color=RENK_PANEL, height=50, corner_radius=0)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="⚙  Ayarlar",
                     font=("Segoe UI", 16, "bold"),
                     text_color=RENK_YAZI).pack(side="left", padx=16, pady=10)

        # Grup adı
        gf = ctk.CTkFrame(self, fg_color=RENK_PANEL, corner_radius=8)
        gf.pack(fill="x", padx=14, pady=(12, 6))
        ctk.CTkLabel(gf, text="WhatsApp Grup Adı:",
                     text_color=RENK_YAZI2, font=("Segoe UI", 11)
                     ).pack(anchor="w", padx=12, pady=(10, 2))
        self.ent_grup = ctk.CTkEntry(gf, fg_color=RENK_KART,
                                      border_color=RENK_SINIR,
                                      text_color=RENK_YAZI,
                                      font=("Segoe UI", 12))
        self.ent_grup.insert(0, load_config().get("wa_group_name", ""))
        self.ent_grup.pack(fill="x", padx=12, pady=(0, 10))

        # Kural başlık
        kh = ctk.CTkFrame(self, fg_color=RENK_ANA_ARKA)
        kh.pack(fill="x", padx=14, pady=(4, 2))
        ctk.CTkLabel(kh, text="KURALLAR", font=("Segoe UI", 11, "bold"),
                     text_color=RENK_YAZI2).pack(side="left")
        ctk.CTkButton(kh, text="+ Yeni Kural", width=110, height=28,
                      fg_color=RENK_VURGU, hover_color=RENK_VURGU2,
                      command=self._kural_ekle).pack(side="right")

        self.scroll = ctk.CTkScrollableFrame(self, fg_color=RENK_ANA_ARKA)
        self.scroll.pack(fill="both", expand=True, padx=14, pady=4)
        self.scroll.columnconfigure(0, weight=1)

        # Kaydet butonu
        ctk.CTkButton(self, text="💾  Kaydet & Kapat",
                      fg_color=RENK_YESIL, hover_color="#17a844",
                      font=("Segoe UI", 13, "bold"), height=40,
                      command=self._kaydet).pack(fill="x", padx=14, pady=(4, 14))

    def _render(self):
        for w in self._kart_list:
            w.destroy()
        self._kart_list.clear()
        cfg = load_config()
        for i, kural in enumerate(cfg.get("kurallar", [])):
            renk = RENKLER[i % len(RENKLER)]
            kart = self._kural_kart(self.scroll, kural, renk, i)
            kart.grid(row=i, column=0, sticky="ew", pady=(0, 10))
            self._kart_list.append(kart)

    def _kural_kart(self, parent, kural, renk, idx):
        frame = ctk.CTkFrame(parent, fg_color=RENK_KART, corner_radius=10)
        frame.columnconfigure(0, weight=1)

        bar = ctk.CTkFrame(frame, fg_color=renk, width=5, corner_radius=0)
        bar.grid(row=0, column=0, rowspan=99, sticky="ns")

        body = ctk.CTkFrame(frame, fg_color=RENK_KART)
        body.grid(row=0, column=1, sticky="ew", padx=10, pady=8)
        body.columnconfigure(1, weight=1)

        # Kural adı
        ctk.CTkLabel(body, text="Kural Adı:", text_color=RENK_YAZI2,
                     font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w")
        ent_ad = ctk.CTkEntry(body, fg_color=RENK_PANEL, border_color=renk,
                               text_color=RENK_YAZI, font=("Segoe UI", 11, "bold"))
        ent_ad.insert(0, kural.get("ad", ""))
        ent_ad.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        ent_ad.bind("<FocusOut>", lambda e, k=kural, w=ent_ad: k.update({"ad": w.get()}))

        # Anahtar kelimeler
        ctk.CTkLabel(body, text="Kelimeler\n(virgülle):", text_color=RENK_YAZI2,
                     font=("Segoe UI", 10)).grid(row=1, column=0, sticky="w", pady=(6, 0))
        ent_kw = ctk.CTkEntry(body, fg_color=RENK_PANEL, border_color=RENK_SINIR,
                               text_color=RENK_YAZI, font=("Segoe UI", 11))
        ent_kw.insert(0, ", ".join(kural.get("keywords", [])))
        ent_kw.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(6, 0))
        ent_kw.bind("<FocusOut>", lambda e, k=kural, w=ent_kw: k.update({
            "keywords": [x.strip().lower() for x in w.get().split(",") if x.strip()]
        }))

        # Mail listesi
        ml_hdr = ctk.CTkFrame(body, fg_color=RENK_KART)
        ml_hdr.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 2))
        ctk.CTkLabel(ml_hdr, text="Mail Listesi:", text_color=RENK_YAZI2,
                     font=("Segoe UI", 10)).pack(side="left")

        ml_frame = ctk.CTkFrame(body, fg_color=RENK_PANEL, corner_radius=6)
        ml_frame.grid(row=3, column=0, columnspan=2, sticky="ew")

        def render_ml():
            for w in ml_frame.winfo_children():
                w.destroy()
            for mail in kural.get("mail_list", []):
                r = ctk.CTkFrame(ml_frame, fg_color=RENK_PANEL)
                r.pack(fill="x", padx=4, pady=2)
                ctk.CTkLabel(r, text=mail, text_color=RENK_YAZI,
                             font=("Segoe UI", 10)).pack(side="left", padx=8)
                ctk.CTkButton(r, text="✕", width=22, height=20,
                              fg_color=RENK_KIRMIZI, hover_color="#a93226",
                              font=("Segoe UI", 10),
                              command=lambda m=mail: (
                                  kural["mail_list"].remove(m), render_ml()
                              )).pack(side="right", padx=4)
            # Mail ekle satırı
            add_r = ctk.CTkFrame(ml_frame, fg_color=RENK_PANEL)
            add_r.pack(fill="x", padx=4, pady=(2, 4))
            ent_new = ctk.CTkEntry(add_r, fg_color=RENK_KART, border_color=RENK_SINIR,
                                    text_color=RENK_YAZI, font=("Segoe UI", 10),
                                    placeholder_text="yeni@mail.com")
            ent_new.pack(side="left", fill="x", expand=True, padx=(4, 4))
            def ekle_mail(w=ent_new):
                m = w.get().strip()
                if m and "@" in m and m not in kural.get("mail_list", []):
                    kural.setdefault("mail_list", []).append(m)
                    render_ml()
            ctk.CTkButton(add_r, text="+ Ekle", width=60, height=22,
                          fg_color=renk, hover_color=RENK_VURGU2,
                          font=("Segoe UI", 10),
                          command=ekle_mail).pack(side="right", padx=(0, 4))

        render_ml()

        # Sil butonu
        ctk.CTkButton(body, text="🗑 Kuralı Sil", width=100, height=24,
                      fg_color=RENK_KIRMIZI, hover_color="#a93226",
                      font=("Segoe UI", 10),
                      command=lambda k=kural, f=frame: self._kural_sil(k, f)
                      ).grid(row=4, column=1, sticky="e", pady=(8, 0))

        return frame

    def _kural_ekle(self):
        cfg = load_config()
        yeni = {"id": str(uuid.uuid4())[:8], "ad": "Yeni Kural",
                "keywords": [], "mail_list": []}
        cfg.setdefault("kurallar", []).append(yeni)
        save_config(cfg)
        self._render()

    def _kural_sil(self, kural, frame):
        cfg = load_config()
        cfg["kurallar"] = [k for k in cfg.get("kurallar", [])
                           if k.get("id") != kural.get("id")]
        save_config(cfg)
        frame.destroy()
        self._kart_list = [w for w in self._kart_list if w != frame]

    def _kaydet(self):
        cfg = load_config()
        cfg["wa_group_name"] = self.ent_grup.get().strip()
        # Kart verilerini güncelle (FocusOut zaten tetikledi ama emin ol)
        save_config(cfg)
        self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# RAPOR PENCERESİ
# ══════════════════════════════════════════════════════════════════════════════
class RaporPencere(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Aylık Rapor")
        self.geometry("820x580")
        self.configure(fg_color=RENK_ANA_ARKA)
        self.grab_set()
        self._build()
        self._goster(datetime.now().year, datetime.now().month)

    def _build(self):
        hdr = ctk.CTkFrame(self, fg_color=RENK_PANEL, height=50, corner_radius=0)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="📊  Aylık Rapor",
                     font=("Segoe UI", 16, "bold"),
                     text_color=RENK_YAZI).pack(side="left", padx=16)

        # Ay/Yıl seçici
        nav = ctk.CTkFrame(self, fg_color=RENK_PANEL, corner_radius=8)
        nav.pack(fill="x", padx=14, pady=(10, 6))

        now = datetime.now()
        yillar = [str(y) for y in range(now.year - 2, now.year + 1)]
        aylar  = ["Ocak","Şubat","Mart","Nisan","Mayıs","Haziran",
                  "Temmuz","Ağustos","Eylül","Ekim","Kasım","Aralık"]

        self.cmb_yil = ctk.CTkComboBox(nav, values=yillar, width=90,
                                        fg_color=RENK_KART, border_color=RENK_SINIR,
                                        button_color=RENK_VURGU, text_color=RENK_YAZI)
        self.cmb_yil.set(str(now.year))
        self.cmb_yil.pack(side="left", padx=(12, 4), pady=8)

        self.cmb_ay = ctk.CTkComboBox(nav, values=aylar, width=120,
                                       fg_color=RENK_KART, border_color=RENK_SINIR,
                                       button_color=RENK_VURGU, text_color=RENK_YAZI)
        self.cmb_ay.set(aylar[now.month - 1])
        self.cmb_ay.pack(side="left", padx=4, pady=8)

        ctk.CTkButton(nav, text="🔍 Görüntüle", height=32,
                      fg_color=RENK_VURGU, hover_color=RENK_VURGU2,
                      command=self._ara).pack(side="left", padx=8)

        # Özet kartları
        self.frm_ozet = ctk.CTkFrame(self, fg_color=RENK_ANA_ARKA)
        self.frm_ozet.pack(fill="x", padx=14, pady=(0, 6))

        # Tablo
        tbl_frame = ctk.CTkFrame(self, fg_color=RENK_PANEL, corner_radius=8)
        tbl_frame.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Dark.Treeview",
                        background=RENK_KART, foreground=RENK_YAZI,
                        fieldbackground=RENK_KART, rowheight=26,
                        font=("Segoe UI", 10))
        style.configure("Dark.Treeview.Heading",
                        background=RENK_PANEL, foreground=RENK_YAZI2,
                        font=("Segoe UI", 10, "bold"))
        style.map("Dark.Treeview", background=[("selected", RENK_VURGU)])

        cols = ("Tarih", "Gönderen", "Kural", "Mesaj", "Resim")
        self.tree = ttk.Treeview(tbl_frame, columns=cols, show="headings",
                                  style="Dark.Treeview")
        for col, w in zip(cols, [120, 140, 110, 280, 60]):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor="w")

        sb = ttk.Scrollbar(tbl_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=4, pady=4)
        sb.pack(side="right", fill="y")

    def _ara(self):
        aylar = ["Ocak","Şubat","Mart","Nisan","Mayıs","Haziran",
                 "Temmuz","Ağustos","Eylül","Ekim","Kasım","Aralık"]
        yil = int(self.cmb_yil.get())
        ay  = aylar.index(self.cmb_ay.get()) + 1
        self._goster(yil, ay)

    def _goster(self, yil, ay):
        rows, gunluk = db_rapor(yil, ay)

        # Özet kartları temizle
        for w in self.frm_ozet.winfo_children():
            w.destroy()

        toplam = len(rows)
        self._ozet_kart("Toplam Red", str(toplam), RENK_KIRMIZI)
        # Günlük ort
        gun_sayisi = len(gunluk)
        ort = round(toplam / gun_sayisi, 1) if gun_sayisi else 0
        self._ozet_kart("Günlük Ort.", str(ort), RENK_SARI)
        # En yoğun gün
        if gunluk:
            en_yogun = max(gunluk, key=gunluk.get)
            self._ozet_kart("En Yoğun Gün", f"{en_yogun}\n({gunluk[en_yogun]})", RENK_VURGU)

        # Tablo doldur
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            tarih  = r[0][:19].replace("T", " ")
            gonder = r[1]
            kural  = r[2]
            mesaj  = (r[3] or "")[:60]
            resim  = "📎 Var" if r[4] else "—"
            self.tree.insert("", "end", values=(tarih, gonder, kural, mesaj, resim))

    def _ozet_kart(self, baslik, deger, renk):
        f = ctk.CTkFrame(self.frm_ozet, fg_color=RENK_KART, corner_radius=8)
        f.pack(side="left", padx=(0, 8), pady=6, ipadx=16, ipady=8)
        ctk.CTkLabel(f, text=baslik, text_color=RENK_YAZI2,
                     font=("Segoe UI", 10)).pack()
        ctk.CTkLabel(f, text=deger, text_color=renk,
                     font=("Segoe UI", 18, "bold")).pack()


# ══════════════════════════════════════════════════════════════════════════════
# SPLASH SCREEN
# ══════════════════════════════════════════════════════════════════════════════
class SplashEkran(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.overrideredirect(True)          # Çerçevesiz pencere
        self.configure(bg="#071320")
        self.attributes("-topmost", True)

        W, H = 540, 320
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x  = (sw - W) // 2
        y  = (sh - H) // 2
        self.geometry(f"{W}x{H}+{x}+{y}")

        splash_path = os.path.join(BASE_DIR, "splash.png")
        if PIL_OK and os.path.exists(splash_path):
            img = Image.open(splash_path).resize((W, H), Image.LANCZOS)
            self._photo = ImageTk.PhotoImage(img)
            lbl = tk.Label(self, image=self._photo, bg="#071320", bd=0)
            lbl.pack(fill="both", expand=True)
        else:
            # Fallback: sade metin
            tk.Label(self, text="Pregate Kayıt Red\nWA → Mail Botu",
                     font=("Segoe UI", 20, "bold"),
                     fg="#e8f0f7", bg="#071320").pack(expand=True)
            tk.Label(self, text="S.SEYMEN tarafından hazırlanmıştır",
                     font=("Segoe UI", 11),
                     fg="#8faabf", bg="#071320").pack(pady=(0, 20))

        # İnce kenarlık
        self.configure(highlightbackground="#1a6ea8",
                       highlightthickness=1)

        # İlerleme çubuğu
        self._bar_frame = tk.Frame(self, bg="#0d1f30", height=4)
        self._bar_frame.place(x=0, y=H-4, width=W, height=4)
        self._bar = tk.Frame(self._bar_frame, bg="#1a6ea8", height=4)
        self._bar.place(x=0, y=0, width=0, height=4)
        self._W = W
        self._step = 0
        self._animate()

    def _animate(self):
        self._step += 1
        w = int((self._step / 30) * self._W)
        self._bar.place(x=0, y=0, width=min(w, self._W), height=4)
        if self._step < 30:
            self.after(50, self._animate)

    def kapat(self):
        self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# ANA UYGULAMA
# ══════════════════════════════════════════════════════════════════════════════
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Pregate Kayıt Red – WA Mail Botu  |  Poliport")
        self.geometry("900x620")
        self.minsize(800, 540)
        self.configure(fg_color=RENK_ANA_ARKA)

        self.wa_bot    = None
        self.running   = False
        self.mail_say  = 0

        db_init()
        self._build_ui()

    # ── UI ───────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Başlık
        hdr = ctk.CTkFrame(self, fg_color=RENK_PANEL, corner_radius=0, height=56)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="📱  Pregate Kayıt Red – WA Mail Botu",
                     font=("Segoe UI", 16, "bold"),
                     text_color=RENK_YAZI).pack(side="left", padx=20, pady=12)
        self.lbl_durum = ctk.CTkLabel(hdr, text="● Beklemede",
                                       font=("Segoe UI", 12),
                                       text_color=RENK_YAZI2)
        self.lbl_durum.pack(side="right", padx=20)

        content = ctk.CTkFrame(self, fg_color=RENK_ANA_ARKA)
        content.pack(fill="both", expand=True, padx=14, pady=10)
        content.columnconfigure(0, weight=0)
        content.columnconfigure(1, weight=1)
        content.rowconfigure(0, weight=1)

        self._build_sol(content)
        self._build_sag(content)

    def _build_sol(self, parent):
        sol = ctk.CTkFrame(parent, fg_color=RENK_PANEL, corner_radius=10, width=200)
        sol.grid(row=0, column=0, sticky="nsew", padx=(0, 7))
        sol.pack_propagate(False)

        ctk.CTkLabel(sol, text="KONTROL", font=("Segoe UI", 11, "bold"),
                     text_color=RENK_YAZI2).pack(anchor="w", padx=14, pady=(14, 4))

        self.btn_baslat = ctk.CTkButton(
            sol, text="▶  Başlat",
            fg_color=RENK_YESIL, hover_color="#17a844",
            font=("Segoe UI", 13, "bold"), height=42,
            command=self._baslat)
        self.btn_baslat.pack(fill="x", padx=12, pady=4)

        self.btn_durdur = ctk.CTkButton(
            sol, text="■  Durdur",
            fg_color=RENK_KIRMIZI, hover_color="#c0392b",
            font=("Segoe UI", 13, "bold"), height=42,
            state="disabled", command=self._durdur)
        self.btn_durdur.pack(fill="x", padx=12, pady=4)

        ctk.CTkFrame(sol, height=1, fg_color=RENK_SINIR).pack(
            fill="x", padx=12, pady=12)

        ctk.CTkButton(sol, text="⚙  Ayarlar", height=36,
                      fg_color=RENK_VURGU, hover_color=RENK_VURGU2,
                      font=("Segoe UI", 12),
                      command=self._ayarlar).pack(fill="x", padx=12, pady=4)

        ctk.CTkButton(sol, text="📊  Rapor", height=36,
                      fg_color="#2c3e6b", hover_color="#1e2d50",
                      font=("Segoe UI", 12),
                      command=self._rapor).pack(fill="x", padx=12, pady=4)

        ctk.CTkFrame(sol, height=1, fg_color=RENK_SINIR).pack(
            fill="x", padx=12, pady=12)

        # Sayaç
        f1 = ctk.CTkFrame(sol, fg_color=RENK_KART, corner_radius=8)
        f1.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(f1, text="Bugün Gönderilen",
                     font=("Segoe UI", 9), text_color=RENK_YAZI2).pack(pady=(8, 0))
        self.lbl_sayac = ctk.CTkLabel(f1, text="0",
                                       font=("Segoe UI", 26, "bold"),
                                       text_color=RENK_YESIL)
        self.lbl_sayac.pack(pady=(0, 8))

        # Grup adı bilgisi
        cfg = load_config()
        self.lbl_grup = ctk.CTkLabel(sol,
                                      text=f"Grup: {cfg.get('wa_group_name','—')}",
                                      font=("Segoe UI", 9),
                                      text_color=RENK_YAZI2,
                                      wraplength=170)
        self.lbl_grup.pack(padx=12, pady=(16, 4), anchor="w")

        ctk.CTkLabel(sol,
                     text="ℹ️ İlk açılışta Chrome\naçılır, QR taratın.",
                     font=("Segoe UI", 9), text_color=RENK_YAZI2,
                     justify="left").pack(anchor="w", padx=14, pady=(8, 0))

    def _build_sag(self, parent):
        sag = ctk.CTkFrame(parent, fg_color=RENK_PANEL, corner_radius=10)
        sag.grid(row=0, column=1, sticky="nsew")
        sag.columnconfigure(0, weight=1)
        sag.rowconfigure(0, weight=1)

        ctk.CTkLabel(sag, text="İŞLEM LOGU", font=("Segoe UI", 11, "bold"),
                     text_color=RENK_YAZI2).grid(row=0, column=0,
                                                  sticky="w", padx=14, pady=(14, 4))

        self.txt_log = ctk.CTkTextbox(sag, fg_color=RENK_KART,
                                       text_color=RENK_YAZI,
                                       font=("Consolas", 10),
                                       corner_radius=8)
        self.txt_log.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 14))
        sag.rowconfigure(1, weight=1)

    # ── Kontrol ───────────────────────────────────────────────────────────────
    def _baslat(self):
        cfg = load_config()
        if not cfg.get("wa_group_name", "").strip():
            messagebox.showwarning("Uyarı", "Önce Ayarlar'dan WhatsApp grup adını girin!")
            return
        if not cfg.get("kurallar"):
            messagebox.showwarning("Uyarı", "Ayarlar'dan en az bir kural ekleyin!")
            return

        self.running = True
        self.btn_baslat.configure(state="disabled")
        self.btn_durdur.configure(state="normal")
        self._log("🚀 Bot başlatılıyor, Chrome penceresi açılacak…")

        self.wa_bot = WhatsAppBot(
            on_log=self._log,
            on_status=self._set_durum_params,
            on_message=self._on_message
        )
        threading.Thread(target=self.wa_bot.start, daemon=True).start()

    def _durdur(self):
        self.running = False
        if self.wa_bot:
            self.wa_bot.stop()
            self.wa_bot = None
        self.btn_baslat.configure(state="normal")
        self.btn_durdur.configure(state="disabled")
        self._set_durum("● Durduruldu", RENK_KIRMIZI)
        self._log("■ Bot durduruldu.")

    def _ayarlar(self):
        AyarlarPencere(self)
        # Pencere kapanınca grup adını güncelle
        self.after(500, self._grup_label_guncelle)

    def _grup_label_guncelle(self):
        cfg = load_config()
        self.lbl_grup.configure(text=f"Grup: {cfg.get('wa_group_name', '—')}")

    def _rapor(self):
        RaporPencere(self)

    # ── Mesaj işleme ──────────────────────────────────────────────────────────
    def _on_message(self, gonderen, text, img_paths):
        cfg = load_config()
        metin_lower = text.lower()
        eslesen = []
        for kural in cfg.get("kurallar", []):
            for kw in kural.get("keywords", []):
                if kw in metin_lower:
                    eslesen.append(kural)
                    break

        if not eslesen:
            self._log(f"💬 Mesaj geldi, eşleşme yok: {text[:50]}")
            return

        for kural in eslesen:
            ml = kural.get("mail_list", [])
            if not ml:
                self._log(f"⚠ '{kural.get('ad')}' için mail listesi boş!")
                continue

            ok, info = send_mail(
                ml, kural["ad"], gonderen, text,
                img_paths, cfg.get("wa_group_name", "")
            )
            if ok:
                self.mail_say += 1
                self.after(0, lambda: self.lbl_sayac.configure(
                    text=str(self.mail_say)))
                db_ekle(gonderen, kural["ad"], text, bool(img_paths))
                self._log(
                    f"✉ [{kural['ad']}] {gonderen} → "
                    f"{', '.join(ml[:2])}{'…' if len(ml)>2 else ''}"
                    f"{'  📎' if img_paths else ''}"
                )
            else:
                self._log(f"❌ [{kural['ad']}] Mail hatası: {info}")

    # ── Yardımcılar ───────────────────────────────────────────────────────────
    def _log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        def _do():
            self.txt_log.configure(state="normal")
            self.txt_log.insert("end", f"[{ts}] {msg}\n")
            self.txt_log.see("end")
            self.txt_log.configure(state="disabled")
        self.after(0, _do)

    def _set_durum(self, text, color):
        self.after(0, lambda: self.lbl_durum.configure(
            text=text, text_color=color))

    def _set_durum_params(self, _, text, color):
        self._set_durum(text, color)

    def on_close(self):
        if self.wa_bot:
            self.wa_bot.stop()
        self.destroy()


if __name__ == "__main__":
    # Splash için gizli root
    root = tk.Tk()
    root.withdraw()

    splash = SplashEkran(root)
    splash.update()

    # 2.5 sn sonra ana pencereyi aç
    def _ac():
        splash.kapat()
        root.destroy()
        app = App()
        app.protocol("WM_DELETE_WINDOW", app.on_close)
        app.mainloop()

    root.after(2500, _ac)
    root.mainloop()
