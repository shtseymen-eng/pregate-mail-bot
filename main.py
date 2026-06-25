import customtkinter as ctk
import threading
import json, os, sys, time, glob, uuid, tempfile
import sqlite3 as _sqlite3
from datetime import datetime
from tkinter import messagebox, ttk
import tkinter as tk
from collections import defaultdict

try:
    from PIL import Image, ImageTk
    PIL_OK = True
except ImportError:
    PIL_OK = False

try:
    import win32com.client
    OUTLOOK_OK = True
except ImportError:
    OUTLOOK_OK = False

try:
    import win32gui, win32con
    WIN32_OK = True
except ImportError:
    WIN32_OK = False

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

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
DB_FILE     = os.path.join(BASE_DIR, "kayitlar.db")

BEKLEME_SURE = 120   # saniye — aynı kişiden mesaj toplama süresi

# ── Config ──────────────────────────────────────────────────────────────────
def load_config():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {"wa_group_name":"Pregate Kayıt Red","outlook_account":"","kurallar":[]}

def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

# ── Kayıt DB ────────────────────────────────────────────────────────────────
def db_init():
    con = _sqlite3.connect(DB_FILE)
    con.execute("""CREATE TABLE IF NOT EXISTS kayitlar (
        id TEXT PRIMARY KEY, tarih TEXT, gonderen TEXT,
        kural_ad TEXT, mesaj TEXT, resim_var INTEGER DEFAULT 0)""")
    con.commit(); con.close()

def db_ekle(gonderen, kural_ad, mesaj, resim_var):
    con = _sqlite3.connect(DB_FILE)
    con.execute("INSERT INTO kayitlar VALUES (?,?,?,?,?,?)",
        (str(uuid.uuid4()), datetime.now().isoformat(),
         gonderen, kural_ad, mesaj, 1 if resim_var else 0))
    con.commit(); con.close()

def db_rapor(yil, ay):
    con = _sqlite3.connect(DB_FILE)
    rows = con.execute(
        "SELECT tarih,gonderen,kural_ad,mesaj,resim_var FROM kayitlar "
        "WHERE tarih LIKE ? ORDER BY tarih",
        (f"{yil:04d}-{ay:02d}%",)).fetchall()
    gunluk = {}
    for r in rows:
        g = r[0][:10]; gunluk[g] = gunluk.get(g,0)+1
    con.close()
    return rows, gunluk

# ── Outlook ─────────────────────────────────────────────────────────────────
def outlook_hesaplari():
    if not OUTLOOK_OK: return []
    try:
        ol = win32com.client.Dispatch("Outlook.Application")
        return [a.SmtpAddress for a in ol.GetNamespace("MAPI").Accounts]
    except: return []

MAIL_SABLON = """\
Sayın İlgili,

Aşağıda bilgileri paylaşılan araç, {tarih} tarihinde kapı girişinde \
kayıt yapılmamış olup sisteme RED olarak işlenmiştir.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Gönderen  : {gonderen}
Grup      : {grup}
Tarih     : {tarih}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Mesaj İçeriği:
{mesaj}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Bu araç kapı giriş kaydı yapılmamış olduğundan \
ilgili birimler tarafından gerekli işlemlerin yapılması gerekmektedir.

Saygılarımızla,
Pregate Araç Kontrol Sistemi
"""

def send_mail(mail_list, kural_ad, gonderen, mesaj, from_account, grup_adi):
    if not OUTLOOK_OK: return False, "pywin32 yüklü değil"
    try:
        tarih = datetime.now().strftime("%d.%m.%Y %H:%M")
        ol    = win32com.client.Dispatch("Outlook.Application")
        mail  = ol.CreateItem(0)
        mail.Subject = f"[KAYIT RED] {kural_ad} – {tarih}"
        mail.Body    = MAIL_SABLON.format(
            tarih=tarih, gonderen=gonderen,
            grup=grup_adi, mesaj=mesaj)
        mail.To = "; ".join(mail_list)
        if from_account:
            try:
                for acc in ol.GetNamespace("MAPI").Accounts:
                    if acc.SmtpAddress.lower() == from_account.lower():
                        mail.SendUsingAccount = acc; break
            except: pass
        mail.Send()
        return True, "OK"
    except Exception as e:
        return False, str(e)

# ══════════════════════════════════════════════════════════════════════════════
# MESAJ BİRİKTİRİCİ — aynı kişiden 2 dk içinde gelen mesajları toplar
# ══════════════════════════════════════════════════════════════════════════════
class MesajBiriktiric:
    def __init__(self, on_gonder, bekleme=BEKLEME_SURE):
        self.on_gonder  = on_gonder   # (gonderen, birlesik_metin) callback
        self.bekleme    = bekleme
        self._kuyruk    = {}          # gonderen → {"metinler":[], "timer":Timer}
        self._lock      = threading.Lock()

    def ekle(self, gonderen, metin):
        """Yeni mesaj geldi — biriktiriciye ekle, timer sıfırla."""
        with self._lock:
            if gonderen not in self._kuyruk:
                self._kuyruk[gonderen] = {"metinler": [], "timer": None}

            self._kuyruk[gonderen]["metinler"].append(metin)

            # Varolan timer'ı iptal et
            t = self._kuyruk[gonderen]["timer"]
            if t: t.cancel()

            # Yeni timer — 2 dk sonra gönder
            yeni_t = threading.Timer(
                self.bekleme,
                self._gonder,
                args=[gonderen])
            yeni_t.daemon = True
            yeni_t.start()
            self._kuyruk[gonderen]["timer"] = yeni_t

    def _gonder(self, gonderen):
        """2 dk doldu — biriken mesajları gönder."""
        with self._lock:
            if gonderen not in self._kuyruk:
                return
            metinler = self._kuyruk.pop(gonderen)["metinler"]

        if metinler:
            birlesik = "\n".join(metinler)
            self.on_gonder(gonderen, birlesik)

    def temizle(self):
        with self._lock:
            for v in self._kuyruk.values():
                if v["timer"]: v["timer"].cancel()
            self._kuyruk.clear()

# ══════════════════════════════════════════════════════════════════════════════
# WHATSAPP DB OKUYUCU
# ══════════════════════════════════════════════════════════════════════════════
class WABot:
    WA_DB_PATTERNS = [
        r"%LOCALAPPDATA%\Packages\5319275A.WhatsApp_cv1g1gvanyjgm\LocalState\messages.db",
        r"%LOCALAPPDATA%\Packages\5319275A.WhatsApp_cv1g1gvanyjgm\LocalState\default\messages.db",
        r"%LOCALAPPDATA%\WhatsApp\messages.db",
        r"%APPDATA%\WhatsApp\messages.db",
        r"%USERPROFILE%\AppData\Local\WhatsApp\messages.db",
    ]

    def __init__(self, on_log, on_status, on_message):
        self.on_log      = on_log
        self.on_status   = on_status
        self.on_message  = on_message   # (gonderen, metin_birikik) — 2dk sonra tetiklenir
        self.running     = False
        self._son_id     = 0
        self._wa_db      = None
        self._tablo      = None    # hangi tablo kullanılıyor
        self._kolon_map  = {}      # sütun isimleri haritası
        self._biriktiric = None

    def start(self):
        self.running = True
        self.on_status("START", "● Başlatılıyor…", RENK_SARI)
        self._biriktiric = MesajBiriktiric(
            on_gonder=self.on_message,
            bekleme=BEKLEME_SURE)

        db = self._wa_db_bul()
        if db:
            self._wa_db = db
            self.on_log(f"✅ WhatsApp DB bulundu: {os.path.basename(os.path.dirname(db))}")
            self._db_kesfet()
            self._son_id = self._son_id_al()
            self.on_log(f"📌 Başlangıç mesaj ID: {self._son_id}")
            self.on_status("OK", "● Çalışıyor ✓", RENK_YESIL)
            self._dinle()
        else:
            self.on_log("⚠ WhatsApp Desktop DB bulunamadı.")
            self.on_log("   → Microsoft Store'dan WhatsApp indirin ve açın.")
            self.on_log("   → Bir mesaj alındıktan sonra tekrar başlatın.")
            self.on_status("WARN", "● DB Bulunamadı", RENK_KIRMIZI)
            # 60 sn'de bir tekrar dene
            while self.running:
                time.sleep(60)
                db = self._wa_db_bul()
                if db:
                    self._wa_db = db
                    self.on_log("✅ DB bulundu, başlatılıyor…")
                    self._db_kesfet()
                    self._son_id = self._son_id_al()
                    self.on_status("OK", "● Çalışıyor ✓", RENK_YESIL)
                    self._dinle()
                    return

    def _wa_db_bul(self):
        # Doğrudan yollar
        for p in self.WA_DB_PATTERNS:
            expanded = os.path.expandvars(p)
            if os.path.exists(expanded):
                return expanded
        # Wildcard arama
        appdata = os.environ.get("LOCALAPPDATA", "")
        for pat in [
            os.path.join(appdata, "Packages", "5319275A.WhatsApp*",
                         "LocalState", "messages.db"),
            os.path.join(appdata, "Packages", "5319275A.WhatsApp*",
                         "LocalState", "*", "messages.db"),
            os.path.join(appdata, "Packages", "5319275A.WhatsApp*",
                         "LocalCache", "Roaming", "WhatsApp", "messages.db"),
        ]:
            found = glob.glob(pat)
            if found:
                return found[0]
        return None

    def _db_kesfet(self):
        """DB şemasını keşfet — tablo ve sütun isimlerini bul."""
        try:
            con = _sqlite3.connect(
                f"file:{self._wa_db}?mode=ro&nolock=1", uri=True, timeout=5)
            cur = con.cursor()

            tabloler = [r[0] for r in cur.execute(
                "SELECT name FROM sqlite_master WHERE type='table'").fetchall()]
            self.on_log(f"📊 DB tabloları: {', '.join(tabloler)}")

            # Mesaj tablosunu bul
            for tbl in ["messages", "message", "msgs", "chat_messages"]:
                if tbl in tabloler:
                    self._tablo = tbl
                    break

            if not self._tablo and tabloler:
                # En fazla satırı olan tabloyu seç
                max_count = 0
                for tbl in tabloler:
                    try:
                        c = cur.execute(f"SELECT COUNT(*) FROM [{tbl}]").fetchone()[0]
                        if c > max_count:
                            max_count = c; self._tablo = tbl
                    except: pass

            if self._tablo:
                cols = [r[1] for r in cur.execute(
                    f"PRAGMA table_info([{self._tablo}])").fetchall()]
                self.on_log(f"📋 Tablo: {self._tablo} | Sütunlar: {', '.join(cols)}")
                self._kolon_map = self._sütun_haritasi(cols)
                self.on_log(f"🗺 Harita: {self._kolon_map}")
            else:
                self.on_log("⚠ Mesaj tablosu bulunamadı!")
            con.close()
        except Exception as e:
            self.on_log(f"⚠ DB keşif hatası: {e}")

    def _sütun_haritasi(self, cols):
        """Sütun isimlerini standart isimlere eşle."""
        cols_lower = {c.lower(): c for c in cols}
        harita = {}

        # Chat/grup alanı
        for c in ["chat_jid","chat_id","chatid","jid","remote_jid","remotejid",
                  "conversation_jid","group_jid","thread_id"]:
            if c in cols_lower:
                harita["chat"] = cols_lower[c]; break

        # Gönderen
        for c in ["sender","sender_jid","from","participant","push_name",
                  "pushname","author","notify_name","notifyname","name"]:
            if c in cols_lower:
                harita["sender"] = cols_lower[c]; break

        # Metin
        for c in ["text","body","message","content","message_text",
                  "messagetext","data","msg","text_data"]:
            if c in cols_lower:
                harita["text"] = cols_lower[c]; break

        # Zaman
        for c in ["timestamp","time","date","created_at","message_timestamp",
                  "sent_at","received_at"]:
            if c in cols_lower:
                harita["time"] = cols_lower[c]; break

        return harita

    def _son_id_al(self):
        if not self._tablo: return 0
        try:
            con = _sqlite3.connect(
                f"file:{self._wa_db}?mode=ro&nolock=1", uri=True, timeout=5)
            row = con.execute(
                f"SELECT MAX(rowid) FROM [{self._tablo}]").fetchone()
            con.close()
            return row[0] if row and row[0] else 0
        except: return 0

    def _dinle(self):
        """10 sn'de bir DB kontrol et."""
        cfg = load_config()
        hedef = cfg.get("wa_group_name","").strip().lower()

        while self.running:
            try:
                self._kontrol(hedef)
            except Exception as e:
                self.on_log(f"⚠ Kontrol hatası: {e}")
            time.sleep(10)

    def _kontrol(self, hedef_grup):
        if not self._tablo: return
        try:
            con = _sqlite3.connect(
                f"file:{self._wa_db}?mode=ro&nolock=1", uri=True, timeout=5)
            con.row_factory = _sqlite3.Row
            cur = con.cursor()

            rows = cur.execute(
                f"SELECT rowid,* FROM [{self._tablo}] "
                f"WHERE rowid > ? ORDER BY rowid",
                (self._son_id,)).fetchall()
            con.close()

            for row in rows:
                self._son_id = max(self._son_id, row["rowid"])
                d = dict(row)

                # Sütun haritasını kullan
                chat   = str(d.get(self._kolon_map.get("chat",""), "") or "").lower()
                sender = str(d.get(self._kolon_map.get("sender",""), "") or "Bilinmiyor")
                text   = str(d.get(self._kolon_map.get("text",""), "") or "")

                # Eğer harita boşsa tüm string sütunlara bak
                if not self._kolon_map.get("text"):
                    for k, v in d.items():
                        if isinstance(v, str) and len(v) > 2 and k not in (
                                "rowid","id","key_id","_id"):
                            text = v; break

                # Grup filtresi
                if hedef_grup and hedef_grup not in chat:
                    # Chat alanı boşsa yine de işle (bazı DB'lerde join gerekir)
                    if chat:
                        continue

                # Sistem mesajlarını atla
                if not text or text in ("null","None",""):
                    continue

                # Kendi telefon numarasını gönderen göster
                sender_clean = sender.split("@")[0] if "@" in sender else sender

                self.on_log(f"📩 [{sender_clean}]: {text[:60]}")
                self._biriktiric.ekle(sender_clean, text)

        except _sqlite3.OperationalError as e:
            if "locked" in str(e).lower() or "readonly" in str(e).lower():
                pass  # WA DB kilitli olabilir, geç
            else:
                raise

    def stop(self):
        self.running = False
        if self._biriktiric:
            self._biriktiric.temizle()


# ══════════════════════════════════════════════════════════════════════════════
# MAİL ÖNİZLEME
# ══════════════════════════════════════════════════════════════════════════════
class MailOnizleme(ctk.CTkToplevel):
    def __init__(self, master, kural_ad, mail_list, keywords):
        super().__init__(master)
        self.title(f"Mail Önizleme – {kural_ad}")
        self.geometry("620x520")
        self.configure(fg_color=RENK_ANA_ARKA)
        self.grab_set()
        cfg      = load_config()
        tarih    = datetime.now().strftime("%d.%m.%Y %H:%M")
        kw_str   = keywords[0] if keywords else kural_ad
        mesaj    = (f"[{kw_str}] 34 ABC 123 plakalı araç kayıt yaptırmadı.\n"
                    f"Sürücü bilgi vermeden ayrıldı.")
        grup     = cfg.get("wa_group_name","Pregate Kayıt Red")
        from_acc = cfg.get("outlook_account","—")

        hdr = ctk.CTkFrame(self, fg_color=RENK_PANEL, height=48, corner_radius=0)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text=f"✉  Mail Önizleme — {kural_ad}",
                     font=("Segoe UI",14,"bold"),
                     text_color=RENK_YAZI).pack(side="left", padx=16, pady=10)

        meta = ctk.CTkFrame(self, fg_color=RENK_KART, corner_radius=8)
        meta.pack(fill="x", padx=14, pady=(12,4))
        for lbl, val in [("Gönderen Hesap:", from_acc),
                         ("Konu:", f"[KAYIT RED] {kural_ad} – {tarih}"),
                         ("Alıcılar:", ", ".join(mail_list) if mail_list else "—")]:
            r = ctk.CTkFrame(meta, fg_color=RENK_KART)
            r.pack(fill="x", padx=10, pady=3)
            ctk.CTkLabel(r, text=lbl, width=130, anchor="w",
                         text_color=RENK_YAZI2,
                         font=("Segoe UI",10,"bold")).pack(side="left")
            ctk.CTkLabel(r, text=val, anchor="w", text_color=RENK_YAZI,
                         font=("Segoe UI",10), wraplength=400
                         ).pack(side="left", padx=(4,0))

        ctk.CTkLabel(self, text="MAİL İÇERİĞİ", font=("Segoe UI",10,"bold"),
                     text_color=RENK_YAZI2).pack(anchor="w", padx=14, pady=(8,2))
        txt = ctk.CTkTextbox(self, fg_color=RENK_KART, text_color=RENK_YAZI,
                              font=("Consolas",10), corner_radius=8)
        txt.pack(fill="both", expand=True, padx=14, pady=(0,8))
        txt.insert("end", MAIL_SABLON.format(
            tarih=tarih, gonderen="Ahmet Yılmaz", grup=grup, mesaj=mesaj))
        txt.configure(state="disabled")
        ctk.CTkLabel(self, text="⚠ Örnek önizlemedir.",
                     font=("Segoe UI",9), text_color=RENK_SARI).pack(pady=(0,4))
        ctk.CTkButton(self, text="Kapat", height=34,
                      fg_color=RENK_VURGU, hover_color=RENK_VURGU2,
                      command=self.destroy).pack(pady=(0,12))


# ══════════════════════════════════════════════════════════════════════════════
# AYARLAR
# ══════════════════════════════════════════════════════════════════════════════
class AyarlarPencere(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Ayarlar"); self.geometry("780x680")
        self.minsize(700,580); self.configure(fg_color=RENK_ANA_ARKA)
        self.grab_set(); self._kart_list = []; self._build()

    def _build(self):
        hdr = ctk.CTkFrame(self, fg_color=RENK_PANEL, height=52, corner_radius=0)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="⚙  Ayarlar", font=("Segoe UI",16,"bold"),
                     text_color=RENK_YAZI).pack(side="left", padx=16, pady=12)
        self.tab = ctk.CTkTabview(self, fg_color=RENK_PANEL,
            segmented_button_fg_color=RENK_KART,
            segmented_button_selected_color=RENK_VURGU,
            segmented_button_selected_hover_color=RENK_VURGU2,
            segmented_button_unselected_color=RENK_KART,
            segmented_button_unselected_hover_color=RENK_SINIR,
            text_color=RENK_YAZI, corner_radius=10)
        self.tab.pack(fill="both", expand=True, padx=14, pady=(8,4))
        self.tab.add("📱  WhatsApp")
        self.tab.add("🔑  Kurallar")
        self.tab.add("✉   Outlook")
        self._build_wa(self.tab.tab("📱  WhatsApp"))
        self._build_kurallar(self.tab.tab("🔑  Kurallar"))
        self._build_outlook(self.tab.tab("✉   Outlook"))
        ctk.CTkButton(self, text="💾  Kaydet & Kapat",
                      fg_color=RENK_YESIL, hover_color="#17a844",
                      font=("Segoe UI",13,"bold"), height=42,
                      command=self._kaydet).pack(fill="x", padx=14, pady=(4,14))

    def _build_wa(self, parent):
        parent.configure(fg_color=RENK_PANEL)
        info = ctk.CTkFrame(parent, fg_color=RENK_KART, corner_radius=10)
        info.pack(fill="x", padx=10, pady=(10,6))
        ctk.CTkLabel(info, text="WhatsApp Desktop Kurulumu",
                     font=("Segoe UI",12,"bold"), text_color=RENK_YAZI
                     ).pack(anchor="w", padx=14, pady=(12,6))
        for no, m in [
            ("1","Microsoft Store → 'WhatsApp' indirin ve kurun."),
            ("2","WhatsApp Desktop'ı açın ve telefonla QR ile giriş yapın."),
            ("3","Aşağıya WhatsApp'taki grup adını AYNEN yazın."),
            ("4","Bot arka planda çalışır — fareye/klavyeye dokunmaz."),
            ("5","Aynı kişinin 2 dk içindeki mesajları tek mail olarak gönderilir."),
        ]:
            r = ctk.CTkFrame(info, fg_color=RENK_KART)
            r.pack(fill="x", padx=14, pady=3)
            ctk.CTkLabel(r, text=no, width=24, height=24, fg_color=RENK_VURGU,
                         corner_radius=12, font=("Segoe UI",10,"bold"),
                         text_color="white").pack(side="left")
            ctk.CTkLabel(r, text=m, text_color=RENK_YAZI,
                         font=("Segoe UI",10)).pack(side="left", padx=10)
        ctk.CTkFrame(info, height=8, fg_color=RENK_KART).pack()

        gf = ctk.CTkFrame(parent, fg_color=RENK_KART, corner_radius=10)
        gf.pack(fill="x", padx=10, pady=6)
        ctk.CTkLabel(gf, text="Dinlenecek WhatsApp Grubu",
                     font=("Segoe UI",12,"bold"),
                     text_color=RENK_YAZI).pack(anchor="w", padx=14, pady=(12,4))
        self.ent_grup = ctk.CTkEntry(gf, fg_color=RENK_PANEL,
                                      border_color=RENK_VURGU,
                                      text_color=RENK_YAZI, font=("Segoe UI",12),
                                      placeholder_text="örn: Pregate Kayıt Red")
        self.ent_grup.insert(0, load_config().get("wa_group_name",""))
        self.ent_grup.pack(fill="x", padx=14, pady=(0,4))
        ctk.CTkLabel(gf, text="💡 Grup adı WhatsApp'takiyle birebir aynı olmalı.",
                     text_color=RENK_YAZI2, font=("Segoe UI",9)
                     ).pack(anchor="w", padx=14, pady=(2,12))

        # Bekleme süresi
        bf = ctk.CTkFrame(parent, fg_color=RENK_KART, corner_radius=10)
        bf.pack(fill="x", padx=10, pady=6)
        ctk.CTkLabel(bf, text="Mesaj Biriktirme Süresi",
                     font=("Segoe UI",12,"bold"),
                     text_color=RENK_YAZI).pack(anchor="w", padx=14, pady=(12,4))
        row = ctk.CTkFrame(bf, fg_color=RENK_KART)
        row.pack(fill="x", padx=14, pady=(0,12))
        ctk.CTkLabel(row,
                     text="Aynı kişinin mesajlarını şu kadar dk bekle ve tek mailde gönder:",
                     text_color=RENK_YAZI2, font=("Segoe UI",10)
                     ).pack(side="left")
        self.ent_sure = ctk.CTkEntry(row, width=60, fg_color=RENK_PANEL,
                                      border_color=RENK_VURGU,
                                      text_color=RENK_YAZI, font=("Segoe UI",11))
        self.ent_sure.insert(0, str(load_config().get("bekleme_dk", 2)))
        self.ent_sure.pack(side="left", padx=10)
        ctk.CTkLabel(row, text="dakika", text_color=RENK_YAZI2,
                     font=("Segoe UI",10)).pack(side="left")

    def _build_kurallar(self, parent):
        parent.configure(fg_color=RENK_PANEL)
        kh = ctk.CTkFrame(parent, fg_color=RENK_PANEL)
        kh.pack(fill="x", padx=10, pady=(10,4))
        ctk.CTkLabel(kh, text="Anahtar Kelime → Mail Grubu Kuralları",
                     font=("Segoe UI",12,"bold"),
                     text_color=RENK_YAZI).pack(side="left")
        ctk.CTkButton(kh, text="+ Yeni Kural", width=110, height=28,
                      fg_color=RENK_VURGU, hover_color=RENK_VURGU2,
                      command=self._kural_ekle).pack(side="right")
        self.scroll = ctk.CTkScrollableFrame(parent, fg_color=RENK_PANEL)
        self.scroll.pack(fill="both", expand=True, padx=10, pady=4)
        self.scroll.columnconfigure(0, weight=1)
        self._render_kurallar()

    def _render_kurallar(self):
        for item in self._kart_list: item[0].destroy()
        self._kart_list.clear()
        for i, k in enumerate(load_config().get("kurallar",[])):
            renk = RENKLER[i % len(RENKLER)]
            kart, ea, ek = self._kural_kart(self.scroll, k, renk)
            kart.grid(row=i, column=0, sticky="ew", pady=(0,10))
            self._kart_list.append((kart, k, ea, ek))

    def _kural_kart(self, parent, kural, renk):
        frame = ctk.CTkFrame(parent, fg_color=RENK_KART, corner_radius=10)
        frame.columnconfigure(1, weight=1)
        ctk.CTkFrame(frame, fg_color=renk, width=5,
                     corner_radius=0).grid(row=0,column=0,rowspan=99,sticky="ns")
        body = ctk.CTkFrame(frame, fg_color=RENK_KART)
        body.grid(row=0, column=1, sticky="ew", padx=10, pady=10)
        body.columnconfigure(1, weight=1)

        ctk.CTkLabel(body, text="Kural Adı:", text_color=RENK_YAZI2,
                     font=("Segoe UI",10)).grid(row=0,column=0,sticky="w")
        ent_ad = ctk.CTkEntry(body, fg_color=RENK_PANEL, border_color=renk,
                               text_color=RENK_YAZI, font=("Segoe UI",11,"bold"))
        ent_ad.insert(0, kural.get("ad",""))
        ent_ad.grid(row=0, column=1, sticky="ew", padx=(8,0))

        ctk.CTkLabel(body, text="Kelimeler\n(virgülle):", text_color=RENK_YAZI2,
                     font=("Segoe UI",10)).grid(row=1,column=0,sticky="w",pady=(8,0))
        ent_kw = ctk.CTkEntry(body, fg_color=RENK_PANEL, border_color=RENK_SINIR,
                               text_color=RENK_YAZI, font=("Segoe UI",11))
        ent_kw.insert(0, ", ".join(kural.get("keywords",[])))
        ent_kw.grid(row=1, column=1, sticky="ew", padx=(8,0), pady=(8,0))
        ctk.CTkLabel(body, text="💡 Büyük/küçük harf ayrımı yoktur.",
                     text_color=RENK_YAZI2, font=("Segoe UI",9)
                     ).grid(row=2, column=1, sticky="w", padx=(8,0), pady=(2,0))

        ml_hdr = ctk.CTkFrame(body, fg_color=RENK_KART)
        ml_hdr.grid(row=3,column=0,columnspan=2,sticky="ew",pady=(10,2))
        ctk.CTkLabel(ml_hdr, text="Mail Listesi:", text_color=RENK_YAZI2,
                     font=("Segoe UI",10)).pack(side="left")

        ml_frame = ctk.CTkFrame(body, fg_color=RENK_PANEL, corner_radius=6)
        ml_frame.grid(row=4, column=0, columnspan=2, sticky="ew")

        def render_ml():
            for w in ml_frame.winfo_children(): w.destroy()
            for mail in kural.get("mail_list",[]):
                r = ctk.CTkFrame(ml_frame, fg_color=RENK_PANEL)
                r.pack(fill="x", padx=4, pady=2)
                ctk.CTkLabel(r, text=mail, text_color=RENK_YAZI,
                             font=("Segoe UI",10)).pack(side="left", padx=8)
                ctk.CTkButton(r, text="✕", width=22, height=20,
                              fg_color=RENK_KIRMIZI, hover_color="#a93226",
                              font=("Segoe UI",10),
                              command=lambda m=mail: (
                                  kural["mail_list"].remove(m), render_ml())
                              ).pack(side="right", padx=4)
            add_r = ctk.CTkFrame(ml_frame, fg_color=RENK_PANEL)
            add_r.pack(fill="x", padx=4, pady=(2,4))
            ent_new = ctk.CTkEntry(add_r, fg_color=RENK_KART,
                                    border_color=RENK_SINIR,
                                    text_color=RENK_YAZI, font=("Segoe UI",10),
                                    placeholder_text="yeni@mail.com")
            ent_new.pack(side="left", fill="x", expand=True, padx=(4,4))
            def ekle(w=ent_new):
                m = w.get().strip()
                if m and "@" in m and m not in kural.get("mail_list",[]):
                    kural.setdefault("mail_list",[]).append(m); render_ml()
            ctk.CTkButton(add_r, text="+ Ekle", width=60, height=22,
                          fg_color=renk, hover_color=RENK_VURGU2,
                          font=("Segoe UI",10), command=ekle
                          ).pack(side="right", padx=(0,4))
        render_ml()

        btn_row = ctk.CTkFrame(body, fg_color=RENK_KART)
        btn_row.grid(row=5, column=0, columnspan=2, sticky="e", pady=(8,0))
        ctk.CTkButton(btn_row, text="👁 Mail Önizle", width=110, height=26,
                      fg_color=RENK_VURGU2, hover_color=RENK_VURGU,
                      font=("Segoe UI",10),
                      command=lambda k=kural: MailOnizleme(
                          self, k.get("ad","?"),
                          k.get("mail_list",[]), k.get("keywords",[]))
                      ).pack(side="left", padx=(0,6))
        ctk.CTkButton(btn_row, text="🗑 Sil", width=70, height=26,
                      fg_color=RENK_KIRMIZI, hover_color="#a93226",
                      font=("Segoe UI",10),
                      command=lambda k=kural, f=frame: self._kural_sil(k,f)
                      ).pack(side="left")
        return frame, ent_ad, ent_kw

    def _kural_ekle(self):
        cfg = load_config()
        cfg.setdefault("kurallar",[]).append({
            "id": str(uuid.uuid4())[:8], "ad":"Yeni Kural",
            "keywords":[], "mail_list":[]})
        save_config(cfg); self._render_kurallar()

    def _kural_sil(self, kural, frame):
        cfg = load_config()
        cfg["kurallar"] = [k for k in cfg.get("kurallar",[])
                           if k.get("id") != kural.get("id")]
        save_config(cfg); frame.destroy()
        self._kart_list = [i for i in self._kart_list if i[0] != frame]

    def _build_outlook(self, parent):
        parent.configure(fg_color=RENK_PANEL)
        ctk.CTkLabel(parent, text="Gönderici Mail Hesabı",
                     font=("Segoe UI",12,"bold"),
                     text_color=RENK_YAZI).pack(anchor="w", padx=10, pady=(12,4))
        kart = ctk.CTkFrame(parent, fg_color=RENK_KART, corner_radius=10)
        kart.pack(fill="x", padx=10, pady=4)
        cfg=load_config(); secili=cfg.get("outlook_account","")
        hesaplar=outlook_hesaplari()
        if hesaplar:
            ctk.CTkLabel(kart, text="Hangi hesaptan mail gönderileceğini seçin:",
                         text_color=RENK_YAZI2, font=("Segoe UI",10)
                         ).pack(anchor="w", padx=14, pady=(12,6))
            self.cmb_outlook = ctk.CTkComboBox(
                kart, values=hesaplar, fg_color=RENK_PANEL,
                border_color=RENK_SINIR, button_color=RENK_VURGU,
                text_color=RENK_YAZI, font=("Segoe UI",11), width=420)
            self.cmb_outlook.set(secili if secili in hesaplar else hesaplar[0])
            self.cmb_outlook.pack(padx=14, pady=(0,4), anchor="w")
            ctk.CTkLabel(kart, text="💡 Seçilmezse Outlook varsayılanı kullanılır.",
                         text_color=RENK_YAZI2, font=("Segoe UI",9)
                         ).pack(anchor="w", padx=14, pady=(2,12))
        else:
            ctk.CTkLabel(kart, text="⚠ Outlook açık değil veya hesap bulunamadı.",
                         text_color=RENK_SARI, font=("Segoe UI",11),
                         justify="left").pack(padx=14, pady=14, anchor="w")
            self.cmb_outlook = None

        test_k = ctk.CTkFrame(parent, fg_color=RENK_KART, corner_radius=10)
        test_k.pack(fill="x", padx=10, pady=(12,4))
        ctk.CTkLabel(test_k, text="Bağlantı Testi", font=("Segoe UI",11,"bold"),
                     text_color=RENK_YAZI).pack(anchor="w", padx=14, pady=(10,4))
        self.lbl_test = ctk.CTkLabel(test_k, text="", font=("Segoe UI",10),
                                      text_color=RENK_YESIL)
        self.lbl_test.pack(anchor="w", padx=14)
        ctk.CTkButton(test_k, text="📧 Test Maili Gönder", height=32,
                      fg_color=RENK_VURGU, hover_color=RENK_VURGU2,
                      font=("Segoe UI",11), command=self._test_mail
                      ).pack(anchor="w", padx=14, pady=(4,12))

    def _test_mail(self):
        cfg=load_config()
        from_acc=cfg.get("outlook_account","")
        hesaplar=outlook_hesaplari()
        test_to=from_acc or (hesaplar[0] if hesaplar else "")
        if not test_to:
            self.lbl_test.configure(text="❌ Hesap bulunamadı.",
                                    text_color=RENK_KIRMIZI); return
        ok,info=send_mail([test_to],"TEST","Test","Test mesajı.",
                          from_acc or None,"Test Grubu")
        if ok:
            self.lbl_test.configure(
                text=f"✅ Test gönderildi → {test_to}", text_color=RENK_YESIL)
        else:
            self.lbl_test.configure(text=f"❌ {info}", text_color=RENK_KIRMIZI)

    def _kaydet(self):
        cfg=load_config()
        cfg["wa_group_name"] = self.ent_grup.get().strip()
        try:
            cfg["bekleme_dk"] = max(1, int(self.ent_sure.get().strip()))
        except: cfg["bekleme_dk"] = 2
        if hasattr(self,"cmb_outlook") and self.cmb_outlook:
            cfg["outlook_account"] = self.cmb_outlook.get().strip()
        kurallar=[]
        for (frame,kural,ea,ek) in self._kart_list:
            if frame.winfo_exists():
                kural["ad"] = ea.get().strip()
                kural["keywords"] = [
                    x.strip().lower() for x in ek.get().split(",") if x.strip()]
                kurallar.append(kural)
        cfg["kurallar"] = kurallar
        save_config(cfg); self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# RAPOR
# ══════════════════════════════════════════════════════════════════════════════
class RaporPencere(ctk.CTkToplevel):
    def __init__(self,master):
        super().__init__(master)
        self.title("Aylık Rapor"); self.geometry("820x580")
        self.configure(fg_color=RENK_ANA_ARKA); self.grab_set()
        self._build(); self._goster(datetime.now().year,datetime.now().month)

    def _build(self):
        hdr=ctk.CTkFrame(self,fg_color=RENK_PANEL,height=50,corner_radius=0)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        ctk.CTkLabel(hdr,text="📊  Aylık Rapor",font=("Segoe UI",16,"bold"),
                     text_color=RENK_YAZI).pack(side="left",padx=16)
        now=datetime.now()
        yillar=[str(y) for y in range(now.year-2,now.year+1)]
        aylar=["Ocak","Şubat","Mart","Nisan","Mayıs","Haziran",
               "Temmuz","Ağustos","Eylül","Ekim","Kasım","Aralık"]
        nav=ctk.CTkFrame(self,fg_color=RENK_PANEL,corner_radius=8)
        nav.pack(fill="x",padx=14,pady=(10,6))
        self.cmb_yil=ctk.CTkComboBox(nav,values=yillar,width=90,
            fg_color=RENK_KART,border_color=RENK_SINIR,
            button_color=RENK_VURGU,text_color=RENK_YAZI)
        self.cmb_yil.set(str(now.year)); self.cmb_yil.pack(side="left",padx=(12,4),pady=8)
        self.cmb_ay=ctk.CTkComboBox(nav,values=aylar,width=120,
            fg_color=RENK_KART,border_color=RENK_SINIR,
            button_color=RENK_VURGU,text_color=RENK_YAZI)
        self.cmb_ay.set(aylar[now.month-1]); self.cmb_ay.pack(side="left",padx=4,pady=8)
        ctk.CTkButton(nav,text="🔍 Görüntüle",height=32,
                      fg_color=RENK_VURGU,hover_color=RENK_VURGU2,
                      command=self._ara).pack(side="left",padx=8)
        self.frm_ozet=ctk.CTkFrame(self,fg_color=RENK_ANA_ARKA)
        self.frm_ozet.pack(fill="x",padx=14,pady=(0,6))
        tbl=ctk.CTkFrame(self,fg_color=RENK_PANEL,corner_radius=8)
        tbl.pack(fill="both",expand=True,padx=14,pady=(0,14))
        st=ttk.Style(); st.theme_use("default")
        st.configure("D.Treeview",background=RENK_KART,foreground=RENK_YAZI,
            fieldbackground=RENK_KART,rowheight=26,font=("Segoe UI",10))
        st.configure("D.Treeview.Heading",background=RENK_PANEL,
            foreground=RENK_YAZI2,font=("Segoe UI",10,"bold"))
        st.map("D.Treeview",background=[("selected",RENK_VURGU)])
        cols=("Tarih","Gönderen","Kural","Mesaj")
        self.tree=ttk.Treeview(tbl,columns=cols,show="headings",style="D.Treeview")
        for col,w in zip(cols,[120,140,110,350]):
            self.tree.heading(col,text=col); self.tree.column(col,width=w,anchor="w")
        sb=ttk.Scrollbar(tbl,orient="vertical",command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left",fill="both",expand=True,padx=4,pady=4)
        sb.pack(side="right",fill="y")

    def _ara(self):
        aylar=["Ocak","Şubat","Mart","Nisan","Mayıs","Haziran",
               "Temmuz","Ağustos","Eylül","Ekim","Kasım","Aralık"]
        self._goster(int(self.cmb_yil.get()),aylar.index(self.cmb_ay.get())+1)

    def _goster(self,yil,ay):
        rows,gunluk=db_rapor(yil,ay)
        for w in self.frm_ozet.winfo_children(): w.destroy()
        t=len(rows); self._krt("Toplam Red",str(t),RENK_KIRMIZI)
        gs=len(gunluk)
        self._krt("Günlük Ort.",str(round(t/gs,1) if gs else 0),RENK_SARI)
        if gunluk:
            en=max(gunluk,key=gunluk.get)
            self._krt("En Yoğun Gün",f"{en}\n({gunluk[en]})",RENK_VURGU)
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            self.tree.insert("","end",values=(
                r[0][:19].replace("T"," "),r[1],r[2],(r[3] or "")[:80]))

    def _krt(self,b,d,r):
        f=ctk.CTkFrame(self.frm_ozet,fg_color=RENK_KART,corner_radius=8)
        f.pack(side="left",padx=(0,8),pady=6,ipadx=16,ipady=8)
        ctk.CTkLabel(f,text=b,text_color=RENK_YAZI2,font=("Segoe UI",10)).pack()
        ctk.CTkLabel(f,text=d,text_color=r,font=("Segoe UI",18,"bold")).pack()


# ══════════════════════════════════════════════════════════════════════════════
# SPLASH
# ══════════════════════════════════════════════════════════════════════════════
class SplashEkran(tk.Toplevel):
    def __init__(self,master):
        super().__init__(master)
        self.overrideredirect(True); self.configure(bg="#071320")
        self.attributes("-topmost",True)
        W,H=600,340
        self.geometry(f"{W}x{H}+{(self.winfo_screenwidth()-W)//2}+"
                      f"{(self.winfo_screenheight()-H)//2}")
        p=os.path.join(BASE_DIR,"splash.png")
        if PIL_OK and os.path.exists(p):
            img=Image.open(p).resize((W,H),Image.LANCZOS)
            self._ph=ImageTk.PhotoImage(img)
            tk.Label(self,image=self._ph,bg="#071320",bd=0
                     ).pack(fill="both",expand=True)
        else:
            tk.Label(self,text="Pregate Kayıt Red\nWA → Mail Botu",
                     font=("Segoe UI",20,"bold"),fg="#e8f0f7",
                     bg="#071320").pack(expand=True)
            tk.Label(self,text="S.SEYMEN tarafından hazırlanmıştır",
                     font=("Segoe UI",11),fg="#8faabf",
                     bg="#071320").pack(pady=(0,20))
        self.configure(highlightbackground="#1a6ea8",highlightthickness=1)
        bf=tk.Frame(self,bg="#0d1f30",height=4)
        bf.place(x=0,y=H-4,width=W,height=4)
        self._bar=tk.Frame(bf,bg="#1a6ea8",height=4)
        self._bar.place(x=0,y=0,width=0,height=4)
        self._W=W; self._step=0; self._anim()

    def _anim(self):
        self._step+=1
        self._bar.place(x=0,y=0,
                        width=min(int(self._step/30*self._W),self._W),height=4)
        if self._step<30: self.after(50,self._anim)

    def kapat(self): self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# ANA UYGULAMA
# ══════════════════════════════════════════════════════════════════════════════
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Pregate Kayıt Red – WA Mail Botu  |  Poliport")
        self.geometry("900x640"); self.minsize(800,540)
        self.configure(fg_color=RENK_ANA_ARKA)
        self.wa_bot=None; self.running=False; self.mail_say=0
        db_init(); self._build_ui()

    def _build_ui(self):
        hdr=ctk.CTkFrame(self,fg_color=RENK_PANEL,corner_radius=0,height=56)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        ctk.CTkLabel(hdr,text="📱  Pregate Kayıt Red – WA Mail Botu",
                     font=("Segoe UI",16,"bold"),
                     text_color=RENK_YAZI).pack(side="left",padx=20,pady=12)
        self.lbl_durum=ctk.CTkLabel(hdr,text="● Beklemede",
                                     font=("Segoe UI",12),text_color=RENK_YAZI2)
        self.lbl_durum.pack(side="right",padx=20)
        content=ctk.CTkFrame(self,fg_color=RENK_ANA_ARKA)
        content.pack(fill="both",expand=True,padx=14,pady=10)
        content.columnconfigure(0,weight=0); content.columnconfigure(1,weight=1)
        content.rowconfigure(0,weight=1)
        self._build_sol(content); self._build_sag(content)

    def _build_sol(self,parent):
        sol=ctk.CTkFrame(parent,fg_color=RENK_PANEL,corner_radius=10,width=200)
        sol.grid(row=0,column=0,sticky="nsew",padx=(0,7)); sol.pack_propagate(False)
        ctk.CTkLabel(sol,text="KONTROL",font=("Segoe UI",11,"bold"),
                     text_color=RENK_YAZI2).pack(anchor="w",padx=14,pady=(14,4))
        self.btn_baslat=ctk.CTkButton(sol,text="▶  Başlat",
            fg_color=RENK_YESIL,hover_color="#17a844",
            font=("Segoe UI",13,"bold"),height=42,command=self._baslat)
        self.btn_baslat.pack(fill="x",padx=12,pady=4)
        self.btn_durdur=ctk.CTkButton(sol,text="■  Durdur",
            fg_color=RENK_KIRMIZI,hover_color="#c0392b",
            font=("Segoe UI",13,"bold"),height=42,
            state="disabled",command=self._durdur)
        self.btn_durdur.pack(fill="x",padx=12,pady=4)
        ctk.CTkFrame(sol,height=1,fg_color=RENK_SINIR).pack(fill="x",padx=12,pady=12)
        ctk.CTkButton(sol,text="⚙  Ayarlar",height=36,
                      fg_color=RENK_VURGU,hover_color=RENK_VURGU2,
                      font=("Segoe UI",12),command=self._ayarlar
                      ).pack(fill="x",padx=12,pady=4)
        ctk.CTkButton(sol,text="📊  Rapor",height=36,
                      fg_color="#2c3e6b",hover_color="#1e2d50",
                      font=("Segoe UI",12),command=self._rapor
                      ).pack(fill="x",padx=12,pady=4)
        ctk.CTkFrame(sol,height=1,fg_color=RENK_SINIR).pack(fill="x",padx=12,pady=12)
        f1=ctk.CTkFrame(sol,fg_color=RENK_KART,corner_radius=8)
        f1.pack(fill="x",padx=12,pady=4)
        ctk.CTkLabel(f1,text="Bugün Gönderilen",
                     font=("Segoe UI",9),text_color=RENK_YAZI2).pack(pady=(8,0))
        self.lbl_sayac=ctk.CTkLabel(f1,text="0",
                                     font=("Segoe UI",26,"bold"),text_color=RENK_YESIL)
        self.lbl_sayac.pack(pady=(0,8))
        cfg=load_config()
        self.lbl_grup=ctk.CTkLabel(sol,
            text=f"Grup: {cfg.get('wa_group_name','—')}",
            font=("Segoe UI",9),text_color=RENK_YAZI2,wraplength=170)
        self.lbl_grup.pack(padx=12,pady=(16,4),anchor="w")
        dk=cfg.get("bekleme_dk",2)
        self.lbl_sure=ctk.CTkLabel(sol,
            text=f"Biriktirme: {dk} dk",
            font=("Segoe UI",9),text_color=RENK_YAZI2)
        self.lbl_sure.pack(padx=12,pady=(0,4),anchor="w")

    def _build_sag(self,parent):
        sag=ctk.CTkFrame(parent,fg_color=RENK_PANEL,corner_radius=10)
        sag.grid(row=0,column=1,sticky="nsew")
        sag.columnconfigure(0,weight=1); sag.rowconfigure(1,weight=1)
        ctk.CTkLabel(sag,text="İŞLEM LOGU",font=("Segoe UI",11,"bold"),
                     text_color=RENK_YAZI2).grid(row=0,column=0,
                                                  sticky="w",padx=14,pady=(14,4))
        self.txt_log=ctk.CTkTextbox(sag,fg_color=RENK_KART,text_color=RENK_YAZI,
                                     font=("Consolas",10),corner_radius=8)
        self.txt_log.grid(row=1,column=0,sticky="nsew",padx=14,pady=(0,14))

    def _baslat(self):
        cfg=load_config()
        if not cfg.get("wa_group_name","").strip():
            messagebox.showwarning("Uyarı","Ayarlar → WhatsApp'tan grup adını girin!")
            return
        if not cfg.get("kurallar"):
            messagebox.showwarning("Uyarı","Ayarlar → Kurallar'dan en az bir kural girin!")
            return
        self.running=True
        self.btn_baslat.configure(state="disabled")
        self.btn_durdur.configure(state="normal")
        self._log("🚀 Bot başlatılıyor…")
        # Bekleme süresini güncelle
        global BEKLEME_SURE
        BEKLEME_SURE = cfg.get("bekleme_dk", 2) * 60
        self.wa_bot=WABot(on_log=self._log,
                          on_status=self._set_durum_p,
                          on_message=self._on_mesaj_bitti)
        threading.Thread(target=self.wa_bot.start, daemon=True).start()

    def _durdur(self):
        self.running=False
        if self.wa_bot: self.wa_bot.stop(); self.wa_bot=None
        self.btn_baslat.configure(state="normal")
        self.btn_durdur.configure(state="disabled")
        self._set_durum("● Durduruldu",RENK_KIRMIZI)
        self._log("■ Bot durduruldu.")

    def _ayarlar(self):
        AyarlarPencere(self)
        self.after(500, self._guncelle_labels)

    def _guncelle_labels(self):
        cfg=load_config()
        self.lbl_grup.configure(text=f"Grup: {cfg.get('wa_group_name','—')}")
        self.lbl_sure.configure(text=f"Biriktirme: {cfg.get('bekleme_dk',2)} dk")

    def _rapor(self): RaporPencere(self)

    def _on_mesaj_bitti(self, gonderen, birlesik_metin):
        """
        Biriktiriciden gelen callback — 2 dk doldu, gönder.
        birlesik_metin = kişinin 2 dk içinde attığı tüm mesajlar \n ile birleşik.
        """
        cfg = load_config()
        metin_lower = birlesik_metin.lower()
        eslesen = []
        for kural in cfg.get("kurallar",[]):
            for kw in kural.get("keywords",[]):
                if kw.lower() in metin_lower:
                    eslesen.append(kural); break

        if not eslesen:
            self._log(f"💬 [{gonderen}]: eşleşme yok — {birlesik_metin[:60]}")
            return

        from_acc = cfg.get("outlook_account","") or None
        for kural in eslesen:
            ml = kural.get("mail_list",[])
            if not ml:
                self._log(f"⚠ '{kural.get('ad')}' mail listesi boş!"); continue
            ok, info = send_mail(ml, kural["ad"], gonderen,
                                 birlesik_metin, from_acc,
                                 cfg.get("wa_group_name",""))
            if ok:
                self.mail_say+=1
                self.after(0, lambda: self.lbl_sayac.configure(
                    text=str(self.mail_say)))
                db_ekle(gonderen, kural["ad"], birlesik_metin, False)
                self._log(f"✉ [{kural['ad']}] {gonderen} → "
                          f"{', '.join(ml[:2])}{'…' if len(ml)>2 else ''}")
            else:
                self._log(f"❌ [{kural['ad']}] Mail hatası: {info}")

    def _log(self,msg):
        ts=datetime.now().strftime("%H:%M:%S")
        def _do():
            self.txt_log.configure(state="normal")
            self.txt_log.insert("end",f"[{ts}] {msg}\n")
            self.txt_log.see("end")
            self.txt_log.configure(state="disabled")
        self.after(0,_do)

    def _set_durum(self,text,color):
        self.after(0,lambda: self.lbl_durum.configure(text=text,text_color=color))

    def _set_durum_p(self,_,text,color): self._set_durum(text,color)

    def on_close(self):
        if self.wa_bot: self.wa_bot.stop()
        self.destroy()


if __name__=="__main__":
    root=tk.Tk(); root.withdraw()
    splash=SplashEkran(root); splash.update()
    def _ac():
        splash.kapat(); root.destroy()
        app=App()
        app.protocol("WM_DELETE_WINDOW",app.on_close)
        app.mainloop()
    root.after(2500,_ac)
    root.mainloop()
