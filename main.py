import customtkinter as ctk
import threading
import json, os, sys, time, uuid, tempfile, glob
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
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException, WebDriverException
    SELENIUM_OK = True
except ImportError:
    SELENIUM_OK = False

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

CONFIG_FILE  = os.path.join(BASE_DIR, "config.json")
DB_FILE      = os.path.join(BASE_DIR, "kayitlar.db")
PROFILE_DIR  = os.path.join(BASE_DIR, "chrome_profile")
os.makedirs(PROFILE_DIR, exist_ok=True)

BEKLEME_SURE = 120

# ── Türkçe-güvenli normalize ──────────────────────────────────────────────
# Python'un .lower() metodu Türkçe büyük 'İ' harfini 'i' yerine
# 'i' + birleşik nokta işaretine (U+0307) çevirir, bu da anahtar kelime
# eşleşmesini bozar (örn. "KİMYA" -> "ki̇mya" olur, "kimya" ile eşleşmez).
# Bu yüzden küçültmeden ÖNCE Türkçe büyük harfleri elle çeviriyoruz.
def turkce_norm(s):
    if not s:
        return ""
    s = (s.replace("İ", "i").replace("I", "ı")
           .replace("Ğ", "g").replace("Ü", "u")
           .replace("Ş", "s").replace("Ö", "o")
           .replace("Ç", "c"))
    s = s.lower()
    s = (s.replace("ı", "i").replace("ğ", "g").replace("ü", "u")
           .replace("ş", "s").replace("ö", "o").replace("ç", "c")
           .replace("\u0307", ""))  # kalan olası birleşik nokta işaretlerini temizle
    return s

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

    # ── kayitlar tablosu ──────────────────────────────────────────────────────
    con.execute("""CREATE TABLE IF NOT EXISTS kayitlar (
        id TEXT PRIMARY KEY, tarih TEXT, gonderen TEXT,
        kural_ad TEXT, mesaj TEXT, resim_var INTEGER DEFAULT 0)""")

    # ── seen_mesajlar tablosu ─────────────────────────────────────────────────
    # Eski sürümlerde bu tablo sadece msg_id ile oluşturulmuştu (zaman kolonu yok).
    # CREATE TABLE IF NOT EXISTS şema değişikliği yapmaz, bu yüzden migration gerekli.
    con.execute("""CREATE TABLE IF NOT EXISTS seen_mesajlar (
        msg_id TEXT PRIMARY KEY, zaman REAL DEFAULT 0)""")

    # Migration: zaman kolonu yoksa ekle (eski DB'lerle uyumluluk)
    try:
        con.execute("SELECT zaman FROM seen_mesajlar LIMIT 1")
    except _sqlite3.OperationalError:
        # zaman kolonu yok — ekle ve tüm mevcut kayıtlara şimdiki zamanı yaz
        con.execute("ALTER TABLE seen_mesajlar ADD COLUMN zaman REAL DEFAULT 0")
        con.execute("UPDATE seen_mesajlar SET zaman = ?", (time.time(),))

    # 48 saatten eski seen kayıtlarını temizle
    con.execute("DELETE FROM seen_mesajlar WHERE zaman < ? AND zaman != 0",
                (time.time() - 172800,))

    con.commit()
    con.close()

def db_seen_yukle():
    """Uygulama başlarken son 48 saatin seen msg_id'lerini belleğe al."""
    try:
        con = _sqlite3.connect(DB_FILE)
        rows = con.execute(
            "SELECT msg_id FROM seen_mesajlar WHERE zaman > ?",
            (time.time() - 172800,)).fetchall()
        con.close()
        return {r[0]: time.time() for r in rows}
    except:
        return {}

def db_seen_kaydet(msg_id):
    """Bir mesaj işlendiğinde DB'ye kaydet (restart'a karşı dayanıklı)."""
    try:
        con = _sqlite3.connect(DB_FILE)
        con.execute("INSERT OR REPLACE INTO seen_mesajlar VALUES (?,?)",
                    (msg_id, time.time()))
        con.commit(); con.close()
    except: pass

def db_ekle(gonderen, kural_ad, mesaj, resim_var):
    con = _sqlite3.connect(DB_FILE)
    con.execute("INSERT INTO kayitlar VALUES (?,?,?,?,?,?)",
        (str(uuid.uuid4()), datetime.now().isoformat(),
         gonderen, kural_ad, mesaj, 1 if resim_var else 0))
    con.commit(); con.close()

def gonderen_goster(sender: str) -> str:
    """WhatsApp'tan gelen gönderen adını/numarasını mailde gösterilecek şekle getirir.
    Eğer numara formatındaysa (başında + ya da 7+ rakam) → son 4 haneyi göster (****1234).
    Kayıtlı kişiyse (metin / isim) → ismi olduğu gibi kullan."""
    import re
    digits = re.sub(r'\D', '', sender)
    if len(digits) >= 7:
        return f"****{digits[-4:]}"
    return sender.strip() or "Bilinmiyor"

def db_bugun_gonderildi_mi(gonderen: str, metin: str) -> bool:
    """Bugün (00:00'dan itibaren) aynı gönderen + aynı içerik için zaten
    mail atıldıysa True döner — tekrar gönderimi engeller."""
    try:
        bugun = datetime.now().strftime("%Y-%m-%d")
        norm_metin = turkce_norm(metin)[:200]
        con = _sqlite3.connect(DB_FILE)
        rows = con.execute(
            "SELECT mesaj FROM kayitlar WHERE gonderen=? AND tarih LIKE ?",
            (gonderen, f"{bugun}%")).fetchall()
        con.close()
        for (m,) in rows:
            if turkce_norm(m or "")[:200] == norm_metin:
                return True
        return False
    except:
        return False

def db_rapor(yil, ay):
    con = _sqlite3.connect(DB_FILE)
    rows = con.execute(
        "SELECT tarih,gonderen,kural_ad,mesaj,resim_var FROM kayitlar "
        "WHERE tarih LIKE ? ORDER BY tarih",
        (f"{yil:04d}-{ay:02d}%",)).fetchall()
    gunluk = {}
    for r in rows:
        g=r[0][:10]; gunluk[g]=gunluk.get(g,0)+1
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

def send_mail(mail_list, kural_ad, gonderen, mesaj, from_account, grup_adi, img_paths=None):
    if not OUTLOOK_OK: return False, "pywin32 yüklü değil"
    sonuc = [False, ""]
    def _gonder():
        try:
            import pythoncom
            pythoncom.CoInitialize()
            tarih = datetime.now().strftime("%d.%m.%Y %H:%M")
            ol    = win32com.client.Dispatch("Outlook.Application")
            mail  = ol.CreateItem(0)
            mail.Subject = f"[KAYIT RED] {kural_ad} – {tarih}"
            mail.Body    = MAIL_SABLON.format(
                tarih=tarih, gonderen=gonderen,
                grup=grup_adi, mesaj=mesaj)
            mail.To = "; ".join(mail_list)
            # Resimleri ek olarak ekle
            if img_paths:
                for p in img_paths:
                    if p and os.path.exists(p):
                        try: mail.Attachments.Add(p)
                        except: pass
            if from_account:
                try:
                    for acc in ol.GetNamespace("MAPI").Accounts:
                        if acc.SmtpAddress.lower() == from_account.lower():
                            mail.SendUsingAccount = acc; break
                except: pass
            mail.Send()
            sonuc[0] = True; sonuc[1] = "OK"
        except Exception as e:
            sonuc[0] = False; sonuc[1] = str(e)
        finally:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except: pass
    t = threading.Thread(target=_gonder)
    t.start(); t.join(timeout=30)
    return sonuc[0], sonuc[1]

# ══════════════════════════════════════════════════════════════════════════════
# MESAJ BİRİKTİRİCİ
# ══════════════════════════════════════════════════════════════════════════════
class MesajBiriktiric:
    def __init__(self, on_gonder, bekleme=BEKLEME_SURE):
        self.on_gonder   = on_gonder
        self.bekleme     = bekleme
        self._kuyruk     = {}
        self._gonderildi = set()
        self._lock       = threading.Lock()

    def ekle(self, gonderen, metin, img_paths=None):
        """Yeni mesaj geldi — biriktiriciye ekle, timer sıfırla."""
        with self._lock:
            self._gonderildi.discard(gonderen)
            if gonderen not in self._kuyruk:
                self._kuyruk[gonderen] = {
                    "metinler": [], "resimler": [], "timer": None}
            self._kuyruk[gonderen]["metinler"].append(metin)
            if img_paths:
                self._kuyruk[gonderen]["resimler"].extend(img_paths)
            t = self._kuyruk[gonderen]["timer"]
            if t: t.cancel()
            yeni_t = threading.Timer(
                self.bekleme, self._gonder, args=[gonderen])
            yeni_t.daemon = True
            yeni_t.start()
            self._kuyruk[gonderen]["timer"] = yeni_t

    def _gonder(self, gonderen):
        """Bekleme süresi doldu — biriken mesajları ve resimleri gönder."""
        with self._lock:
            if gonderen not in self._kuyruk: return
            veri     = self._kuyruk.pop(gonderen)
            metinler = veri["metinler"]
            resimler = veri["resimler"]
            self._gonderildi.add(gonderen)
        if metinler:
            # on_gonder çağrısı sonucunda WA yanıtı ne olacağını döndürür
            self.on_gonder(gonderen, "\n".join(metinler), resimler)

    def temizle(self):
        with self._lock:
            for v in self._kuyruk.values():
                if v["timer"]: v["timer"].cancel()
            self._kuyruk.clear()

# ══════════════════════════════════════════════════════════════════════════════
# WHATSAPP WEB BOT — Selenium + JavaScript ile mesaj okuma
# QR bir kez taranır, sonra profil kaydedilir
# ══════════════════════════════════════════════════════════════════════════════
class WABot:
    # JavaScript: sayfadaki son gelen mesajları oku (metin + resim URL + zaman)
    JS_MESAJLARI_OKU = r"""
    const mesajlar = [];
    const rows = document.querySelectorAll('[data-testid="msg-container"]');
    let sonGonderen = ''; // WhatsApp art arda gelen mesajlarda gönderen adını
                          // sadece ilk mesajda gösterir; sonrakiler için bunu kullan
    rows.forEach((el, idx) => {
        // ── Giden mesajları atla (botun kendi yazdıkları) ──────────────────
        // WhatsApp Web farklı versiyonlarda farklı class/attribute kullanır,
        // bu yüzden birden fazla yöntemle kontrol ediyoruz:
        // 1) Eski sürüm class adı
        if (el.closest('[class*="message-out"]')) return;
        // 2) Yeni sürüm: data-id "true_" ile başlıyorsa giden mesaj
        const dataId = el.getAttribute('data-id') || '';
        if (dataId.startsWith('true_')) return;
        // 3) focusable wrapper'da "out" geçiyorsa
        const wrapper = el.closest('[class*="focusable"]') ||
                        el.closest('[class*="_akbu"]') ||
                        el.parentElement;
        if (wrapper) {
            const cls = wrapper.className || '';
            if (cls.includes('out') || cls.includes('_akbu')) return;
        }
        // 4) Mesaj kutusunun sağ tarafta olup olmadığını kontrol et
        //    (giden mesajlar sağa yaslanır)
        const rect = el.getBoundingClientRect();
        const parentRect = el.parentElement ?
                           el.parentElement.getBoundingClientRect() : null;
        if (parentRect && rect.left > parentRect.left + parentRect.width * 0.5) return;

        // Metin
        let text = '';
        const selectors = [
            '[data-testid="msg-text"]',
            'span.selectable-text.copyable-text',
            'span[class*="selectable-text"]',
            'div[class*="copyable-text"] span[dir]',
        ];
        for (const sel of selectors) {
            const el2 = el.querySelector(sel);
            if (el2 && el2.innerText) { text = el2.innerText.trim(); break; }
        }

        // Gönderen
        let sender = '';
        const copyable = el.querySelector('[data-pre-plain-text]');
        if (copyable) {
            const pre = copyable.getAttribute('data-pre-plain-text') || '';
            const m = pre.match(/] (.+?):\s*$/);
            if (m) sender = m[1].trim();
        }
        if (!sender) {
            const sEl = el.querySelector('[data-testid="author"]');
            if (sEl) sender = sEl.innerText || '';
        }
        if (sender) {
            sonGonderen = sender; // bilinen gönderen güncellendi
        } else if (sonGonderen) {
            // Art arda gelen mesajda WhatsApp gönderen göstermiyor —
            // bir önceki mesajdaki gönderenden devam et.
            sender = sonGonderen;
        }

        // Mesaj zamanı (data-pre-plain-text içinde: "[SS:DD, GG.AA.YYYY]")
        let msgTime = null;
        const copyable2 = el.querySelector('[data-pre-plain-text]');
        if (copyable2) {
            const pre = copyable2.getAttribute('data-pre-plain-text') || '';
            const tm = pre.match(/\[(\d{1,2}:\d{2}),\s*(\d{1,2}\.\d{1,2}\.\d{4})\]/);
            if (tm) {
                const [h, min] = tm[1].split(':').map(Number);
                const [d, mo, y] = tm[2].split('.').map(Number);
                msgTime = new Date(y, mo-1, d, h, min).getTime() / 1000;
            }
        }

        // Resim URL'leri
        const imgUrls = [];
        el.querySelectorAll('img').forEach(img => {
            const src = img.src || '';
            if (src.startsWith('blob:') || src.startsWith('https://')) {
                imgUrls.push(src);
            }
        });

        const msgId = el.getAttribute('data-id') ||
                      el.getAttribute('data-key') ||
                      (sender + '::' + (msgTime || idx) + '::' + text);

        if (text || imgUrls.length > 0) {
            mesajlar.push({
                id: msgId,
                text: text,
                sender: sender || 'Bilinmiyor',
                imgUrls: imgUrls,
                msgTime: msgTime
            });
        }
    });
    return mesajlar.slice(-30);
    """

    # JavaScript: blob URL'yi base64'e çevir
    JS_BLOB_TO_B64 = """
    const url = arguments[0];
    return new Promise((resolve) => {
        fetch(url)
            .then(r => r.blob())
            .then(b => {
                const reader = new FileReader();
                reader.onload = () => resolve(reader.result);
                reader.onerror = () => resolve(null);
                reader.readAsDataURL(b);
            })
            .catch(() => resolve(null));
    });
    """

    def __init__(self, on_log, on_status, on_message):
        self.on_log     = on_log
        self.on_status  = on_status
        self.on_message = on_message
        self.running          = False
        self._driver          = None
        self._manuel_durdurma = False  # True ise otomatik restart yapılmaz
        self._baglanti_hata_say = 0    # ard arda bağlantı hata sayacı
        # Restart sonrasında eski mesajları tekrar işlememek için
        # seen listesi DB'den yükleniyor (48 saatlik hafıza)
        self._seen       = db_seen_yukle()
        self._biriktiric = None
        self._baslangic  = None  # bot başlama zamanı

    def start(self):
        self.running = True
        self._manuel_durdurma = False  # yeni başlatmada flag sıfırla
        global BEKLEME_SURE
        cfg = load_config()
        BEKLEME_SURE = cfg.get("bekleme_dk", 2) * 60
        self._biriktiric = MesajBiriktiric(
            on_gonder=self.on_message,
            bekleme=BEKLEME_SURE)

        self.on_status("START", "● Başlatılıyor…", RENK_SARI)
        self._baslangic = time.time()  # Bu andan önceki mesajları işleme
        self.on_log("🌐 Chrome açılıyor (WhatsApp Web)…")

        try:
            opts = Options()
            opts.add_argument(f"--user-data-dir={PROFILE_DIR}")
            opts.add_argument("--profile-directory=WABot")
            opts.add_argument("--no-sandbox")
            opts.add_argument("--disable-dev-shm-usage")
            opts.add_argument("--disable-notifications")
            opts.add_argument("--disable-popup-blocking")
            opts.add_argument("--no-first-run")
            opts.add_argument("--no-default-browser-check")
            opts.add_argument("--disable-background-timer-throttling")
            opts.add_argument("--disable-backgrounding-occluded-windows")
            opts.add_argument("--disable-renderer-backgrounding")
            opts.add_experimental_option("excludeSwitches", ["enable-automation"])
            opts.add_experimental_option("useAutomationExtension", False)

            # ChromeDriver otomatik versiyon eşleşmesi
            try:
                from selenium.webdriver.chrome.service import Service
                from webdriver_manager.chrome import ChromeDriverManager
                service = Service(ChromeDriverManager().install())
                self._driver = webdriver.Chrome(service=service, options=opts)
            except Exception:
                self._driver = webdriver.Chrome(options=opts)
            self._driver.get("https://web.whatsapp.com")

            self.on_log("📷 İlk kullanımda QR taratın — sonra otomatik giriş yapar.")
            self.on_status("QR", "● QR Bekleniyor", RENK_SARI)

            # Giriş bekle (maks 3 dk)
            self.on_log("⏳ WhatsApp Web yükleniyor…")
            WebDriverWait(self._driver, 180).until(
                lambda d: d.execute_script(
                    "return document.querySelector('[data-testid=\"chat-list\"],"
                    "[aria-label*=\"Sohbet\"], [aria-label*=\"Chat\"]') !== null"
                )
            )

            self.on_log("✅ WhatsApp Web bağlandı!")
            self._baglanti_hata_say = 0  # başarılı bağlantıda sayacı sıfırla
            self.on_status("OK", "● Bağlandı ✓", RENK_YESIL)

            # Gruba git
            self._gruba_git()

        except TimeoutException:
            self._baglanti_hata_say += 1
            self.on_log(f"⏱ WA bağlantısı kurulamadı (deneme {self._baglanti_hata_say}) — yeniden deneniyor…")
            self.on_status("ERR", "● Yeniden bağlanıyor…", RENK_SARI)
            if self._baglanti_hata_say >= 3:
                self._wa_erisim_uyarisi_gonder()
            self._otomatik_yeniden_baslat()
        except WebDriverException as e:
            self._baglanti_hata_say += 1
            self.on_log(f"❌ Chrome hatası (deneme {self._baglanti_hata_say}): {str(e)[:80]} — yeniden deneniyor…")
            self.on_status("ERR", "● Yeniden bağlanıyor…", RENK_SARI)
            if self._baglanti_hata_say >= 3:
                self._wa_erisim_uyarisi_gonder()
            self._otomatik_yeniden_baslat()

    def _gruba_git(self):
        cfg  = load_config()
        grup = cfg.get("wa_group_name","").strip()
        if not grup:
            self.on_log("⚠ Grup adı boş! Ayarlar'dan girin.")
            return

        self.on_log(f"🔍 '{grup}' grubu aranıyor…")
        try:
            # Arama kutusunu bul — birden fazla selector dene
            sb = None
            for sel in [
                '[data-testid="chat-list-search"]',
                '[aria-label="Sohbet veya kişi ara"]',
                '[aria-label="Search or start new chat"]',
                'div[contenteditable="true"]',
            ]:
                try:
                    sb = WebDriverWait(self._driver, 5).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, sel)))
                    break
                except: continue

            if sb is None:
                raise Exception("Arama kutusu bulunamadı — tüm selector'lar denendi")

            sb.click()
            time.sleep(0.5)
            sb.send_keys(grup)
            time.sleep(2)

            # İlk sonucu tıkla
            for sel in [
                f'[title="{grup}"]',
                '[data-testid="cell-frame-container"]',
                'div[role="listitem"]',
            ]:
                try:
                    el = WebDriverWait(self._driver, 5).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, sel)))
                    el.click()
                    break
                except: continue

            time.sleep(2)
            self.on_log(f"📌 '{grup}' grubuna girildi.")
            self.on_log("👂 Mesajlar izleniyor (10 sn'de bir kontrol)…")
            self._dinle()

        except Exception as e:
            self.on_log(f"⚠ Grup açma hatası: {e}")
            self.on_log("ℹ Grubu WhatsApp Web'de manuel açın, bot dinlemeye devam eder.")
            self._dinle()

    def _dinle(self):
        """Her 10 sn'de JS ile mesajları oku."""
        # Başlangıçta mevcut mesajları seen'e al — mail gönderme
        self.on_log("📋 Mevcut mesajlar taranıyor (mail gönderilmeyecek)…")
        try:
            mesajlar = self._driver.execute_script(self.JS_MESAJLARI_OKU)
            if mesajlar:
                now_ts = time.time()
                for m in mesajlar:
                    mid = m.get("id","")
                    if mid:
                        self._seen[mid] = now_ts
                self.on_log(f"✅ {len(mesajlar)} eski mesaj atlandı. Yeni mesajlar bekleniyor…")
        except: pass
        dongu = 0
        while self.running:
            try:
                self._mesajlari_oku()
                dongu += 1
                # Her ~5 dakikada bir sayfa sağlık kontrolü
                if dongu % 30 == 0:
                    self._saglik_kontrol()
                # Her ~1 dakikada bir gece yarısı temizlik kontrolü
                if dongu % 6 == 0:
                    self._gece_yarisi_kontrol()
            except WebDriverException as e:
                if "chrome not reachable" in str(e).lower() or \
                   "disconnected" in str(e).lower() or \
                   "session deleted" in str(e).lower():
                    self._baglanti_hata_say += 1
                    self.on_log(f"❌ Chrome bağlantısı koptu (deneme {self._baglanti_hata_say}) — otomatik yeniden başlatılıyor…")
                    self.on_status("ERR", "● Yeniden bağlanıyor…", RENK_SARI)
                    if self._baglanti_hata_say >= 3:
                        self._wa_erisim_uyarisi_gonder()
                    self._otomatik_yeniden_baslat()
                    return
                self.on_log(f"⚠ {str(e)[:60]}")
            except Exception as e:
                self.on_log(f"⚠ Hata: {str(e)[:60]}")
            time.sleep(10)

    def _mesajlari_oku(self):
        """JavaScript ile sayfadaki mesajları oku.
        Önce aktif sohbetin doğru grup olduğunu kontrol eder."""
        try:
            # Aktif sohbet başlığını kontrol et — yanlış grupta mesaj okuma
            cfg_now = load_config()
            hedef_grup = cfg_now.get("wa_group_name","").strip()
            if hedef_grup:
                try:
                    baslik = self._driver.execute_script("""
                        const el = document.querySelector(
                            '[data-testid="conversation-info-header-chat-title"],' +
                            'header span[title],' +
                            'div[data-testid="conversation-header"] span'
                        );
                        return el ? (el.innerText || el.getAttribute('title') || '') : '';
                    """) or ""
                    if baslik.strip() and baslik.strip() != hedef_grup:
                        # Yanlış sohbetteyiz — sessizce doğru gruba geri dön
                        self._gruba_git_sessiz()
                        return
                except: pass

            mesajlar = self._driver.execute_script(self.JS_MESAJLARI_OKU)
            if not mesajlar:
                return

            now_ts = time.time()
            self._seen = {k:v for k,v in self._seen.items() if now_ts-v < 86400}  # 24 saat

            _norm = turkce_norm

            for m in mesajlar:
                msg_id   = m.get("id","")
                text     = m.get("text","").strip()
                sender   = m.get("sender","Bilinmiyor").strip()
                img_urls = m.get("imgUrls",[])
                msg_time = m.get("msgTime")  # mesajın gerçek zamanı

                if not msg_id or msg_id in self._seen:
                    continue
                self._seen[msg_id] = now_ts
                db_seen_kaydet(msg_id)  # restart'a karşı kalıcı hafıza

                # Bot başlamadan önceki mesajları atla
                if msg_time and self._baslangic:
                    if msg_time < self._baslangic - 30:  # 30sn tolerans
                        continue

                # Boş göndereni atla
                if not sender or sender == "Bilinmiyor":
                    continue

                # Botun kendi yazdığı yanıt mesajlarını atla:
                # WA'daki bot hesabı genelde "Pregate Kayıt Red" (grup adı) 
                # ya da config'deki grup adıyla aynı isimde görünür.
                # data-id "true_" ile başlıyorsa zaten JS'te filtrelendi,
                # ama güvenlik için Python'da da kontrol et.
                grup_adi_norm = turkce_norm(
                    load_config().get("wa_group_name",""))
                if turkce_norm(sender) == grup_adi_norm:
                    continue

                # Resimleri indir
                img_paths = []
                if img_urls:
                    try:
                        import base64 as _b64, tempfile as _tmp
                        # NOT: execute_script (senkron) bir Promise döndüğünde
                        # onu BEKLEMEDEN hemen geri döner — bu yüzden eski kod
                        # neredeyse her zaman boş sonuç alıyordu ve resimler
                        # sessizce ekrana eklenmeden geçiliyordu. Doğrusu
                        # execute_async_script kullanmaktır (callback'i bekler).
                        JS_MULTI_ASYNC = """
                        const urls = arguments[0];
                        const callback = arguments[arguments.length - 1];
                        (async () => {
                            const results = [];
                            for (const url of urls.slice(0,3)) {
                                try {
                                    const r = await fetch(url);
                                    const b = await r.blob();
                                    const buf = await b.arrayBuffer();
                                    const arr = Array.from(new Uint8Array(buf));
                                    let bin = '';
                                    for (const byte of arr) bin += String.fromCharCode(byte);
                                    results.push({data: btoa(bin), type: b.type});
                                } catch (e) { results.push(null); }
                            }
                            callback(results);
                        })();
                        """
                        self._driver.set_script_timeout(20)
                        results = self._driver.execute_async_script(
                            JS_MULTI_ASYNC, img_urls) or []
                        for res in results:
                            if res and res.get("data"):
                                try:
                                    raw = _b64.b64decode(res["data"])
                                    mime = res.get("type","image/jpeg")
                                    ext = ".png" if "png" in mime else ".jpg"
                                    p = os.path.join(_tmp.gettempdir(),
                                        f"wa_{int(time.time()*1000)}{ext}")
                                    with open(p,"wb") as f: f.write(raw)
                                    img_paths.append(p)
                                except: pass
                    except Exception as e:
                        self.on_log(f"⚠ Resim indirme hatası: {str(e)[:60]}")

                # Metin yoksa resim var mı bak
                if not text and img_paths:
                    text = "(Resim gönderildi)"
                if not text:
                    continue

                # Anahtar kelime kontrolü
                cfg_check = load_config()
                metin_n = _norm(text)
                eslesen_kw = False
                for kural in cfg_check.get("kurallar",[]):
                    for kw in kural.get("keywords",[]):
                        if _norm(kw) in metin_n:
                            eslesen_kw = True; break
                    if eslesen_kw: break

                if not eslesen_kw:
                    self.on_log(f"💬 [{sender}]: eşleşme yok — {text[:40]}")
                    continue

                self.on_log(f"📩 [{sender}]: {text[:60]}"
                            + (f"  📎{len(img_paths)}" if img_paths else ""))
                self._biriktiric.ekle(sender, text, img_paths)

        except Exception as e:
            raise e

    def _gece_yarisi_kontrol(self):
        """Her dakika çağrılır. Saat 00:00–00:01 arasında sohbeti temizler.
        Aynı gün içinde birden fazla temizlik yapılmaması için tarih takibi yapar."""
        now = datetime.now()
        bugun = now.strftime("%Y-%m-%d")
        # Sadece 00:00–00:01 arasında ve bugün henüz temizlemediyse
        if now.hour == 0 and now.minute == 0:
            if getattr(self, "_son_temizlik_tarihi", "") != bugun:
                self._son_temizlik_tarihi = bugun
                self.on_log("🌙 Gece yarısı — grup sohbeti temizleniyor…")
                self._sohbeti_temizle()

    def _sohbeti_temizle(self):
        """WhatsApp Web'de açık olan grup sohbetini 'Sohbeti Temizle' menüsü
        aracılığıyla yalnızca bu cihazdan siler (karşı tarafa görünmez)."""
        try:
            from selenium.webdriver.common.keys import Keys

            # 1) Grup başlığına tıkla (menüyü açmak için)
            baslik_selectors = [
                '[data-testid="conversation-info-header"]',
                'header [data-testid="conversation-info-header-chat-title"]',
                'header span[title]',
            ]
            baslik_el = None
            for sel in baslik_selectors:
                try:
                    baslik_el = WebDriverWait(self._driver, 4).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, sel)))
                    break
                except: continue

            # 2) Üç nokta menüsünü aç
            menu_selectors = [
                '[data-testid="menu"]',
                '[data-testid="conversation-menu"]',
                'div[title="Menü"]',
                'span[data-testid="menu"]',
            ]
            menu_el = None
            for sel in menu_selectors:
                try:
                    menu_el = WebDriverWait(self._driver, 4).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, sel)))
                    menu_el.click()
                    time.sleep(0.5)
                    break
                except: continue

            if not menu_el:
                self.on_log("⚠ Menü butonu bulunamadı — sohbet temizlenemedi")
                return

            # 3) 'Sohbeti Temizle' seçeneğini bul ve tıkla
            temizle_selectors = [
                '[aria-label="Sohbeti temizle"]',
                '[aria-label="Clear chat"]',
            ]
            temizle_el = None
            for sel in temizle_selectors:
                try:
                    temizle_el = WebDriverWait(self._driver, 4).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, sel)))
                    temizle_el.click()
                    time.sleep(0.5)
                    break
                except: continue

            if not temizle_el:
                # Metin ile ara
                try:
                    items = self._driver.find_elements(
                        By.CSS_SELECTOR, '[role="menuitem"], li[data-animate-dropdown-item]')
                    for item in items:
                        if "temizle" in (item.text or "").lower() or \
                           "clear" in (item.text or "").lower():
                            item.click()
                            time.sleep(0.5)
                            temizle_el = item
                            break
                except: pass

            if not temizle_el:
                self.on_log("⚠ 'Sohbeti Temizle' seçeneği bulunamadı")
                return

            # 4) Onay diyaloğunu onayla
            onay_selectors = [
                '[data-testid="confirm-btn"]',
                'div[role="button"][aria-label*="Temizle"]',
                'div[role="button"][aria-label*="Clear"]',
            ]
            for sel in onay_selectors:
                try:
                    onay = WebDriverWait(self._driver, 4).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, sel)))
                    onay.click()
                    time.sleep(1)
                    break
                except: continue

            # 5) "Mesajları sil" onayı (bazen iki adımlı)
            try:
                sil_selectors = [
                    'button[aria-label*="Sil"]',
                    'div[role="button"][aria-label*="Sil"]',
                ]
                for sel in sil_selectors:
                    try:
                        sil = WebDriverWait(self._driver, 3).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, sel)))
                        sil.click()
                        time.sleep(1)
                        break
                    except: continue
            except: pass

            self.on_log("🧹 Grup sohbeti temizlendi (sadece bu cihazda).")

        except Exception as e:
            self.on_log(f"⚠ Sohbet temizleme hatası: {str(e)[:80]}")

    def _saglik_kontrol(self):
        """Sayfa hâlâ yanıt veriyor mu ve sohbet listesi görünür mü kontrol et.
        Donmuşsa veya oturum düşmüşse sayfayı yenile, olmazsa botu yeniden başlat."""
        try:
            ok = self._driver.execute_script(
                "return document.querySelector('[data-testid=\"chat-list\"],"
                "[aria-label*=\"Sohbet\"], [aria-label*=\"Chat\"]') !== null"
            )
            if not ok:
                self.on_log("⚠ Sohbet listesi görünmüyor, sayfa yenileniyor…")
                self._driver.refresh()
                time.sleep(8)
                self._gruba_git_sessiz()
        except Exception:
            self.on_log("❌ Sağlık kontrolü başarısız — bot otomatik yeniden başlatılıyor…")
            self._otomatik_yeniden_baslat()

    def _gruba_git_sessiz(self):
        """Sayfa yenilendikten sonra gruba tekrar gir (log kirletmeden)."""
        try:
            cfg = load_config()
            grup = cfg.get("wa_group_name","").strip()
            if not grup: return
            for sel in ['[data-testid="chat-list-search"]',
                        '[aria-label="Sohbet veya kişi ara"]',
                        '[aria-label="Search or start new chat"]',
                        'div[contenteditable="true"]']:
                try:
                    sb = WebDriverWait(self._driver, 5).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, sel)))
                    sb.click(); time.sleep(0.5)
                    sb.send_keys(grup); time.sleep(2)
                    break
                except: continue
            for sel in [f'[title="{grup}"]',
                        '[data-testid="cell-frame-container"]',
                        'div[role="listitem"]']:
                try:
                    el = WebDriverWait(self._driver, 5).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, sel)))
                    el.click(); break
                except: continue
            time.sleep(1)
        except: pass

    def wa_yanit_gonder(self, metin):
        """WhatsApp Web'deki açık gruba otomatik yanıt yazar ve gönderir."""
        if not metin or not self._driver:
            return False
        try:
            from selenium.webdriver.common.keys import Keys
            import tempfile, subprocess, os as _os

            # ── 1) Panoyu ayarla ────────────────────────────────────────────
            # PowerShell temp-file yöntemi: özel karakter/Türkçe sorunsuz,
            # izin gerektirmez, her Windows'ta çalışır.
            _pano_ok = False
            try:
                with tempfile.NamedTemporaryFile(
                        mode='w', encoding='utf-8-sig',
                        suffix='.txt', delete=False) as _f:
                    _f.write(metin); _tmp = _f.name
                subprocess.run(
                    ['powershell', '-NoProfile', '-Command',
                     f'[System.IO.File]::ReadAllText("{_tmp}",'
                     f'[System.Text.Encoding]::UTF8) | Set-Clipboard'],
                    capture_output=True, timeout=5)
                _os.unlink(_tmp)
                _pano_ok = True
            except Exception:
                pass

            if not _pano_ok:
                # Fallback: win32clipboard
                try:
                    import win32clipboard as _wc
                    _wc.OpenClipboard(0)
                    _wc.EmptyClipboard()
                    _wc.SetClipboardText(metin, _wc.CF_UNICODETEXT)
                    _wc.CloseClipboard()
                    _pano_ok = True
                except Exception:
                    pass

            if not _pano_ok:
                self.on_log("⚠ Pano ayarlanamadı — WA yanıtı gönderilemedi")
                return False

            time.sleep(0.2)

            # ── 2) Compose kutusunu bul ──────────────────────────────────────
            # Daha spesifik selectors — footer içindeki compose kutusu,
            # arama kutusunu ASLA eşleştirmez
            selectors = [
                'footer [data-testid="conversation-compose-box-input"]',
                '[data-testid="conversation-compose-box-input"]',
                'footer div[contenteditable="true"]',
                'div[contenteditable="true"][data-tab="10"]',
            ]
            input_box = None
            for sel in selectors:
                try:
                    input_box = WebDriverWait(self._driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, sel)))
                    # Doğrulama: footer içinde mi?
                    in_footer = self._driver.execute_script(
                        "return arguments[0].closest('footer') !== null "
                        "|| arguments[0].getAttribute('data-tab') === '10'",
                        input_box)
                    if in_footer:
                        break
                    input_box = None
                except:
                    continue

            if not input_box:
                self.on_log("⚠ WA compose kutusu bulunamadı")
                return False

            # ── 3) Kutuyu temizle ve yapıştır ───────────────────────────────
            input_box.click()
            time.sleep(0.3)
            # Önce içeriği JS ile temizle, sonra Ctrl+V ile yapıştır
            self._driver.execute_script(
                "arguments[0].focus();"
                "document.execCommand('selectAll', false, null);"
                "document.execCommand('delete', false, null);",
                input_box)
            time.sleep(0.15)
            input_box.send_keys(Keys.CONTROL + 'v')
            time.sleep(0.5)

            # ── 4) Gönder butonuna tıkla ─────────────────────────────────────
            send_selectors = [
                '[data-testid="send"]',
                '[data-testid="compose-btn-send"]',
                'button[aria-label="Gönder"]',
                'button[aria-label="Send"]',
                'span[data-testid="send"]',
            ]
            sent = False
            for sel in send_selectors:
                try:
                    btn = WebDriverWait(self._driver, 3).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, sel)))
                    btn.click()
                    sent = True
                    break
                except:
                    continue

            if not sent:
                input_box.send_keys(Keys.ENTER)

            time.sleep(0.5)
            return True

        except Exception as e:
            self.on_log(f"⚠ WA yanıt gönderilemedi: {str(e)[:80]}")
            return False

    def _wa_erisim_uyarisi_gonder(self):
        """3 ard arda başarısız bağlantıdan sonra config'deki ilk kuralın
        mail listesine WhatsApp Web erişilemiyor uyarısı gönderir."""
        try:
            cfg = load_config()
            from_acc = cfg.get("outlook_account","") or None
            grup_adi = cfg.get("wa_group_name","")
            # Tüm kuralların mail listelerini birleştir, tekrarsız
            tum_mailler = []
            seen_ml = set()
            for k in cfg.get("kurallar",[]):
                for m in k.get("mail_list",[]):
                    if m not in seen_ml:
                        tum_mailler.append(m); seen_ml.add(m)
            if not tum_mailler:
                self.on_log("⚠ Uyarı maili gönderilemedi: mail listesi boş")
                return
            tarih = datetime.now().strftime("%d.%m.%Y %H:%M")
            sablon = (
                f"Sayın İlgili,\n\n"
                f"Pregate WA Mail Botu, WhatsApp Web'e {self._baglanti_hata_say} kez ard arda "
                f"bağlanamadı.\n\n"
                f"⚠ El ile müdahale gereklidir.\n\n"
                f"Olası nedenler:\n"
                f"  • WhatsApp Web oturumu düştü (QR yeniden okutulması gerekebilir)\n"
                f"  • İnternet bağlantısı kesildi\n"
                f"  • WhatsApp Web'de güncelleme / kesinti var\n\n"
                f"Lütfen botu çalıştıran bilgisayarda uygulamayı kontrol edin.\n\n"
                f"Tarih: {tarih}\n"
                f"Grup : {grup_adi}\n\n"
                f"Saygılarımızla,\nPregate Araç Kontrol Sistemi"
            )
            import pythoncom, win32com.client as _win32
            def _gonder():
                try:
                    pythoncom.CoInitialize()
                    ol   = _win32.Dispatch("Outlook.Application")
                    mail = ol.CreateItem(0)
                    mail.Subject = f"[UYARI] WhatsApp Web'e Erişilemiyor – {tarih}"
                    mail.Body    = sablon
                    mail.To      = "; ".join(tum_mailler)
                    if from_acc:
                        for acc in ol.GetNamespace("MAPI").Accounts:
                            if acc.SmtpAddress.lower() == from_acc.lower():
                                mail.SendUsingAccount = acc; break
                    mail.Send()
                    self.on_log(f"📧 WA erişim uyarısı maili gönderildi → {', '.join(tum_mailler[:2])}")
                except Exception as ex:
                    self.on_log(f"⚠ Uyarı maili gönderilemedi: {str(ex)[:80]}")
                finally:
                    pythoncom.CoUninitialize()
            import threading
            threading.Thread(target=_gonder, daemon=True).start()
        except Exception as ex:
            self.on_log(f"⚠ Uyarı maili hatası: {str(ex)[:80]}")

    def _otomatik_yeniden_baslat(self):
        """Chrome çöktüğünde veya bağlantı koptuğunda kullanıcı müdahalesi
        olmadan botu kendi kendine yeniden başlatır.
        Elle Durdur butonuna basıldıysa çalışmaz."""
        if self._manuel_durdurma:
            self.on_log("ℹ️ Manuel durdurma — otomatik restart yapılmıyor.")
            return
        if not self.running:
            return
        try:
            if self._driver: self._driver.quit()
        except: pass
        self._driver = None
        time.sleep(5)
        if not self.running or self._manuel_durdurma:
            return
        self.on_log("🔄 Bot otomatik olarak yeniden başlatılıyor…")
        self.start()

    def stop(self, manuel=True):
        """Botu durdur. manuel=True ise Durdur butonuna basıldı demek —
        otomatik restart bu durumda çalışmaz."""
        self._manuel_durdurma = manuel
        self.running = False
        if self._biriktiric:
            self._biriktiric.temizle()
        try:
            if self._driver:
                self._driver.quit()
        except: pass
        self._driver = None


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
        mesaj    = f"[{kw_str}] 34 ABC 123 plakalı araç kayıt yaptırmadı.\nSürücü bilgi vermeden ayrıldı."
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
        ctk.CTkLabel(info, text="WhatsApp Web Bağlantısı",
                     font=("Segoe UI",12,"bold"),
                     text_color=RENK_YAZI).pack(anchor="w", padx=14, pady=(12,6))
        for no, m in [
            ("1","Başlat'a basın — Chrome otomatik açılır."),
            ("2","İlk seferde QR taratın (WhatsApp → Bağlı Cihazlar → Cihaz Ekle)."),
            ("3","Sonraki açılışlarda otomatik giriş yapar, QR sormaz."),
            ("4","Chrome arka planda açık kalır — bot 10 sn'de bir mesaj okur."),
            ("5","Aynı kişinin mesajları biriktirilerek tek mail olarak gönderilir."),
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

        bf = ctk.CTkFrame(parent, fg_color=RENK_KART, corner_radius=10)
        bf.pack(fill="x", padx=10, pady=6)
        ctk.CTkLabel(bf, text="Mesaj Biriktirme Süresi",
                     font=("Segoe UI",12,"bold"),
                     text_color=RENK_YAZI).pack(anchor="w", padx=14, pady=(12,4))
        row = ctk.CTkFrame(bf, fg_color=RENK_KART)
        row.pack(fill="x", padx=14, pady=(0,12))
        ctk.CTkLabel(row,
                     text="Aynı kişinin mesajlarını kaç dakika biriktir:",
                     text_color=RENK_YAZI2, font=("Segoe UI",10)).pack(side="left")
        self.ent_sure = ctk.CTkEntry(row, width=60, fg_color=RENK_PANEL,
                                      border_color=RENK_VURGU,
                                      text_color=RENK_YAZI, font=("Segoe UI",11))
        self.ent_sure.insert(0, str(load_config().get("bekleme_dk",2)))
        self.ent_sure.pack(side="left", padx=10)
        ctk.CTkLabel(row, text="dakika", text_color=RENK_YAZI2,
                     font=("Segoe UI",10)).pack(side="left")

        # Uygunsuz İçerik Yanıtı alanı
        uf = ctk.CTkFrame(parent, fg_color=RENK_KART, corner_radius=10)
        uf.pack(fill="x", padx=10, pady=6)
        ctk.CTkLabel(uf, text="Uygunsuz İçerik WA Yanıtı",
                     font=("Segoe UI",12,"bold"),
                     text_color=RENK_YAZI).pack(anchor="w", padx=14, pady=(12,4))
        ctk.CTkLabel(uf,
                     text="Anahtar kelime eşleşmezse gruba otomatik bu mesaj yazılır. Boş bırakılırsa yanıt gönderilmez.",
                     text_color=RENK_YAZI2, font=("Segoe UI",9), wraplength=480
                     ).pack(anchor="w", padx=14, pady=(0,4))
        self.txt_uygunsuz = ctk.CTkTextbox(uf, fg_color=RENK_PANEL,
                                            border_color=RENK_SINIR,
                                            text_color=RENK_YAZI,
                                            font=("Segoe UI",10), height=60,
                                            border_width=2, corner_radius=6,
                                            wrap="word")
        self.txt_uygunsuz.insert("1.0",
            load_config().get("uygunsuz_cevap",
                "❌ Bu mesaj içeriği uygun formatta değil.\n"
                "Lütfen araç plakası ve yükleme bilgilerini içerecek şekilde düzenleyip tekrar gönderin."))
        self.txt_uygunsuz.pack(fill="x", padx=14, pady=(0,12))

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
            kart, ea, ek, ec = self._kural_kart(self.scroll, k, renk)
            kart.grid(row=i, column=0, sticky="ew", pady=(0,10))
            self._kart_list.append((kart, k, ea, ek, ec))

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
        # WA Otomatik Yanıt
        ctk.CTkLabel(body, text="WA Yanıt\n(opsiyonel):", text_color=RENK_YAZI2,
                     font=("Segoe UI",10)).grid(row=3,column=0,sticky="nw",pady=(8,0))
        ent_cevap = ctk.CTkTextbox(body, fg_color=RENK_PANEL, border_color=RENK_SINIR,
                                    text_color=RENK_YAZI, font=("Segoe UI",10),
                                    height=52, border_width=2, corner_radius=6,
                                    wrap="word")
        ent_cevap.insert("1.0", kural.get("wa_cevap",""))
        ent_cevap.grid(row=3, column=1, sticky="ew", padx=(8,0), pady=(8,0))
        ctk.CTkLabel(body, text="💡 Mail gönderilince gruba otomatik bu mesaj yazılıp gönderilir. Boş bırakılırsa yanıt gönderilmez.",
                     text_color=RENK_YAZI2, font=("Segoe UI",9), wraplength=340
                     ).grid(row=4, column=1, sticky="w", padx=(8,0), pady=(2,0))
        ml_hdr = ctk.CTkFrame(body, fg_color=RENK_KART)
        ml_hdr.grid(row=5,column=0,columnspan=2,sticky="ew",pady=(10,2))
        ctk.CTkLabel(ml_hdr, text="Mail Listesi:", text_color=RENK_YAZI2,
                     font=("Segoe UI",10)).pack(side="left")
        ml_frame = ctk.CTkFrame(body, fg_color=RENK_PANEL, corner_radius=6)
        ml_frame.grid(row=6, column=0, columnspan=2, sticky="ew")
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
                                    border_color=RENK_SINIR, text_color=RENK_YAZI,
                                    font=("Segoe UI",10),
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
        btn_row.grid(row=7, column=0, columnspan=2, sticky="e", pady=(8,0))
        ctk.CTkButton(btn_row, text="👁 Mail Önizle", width=110, height=26,
                      fg_color=RENK_VURGU2, hover_color=RENK_VURGU,
                      font=("Segoe UI",10),
                      command=lambda k=kural: MailOnizleme(
                          self,k.get("ad","?"),k.get("mail_list",[]),k.get("keywords",[]))
                      ).pack(side="left", padx=(0,6))
        ctk.CTkButton(btn_row, text="🗑 Sil", width=70, height=26,
                      fg_color=RENK_KIRMIZI, hover_color="#a93226",
                      font=("Segoe UI",10),
                      command=lambda k=kural,f=frame: self._kural_sil(k,f)
                      ).pack(side="left")
        return frame, ent_ad, ent_kw, ent_cevap

    def _kural_ekle(self):
        cfg = load_config()
        cfg.setdefault("kurallar",[]).append({
            "id":str(uuid.uuid4())[:8],"ad":"Yeni Kural",
            "keywords":[],"mail_list":[]})
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
            ctk.CTkLabel(kart,
                         text="⚠ Outlook açık değil veya hesap bulunamadı.",
                         text_color=RENK_SARI, font=("Segoe UI",11),
                         justify="left").pack(padx=14, pady=14, anchor="w")
            self.cmb_outlook = None
        test_k = ctk.CTkFrame(parent, fg_color=RENK_KART, corner_radius=10)
        test_k.pack(fill="x", padx=10, pady=(12,4))
        ctk.CTkLabel(test_k, text="Bağlantı Testi",
                     font=("Segoe UI",11,"bold"),
                     text_color=RENK_YAZI).pack(anchor="w", padx=14, pady=(10,4))
        self.lbl_test = ctk.CTkLabel(test_k, text="",
                                      font=("Segoe UI",10), text_color=RENK_YESIL)
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
        try: cfg["bekleme_dk"] = max(1,int(self.ent_sure.get().strip()))
        except: cfg["bekleme_dk"] = 2
        cfg["uygunsuz_cevap"] = self.txt_uygunsuz.get("1.0","end").strip()
        if hasattr(self,"cmb_outlook") and self.cmb_outlook:
            cfg["outlook_account"] = self.cmb_outlook.get().strip()
        kurallar=[]
        for item in self._kart_list:
            frame,kural,ea,ek,ec = item[0],item[1],item[2],item[3],item[4]
            if frame.winfo_exists():
                kural["ad"] = ea.get().strip()
                kural["keywords"] = [
                    x.strip().lower() for x in ek.get().split(",") if x.strip()]
                kural["wa_cevap"] = ec.get("1.0","end").strip()
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
        t=len(rows)
        self._krt("Toplam Red",str(t),RENK_KIRMIZI)
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
        self.state("normal")
        self.deiconify()
        self.wa_bot=None; self.running=False; self.mail_say=0
        db_init(); self._build_ui()
        # Siyah pencere sorununu önlemek için render zorla
        self.after(100, self._render_zorla)

    def _render_zorla(self):
        """CustomTkinter bazen ilk açılışta render etmez — zorla yenile."""
        try:
            self.update_idletasks()
            self.update()
            self.lift()
            self.focus_force()
        except: pass

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
        self.lbl_sure=ctk.CTkLabel(sol,
            text=f"Biriktirme: {cfg.get('bekleme_dk',2)} dk",
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
        self.wa_bot=WABot(on_log=self._log,
                          on_status=self._set_durum_p,
                          on_message=self._on_mesaj_bitti)
        threading.Thread(target=self.wa_bot.start,daemon=True).start()

    def _durdur(self):
        self.running=False
        if self.wa_bot: self.wa_bot.stop(manuel=True); self.wa_bot=None
        self.btn_baslat.configure(state="normal")
        self.btn_durdur.configure(state="disabled")
        self._set_durum("● Durduruldu",RENK_KIRMIZI)
        self._log("■ Bot durduruldu.")

    def _ayarlar(self):
        AyarlarPencere(self)
        self.after(500,self._guncelle_labels)

    def _guncelle_labels(self):
        cfg=load_config()
        self.lbl_grup.configure(text=f"Grup: {cfg.get('wa_group_name','—')}")
        self.lbl_sure.configure(text=f"Biriktirme: {cfg.get('bekleme_dk',2)} dk")

    def _rapor(self): RaporPencere(self)

    def _wa_yanit_gonder(self, metin):
        """WA grubuna yanıt gönder — thread'de çalışır."""
        if not metin or not self.wa_bot:
            return
        def _g(m=metin):
            ok = self.wa_bot.wa_yanit_gonder(m)
            if ok:
                self._log(f"💬 WA yanıtı gönderildi: {m[:60]}")
            else:
                self._log(f"⚠ WA yanıtı gönderilemedi")
        threading.Thread(target=_g, daemon=True).start()

    def _on_mesaj_bitti(self, gonderen, birlesik_metin, resimler=None):
        cfg  = load_config()
        norm = turkce_norm

        # Göndereni mailde gösterilecek forma getir (numara→****XXXX, isim→isim)
        gonderen_mail = gonderen_goster(gonderen)

        # ── Gün bazlı duplicate kontrolü ────────────────────────────────────
        # Bugün aynı gönderen + aynı içerik için zaten mail atıldıysa atla.
        # Program kapansa bile DB'den kontrol edilir (restart güvenli).
        if db_bugun_gonderildi_mi(gonderen, birlesik_metin):
            self._log(f"⏭ Bugün zaten gönderildi, atlanıyor: {gonderen_mail}")
            self._wa_yanit_gonder(
                "ℹ️ Bu mesaj bugün daha önce işlendi, tekrar mail gönderilmedi.")
            return

        # ── Kural eşleştirme ────────────────────────────────────────────────
        metin_lower = norm(birlesik_metin)
        eslesen = []
        for kural in cfg.get("kurallar", []):
            for kw in kural.get("keywords", []):
                if norm(kw) in metin_lower:
                    eslesen.append(kural); break

        if not eslesen:
            # Ayarlarda özel mesaj varsa onu kullan, yoksa otomatik oluştur
            uygunsuz = cfg.get("uygunsuz_cevap", "").strip()
            if not uygunsuz:
                # Config'deki kurallardan komut listesini dinamik oluştur
                kural_listesi = []
                for k in cfg.get("kurallar", []):
                    kw_str = ", ".join(k.get("keywords", []))
                    kural_listesi.append(f"• {k.get('ad','?')} → {kw_str}")
                komutlar_str = "\n".join(kural_listesi) if kural_listesi else "• (Kural tanımlanmamış)"
                uygunsuz = (
                    "❌ Uygunsuz mesaj — mail atılamaz.\n\n"
                    "Mail atılabilmesi için mesajınızda\n"
                    "aşağıdaki departman komutlarından\n"
                    "birinin bulunması gerekiyor:\n\n"
                    f"{komutlar_str}"
                )
            self._log(f"💬 [{gonderen_mail}]: eşleşme yok — komut hatası WA yanıtı gönderildi")
            self._wa_yanit_gonder(uygunsuz)
            return

        # ── Mail gönderimi ───────────────────────────────────────────────────
        from_acc          = cfg.get("outlook_account", "") or None
        grup_adi          = cfg.get("wa_group_name", "")
        basarili_kurallar = []   # (kural_ad, wa_cevap_ozel)
        basarisiz_kurallar= []

        for kural in eslesen:
            ml = kural.get("mail_list", [])
            if not ml:
                self._log(f"⚠ '{kural.get('ad')}' mail listesi boş!")
                basarisiz_kurallar.append(kural.get("ad","?"))
                continue

            ok, info = send_mail(
                ml, kural["ad"], gonderen_mail, birlesik_metin,
                from_acc, grup_adi, resimler or [])

            if ok:
                self.mail_say += 1
                self.after(0, lambda: self.lbl_sayac.configure(
                    text=str(self.mail_say)))
                db_ekle(gonderen, kural["ad"], birlesik_metin, bool(resimler))
                self._log(f"✉ [{kural['ad']}] {gonderen_mail} → "
                          f"{', '.join(ml[:2])}{'…' if len(ml)>2 else ''}")
                wa_cevap_ozel = kural.get("wa_cevap", "").strip()
                basarili_kurallar.append((kural["ad"], wa_cevap_ozel))
            else:
                self._log(f"❌ [{kural['ad']}] Mail hatası: {info}")
                basarisiz_kurallar.append(kural.get("ad","?"))

        # ── Tüm kurallar bittikten sonra TEK WA yanıtı gönder ───────────────
        tarih_str = datetime.now().strftime("%d.%m.%Y %H:%M")
        parcalar  = []

        if basarili_kurallar:
            for (kural_ad, wa_ozel) in basarili_kurallar:
                if wa_ozel:
                    parcalar.append(wa_ozel)
                else:
                    parcalar.append(
                        f"✅ Komutlar alındı.\n"
                        f"{kural_ad} grubuna mail gönderildi.\n"
                        f"Tarih: {tarih_str}"
                    )

        if basarisiz_kurallar:
            gruplar_hata = ", ".join(basarisiz_kurallar)
            parcalar.append(
                f"❌ Mail gönderilemedi: {gruplar_hata}\n"
                f"Sistem yöneticisini bilgilendirin."
            )

        if parcalar:
            self._wa_yanit_gonder("\n\n".join(parcalar))

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
        if self.wa_bot: self.wa_bot.stop(manuel=True)
        self.destroy()


if __name__=="__main__":
    # App doğrudan başlar — tek mainloop, siyah pencere sorunu yok.
    # Splash ayrı bir tk.Tk() kökü içinde açılmıyor; App'in kendi
    # Toplevel'ı olarak açılıp 2.5 saniye sonra kapanıyor.
    app = App()
    splash = SplashEkran(app)
    app.after(2500, splash.kapat)
    app.protocol("WM_DELETE_WINDOW", app.on_close)
    app.mainloop()
