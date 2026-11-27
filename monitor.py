import base64
import os
import re
import sqlite3
import tempfile
import threading
import time
import unicodedata
from contextlib import closing
from datetime import datetime
from pathlib import Path

from selenium import webdriver
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    ElementNotInteractableException,
    StaleElementReferenceException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from command_queue import CommandQueue


READ = """
const main = document.querySelector('#main');
const messages = new Map();
const candidates = main?.querySelectorAll(
    '[data-testid="msg-container"], .message-in, .message-out, [data-pre-plain-text], [data-id]'
) || [];
for (const node of candidates) {
    const row = node.closest('.message-in, .message-out') ||
                node.closest('[data-id]') || node.closest('[data-testid="msg-container"]');
    if (!row || row.closest('[data-testid="quoted-message"]')) continue;
    const id = (row.closest('[data-id]') || row.querySelector('[data-id]'))?.getAttribute('data-id');
    const metadata = (row.matches('[data-pre-plain-text]') ? row : row.querySelector('[data-pre-plain-text]'))?.getAttribute('data-pre-plain-text') || '';
    const body = row.querySelector('[data-testid="msg-text"]') ||
                 row.querySelector('span.selectable-text.copyable-text, span.selectable-text, [data-lexical-text="true"]');
    const text = body?.innerText || '';
    const urls = [...new Set([...row.querySelectorAll('img[src^="blob:"]')].map(image => image.src))];
    const key = id || (metadata && text ? 'meta:' + metadata + text : '');
    if (key && (text || urls.length)) {
        messages.set(key, [key, text, urls, row.matches('.message-out'), metadata]);
    }
}
return [...messages.values()];
"""


ACTIVE_CHAT_LABELS = """
const main = document.querySelector('#main');
const header = main?.querySelector('header, [data-testid="conversation-header"], [role="banner"]');
if (!header) return [];
const titles = header.querySelectorAll(
  '[data-testid="conversation-info-header-chat-title"], [dir="auto"], [title]'
);
return [...titles].flatMap(el => [el.getAttribute('title'), el.innerText]).filter(Boolean);
"""


COMMON = {
    "?": "Komutları göster", "durum": "Gönderim durumunu göster",
    "start": "Gönderimi başlat (başla)", "stop": "Gönderimi durdur (dur)",
    "deneme": "Ayarlardaki test adresine gönder",
}
STAFF = {"kontrol": "Bekleyen kaydı göster", "gönder": "Bekleyen kaydı şimdi gönder", "iptal": "Bekleyen kaydı iptal et", "son": "Son işlem sonucunu göster", "gruplar": "Departmanları listele"}
ALIASES = {"başla": "start", "basla": "start", "dur": "stop", "durdur": "stop"}


def turkish_phone(value):
    cleaned = re.sub(r"[\s()+-]", "", value)
    if cleaned.startswith("0090"):
        cleaned = cleaned[2:]
    elif re.fullmatch(r"0?5[0-9]{9}", cleaned):
        cleaned = "90" + cleaned.lstrip("0")
    return cleaned if re.fullmatch(r"905[0-9]{9}", cleaned) else ""


def command_list(cfg, private):
    commands = dict(COMMON, **STAFF)
    if not private:
        for rule in cfg.get("kurallar", []):
            for word in rule.get("keywords", []):
                name = word.strip().lstrip("!").casefold()
                if name and name not in ALIASES:
                    commands.setdefault(name, rule.get("ad", "") + " grubuna gönder")
    return commands


def message_timestamp(metadata):
    # WhatsApp Web's Turkish message metadata exposes time at minute resolution.
    match = re.match(
        r"^\[\s*(?:(ÖÖ|ÖS|AM|PM)\s*)?(\d{1,2}):(\d{2})"
        r"(?:\s*(ÖÖ|ÖS|AM|PM))?\s*,?\s+(\d{1,2})[./](\d{1,2})[./](\d{4}|\d{2})\s*\]",
        metadata, re.IGNORECASE,
    )
    if not match:
        return None
    before, hour, minute, after, day, month, year = match.groups()
    marker = (before or after or "").upper()
    hour = int(hour)
    if marker:
        if not 1 <= hour <= 12:
            return None
        hour = hour % 12 + (12 if marker in ("ÖS", "PM") else 0)
    try:
        year = int(year) + (2000 if len(year) == 2 else 0)
        return datetime(year, int(month), int(day), hour, int(minute)).timestamp()
    except ValueError:
        return None


class Monitor:
    def __init__(self, config, db, profile, deliver, log, status):
        self.config, self.profile, self.deliver = config, profile, deliver
        self.log, self.status = log, status
        self.queue = CommandQueue(db)
        self.db = db
        with closing(sqlite3.connect(db)) as conn:
            self.active = bool(conn.execute("SELECT active FROM mail_switch").fetchone()[0])
        self.pending, self.photos, self.seen = {}, {}, set()
        self.last = {}
        self.targets = {}
        self.last_chat_warning = None
        self.initialized = set()
        self.lock = threading.RLock()
        self.closed = threading.Event()
        self.since = self.session_since = time.time()

    def switch(self, active):
        with self.lock:
            if self.active != active:
                self.since = time.time()
                self.queue.switch(active, self.since)
                self.active = active
                self.pending.clear()
                self.photos.clear()
            self.status(active)

    def stop(self, manuel=True):
        self.closed.set()

    def run(self):
        while not self.closed.is_set():
            driver = None
            try:
                options = webdriver.ChromeOptions()
                options.add_argument("--user-data-dir=" + self.profile)
                options.add_argument("--lang=tr")
                driver = webdriver.Chrome(options=options)
                driver.get("https://web.whatsapp.com/")
                self.initialized.clear()
                self.log("WhatsApp Web açıldı. Gerekirse QR kodunu okutun.")
                while not self.closed.wait(1):
                    cfg = self.config()
                    admins = list(dict.fromkeys(filter(None, map(turkish_phone, cfg.get("admins", [])))))
                    for target in [cfg.get("wa_group_name", "")] + admins:
                        if target and self.open_chat(driver, target):
                            self.scan(driver, target, target in admins, cfg)
                    self.flush(driver)
            except Exception as error:
                self.log("WhatsApp bağlantısı yeniden kurulacak: " + type(error).__name__)
            finally:
                if driver:
                    try:
                        driver.quit()
                    except Exception:
                        pass
            self.closed.wait(10)

    def open_chat(self, driver, title):
        try:
            return self._open_chat(driver, title)
        except (ElementClickInterceptedException, ElementNotInteractableException):
            now = time.monotonic()
            if self.last_chat_warning is None or now - self.last_chat_warning >= 60:
                self.log("WhatsApp sohbeti açılamadı. Açık bir pencere varsa kapatın. Chrome açık kalacak.")
                self.last_chat_warning = now
        except StaleElementReferenceException:
            # WhatsApp refreshed the list. The next scan will locate the new row.
            pass
        return False

    def _open_chat(self, driver, title):
        def normalize(value):
            clean = unicodedata.normalize("NFC", value).translate(
                dict.fromkeys(map(ord, "\u200b\u200e\u200f\ufeff"))
            )
            return " ".join(clean.split()).casefold()

        expected = normalize(title)
        phone = turkish_phone(title)

        def same_title(value):
            return normalize(value) == expected or bool(phone and turkish_phone(value) == phone)

        def matches():
            return any(same_title(value) for value in driver.execute_script(ACTIVE_CHAT_LABELS))

        if matches():
            return True
        for element in driver.find_elements(By.CSS_SELECTOR, "#pane-side [title]"):
            if not element.is_displayed() or not same_title(element.get_attribute("title") or ""):
                continue
            row = driver.execute_script(
                "return arguments[0].closest('[role=row], [role=listitem], [data-testid=cell-frame-container]') || arguments[0]",
                element,
            )
            driver.execute_script("arguments[0].scrollIntoView({block: 'nearest'})", row)
            row.click()
            for _ in range(20):
                if matches():
                    return True
                time.sleep(.1)
            break
        return False

    def scan(self, driver, target, private, cfg):
        rows = driver.execute_script(READ)
        if not rows:
            return
        with self.lock:
            baseline = target not in self.initialized
            self.initialized.add(target)
            if baseline and not private:
                self.log(f"WhatsApp grubu doğrulandı. Yeni mesajlar izleniyor ({len(rows)} mevcut mesaj atlandı). Başlatma dakikasından sonraki mesajlar alınır.")
            for mid, body, urls, outgoing, metadata in rows:
                if mid in self.seen:
                    continue
                self.seen.add(mid)
                if baseline:
                    continue
                timestamp = message_timestamp(metadata)
                now = time.time()
                if timestamp is None or not self.session_since <= timestamp <= now:
                    continue
                address = re.match(r"^(true|false)_([^_]+)_[^_]+(?:_(.+))?$", mid)
                if address:
                    direction, chat_id, participant = address.groups()
                    outgoing = outgoing or direction == "true"
                else:
                    chat_id, participant = "", None
                if private:
                    phone = chat_id.partition("@")[0]
                    if outgoing or not chat_id.endswith("@c.us") or phone != turkish_phone(target):
                        continue
                    chat, sender = chat_id, phone
                else:
                    if "@" in chat_id and not chat_id.endswith("@g.us"):
                        continue
                    # open_chat already verified the configured group by its title.
                    chat = "group:" + target
                    author = metadata.partition("] ")[2].removesuffix(": ")
                    sender = "self" if outgoing else participant or author or mid
                self.targets[chat] = target
                prefix = cfg.get("command_prefix") or "!"
                key = chat, sender
                def reply(message):
                    self.reply(driver, message)
                token, _, description = " ".join(body.split()).partition(" ")
                command = token.startswith(prefix)
                token = token[len(prefix):].casefold() if command else ""
                token = ALIASES.get(token, token)
                allowed = command_list(cfg, private)
                if command:
                    self.log("WhatsApp komutu algılandı: " + (token if token in allowed else "tanınmayan komut"))
                if command and token not in allowed:
                    reply("Komut tanınmadı veya yetkiniz yok. Liste için " + prefix + "? yazın.")
                    continue
                if command:
                    reply("Komut alındı.")
                if command and token == "?":
                    reply("Komutlar • " + " | ".join(prefix + k + " — " + v for k, v in allowed.items()))
                    continue
                if command and token in ("start", "stop", "durum"):
                    if token != "durum":
                        self.switch(token == "start")
                    if token == "start":
                        reply("Çalışmaya başladım. Yeni red kayıtlarını alıyorum.")
                    elif token == "stop":
                        reply("Durdum. Mail gönderimi durduruldu.")
                    else:
                        reply("Mail gönderimi " + ("aktif." if self.active else "durduruldu."))
                    continue
                if command and token == "gruplar":
                    reply("Gruplar: " + (", ".join(r["ad"] for r in cfg.get("kurallar", [])) or "Tanımlı grup yok."))
                    continue
                if command and token == "son":
                    reply(self.last.get(key, "Bu oturumda tamamlanan işleminiz yok."))
                    continue
                if command and token in ("kontrol", "gönder", "iptal"):
                    ident = next((i for i in reversed(self.pending) if self.pending[i][0] == key), None)
                    if not ident:
                        reply("Bekleyen kaydınız yok.")
                    else:
                        row = self.pending[ident]
                        if token == "kontrol":
                            reply(f"{row[2][0]} • {len(row[4])} fotoğraf • {row[3]}")
                        elif token == "gönder":
                            row[1] = 0
                        else:
                            del self.pending[ident]
                            self.queue.cancel(*ident)
                            self.last[key] = "Kayıt iptal edildi."
                            reply(self.last[key])
                    continue
                if not self.active:
                    if command:
                        reply("Gönderim durduruldu.")
                    continue
                if timestamp < self.since:
                    if command:
                        reply("Kayıt alınmadı. Başlatma dakikası geçtikten sonra komutu yeniden yazın.")
                    continue
                window = max(1, int(cfg.get("bekleme_dk", 2))) * 60
                key = chat, sender
                groups = [k for k in cfg.get("kurallar", []) if token in
                          {word.strip().lstrip("!").casefold() for word in k.get("keywords", [])}]
                group = None
                if token == "deneme":
                    group = ("DENEME", [cfg.get("test_email", "sseymen@poliport.com")])
                    description = description or "Pregate Mail Botu deneme gönderimi."
                elif not private and len(groups) == 1:
                    group = (groups[0]["ad"], groups[0]["mail_list"])
                if group and description.strip() and all(group[1]) and group[1]:
                    if self.queue.admit(chat, mid, timestamp):
                        self.pending[(chat, mid)] = [key, now + window, group, description.strip(),
                            [p for when, p in self.photos.pop(key, []) if now - when <= window]]
                        self.log(f"{group[0]} kayda alındı. Gönderim {window // 60} dk içinde planlandı.")
                        reply(f"{group[0]} kaydı alındı. Fotoğraflar için {window // 60} dk bekleniyor.")
                elif command:
                    reply("Kayıt açılamadı. Açıklamayı ve departman alıcılarını kontrol edin.")
                if urls:
                    latest = next((v for v in reversed(list(self.pending.values()))
                                   if v[0] == key and v[1] >= now), None)
                    for url in urls:
                        encoded = driver.execute_async_script("""
                            const done=arguments[arguments.length-1];
                            fetch(arguments[0]).then(r=>r.blob()).then(b=>{
                            const reader=new FileReader();reader.onload=()=>done(reader.result);
                            reader.onerror=()=>done(null);reader.readAsDataURL(b);}).catch(()=>done(null));
                        """, url)
                        if encoded:
                            if latest:
                                latest[4].append(encoded)
                            else:
                                self.photos.setdefault(key, []).append((now, encoded))
                self.photos[key] = [(t,p) for t,p in self.photos.get(key,[]) if now-t<=window]

    def reply(self, driver, message):
        editors = driver.find_elements(By.CSS_SELECTOR, "#main footer [contenteditable=true], #main [contenteditable=true][role=textbox]")
        for editor in reversed(editors):
            try:
                if editor.is_displayed():
                    editor.send_keys(message, Keys.ENTER)
                    return
            except (ElementNotInteractableException, StaleElementReferenceException):
                continue
        self.log("WhatsApp yanıtı yazılamadı. Mesaj kutusu bulunamadı veya kullanılamıyor.")

    def flush(self, driver):
        with self.lock:
            for identity, row in list(self.pending.items()):
                _, deadline, (group, recipients), body, photos = row
                if time.time() < deadline:
                    continue
                del self.pending[identity]
                try:
                    with tempfile.TemporaryDirectory() as folder:
                        paths = []
                        for n, photo in enumerate(photos, 1):
                            header, data = photo.split(",", 1)
                            ext = {"image/png":"png", "image/webp":"webp"}.get(header[5:].split(";")[0], "jpg")
                            path = Path(folder) / f"Kontrol_Fotografi_{n:02}.{ext}"
                            path.write_bytes(base64.b64decode(data))
                            paths.append(str(path))
                        ok = self.queue.submit(*identity, lambda: self.deliver(recipients, group, body, paths))
                except Exception as error:
                    ok = False
                    self.log("Gönderim kontrol edilmeli: " + type(error).__name__)
                result = (f"Mail atıldı. {group} • {len(photos)} fotoğraf." if ok
                          else f"{group} • Gönderim doğrulanamadı. Outlook'u kontrol edin.")
                self.last[row[0]] = result
                self.log(result)
                if self.open_chat(driver, self.targets[identity[0]]):
                    self.reply(driver, result)
