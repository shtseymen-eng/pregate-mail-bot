import base64
import os
import re
import sqlite3
import tempfile
import threading
import time
from contextlib import closing
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
return [...document.querySelectorAll('#main [data-id]')].map(row => {
 const id = row.getAttribute('data-id');
 const body = row.querySelector('span.selectable-text');
 const urls = [...row.querySelectorAll('img[src^="blob:"]')].map(i => i.src);
 return [id, body?.innerText || '', urls];
});
"""


COMMON = {"?": "Komutları göster", "durum": "Gönderim durumunu göster"}
STAFF = {"kontrol": "Bekleyen kaydı göster", "gönder": "Bekleyen kaydı şimdi gönder", "iptal": "Bekleyen kaydı iptal et", "son": "Son işlem sonucunu göster", "gruplar": "Departmanları listele"}
ADMIN = {"start": "Gönderimi başlat", "stop": "Gönderimi durdur", "deneme": "Test adresine gönder"}


def command_list(cfg, private):
    commands = dict(COMMON, **STAFF)
    if private:
        commands.update(ADMIN)
    else:
        for rule in cfg.get("kurallar", []):
            for word in rule.get("keywords", []):
                name = word.strip().lstrip("!").casefold()
                if name and name not in ADMIN:
                    commands.setdefault(name, rule.get("ad", "") + " grubuna gönder")
    return commands


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
        self.since = time.time()

    def switch(self, active):
        with self.lock:
            if self.active != active:
                self.since = time.time()
                self.queue.switch(active, self.since)
                self.active = active
                self.pending.clear()
                self.photos.clear()
                self.initialized.clear()
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
                    admins = [x.strip() for x in cfg.get("admins", []) if x.strip()]
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
        def matches():
            return any(el.get_attribute("title") == title for el in
                       driver.find_elements(By.CSS_SELECTOR, "#main header span[title]"))
        if matches():
            return True
        results = driver.find_elements(By.CSS_SELECTOR, "#pane-side span[title]")
        for el in results:
            if el.is_displayed() and el.get_attribute("title") == title:
                el.click()
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
            for mid, body, urls in rows:
                if mid in self.seen or not mid.startswith("false_"):
                    continue
                self.seen.add(mid)
                if baseline:
                    continue
                parts = mid.split("_")
                chat = parts[1]
                if private:
                    phone = chat.partition("@")[0]
                    if not chat.endswith("@c.us") or phone != re.sub(r"\D", "", target):
                        continue
                    sender = phone
                else:
                    if not chat.endswith("@g.us") or len(parts) < 4:
                        continue
                    sender = parts[-1]
                self.targets[chat] = target
                prefix = cfg.get("command_prefix") or "!"
                key = chat, sender
                def reply(message):
                    self.reply(driver, message)
                token, _, description = " ".join(body.split()).partition(" ")
                command = token.startswith(prefix)
                token = token[len(prefix):].casefold() if command else ""
                allowed = command_list(cfg, private)
                if command and token not in allowed:
                    reply("Komut tanınmadı veya yetkiniz yok. Liste için " + prefix + "? yazın.")
                    continue
                if command:
                    reply("Komut alındı.")
                if command and token == "?":
                    reply(" | ".join(prefix + k + " — " + v for k, v in allowed.items()))
                    continue
                if command and token in ("start", "stop", "durum"):
                    if token != "durum":
                        self.switch(token == "start")
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
                now = time.time()
                window = max(1, int(cfg.get("bekleme_dk", 2))) * 60
                key = chat, sender
                groups = [k for k in cfg.get("kurallar", []) if token in
                          {word.strip().lstrip("!").casefold() for word in k.get("keywords", [])}]
                group = None
                if private and token == "deneme":
                    group = ("DENEME", [cfg.get("test_email", "")])
                elif not private and len(groups) == 1:
                    group = (groups[0]["ad"], groups[0]["mail_list"])
                if group and description.strip() and all(group[1]) and group[1]:
                    if self.queue.admit(chat, mid, now):
                        self.pending[(chat, mid)] = [key, now + window, group, description.strip(),
                            [p for when, p in self.photos.pop(key, []) if now - when <= window]]
                        reply(group[0] + " kaydı alındı. Fotoğraflar bekleniyor.")
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
        editors = driver.find_elements(By.CSS_SELECTOR, "#main footer [contenteditable=true]")
        if editors:
            editors[-1].send_keys(message, Keys.ENTER)

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
