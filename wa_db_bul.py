"""
Bu script PC'de WhatsApp Desktop DB dosyasını arar.
Sonucu ekrana yazar - o yolu programa gireceğiz.
"""
import os, glob

print("WhatsApp DB aranıyor...\n")

ARAMALAR = [
    r"%LOCALAPPDATA%\Packages\5319275A.WhatsApp_cv1g1gvanyjgm\LocalState",
    r"%LOCALAPPDATA%\Packages\5319275A.WhatsApp_cv1g1gvanyjgm\LocalCache",
    r"%LOCALAPPDATA%\WhatsApp",
    r"%APPDATA%\WhatsApp",
    r"%USERPROFILE%\AppData\Local\WhatsApp",
    r"%USERPROFILE%\AppData\Roaming\WhatsApp",
]

# 1. Bilinen dizinleri tara
for p in ARAMALAR:
    exp = os.path.expandvars(p)
    if os.path.exists(exp):
        print(f"[VAR] {exp}")
        for root, dirs, files in os.walk(exp):
            for f in files:
                tam = os.path.join(root, f)
                print(f"  {tam}")
    else:
        print(f"[YOK] {exp}")

# 2. Wildcard ile Packages altını tara
print("\n--- Packages wildcard ---")
appdata = os.environ.get("LOCALAPPDATA","")
for pattern in [
    os.path.join(appdata, "Packages", "5319275A*", "**", "*.db"),
    os.path.join(appdata, "Packages", "5319275A*", "**", "*.sqlite"),
    os.path.join(appdata, "Packages", "5319275A*", "**", "*.sqlite3"),
]:
    found = glob.glob(pattern, recursive=True)
    for f in found:
        print(f"  [DB] {f}")

# 3. Genel AppData tarama (*.db uzantılı whatsapp dosyaları)
print("\n--- AppData WA db dosyaları ---")
for base in [os.environ.get("LOCALAPPDATA",""),
             os.environ.get("APPDATA","")]:
    for root, dirs, files in os.walk(base):
        # WhatsApp klasörü varsa
        if "whatsapp" in root.lower():
            for f in files:
                if f.endswith((".db",".sqlite",".sqlite3")):
                    print(f"  {os.path.join(root,f)}")
        # Çok derin gitme
        dirs[:] = [d for d in dirs if "whatsapp" in d.lower() 
                   or d.startswith("5319275A")]

print("\nTAMAM")
