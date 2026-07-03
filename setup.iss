[Setup]
AppName=Pregate Kayıt Red Mail Botu
AppVersion=1.21
AppPublisher=Poliport
DefaultDirName={autopf}\PregateMail
DefaultGroupName=Pregate Mail Botu
OutputBaseFilename=PregateMail_Kurulum
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
OutputDir=installer_output

[Languages]
Name: "turkish"; MessagesFile: "compiler:Languages\Turkish.isl"

[Tasks]
Name: "desktopicon"; Description: "Masaüstüne kısayol oluştur"; GroupDescription: "Ek görevler:"
Name: "startupicon"; Description: "Windows başlangıcında otomatik çalıştır"; GroupDescription: "Ek görevler:"; Flags: unchecked

[Files]
Source: "dist\PregateMail\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Pregate Mail Botu"; Filename: "{app}\PregateMail.exe"
Name: "{group}\Kaldır"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Pregate Mail Botu"; Filename: "{app}\PregateMail.exe"; Tasks: desktopicon
; Registry yerine Başlangıç klasörüne kısayol — izin gerektirmez
Name: "{userstartup}\Pregate Mail Botu"; Filename: "{app}\PregateMail.exe"; Tasks: startupicon

[Run]
Filename: "{app}\PregateMail.exe"; Description: "Programı şimdi başlat"; Flags: nowait postinstall skipifsilent
