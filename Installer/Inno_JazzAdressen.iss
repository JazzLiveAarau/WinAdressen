; Installation of JAZZ live AARAU Adressen

[Setup]
AppPublisher=JAZZ live AARAU
AppPublisherURL=https://jazzliveaarau.ch/
AppName=JAZZ live AARAU Adressen
AppVerName=JAZZ live AARAU Adressen version 2.7
DefaultDirName={sd}\Apps\JazzLiveAarau\Adressen
DefaultGroupName=JAZZ live AARAU Adressen
UninstallDisplayIcon={app}\AddressesJazz.exe
Compression=lzma
SolidCompression=yes
OutputDir= NeueVersion
OutputBaseFilename= SetupJazzLiveAarauAdressen-version-2-7

[Dirs]
Name: "{app}\Output"; Permissions: users-modify
Name: "{app}\Excel"; Permissions: users-modify
Name: "{app}\Help"; Permissions: users-modify
Name: "{app}\Backups"; Permissions: users-modify
Name: "{app}\NeueVersion"; Permissions: users-modify

[Files]
Source: "AddressesJazz.exe"; DestDir: "{app}"
Source: "ExcelUtil.dll"; DestDir: "{app}"
Source: "Ftp.dll"; DestDir: "{app}"
Source: "JazzFtp.dll"; DestDir: "{app}"
Source: "AdressesUtility.dll"; DestDir: "{app}"
Source: "JazzVersion.dll"; DestDir: "{app}"
Source: "EncodingTools.dll"; DestDir: "{app}"
Source: "Help\JAZZ_live_AARAU_Adressen.rtf"; DestDir: "{app}\Help"; Flags: isreadme; Permissions: users-modify

[Icons]
Name: "{group}\JAZZ live AARAU Adressen"; Filename: "{app}\AddressesJazz.exe"

[InstallDelete]
Type: files; Name: "{app}\AddressesJazz.exe"
Type: files; Name: "{app}\Ftp.dll"
Type: files; Name: "{app}\JazzFtp.dll"
Type: files; Name: "{app}\AdressesUtility.dll"
Type: files; Name: "{app}\EncodingTools.dll"
Type: files; Name: "{app}\JazzVersion.dll"
Type: files; Name: "{app}\ExcelUtil.dll"

[UninstallDelete]
Type: files; Name: "{app}\AddressesJazzSettings.config"
