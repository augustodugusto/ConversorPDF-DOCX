[Setup]
AppName=Conversor PDF para DOCX
AppVersion=1.0
AppPublisher=Augusto Duran
DefaultDirName={autopf}\ConversorPDF
DefaultGroupName=Conversor PDF para DOCX
OutputBaseFilename=setup_conversor_pdf_v1.0
Compression=lzma2/ultra
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\ConversorPDF.exe

[Files]
; Este comando crucial copia TUDO da pasta que o PyInstaller criou.
Source: "dist\ConversorPDF\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Conversor PDF para DOCX"; Filename: "{app}\ConversorPDF.exe"; IconFilename: "{app}\assets\icon.ico"
Name: "{autodesktop}\Conversor PDF para DOCX"; Filename: "{app}\ConversorPDF.exe"; IconFilename: "{app}\assets\icon.ico"

[Run]
Filename: "{app}\ConversorPDF.exe"; Description: "Executar a aplicação"; Flags: nowait postinstall skipifsilent