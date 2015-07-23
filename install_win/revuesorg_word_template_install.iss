; INSTALLATEUR DES MODELES ET MACROS WORD POUR REVUES.ORG

#define AppVersion "4.0.2"
#define SetupVersion "4.0.2.1"
#define AppPublisher "OpenEdition"
#define AppURL "http://www.openedition.org"
#define AppSource "src"
#define ModelesDir "templates"
#define StartupDir "startup"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{35CB58FF-2EAF-428F-AAC7-D56AB61A7DC2}
AppName={cm:AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
OutputBaseFilename=setup_modeles_revuesorg_{#SetupVersion}
VersionInfoProductVersion={#AppVersion}
VersionInfoVersion={#SetupVersion}
DefaultDirName="{userpf}\RevuesOrgForWord"
DisableDirPage=yes
PrivilegesRequired=lowest
Compression=lzma
SolidCompression=yes
WizardImageFile=img\logo.bmp
WizardSmallImageFile=img\icon.bmp
SetupIconFile=img\revuesorg.ico

[Messages]
BeveledLabel= {#SetupVersion}

[CustomMessages]
fr.AppName=Modèles pour Revues.org
fr.TasksWord2007=Installer les modèles et macros pour Word 2007 et ultérieurs
fr.TasksWord2003=Installer les modèles et macros pour Word 2003
fr.TasksVersion=Version de Word :
fr.WordNotFound=Désolé, la détection de Microsoft Word a échoué. L’installation va être abandonnée.

en.AppName=Revues.org Templates
en.TasksWord2007=Install templates and macros for Word 2007 and later
en.TasksWord2003=Install templates and macros for Word 2003
en.TasksVersion=Word version:
en.WordNotFound=Sorry, detection of Microsoft Word has failed. The setup will be aborted.

[Tasks]
Name: "word2007"; Description: "{cm:TasksWord2007}"; GroupDescription: "{cm:TasksVersion}"; Check: Word2007Exists()
Name: "word2003"; Description: "{cm:TasksWord2003}"; GroupDescription: "{cm:TasksVersion}"; Check: Word2003Exists()

[Languages]
Name: "fr"; MessagesFile: "compiler:Languages\French.isl"
Name: "en"; MessagesFile: "compiler:Default.isl"

[Files]
; Word 2003
; FIXME : ne fonctionnera pas avec les versions non francophones de Word. Il faut récupérer le chemin dans le registre (ou abandonner le support de W2003)
Source: "{#AppSource}\{#ModelesDir}\*.dot"; DestDir: "{userappdata}\Microsoft\Modèles"; Flags: ignoreversion overwritereadonly; Tasks: word2003
Source: "{#AppSource}\{#StartupDir}\*.dot"; DestDir: "{userappdata}\Microsoft\Word\DÉMARRAGE"; Flags: ignoreversion overwritereadonly; Tasks: word2003

; Word 2007 et supérieur
Source: "{#AppSource}\{#ModelesDir}\*.dot"; DestDir: "{userappdata}\Microsoft\Templates"; Flags: ignoreversion overwritereadonly; Tasks: word2007
Source: "{#AppSource}\{#StartupDir}\*.dot"; DestDir: "{userappdata}\Microsoft\Word\STARTUP"; Flags: ignoreversion overwritereadonly; Tasks: word2007

[InstallDelete]
; Anciens fichiers à supprimer
Type: files; Name: "{userappdata}\Microsoft\Modèles\revuesorg_complet.dot"; Tasks: word2003
Type: files; Name: "{userappdata}\Microsoft\Templates\revuesorg_complet.dot"; Tasks: word2007

[Code]
var
  WordExists: Boolean;
  Msg: String;

function Word2007Exists(): Boolean;
begin
  Result := RegKeyExists(HKEY_CURRENT_USER, 'Software\Microsoft\Office') and DirExists(ExpandConstant('{userappdata}\Microsoft\Templates')) and DirExists(ExpandConstant('{userappdata}\Microsoft\Word\STARTUP'));
end;

function Word2003Exists(): Boolean;
begin
  Result := RegKeyExists(HKEY_CURRENT_USER, 'Software\Microsoft\Office\11.0\Word') and DirExists(ExpandConstant('{userappdata}\Microsoft\Modèles')) and DirExists(ExpandConstant('{userappdata}\Microsoft\Word\DÉMARRAGE'));
end;

function InitializeSetup(): Boolean;
begin
  Log('InitializeSetup called');
  Result := True;
  WordExists := Word2007Exists() or Word2003Exists();
  if WordExists = False then
  begin
    Msg := CustomMessage('WordNotFound');
    MsgBox(Msg, mbError, MB_OK);
    Result := False;
  end;
end;