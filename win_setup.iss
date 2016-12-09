; INSTALLATEUR DES MODELES ET MACROS WORD POUR REVUES.ORG

#define AppVersion ReadIni(AddBackslash(SourcePath) + "src\translations.ini", "_configuration", "version", '0')
#define SetupVersion "3"
#define AppPublisher "OpenEdition"
#define AppURL "http://www.openedition.org"
#define SrcStartupDir "src\startup"
#define SrcMacrosDir "src\macros"
#define SrcModelesDir "build\templates"


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
OutputBaseFilename=setup_modeles_revuesorg_{#AppVersion}.{#SetupVersion}
VersionInfoProductVersion={#AppVersion}
VersionInfoVersion={#AppVersion}.{#SetupVersion}
DefaultDirName="{userpf}\RevuesOrgForWord"
DisableDirPage=yes
PrivilegesRequired=lowest
Compression=lzma
SolidCompression=yes
WizardImageFile=src\img\logo.bmp
WizardSmallImageFile=src\img\icon.bmp
SetupIconFile=src\img\revuesorg.ico
OutputDir=build\win_setup

[Messages]
BeveledLabel= {#AppVersion}.{#SetupVersion}

[CustomMessages]
fr.AppName=Modèles pour Revues.org
fr.TasksWord2007=Installer les modèles et macros (Word 2007 et ultérieurs)
fr.TasksVersion=Version de Word :
fr.WordNotFound=Désolé, la détection de Microsoft Word a échoué. L'installation va être abandonnée.

en.AppName=Revues.org Templates
en.TasksWord2007=Install templates and macros (Word 2007 and later)
en.TasksVersion=Word version:
en.WordNotFound=Sorry, detection of Microsoft Word has failed. The setup will be aborted.

[Tasks]
Name: "word2007"; Description: "{cm:TasksWord2007}"; GroupDescription: "{cm:TasksVersion}"; Check: Word2007Exists()

[Languages]
Name: "fr"; MessagesFile: "compiler:Languages\French.isl"
Name: "en"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "{#SrcModelesDir}\*.dot"; DestDir: "{userappdata}\Microsoft\Templates"; Flags: ignoreversion overwritereadonly; Tasks: word2007
Source: "{#SrcMacrosDir}\*.dot"; DestDir: "{userappdata}\Microsoft\Templates"; Flags: ignoreversion overwritereadonly; Tasks: word2007
Source: "{#SrcStartupDir}\*.dot"; DestDir: "{userappdata}\Microsoft\Word\STARTUP"; Flags: ignoreversion overwritereadonly; Tasks: word2007

[InstallDelete]
; Anciens fichiers à supprimer
Type: files; Name: "{userappdata}\Microsoft\Templates\revuesorg_complet.dot"; Tasks: word2007

[Code]
var
  WordExists: Boolean;
  Msg: String;

function Word2007Exists(): Boolean;
begin
  Result := DirExists(ExpandConstant('{userappdata}\Microsoft\Templates')) and DirExists(ExpandConstant('{userappdata}\Microsoft\Word\STARTUP'));
end;

function InitializeSetup(): Boolean;
begin
  Log('InitializeSetup called');
  Result := True;
  WordExists := Word2007Exists();
  if WordExists = False then
  begin
    Msg := CustomMessage('WordNotFound');
    MsgBox(Msg, mbError, MB_OK);
    Result := False;
  end;
end;
