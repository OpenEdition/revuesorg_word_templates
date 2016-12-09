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
UninstallDisplayIcon={app}\revuesorg.ico
OutputDir=build\win_setup
CloseApplications=yes

[Messages]
BeveledLabel= {#AppVersion}.{#SetupVersion}

[CustomMessages]
fr.AppName=Modèles pour Revues.org
fr.WordNotFound=Attention! La détection de Microsoft Word a échoué. Souhaitez-vous quand même poursuivre l'installation ?

en.AppName=Revues.org Templates
en.WordNotFound=Warning! Detection of Microsoft Word has failed. Do you want to force installation?

[Languages]
Name: "fr"; MessagesFile: "compiler:Languages\French.isl"
Name: "en"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "{#SrcModelesDir}\*.dot"; DestDir: "{userappdata}\Microsoft\Templates"; Flags: ignoreversion overwritereadonly
Source: "{#SrcMacrosDir}\*.dot"; DestDir: "{userappdata}\Microsoft\Templates"; Flags: ignoreversion overwritereadonly
Source: "{#SrcStartupDir}\*.dot"; DestDir: "{userappdata}\Microsoft\Word\STARTUP"; Flags: ignoreversion overwritereadonly
Source: "src\img\revuesorg.ico"; DestDir: "{app}"; Flags: ignoreversion overwritereadonly

[InstallDelete]
; Anciens fichiers à supprimer
Type: files; Name: "{userappdata}\Microsoft\Templates\revuesorg_complet.dot"

[Code]
var
  Msg: String;

function WordExists(): Boolean;
begin
  Result := DirExists(ExpandConstant('{userappdata}\Microsoft\Templates')) and DirExists(ExpandConstant('{userappdata}\Microsoft\Word\STARTUP'));
end;

function InitializeSetup(): Boolean;
begin
  Log('InitializeSetup called');
  Result := True;
  if WordExists() = False then
  begin
    Result := False;
    Msg := CustomMessage('WordNotFound');
    if MsgBox(Msg, mbConfirmation, MB_YESNO or MB_DEFBUTTON2) = IDYES then
    begin
      Result := True;
    end;
  end;
end;
