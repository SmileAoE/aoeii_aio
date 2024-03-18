#Requires AutoHotkey v2
#SingleInstance Force

If !A_IsAdmin {
    MsgBox("This application is not being ran as administrator`n"
         . "which can cause an unexpected behaviour on using any of it's options`n"
         . "`nYou must run it as administrator"
         , 'Exit'
         , 0x30)
    ExitApp
}

; Initialization
Server := 'https://raw.githubusercontent.com'
User := 'SmileAoE'
Repo := 'aoeii_aio'
Version := '1.9'
Layers := 'HKLM\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers'
Config := A_AppData '\aoeii_aio\config.ini'
MConfig := A_AppData '\aoeii_aio\mconfig.ini'
RConfig := A_AppData '\aoeii_aio\rconfig.ini'
Startup := A_AppData '\Microsoft\Windows\Start Menu\Programs\Startup\' StrReplace(A_ScriptName, 'ahk', 'lnk')
AppDir := ['DB', A_AppData '\aoeii_aio', A_AppData '\aoeii_aio\Hotkeys', A_AppData '\aoeii_aio\Records']
GRSetting := A_AppData '\GameRanger\GameRanger Prefs\Settings'
GRApp := A_AppData '\GameRanger\GameRanger\GameRanger.exe'
DrsTypes := Map('gra', 'graphics.drs', 'int', 'interfac.drs', 'ter', 'terrain.drs')
DrsRange := Map('gra', [2, 5312], 'int', [50100, 53211], 'ter', [15000, 15031])
IDL := 5
VCodedSlp := '3713EFBE'
NormalSlp := '322E304E'
General := Map()
General['AOK'] := Map()
General['AOK']['VersionsN'] := Map()
General['AOK']['Combine'] := Map('2.0b CD', ['2.0a No CD'])
General['AOC'] := Map()
General['AOC']['VersionsN'] := Map()
General['AOC']['Combine'] := Map('1.0e No CD', ['1.0c No CD'], '1.0e No CD', ['1.0c No CD'], '1.1  No CD', ['1.0c No CD'], '1.5  CD'   , ['1.0c No CD'])
General['FOE'] := Map()
General['FOE']['VersionsN'] := Map()
General['FOE']['Combine'] := Map()
General['LNG'] := Map()
Compatibilities := Map(1 , ["_____Not Set_____" , ""], 2, ["Windows 8", "WIN8RTM"], 3, ["Windows 7", "WIN7RTM"], 4, ["Windows Vista Sp2" , "VISTASP2"], 5, ["Windows Vista Sp1" , "VISTASP1"], 6, ["Windows Vista", "VISTARTM"], 7, ["Windows XP Sp3", "WINXPSP3"], 8, ["Windows XP Sp2", "WINXPSP2"], 9, ["Windows 98", "WIN98"], 10, ["Windows 95", "WIN95"])
BasePackages := ['DB/000.7z.001', 'DB/001.7z.001', 'DB/002.7z.001', 'DB/006.7z.001', 'DB/007.7z.001', 'DB/008.7z.001']
GamePackages := ['DB/003.7z.001', 'DB/003.7z.002', 'DB/003.7z.003', 'DB/003.7z.004', 'DB/004.7z.001', 'DB/004.7z.002', 'DB/004.7z.003', 'DB/005.7z.001']
Dots := 0
Task := 1
TaskNumber := BasePackages.Length
ProgramFiles86 := EnvGet(A_Is64bitOS ? "ProgramFiles(x86)" : "ProgramFiles")
VPNDir := ProgramFiles86 '\Hide ALL IP'
VPNExe := 'HideALLIP.exe'
VPNPath := VPNDir '\' VPNExe
Shortcut1 := '
(
;Fast One Unit Un-Select;
#Requires AutoHotkey v2
#SingleInstance Force
GroupAdd('AOKAOC', 'ahk_exe empires2.exe')
GroupAdd('AOKAOC', 'ahk_exe age2_x1.exe')
GroupAdd('AOKAOC', 'ahk_exe age2_x2.exe')
HotIfWinActive("ahk_group AOKAOC")
Hotkey('!RButton', Action)
Action(*) {
WinGetPos(,, &W, &H, 'ahk_group AOKAOC')
If W != A_ScreenWidth || H != A_ScreenHeight
Return
MouseClick('Right', , , , 0)
MouseGetPos(&X, &Y)
SendInput('{LCtrl Down}')
MouseClick('Left', 315, A_ScreenHeight - 130, , 0)
SendInput('{Ctrl Up}')
MouseMove(X, Y, 0)
}
ProcessWaitClose(A_Args[1])
ExitApp
)'
Shortcut2 := '
(
;Terminates The Game;
#Requires AutoHotkey v2
#SingleInstance Force
GroupAdd('AOKAOC', 'ahk_exe empires2.exe')
GroupAdd('AOKAOC', 'ahk_exe age2_x1.exe')
GroupAdd('AOKAOC', 'ahk_exe age2_x2.exe')
HotIfWinActive("ahk_group AOKAOC")
Hotkey('#q', Action)
Action(*) {
If GameIsRunning()
Msgbox('Game termination failure!', 'Game Terminate', 0x30)
}
GameIsRunning() {
Processes := ['empires2.exe', 'age2_x1.exe', 'age2_x2.exe']
For Each, Process in Processes {
If ProcessExist(Process) {
ProcessClose(Process)
}
ProcessWaitClose(Process, 5)
If ProcessExist(Process) {
Return True
}
}
Return False
}
ProcessWaitClose(A_Args[1])
ExitApp
)'

; Preparation

; Create app folders
For _, Item in AppDir {
    If !DirExist(Item) {
        DirCreate(Item)
    }
}

; Create Default Shortcuts 1
If !FileExist(AppDir[3] '\001.ahk') || FileRead(AppDir[3] '\001.ahk') != Shortcut1 {
    O := FileOpen(AppDir[3] '\001.ahk', 'w')
    O.Write(Shortcut1)
    O.Close()
}

; Create Default Shortcuts 2
If !FileExist(AppDir[3] '\002.ahk') || FileRead(AppDir[3] '\002.ahk') != Shortcut2 {
    O := FileOpen(AppDir[3] '\002.ahk', 'w')
    O.Write(Shortcut2)
    O.Close()
}

; Use Gdip
UseGDIP()

; Set Default Gui Color
CreateImageButton("SetDefGuiColor", 0xFFFFFF)

; Show loading GUI
Prepare := Gui(, 'Preparing...')
Prepare.OnEvent('Close', (*) => ExitApp())
HoldOn := Prepare.AddText('Center w400 h25', 'Please Wait...')
HoldOn.SetFont('s12 Bold')
DoneSteps := Prepare.AddProgress('Center w400 h20 -Smooth Range1-' BasePackages.Length + 1)
DoneStepsText := Prepare.AddText('Center wp cBlue')
Prepare.Show()

; Prepare the un-packer
PrepareTheUnpacker() {
    Try {
        If !FileExist('DB\7za.exe') || (HashFile('DB\7za.exe') != '80014d2b38a815f1a6ea220e679111c6') {
            Download(Server '/' User '/' Repo '/main/DB/7za.exe', 'DB\7za.exe')
        }
        If (HashFile('DB\7za.exe') != '80014d2b38a815f1a6ea220e679111c6') {
            MsgBox('Un-valid unpacker hashsum!', 'Oops!', '48')
            ; Run installation help page
            ExitApp
        }
        DoneSteps.Value += 1
        DoneStepsText.Text := 'DB/7za.exe'
    } Catch As Err {
        Choice := MsgBox('An error occured while preparing the unpacker!`nYou may like to open the help page?', 'Oops!', 0x30 + 0x04)
        If Choice = 'Yes' {
            ; Run installation help page
        }
        ExitApp
    }
}
PrepareTheUnpacker()

; Download the base files
PackageIsValid(Package) {
    If FileGetSize(PackagePath) <= 14 {
        MsgBox('An error occured while downloading "' Server '/' User '/' Repo '/main/' Package '" !', 'Oops!', '48')
        Return False
    }
    Return True
}
Try {
    For Each, Package in BasePackages {
        PackagePath := StrReplace(Package, '/', '\')
        If !FileExist(PackagePath) {
            Download(Server '/' User '/' Repo '/main/' Package, PackagePath)
            If !PackageIsValid(PackagePath)
                ExitApp()
        }
        PackageFolder := StrSplit(PackagePath, '.')[1]
        If !DirExist(PackageFolder) {
            RunWait('DB\7za.exe x ' PackagePath ' -o' PackageFolder ' -aoa', , 'Hide')
        }
        DoneSteps.Value += 1
        DoneStepsText.Text := Package
    }
} Catch As Err {
    Choice := MsgBox('An error occured while preparing the unpacker!`nYou may like to open the help page?', 'Oops!', 0x30 + 0x04)
    If Choice = 'Yes' {
        ; Run installation help page
    }
    ExitApp
}
Prepare.Hide()

; Main window
Box := '640x400'
Manager := Gui('-DPIScale Resize MinSize' Box ' MaxSize' Box, 'AGE OF EMPIRES II MANAGER ALL IN ONE v' Version)
Manager.SetFont('s10 Bold', 'Calibri')
Manager.BackColor := 'White'
Manager.OnEvent('Close', (*) => ExitApp())
Manager.OnEvent('Escape', (*) => ExitApp())
SB1 := ScrollBar(Manager, 600, 500)
#HotIf WinActive(Manager.Hwnd)
    WheelUp::
    WheelDown::
    +WheelUp::
    +WheelDown::
    Up::
    Down::
    +Up::
    +Down:: {
        SB1.ScrollMsg((InStr(A_ThisHotkey,"Down") || InStr(A_ThisHotkey,"Dn")) ? 1 : 0, 0, GetKeyState("Shift") ? 0x114 : 0x115, Manager.Hwnd)
    }
    PgUp::
    PgDn::
    +PgUp::
    +PgDn:: {
        SB1.ScrollMsg((InStr(A_ThisHotkey,"Down") || InStr(A_ThisHotkey,"Dn")) ? 3 : 2, 0, GetKeyState("Shift") ? 0x114 : 0x115, Manager.Hwnd)
    }
    Home::SB1.ScrollMsg(6, 0, GetKeyState("Shift") ? 0x114 : 0x115, Manager.Hwnd)
    End::SB1.ScrollMsg(7, 0, GetKeyState("Shift") ? 0x114 : 0x115, Manager.Hwnd)
#HotIf

; Features
Features := Map()

; Progress Bar
ProgressBar := Manager.AddProgress('-Smooth Hidden')

; About
Manager.AddPicture('xm+150', 'DB\000\game.png').Focus()
Manager.AddText('xm w580 h50 Center cGreen', 'AGE OF EMPIRES II MANAGER v' Version).SetFont('Bold s20')

; My Game
Features['The Game'] := []
_Game_ := Manager.AddText('xm yp+100 w600 h26 cBlue', 'GAME LOCATION:')
Features['The Game'].Push(_Game_)
_Game_.SetFont('Bold s14')

ChooseFolder := Manager.AddButton('xm+10 w100', 'Select')
GuiButtonIcon(ChooseFolder, 'DB\000\sfolder.png',, 'a1 r5')
Features['The Game'].Push(ChooseFolder)

ChooseFolder.OnEvent('Click', (*) => SelectTheGame())
SelectTheGame() {
    If SelectAFolder()
        ChargeSettings________(True)
}
SelectAFolder() {
    ChosenDir := FileSelect('D', 'C:\' (A_Is64bitOS ? 'Program Files (x86)' : 'Program Files') '\Microsoft Games')
    If !ChosenDir
        Return False
    IniWrite(ChosenDir, Config, 'Game', 'Path')
    ChosenFolder.Value := ChosenDir
    Return True
}
LoadGRFolder := Manager.AddButton('yp w200', 'Select from GameRanger')
GuiButtonIcon(LoadGRFolder, 'DB\000\gr.png',, 'a1 r5')
Features['The Game'].Push(LoadGRFolder)
LoadGRFolder.OnEvent('Click', (*) => SelectTheGameFromGR())
SelectTheGameFromGR() {
    TextFound := LoadGRSettingText()[1]
    FoundLocations := []
    AOKDir := ''
    AOCDir := ''
    FOEDir := ''
    ChosenDir := ''
    GRGamePath(TextFound, AppName) {
        P := InStr(TextFound, LFE := AppName, , -1)
        Loop {
            Char := SubStr(TextFound, P - (I := A_Index), 1)
            LFE := Char LFE
        } Until (Char = ':' || Ord(Char) = 10 || Ord(Char) = 13)
        Result := SubStr(TextFound, P - (I + 1), 1) LFE
        Return (FileExist(Result) ? Result : '')
    }
    If AOKFile := GRGamePath(TextFound, 'empires2.exe') {
        SplitPath(AOKFile, , &AOKDir)
        FoundLocations.Push(ChosenDir := AOKDir)
    }
    If AOCFile := GRGamePath(TextFound, 'age2_x1.exe') {
        SplitPath(AOCFile, , &AOCDir)
        SplitPath(AOCDir, , &AOCDir)
        FoundLocations.Push(ChosenDir := AOCDir)
    }
    If FOEFile := GRGamePath(TextFound, 'age2_x2.exe') {
        SplitPath(FOEFile, , &FOEDir)
        SplitPath(FOEDir, , &FOEDir)
        FoundLocations.Push(ChosenDir := FOEDir)
    }
    FoundLocations := RemoveDuplications(FoundLocations)
    RemoveDuplications(Arr) {
        E := ''
        For Each, Value in Arr {
            If !InStr(E, Value) {
                E .= E = '' ? Value : ',' Value
            }
        }
        Return StrSplit(E, ',')
    }
    If (FoundLocations.Length > 1) {
        Location := Gui(, 'Pick one location')
        Location.OnEvent('Close', (*) => DoNothing())
        DoNothing() {
            Location.Destroy()
            Return
        }
        Location.AddText('Center r4 w350 cRed', 'Different locations were found for Age of Empires II game in GameRanger setting`n`nPlease pick only one`n').SetFont('Bold')
        For Locatio in FoundLocations {
            Location.AddRadio('wp Center', Locatio).OnEvent('Click', SetLocation)
        }
        SetLocation(Ctrl, Info) {
            PickedLocation := Ctrl.Text
        }
        PickLocation := Location.AddButton('wp', 'OK').OnEvent('Click', (*) => ApplyLocation())
        PickedLocation := ''
        ApplyLocation() {
            If !DirExist(PickedLocation) {
                MsgBox('Please choose a location first!', 'Nothing was chosen', 48)
                Return
            }
            Location.Destroy()
            IniWrite(PickedLocation, Config, 'Game', 'Path')
            ChosenFolder.Value := PickedLocation
            ChargeSettings________()
            Return
        }
        Location.Show()
        Return
    }
    If !ChosenDir
        Return
    IniWrite(ChosenDir, Config, 'Game', 'Path')
    ChosenFolder.Value := ChosenDir
    ChargeSettings________()
}
LoadGRSettingText() {
    Setting := FileRead(GRSetting, 'RAW')
    TextFound := ''
    LastTextFound := ''
    ExeAddrs := Map('age2_x1.exe', []
                  , 'age2_x2.exe', []
                  , 'empires2.exe', [])
    Loop Setting.Size {
        Address := A_Index - 1
        Byte := NumGet(Setting, Address, 'UChar')
        If (32 <= Byte && Byte <= 126) || (Byte = 10) || (Byte = 13) {
            TextFound .= Chr(Byte)
            LastTextFound .= Chr(Byte)
        } Else {
            TextFound .= '`n'
            For Exe, Addrs in ExeAddrs {
                If InStr(LastTextFound, Exe) {
                    Addrs.Push([Address - StrLen(LastTextFound), LastTextFound])
                }
            }
            LastTextFound := ''
        }
    }
    Return [TextFound, ExeAddrs]
}
ChosenFolder := Manager.AddEdit('xm+10 w500 ReadOnly -E0x200 BackgroundFFDBB7')
Features['The Game'].Push(ChosenFolder)
ChosenFolder.SetFont('Bold')
OpenTheGameFolder := Manager.AddButton('xm+10 w200', 'Open the selected')
GuiButtonIcon(OpenTheGameFolder, 'DB\000\folder.png',, 'a1 r5')
Features['The Game'].Push(OpenTheGameFolder)
OpenTheGameFolder.OnEvent('Click', (*) => Run(ChosenFolder.Value))
H := Manager.AddText('xm+10 yp+50', "Don't have it? Download now! ")
GetTheGame := Manager.AddButton('wp', 'Download')
GuiButtonIcon(GetTheGame, 'DB\000\download.png',, 'a1 r5')
Features['The Game'].Push(GetTheGame)
GetTheGame.OnEvent('Click', (*) => DownloadInstallGame())
ProgressBar := Manager.AddProgress('xp yp wp h25 Hidden -Smooth', 0)
Features['The Game'].Push(ProgressBar)
ProgressInfo := Manager.AddText('xp yp+25 wp Hidden Center BackgroundTrans cRed')
Features['The Game'].Push(ProgressInfo)
DownloadInstallGame() {
    ExportRange := 4
    ProgressInfo.Value := ''
    ProgressBar.Value := 0
    ProgressBar.Opt('Range0-' GamePackages.Length + 3)
    GameSectionInstallView()
    GameSectionInstallView() {
        GetTheGame.Visible := False
        ProgressBar.Visible := True
        ProgressInfo.Visible := True
    }
    If !ExportDir := FileSelect('D') {
        GameSectionNormalView()
        GameSectionNormalView() {
            GetTheGame.Visible := True
            ProgressBar.Visible := False
            ProgressInfo.Visible := False
        }
        Return
    }
    ExportDir := RTrim(ExportDir, '\')
    If DirExist(ExportDir '\Age of Empires II') {
        Choice := MsgBox('Game seems to be already exported at this location!`n`nOverwrite?', 'Game Exist', 0x4 + 0x30)
        If Choice != 'Yes' {
            GameSectionNormalView()
            Return
        }
    }
    Try {
        ProgressInfo.Value := 'Downloading...'
        For Each, Package in GamePackages {
            PackagePath := StrReplace(Package, '/', '\')
            If !FileExist(PackagePath) {
                Download(Server '/' User '/' Repo '/main/' Package, PackagePath)
                If !PackageIsValid(PackagePath)
                    ExitApp()
            }
            ProgressBar.Value += 1
            ProgressInfo.Value := Package ' - Downloaded'
        }
        ProgressInfo.Value := 'Unpacking...'
        Loop 3 {
            RunWait('DB\7za.exe x DB\00' (2 + A_Index) '.7z.001 -o"' ExportDir '\Age of Empires II" -aoa', , 'Hide')
            ProgressBar.Value += 1
            ProgressInfo.Value := 'DB\00' (2 + A_Index) '.7z.001 - Unpacked'
        }
    } Catch As Err {
        GameSectionNormalView()
        MsgBox('Unable to get the game!', 'Oops!', '48')
        Return
    }
    Choice := MsgBox('Done!`n`nGame located at: "' ExportDir '\Age of Empires II"`n`nWanna create shortcuts on your desktop?', 'Question', 0x20 + 0x4)
    If Choice = 'Yes' {
        ChosenFolder.Value := ExportDir '\Age of Empires II'
        CreateGameShortcuts()
    }
    Choice := MsgBox('Done!`n`nGame located at: "' ExportDir '\Age of Empires II"`n`nWanna select this game location?', 'Question', 0x20 + 0x4)
    If Choice = 'Yes' {
        ChosenFolder.Value := ExportDir '\Age of Empires II'
        IniWrite(ChosenFolder.Value, Config, 'Game', 'Path')
        ChargeSettings________()
    }
    GameSectionNormalView()
}
CreateGameShortcuts() {
    If FileExist(ChosenFolder.Value '\empires2.exe') && !FileExist(A_Desktop '\Age of Empires II.lnk')
        FileCreateShortcut(ChosenFolder.Value '\empires2.exe', A_Desktop '\Age of Empires II.lnk')
    If FileExist(ChosenFolder.Value '\age2_x1\age2_x1.exe') && !FileExist(A_Desktop '\The Conquerors.lnk')
        FileCreateShortcut(ChosenFolder.Value '\age2_x1\age2_x1.exe', A_Desktop '\The Conquerors.lnk')
    If FileExist(ChosenFolder.Value '\age2_x1\age2_x2.exe') && !FileExist(A_Desktop '\Forgotten Empires.lnk')
        FileCreateShortcut(ChosenFolder.Value '\age2_x1\age2_x2.exe', A_Desktop '\Forgotten Empires.lnk')
}
; # App
Features['App'] := []
_App_ := Manager.AddText('xm yp+50 w600 h26 cBlue', 'OPTIONS:')
Features['App'].Push(_App_)
_App_.SetFont('Bold s14')
OpenDB := Manager.AddButton('xm+10 yp+30', 'Open the app DB folder')
Features['App'].Push(OpenDB)
OpenDB.OnEvent('Click', (*) => Run(AppDir[1]))
OpenSetting := Manager.AddButton(, 'Open the setting DB folder')
Features['App'].Push(OpenSetting)
OpenSetting.OnEvent('Click', (*) => Run(AppDir[2]))
AtStartUp := Manager.AddCheckbox(, 'Auto launch the app when windows starts')
If IniRead(Config, 'Game', 'Startup', 0) {
    AtStartUp.Value := 1
}
AtStartUp.OnEvent('Click', (*) => StartUpUpdate())
StartUpUpdate() {
    IniWrite(AtStartUp.Value, Config, 'Game', 'Startup')
    If AtStartUp.Value {
        FileCreateShortcut(A_ScriptFullPath, Startup)
    } Else {
        FileDelete(Startup)
    }
}
UpdateChk := Manager.AddCheckbox(, 'Check for updates when the app starts')
If IniRead(Config, 'Game', 'UpdateChk', 0) {
    UpdateChk.Value := 1
}
UpdateChk.OnEvent('Click', (*) => IniWrite(UpdateChk.Value, Config, 'Game', 'UpdateChk'))
; # Versions
Features['Versions'] := []
; # Compatibilities
Features['Compatibilities'] := []
_Version_ := Manager.AddText('xm yp+50 w600 h26 cBlue', 'GAME VERSIONS:')
Features['Versions'].Push(_Version_)
_Version_.SetFont('Bold s14')
H := Manager.AddButton('xm+69 yp+30 w36 h36')
CreateImageButton(H, 0, [['DB\000\aok_normal.png'], ['DB\000\aok_hover.png'], ['DB\000\aok_click.png'], ['DB\000\aok_disable.png']]*)
H.OnEvent('Click', (*) => Run(ChosenFolder.Value '\empires2.exe', ChosenFolder.Value))
Features['Versions'].Push(H)
H := Manager.AddText('xp-59 yp+40 cRed w150 Center BackgroundTrans', 'The Age of Kings')
Features['Versions'].Push(H)
H.SetFont('Bold')
AoKCom := Manager.AddDropDownList('w150')
Features['Compatibilities'].Push(AoKCom)
For Each, Compat in Compatibilities {
    AoKCom.Add([Compat[1]])
}
AoKCom.Choose(1)
AoKCom.OnEvent("Change", (*) => AoKComReg())
AoKRun := Manager.AddCheckbox('xp yp+30 wp hp', 'Run as administrator')
Features['Compatibilities'].Push(AoKRun)
AoKRun.OnEvent("Click", (*) => AoKComReg())
AoKComReg() {
    RegVal := Compatibilities[AoKCom.Value][2] (Compatibilities[AoKCom.Value][2] ? ' ' : '') (AoKRun.Value ? 'RUNASADMIN' : '')
    If !RegVal {
        Try {
            RegDelete(Layers, ChosenFolder.Value '\empires2.exe')
        }
        Return
    }
    RegWrite(RegVal, 'REG_SZ', Layers, ChosenFolder.Value '\empires2.exe')
}
Loop Files, 'DB\002\2*', 'D' {
    Handle := Manager.AddRadio('w100 BackgroundFFFFFF', A_LoopFileName)
    Features['Versions'].Push(Handle)
    Handle.SetFont('s10', 'Consolas')
    Handle.OnEvent('Click', ApplyVersion)
    General['AOK']['VersionsN'][A_LoopFileName] := Handle
}
ApplyVersion(Ctrl, Info?, Wololo := 1) {
    SectionInteract(Features['Versions'], False)
    If GameIsRunning() {
        ChargeSettings________()
        Return
    }
    Try {
        CleanUp(Ctrl.Text)
        SetVersion(Ctrl.Text, !Wololo ? 'AOC' : '')
    } Catch As Err {
        MsgBox('An error occured while trying to set v' Ctrl.Text, 'Version apply error!', 0x20)
    }
    SectionInteract(Features['Versions'])
    UpdateVersionRadio()
    If Wololo
        SoundPlay('DB\000\30 wololo.mp3')
}
UpdateVersionRadio() {
    Fix := IniRead(Config, 'Game', 'Fix', '')
    Fix := StrSplit(Fix)
    If Fix.Length != 2 {
        NoPatch.Value := 1
    }
    If Fix[1] = 1 {
        Patch1.Value := 1
    }
    If Fix[1] = 2 {
        Patch2.Value := 1
        If Fix[2] = 1 {
            WideScreen.Value := 1
        }
        If Fix[2] = 2 {
            CWideScreen.Value := 1
        }
        If Fix[2] = 3 {
            AIWideScreen.Value := 1
        }
        If Fix[2] = 4 {
            AIOWideScreen.Value := 1
        }
    }
    If NoPatch.Value || Patch1.Value {
        DisableSubRadio()
    }
    If Patch2.Value {
        EnableSubRadio()
    }
}
CleanUp(Patch) {
    Edition := SubStr(Patch, 1, 1)
    Loop Files, 'DB\002\' Edition '*', 'D' {
        Version := A_LoopFileName
        Loop Files, 'DB\002\' Version '\*.*', 'R' {
            PatchFile := A_LoopFileDir '\' A_LoopFileName
            GameFile := ChosenFolder.Value StrReplace(PatchFile, 'DB\002\' Version)
            If FileExist(GameFile) {
                FileDelete(GameFile)
            }
        }
    }
    Loop Files, 'DB\001\*', 'D' {
        Fix := A_LoopFileName
        Loop Files, 'DB\001\' Fix '\' Edition '*', 'D' {
            Version := A_LoopFileName
            Loop Files, 'DB\001\' Fix '\' Version '\*.*', 'R' {
                PatchFile := A_LoopFileDir '\' A_LoopFileName
                GameFile := ChosenFolder.Value StrReplace(PatchFile, 'DB\001\' Fix '\' Version)
                If FileExist(GameFile) {
                    FileDelete(GameFile)
                }
            }
        }
    }
    If RegRead('HKEY_CURRENT_USER\SOFTWARE\Microsoft\Microsoft Games\Age of Empires', 'Aoe2Patch', '')
        RegDelete('HKEY_CURRENT_USER\SOFTWARE\Microsoft\Microsoft Games\Age of Empires', 'Aoe2Patch')
}
GameIsRunning() {
    Processes := ['empires2.exe', 'age2_x1.exe', 'age2_x2.exe']
    For Each, Process in Processes {
        If ProcessExist(Process) {
            ProcessClose(Process)
        }
        ProcessWaitClose(Process, 5)
        If ProcessExist(Process) {
            Return True
        }
    }
    Return False
}
SetVersion(Version, Select := '') {
    If General['AOC']['Combine'].Has(Version) {
        For Each, pVersion in General['AOC']['Combine'][Version] {
            DirCopy('DB\002\' pVersion, ChosenFolder.Value, 1)
        }
    }
    If General['AOK']['Combine'].Has(Version) {
        For Each, pVersion in General['AOK']['Combine'][Version] {
            DirCopy('DB\002\' pVersion, ChosenFolder.Value, 1)
        }
    }
    DirCopy('DB\002\' Version, ChosenFolder.Value, 1)
    If (Select != '') {
        General[Select]['VersionsN'][Version].Value := 1
    }
    If NoPatch.Value {
        Return
    }
    If Patch1.Value {
        If !DirExist('DB\001\Enable Fix v1\' Version) {
            NoPatch.Value := 1
            IniWrite('0', Config, 'Game', 'Fix')
            Return
        }
        DirCopy('DB\001\Enable Fix v1\' Version, ChosenFolder.Value, 1)
        DirCopy('DB\001\Enable Fix v1\Static', ChosenFolder.Value, 1)
    }
    If Patch2.Value {
        If !DirExist('DB\001\Enable Fix v2\' Version) {
            NoPatch.Value := 1
            IniWrite('0', Config, 'Game', 'Fix')
            Return
        }
        DirCopy('DB\001\Enable Fix v2\' Version, ChosenFolder.Value, 1)
        DirCopy('DB\001\Enable Fix v2\Static', ChosenFolder.Value, 1)
        If AIWideScreen.Value {
            RegWrite(1, 'REG_DWORD', 'HKEY_CURRENT_USER\SOFTWARE\Microsoft\Microsoft Games\Age of Empires', 'Aoe2Patch')
        }
        If AIOWideScreen.Value {
            RegWrite(2, 'REG_DWORD', 'HKEY_CURRENT_USER\SOFTWARE\Microsoft\Microsoft Games\Age of Empires', 'Aoe2Patch')
        }
        If WideScreen.Value {
            RegWrite(3, 'REG_DWORD', 'HKEY_CURRENT_USER\SOFTWARE\Microsoft\Microsoft Games\Age of Empires', 'Aoe2Patch')
        }
        If CWideScreen.Value {
            RegWrite(4, 'REG_DWORD', 'HKEY_CURRENT_USER\SOFTWARE\Microsoft\Microsoft Games\Age of Empires', 'Aoe2Patch')
        }
    }
}
ChargeVersions________() {
    FoundVersions := 0
    Loop Files, 'DB\002\*', 'D' {
        Version := A_LoopFileName
        CountedFiles := 0
        EqualHashCount := 0
        Loop Files, 'DB\002\' Version '\*.*', 'R' {
            ++CountedFiles
        }
        Loop Files, 'DB\002\' Version '\*.*', 'R' {
            PatchFile := A_LoopFileDir '\' A_LoopFileName
            GameFile := ChosenFolder.Value StrReplace(PatchFile, 'DB\002\' Version)
            If FileExist(GameFile) && (HashFile(PatchFile) = HashFile(GameFile)) {
                ++EqualHashCount
            }
        }
        If CountedFiles = EqualHashCount {
            Flag := SubStr(Version, 1, 1)
            Switch Flag {
                Case 1:
                    General['AOC']['VersionsN'][Version].Value := 1
                    ++FoundVersions
                Case 2:
                    General['AOK']['VersionsN'][Version].Value := 1
                    ++FoundVersions
            }
        }
    }
    If FoundVersions < 2 {
        Choice := MsgBox('Unable to find out the game version(s)`nWould you like to apply the defaults?', 'Info', 0x4 + 0x20)
        If Choice = 'Yes' {
            General['AOK']['VersionsN']['2.0  CD'].Value := 1
            SetVersion('2.0  CD')
            General['AOC']['VersionsN']['1.0  CD'].Value := 1
            SetVersion('1.0  CD')
        }
    }
}
_Version_.GetPos(&X, &Y)
H := Manager.AddButton('xp+219 y' (Y + 30) ' w36 h36')
CreateImageButton(H, 0, [['DB\000\aoc_normal.png'], ['DB\000\aoc_hover.png'], ['DB\000\aoc_click.png'], ['DB\000\aoc_disable.png']]*)
H.OnEvent('Click', (*) => Run(ChosenFolder.Value '\age2_x1\age2_x1.exe', ChosenFolder.Value '\age2_x1'))
Features['Versions'].Push(H)
H := Manager.AddText('xp-59 yp+40 cBlue w150 Center BackgroundTrans', 'The Conquerors')
Features['Versions'].Push(H)
H.SetFont('Bold')
AoCCom := Manager.AddDropDownList('w150')
Features['Compatibilities'].Push(AoCCom)
For Each, Compat in Compatibilities {
    AoCCom.Add([Compat[1]])
}
AoCCom.Choose(1)
AoCCom.OnEvent("Change", (*) => AoCComReg())
AoCRun := Manager.AddCheckbox('xp yp+30 wp hp', 'Run as administrator')
Features['Compatibilities'].Push(AoCRun)
AoCRun.OnEvent("Click", (*) => AoCComReg())
AoCComReg() {
    RegVal := Compatibilities[AoCCom.Value][2] (Compatibilities[AoCCom.Value][2] ? ' ' : '') (AoCRun.Value ? 'RUNASADMIN' : '')
    If !RegVal {
        Try {
            RegDelete(Layers, ChosenFolder.Value '\age2_x1\age2_x1.exe')
        }
        Return
    }
    RegWrite(RegVal, 'REG_SZ', Layers, ChosenFolder.Value '\age2_x1\age2_x1.exe')
}
Loop Files, 'DB\002\1*', 'D' {
    Handle := Manager.AddRadio('w30 w100', A_LoopFileName)
    Features['Versions'].Push(Handle)
    Handle.SetFont('s10', 'Consolas')
    Handle.OnEvent('Click', ApplyVersion)
    General['AOC']['VersionsN'][A_LoopFileName] := Handle
}
H := Manager.AddButton('xp+219 y' (Y + 30) ' w36 h36')
CreateImageButton(H, 0, [['DB\000\fe_normal.png'], ['DB\000\fe_hover.png'], ['DB\000\fe_click.png'], ['DB\000\fe_disable.png']]*)
H.OnEvent('Click', (*) => Run(ChosenFolder.Value '\age2_x1\age2_x2.exe', ChosenFolder.Value '\age2_x1'))
Features['Versions'].Push(H)
H := Manager.AddText('xp-59 yp+40 cGreen w150 Center BackgroundTrans', 'Forgotten Empires')
Features['Versions'].Push(H)
H.SetFont('Bold')
FOECom := Manager.AddDropDownList('w150')
Features['Compatibilities'].Push(FOECom)
For Each, Compat in Compatibilities {
    FOECom.Add([Compat[1]])
}
FOECom.Choose(1)
FOECom.OnEvent("Change", (*) => FOEComReg())
FOERun := Manager.AddCheckbox('xp yp+30 wp hp', 'Run as administrator')
Features['Compatibilities'].Push(FOERun)
FOERun.OnEvent("Click", (*) => FOEComReg())
FOEComReg() {
    RegVal := Compatibilities[FOECom.Value][2] (Compatibilities[FOECom.Value][2] ? ' ' : '') (FOERun.Value ? 'RUNASADMIN' : '')
    If !RegVal {
        Try {
            RegDelete(Layers, ChosenFolder.Value '\age2_x1\age2_x2.exe')
        }
        Return
    }
    RegWrite(RegVal, 'REG_SZ', Layers, ChosenFolder.Value '\age2_x1\age2_x2.exe')
}
Handle := Manager.AddRadio('w30 w100 Checked', '2.2  CD')
Features['Versions'].Push(Handle)
Handle.SetFont('s10', 'Consolas')
General['FOE']['VersionsN']['2.2  CD'] := Handle
H := Manager.AddText('xm+10 cBlue', 'Updates to be applied after changing the version:')
Features['Versions'].Push(H)
NoPatch := Manager.AddRadio(, 'None')
NoPatch.OnEvent('Click', (*) => DisableSubRadio())
DisableSubRadio() {
    WideScreen.Enabled := False
    CWideScreen.Enabled := False
    AIWideScreen.Enabled := False
    AIOWideScreen.Enabled := False
    If NoPatch.Value {
        IniWrite('0', Config, 'Game', 'Fix')
    }
    If Patch1.Value {
        IniWrite('10', Config, 'Game', 'Fix')
    }
}
Features['Versions'].Push(NoPatch)
Patch1 := Manager.AddRadio(, 'Enable Fix v1')
Patch1.OnEvent('Click', (*) => DisableSubRadio())
Features['Versions'].Push(Patch1)
Patch2 := Manager.AddRadio(, 'Enable Fix v2')
Features['Versions'].Push(Patch2)
Patch2.OnEvent('Click', (*) => EnableSubRadio())
EnableSubRadio() {
    WideScreen.Enabled := True
    CWideScreen.Enabled := True
    AIWideScreen.Enabled := True
    AIOWideScreen.Enabled := True
}
WideScreen := Manager.AddRadio('xp+20 yp+25 Group', 'Wide Screen')
WideScreen.OnEvent('Click', (*) => IniWrite('21', Config, 'Game', 'Fix'))
Features['Versions'].Push(WideScreen)
CWideScreen := Manager.AddRadio(, 'Centred Wide Screen')
CWideScreen.OnEvent('Click', (*) => IniWrite('22', Config, 'Game', 'Fix'))
Features['Versions'].Push(CWideScreen)
AIWideScreen := Manager.AddRadio(, 'Tech Overlay + Wide Screen')
AIWideScreen.OnEvent('Click', (*) => IniWrite('23', Config, 'Game', 'Fix'))
Features['Versions'].Push(AIWideScreen)
AIOWideScreen := Manager.AddRadio(, 'Tasks Overlay + Tech Overlay + Wide Screen')
AIOWideScreen.OnEvent('Click', (*) => IniWrite('24', Config, 'Game', 'Fix'))
Features['Versions'].Push(AIOWideScreen)
; # Language
Features['Language'] := []
_Language_ := Manager.AddText('xm yp+50 cBlue w600 h26', 'GAME INTERFACE LANGUAGE:')
Features['Language'].Push(_Language_)
_Language_.SetFont('Bold s14')
Loop Files, 'DB\006\*', 'D' {
    Handle := Manager.AddRadio(, A_LoopFileName)
    Features['Language'].Push(Handle)
    Handle.SetFont('Bold')
    Handle.OnEvent('Click', ApplyLanguage)
    General['LNG'][A_LoopFileName] := Handle
}
ApplyLanguage(Ctrl, Info) {
    SectionInteract(Features['Language'], False)
    Sleep(500)
    If !((Time := FoundDefaultLanguage()) && ((Ctrl.Text = '___Default___'))) {
        DirCopy('DB\006\' Ctrl.Text, ChosenFolder.Value, 1)
    } Else {
        DirCopy(AppDir[2] '\' Time, ChosenFolder.Value, 1)
    }
    SectionInteract(Features['Language'])
    SoundPlay('DB\000\30 wololo.mp3')
}
FoundDefaultLanguage() {
    Now := ''
    If NowKeys := IniRead(Config, 'FoundLanguage', , '') {
        For Every, NowPath in StrSplit(NowKeys, '`n') {
            NowPathArr := StrSplit(NowPath, '=')
            If NowPathArr[2] = ChosenFolder.Value {
                Now := NowPathArr[1]
                Break
            }
        }
    }
    Return Now
}
BackupDefaultLanguage_() {
    If !Now := FoundDefaultLanguage()
        IniWrite(ChosenFolder.Value, Config, 'FoundLanguage', Now := A_Now)
    Loop Files, 'DB\006\*', 'D' {
        LanguageName := A_LoopFileName
        Loop Files, 'DB\006\' LanguageName '\*.*', 'R' {
            LngPath := A_LoopFileDir '\' A_LoopFileName
            GamePath := ChosenFolder.Value StrReplace(LngPath, 'DB\006\' LanguageName)
            If FileExist(GamePath) {
                DefPath := StrReplace(LngPath, 'DB\006\' LanguageName, AppDir[2] '\' Now)
                SplitPath(DefPath, , &OutDir)
                If !DirExist(OutDir) {
                    DirCreate(OutDir)
                }
                If !FileExist(DefPath) {
                    FileCopy(GamePath, DefPath)
                }
            }
        }
    }
}
ChargeLanguage________() {
    Loop Files, 'DB\006\*', 'D' {
        If (A_LoopFileName = '___Default___') && (Time := FoundDefaultLanguage()) {
            Language := AppDir[2] '\' Time
        } Else {
            Language := 'DB\006\' A_LoopFileName
        }
        Found := True
        CountedFiles := 0
        EqualHashCount := 0
        Loop Files, Language '\*.dll', 'R' {
            ++CountedFiles
        }
        Loop Files, Language '\*.dll', 'R' {
            LngFile := A_LoopFileDir '\' A_LoopFileName
            GameFile := ChosenFolder.Value StrReplace(LngFile, Language)
            If FileExist(GameFile) && (HashFile(LngFile) = HashFile(GameFile)) {
                ++EqualHashCount
            }
        }
        If CountedFiles && (CountedFiles = EqualHashCount) {
            General['LNG'][A_LoopFileName].Value := 1
            Break
        }
    }
}
; # Visual Mods
Features['Visual Modes'] := []
_VisualMods_ := Manager.AddText('yp+50 w600 h26 cBlue', 'GAME VISUAL MODS:')
Features['Visual Modes'].Push(_VisualMods_)
_VisualMods_.SetFont('Bold s14')
Features['Visual Modes'].Push(H)
Loop Files, 'DB\007\*', 'D' {
    M := Manager.AddText('xm+10 cRed yp' (A_Index = 1 ? 25 : 50), A_Index ' - ' A_LoopFileName)
    M.SetFont('Bold s10')
    Features['Visual Modes'].Push(M)
    M := Manager.AddPicture('Border', 'DB\007\' A_LoopFileName '\img.png')
    Features['Visual Modes'].Push(M)
    M := Manager.AddButton('xm+10 w200', 'Install ' A_LoopFileName)
    M.OnEvent('Click', ApplyVM)
    Features['Visual Modes'].Push(M)
    M := Manager.AddButton('yp w200', 'Uninstall ' A_LoopFileName)
    M.OnEvent('Click', ApplyVM)
    Features['Visual Modes'].Push(M)
}
ApplyVM(Ctrl, Info) {
    Ctrl.GetPos(&X, &Y, &W, &H)
    ProgressBar.Move(X, Y, W, H)
    Ctrl.Visible := False
    ProgressBar.Visible := True
    ProgressBar.Value := 0
    SectionInteract(Features['Visual Modes'], False)
    VMName := SubStr(Ctrl.Text, InStr(Ctrl.Text, ' ') + 1)
    SlpDir := InStr(Ctrl.Text, 'Uninstall') ? 'DB\007\' VMName '\U' : 'DB\007\' VMName
    RunWait('DB\000\DrsBuild.exe /a "' ChosenFolder.Value '\Data\' DrsTypes['gra'] '" "' SlpDir '\gra*.slp"',, 'Hide')
    ProgressBar.Value := 20
    RunWait('DB\000\DrsBuild.exe /a "' ChosenFolder.Value '\Data\' DrsTypes['int'] '" "' SlpDir '\int*.slp"',, 'Hide')
    ProgressBar.Value := 40
    RunWait('DB\000\DrsBuild.exe /a "' ChosenFolder.Value '\Data\' DrsTypes['ter'] '" "' SlpDir '\ter*.slp"',, 'Hide')
    ProgressBar.Value := 60
    If FileExist(SlpDir '\Info.ini') {
        Drs := IniRead(SlpDir '\Info.ini', 'Info', 'Drs', '')
        FileN := IniRead(SlpDir '\Info.ini', 'Info', 'File', '')
        Lines := StrSplit(IniRead(SlpDir '\Info.ini', 'Info', 'Line', ''), ',')
        Values := StrSplit(IniRead(SlpDir '\Info.ini', 'Info', 'Value', ''), ',')
        RunWait('DB\000\DrsBuild.exe /e "' ChosenFolder.Value '\Data\' Drs '" ' FileN ' /o "' ChosenFolder.Value '\Data"',, 'Hide')
        OBJ := FileOpen(ChosenFolder.Value '\Data\' FileN, 'r')
        NValues := Map()
        While !OBJ.AtEOF {
            Index := Format('{:03}', A_Index)
            NValues[Index] := OBJ.ReadLine()
        }
        OBJ.Close()
        For Index, Line in Lines {
            NValues[Line] := Values[Index]
        }
        OBJ := FileOpen(ChosenFolder.Value '\Data\' FileN, 'w')
        For Index, Line in NValues {
            OBJ.WriteLine(Line)
        }
        OBJ.Close()
        RunWait('DB\000\DrsBuild.exe /a "' ChosenFolder.Value '\Data\' Drs '" "' ChosenFolder.Value '\Data\' FileN '"',, 'Hide')
        FileDelete(ChosenFolder.Value '\Data\' FileN)
    }
    ProgressBar.Value := 80
    If FileExist(ChosenFolder.Value '\Games\age2_x1.xml') {
        If RegExMatch(FileRead(ChosenFolder.Value '\Games\age2_x1.xml'), '\Q<path>\E(.*)\Q</path>\E', &DName) {
            SectionInteract(Features['Data Mods'], False)
            RunWait('DB\000\DrsBuild.exe /a "' ChosenFolder.Value '\Games\' DName[1] '\Data\gamedata_x1_p1.drs" "' SlpDir '\*.slp"',, 'Hide')
            SectionInteract(Features['Data Mods'])
        }
    }
    ProgressBar.Value := 100
    SectionInteract(Features['Visual Modes'])
    SoundPlay('DB\000\30 wololo.mp3')
    Ctrl.Visible := True
    ProgressBar.Visible := False
}
;LoadVM := Manager.AddButton('wp Disabled', 'Import')
;LoadVM.SetFont('Bold')
;LoadVM.OnEvent('Click', (*) => ImportVisualMod())
ImportVisualMod() {
    MsgBox('Make sure the visual mod you select is compatible with your game!', 'Notice', 0x40)
    If Selected := FileSelect('D') {
        SplitPath(Selected, &ModeName)
        If DirExist('DB\007\' ModeName) {
            MsgBox(ModeName ' is already imported!', ModeName, 0x30)
            Return
        }
        SectionInteract(Features['Visual Modes'], False)
        DirCreate('DB\007\' ModeName '\U')
        Loop Files, Selected '\*.slp', 'R' {
            ID := SubStr(A_LoopFileName, 1, -4)
            If !IsDigit(ID) {
                Continue
            }
            LZID := Format("{:0" IDL "}", ID)
            If ID >= DrsRange['gra'][1] && ID <= DrsRange['gra'][2] {
                Name := 'gra' LZID '.slp'
            }
            If ID >= DrsRange['int'][1] && ID <= DrsRange['int'][2] {
                Name := 'int' LZID '.slp'
            }
            If ID >= DrsRange['ter'][1] && ID <= DrsRange['ter'][2] {
                Name := 'ter' LZID '.slp'
            }
            FileCopy(A_LoopFileFullPath, 'DB\007\' ModeName '\' Name, 1)
            DecodeSlp(FileName) {
                F := FileRead(FileName, 'RAW m4')
                H := ''
                Loop 4 {
                    H .= Format('{:02X}', NumGet(F, A_Index - 1, 'UChar'))
                }
                If H = NormalSlp {
                    Return True
                }
                If H = VCodedSlp {
                    F := FileRead(FileName, 'RAW')
                    NF := Buffer(F.Size - 4)
                    Loop NF.Size {
                        Byte := NumGet(F, (A_Index - 1) + 4, 'UChar')
                        Val := (Byte - 17) ^ 0x23
                        UChar := Val & 0xFF
                        NByte := (0x20 * (Val) | (UChar >> 3))
                        NumPut('UChar', NByte, NF, A_Index - 1)
                    }
                    FileOpen(FileName, 'w').RawWrite(NF, NF.Size)
                    Return True
                }
                Return True
            }
            If !DecodeSlp('DB\007\' ModeName '\' Name) {
                FileDelete('DB\007\' ModeName '\' Name)
            }
        }
        Loop Files, 'DB\007\U', 'DR' {
            If InStr(A_LoopFileDir, ModeName) {
                Continue
            }
            Loop Files, A_LoopFileDir '\U\*.slp' {
                If FileExist('DB\007\' ModeName '\' A_LoopFileName) {
                    FileCopy(A_LoopFileFullPath, 'DB\007\' ModeName '\U\' A_LoopFileName, 1)
                }
            }
        }
        Loop Files, 'DB\007\' ModeName '\*.slp' {
            Flag := SubStr(A_LoopFileName, 1, 3)
            If !FileExist('DB\007\' ModeName '\U\' A_LoopFileName) {
                RunWait('DB\000\DrsBuild.exe /e "' ChosenFolder.Value '\Data\' DrsTypes[Flag] '" ' A_LoopFileName ' /o "DB\007\' ModeName '\U"', , 'Hide')
            }
        }
        MsgBox(ModeName ' should be added to the list by now!', 'Info', 0x40)
        SectionInteract(Features['Visual Modes'])
    }
}

; # Data Mods
Features['Data Mods'] := []
_DataMods_ := Manager.AddText('xm yp+50 w600 h26 cBlue', 'GAME DATA MODS:')
Features['Data Mods'].Push(_DataMods_)
_DataMods_.SetFont('Bold s14')
Features['Data Mods'].Push(H)
For Each, _Mod in StrSplit(IniRead('DB\008\DataMod.ini', 'DataMod',, ''), '`n') {
    ModName := StrSplit(_Mod, '=')[1]
    M := Manager.AddText('xm+10 cRed yp' (A_Index = 1 ? 25 : 50), A_Index ' - ' ModName)
    M.SetFont('Bold s10 Underline')
    Features['Data Mods'].Push(M)
    M := Manager.AddPicture('Border', 'DB\008\' ModName '.png')
    Features['Data Mods'].Push(M)
    M := Manager.AddButton('xm+10 w200', 'Install ' ModName)
    M.OnEvent('Click', ApplyDM)
    Features['Data Mods'].Push(M)
    M := Manager.AddButton('yp w200', 'Uninstall ' ModName)
    M.OnEvent('Click', ApplyDM)
    Features['Data Mods'].Push(M)
}
ApplyDM(Ctrl, Info) {
    Ctrl.GetPos(&X, &Y, &W, &H)
    ProgressBar.Move(X, Y, W, H)
    Ctrl.Visible := False
    ProgressBar.Visible := True
    ProgressBar.Value := 0
    DMName := SubStr(Ctrl.Text, InStr(Ctrl.Text, ' ') + 1)
    If !InStr(Ctrl.Text, 'Uninstall') {
        SectionInteract(Features['Data Mods'], False)
        ModeDir := IniRead('DB\008\DataMod.ini', 'DataMod', DMName, '')
        ModeDir := StrSplit(ModeDir, '|')
        Parts := StrSplit(ModeDir[2], ',')
        PrepareTheDataMod()
        ProgressBar.Value := 25
        PrepareTheDataMod() {
            Try {
                For Each, Part in Parts {
                    If !FileExist('DB\' ModeDir[1] '.7z.' Part) {
                        Download(Server '/' User '/' Repo '/main/DB/' ModeDir[1] '.7z.' Part, 'DB\' ModeDir[1] '.7z.' Part)
                        If !PackageIsValid('DB\' ModeDir[1] '.7z.' Part) {
                            Return False
                        }
                    }
                }
                If !DirExist('DB\' ModeDir[1])
                    && RunWait('DB\7za.exe x ' 'DB\' ModeDir[1] '.7z.001 -oDB\' ModeDir[1], , 'Hide') {
                        Return False
                }
                Return True
            } Catch As Err {
                Return False
            }
        }
        ApplyVersion(General['AOC']['VersionsN']['1.5  CD'],, 0)
        ProgressBar.Value := 50
        If !DirExist(ChosenFolder.Value '\Games') {
            DirCreate(ChosenFolder.Value '\Games')
        }
        ProgressBar.Value := 75
        If FileExist('DB\' ModeDir[1] '\Games\age2_x1.xml') {
            FileCopy('DB\' ModeDir[1] '\Games\age2_x1.xml', ChosenFolder.Value '\Games\age2_x1.xml', 1)
            If !DirExist(ChosenFolder.Value '\Games\' ModeDir[1])
                DirCopy('DB\' ModeDir[1], ChosenFolder.Value, 1)
        } Else {
            DirCopy('DB\' ModeDir[1], ChosenFolder.Value, 1)
        }
        ProgressBar.Value := 100
    } Else {
        ApplyVersion(General['AOC']['VersionsN']['1.5  CD'],, 0)
        ProgressBar.Value := 20
        If FileExist(ChosenFolder.Value '\Games\age2_x1.xml') {
            FileDelete(ChosenFolder.Value '\Games\age2_x1.xml')
        }
        ProgressBar.Value := 40
        If DirExist(ChosenFolder.Value '\Games\' DMName) {
            DirDelete(ChosenFolder.Value '\Games\' DMName, 1)
        }
        ProgressBar.Value := 60
        If DMName = 'Sheep vs Wolf 2' {
            DirCopy('DB\008\Sound\stream', ChosenFolder.Value '\Sound\stream', 1)
        }
        ProgressBar.Value := 80
        IniDelete(Config, 'Game', 'CurrDM')
        ProgressBar.Value := 100
    }
    SectionInteract(Features['Data Mods'])
    SoundPlay('DB\000\30 wololo.mp3')
    Ctrl.Visible := True
    ProgressBar.Visible := False
}
;ApplyVMDM(Ctrl, Info) {
;    SectionInteract(Features['Data Mods'], False)
;    VMName := SubStr(Ctrl.Text, InStr(Ctrl.Text, ' ') + 1)
;    SlpDir := InStr(Ctrl.Text, 'Uninstall') ? 'DB\007\' VMName '\U' : 'DB\007\' VMName
;    RunWait(A_Clipboard := 'DB\000\DrsBuild.exe /a "' ChosenFolder.Value '\Games\' VMDMTitle.Text '\Data\gamedata_x1_p1.drs" "' SlpDir '\*.slp"', , 'Hide')
;    SectionInteract(Features['Data Mods'])
;    SoundPlay('DB\000\30 wololo.mp3')
;}
;ImportDM := Manager.AddButton('wp Disabled', 'Import')
;ImportDM.SetFont('Bold')
;ImportDM.OnEvent('Click', (*) => ImportDataMod())
ImportDataMod() {
    If Selected := FileSelect('D') {
        SplitPath(Selected, &ModeName)
        If IniRead('DB\008\DataMod.ini', 'DataMod', ModeName, '') {
            MsgBox(ModeName ' is already imported!', ModeName, 0x30)
            Return
        }
        AfterLastDBDir() {
            Loop 100 {
                DID := Format('{:03}', A_Index)
                If !DirExist('DB\' DID) {
                    Return DID
                }
            }
            Return 0
        }
        If !DID := AfterLastDBDir()
            Return
        Dir := 'DB\' DID
        If FileExist('.gitignore') {
            FileAppend('`n' Dir '/', '.gitignore')
        }
        DirCreate(Dir '\' ModeName)
        If !FileExist(Selected '\age2_x1.xml') {
            MsgBox('Unable to find the "age2_x1.xml"!', 'Import failed!', 0x30)
            Return
        }
        SectionInteract(Features['Data Mods'], False)
        DirCopy(Selected, Dir '\' ModeName, 1)
        If DirExist(Dir '\' ModeName '\Drs') {
            Loop Files, Dir '\' ModeName '\Drs\*.*', 'R' {
                ID := SubStr(A_LoopFileName, 1, -4)
                If !IsDigit(ID) {
                    Continue
                }
                LZID := 'gam' Format("{:0" IDL "}", ID)
                FileMove(A_LoopFileFullPath, Dir '\' ModeName '\Drs\' LZID '.' A_LoopFileExt)
            }
            RunWait('DB\000\DrsBuild.exe /a "' Dir '\' ModeName '\Data\gamedata_x1_p1.drs" "' Dir '\' ModeName '\Drs\gam*.*"', , 'Hide')
            DirDelete(Dir '\' ModeName '\Drs', 1)
        }
        FileMove(Dir '\' ModeName '\age2_x1.xml', Dir '\age2_x1.xml')
        DirCreate(Dir '\' ModeName '\mmods')
        FileCopy('DB\000\aoc-language-ini.dll', Dir '\' ModeName '\mmods\aoc-language-ini.dll')
        FileCopy('DB\000\language_x1_p1.dll', Dir '\' ModeName '\Data\language_x1_p1.dll')
        For Each, Wildcard in ['*.mgz', '*.msz', '*.scx', '*.nfz', '*.bmp', '*.png'] {
            Loop Files, Dir '\' ModeName '\' Wildcard, 'R' {
                FileDelete(A_LoopFileFullPath)
            }
        }
        IniWrite(DID, 'DB\008\DataMod.ini', 'DataMod', ModeName)
        SectionInteract(Features['Data Mods'])
    }
}

; # Game Hotkeys
Features['Game Hotkeys'] := []
_ATools_ := Manager.AddText('xm yp+50 w600 h26 cBlue', 'GAME HOTKEYS:')
Features['Game Hotkeys'].Push(_ATools_)
_ATools_.SetFont('Bold s14')
; # Shortcuts
Shortcuts := Manager.AddButton('xm+10', 'Hotkeys')
Features['Game Hotkeys'].Push(Shortcuts)
Shortcuts.SetFont('Bold')
ShortcutsG := Gui(, 'Defined Hotkeys')
ShortcutList := ShortcutsG.AddListView('w605 r10 Checked', ['Hotkey', 'Comment', 'ID'])
CLV := LV_Colors(ShortcutList)
CLV.SelectionColors(0x008000, 0xFFFFFF)
CLV.AlternateRows(0xCCCCCC)
ShortcutList.ModifyCol(3, 0)
ShortcutList.ModifyCol(4, 0)
ShortcutList.ModifyCol(1, 200)
ShortcutList.ModifyCol(2, 400)
ShortcutList.SetFont('Bold')
ShortcutAdd := ShortcutsG.AddButton('w100', 'Add')
ShortcutAdd.OnEvent('Click', (*) => AddShortcut())
AddShortcut() {
    ShortcutNameA.Value := ''
    ShortcutCommentA.Value := ''
    ShortcutActionA.Value := ''
    ShortcutsGA.Show()
}
ShortcutsGA := Gui(, 'Add a hotkey')
ShortcutsGA.AddText(, '1 - Hotkey Name')
ShortcutNameA := ShortcutsGA.AddEdit('w300 cRed')
ShortcutNameA.SetFont('Bold')
ShortcutsGA.AddText(, '2 - Hotkey Comment')
ShortcutCommentA := ShortcutsGA.AddEdit('w300 cGreen')
ShortcutCommentA.SetFont('Bold')
ShortcutsGA.AddText(, '3 - Hotkey Action')
ShortcutActionA := ShortcutsGA.AddEdit('w600 r10 cBlue HScroll')
ShortcutActionA.SetFont('s12', 'Calibri')
ShortcutAddOK := ShortcutsGA.AddButton('w600', 'Submit')
ShortcutAddOK.OnEvent('Click', (*) => SaveShortcutA())
SaveShortcutA() {
    If ShortcutNameA.Value = '' || ShortcutActionA.Value = '' || ShortcutCommentA.Value = '' {
        MsgBox('One of the inputs is not being filled!', 'Fill!', 0x30)
        Return
    }
    Value := IniRead(Config, 'Hotkey', ShortcutNameA.Value, 'NA')
    If Value != 'NA' {
        MsgBox('Hotkey already added!', 'Duplication!', 0x30)
        Return
    }
    Loop {
        ShortcutFile := Format('{:03}', A_Index) '.ahk'
        ShortcutFileM := Format('{:03}', A_Index) '.ahkm'
        If !FileExist(AppDir[3] '\' ShortcutFile) && !FileExist(AppDir[3] '\' ShortcutFileM)
            Break
    }
    O := FileOpen(AppDir[3] '\' ShortcutFile, 'w')
    O.WriteLine(';' ShortcutCommentA.Value ';')
    O.WriteLine('#Requires AutoHotkey v2')
    O.WriteLine('#SingleInstance Force')
    O.WriteLine("GroupAdd('AOKAOC', 'ahk_exe empires2.exe')")
    O.WriteLine("GroupAdd('AOKAOC', 'ahk_exe age2_x1.exe')")
    O.WriteLine("GroupAdd('AOKAOC', 'ahk_exe age2_x2.exe')")
    O.WriteLine('HotIfWinActive("ahk_group AOKAOC")')
    O.WriteLine("Hotkey('" ShortcutNameA.Value "', Action)")
    O.WriteLine("Action(*) {")
    O.WriteLine(ShortcutActionA.Value)
    O.WriteLine('}')
    O.WriteLine('ProcessWaitClose(A_Args[1])')
    O.Write('ExitApp')
    O.Close()
    ShortcutList.Add('Check', ShortcutNameA.Value, ShortcutCommentA.Value, ShortcutFile)
    IniWrite(1, Config, 'Hotkey', ShortcutNameA.Value)
    Run(AppDir[3] '\' ShortcutFile ' ' ProcessExist())
    ShortcutsGA.Hide()
}
ShortcutEdit := ShortcutsG.AddButton('w100 yp', 'Edit')
ShortcutEdit.OnEvent('Click', (*) => EditShortcut())
EditShortcut() {
    If !R := ShortcutList.GetNext() {
        Return
    }
    If R <= 2 {
        Msgbox('You can only un-check this shortcut!', 'Unable to delete!', 0x30)
        Return
    }
    ShortcutNameE.Value := ''
    ShortcutCommentE.Value := ''
    ShortcutActionE.Value := ''
    ShortcutFile := ShortcutList.GetText(R, 3)
    Info := ShortcutGetInfo(AppDir[3] '\' ShortcutFile)
    If !Info['ShortcutOK']
        Return
    ShortcutNameE.Value := R '|' Info['ShortcutName']
    ShortcutCommentE.Value := Info['ShortcutComment']
    ShortcutActionE.Value := Info['ShortcutAction']
    ShortcutsGE.Show()
}
ShortcutsGE := Gui(, 'Edit a hotkey')
ShortcutsGE.AddText(, '1 - Hotkey Name')
ShortcutNameE := ShortcutsGE.AddEdit('w300 cRed ReadOnly')
ShortcutNameE.SetFont('Bold')
ShortcutsGE.AddText(, '2 - Hotkey Comment')
ShortcutCommentE := ShortcutsGE.AddEdit('w300 cGreen')
ShortcutCommentE.SetFont('Bold')
ShortcutsGE.AddText(, '3 - Hotkey Action')
ShortcutActionE := ShortcutsGE.AddEdit('w600 r10 cBlue HScroll')
ShortcutActionE.SetFont('s12', 'Calibri')
ShortcutEditOK := ShortcutsGE.AddButton('w600', 'Submit')
ShortcutEditOK.OnEvent('Click', (*) => SaveShortcutE())
SaveShortcutE() {
    If ShortcutNameE.Value = '' || ShortcutActionE.Value = '' || ShortcutCommentE.Value = '' {
        MsgBox('One of the inputs is not being filled!', 'Fill!', 0x30)
        Return
    }
    ShortcutRef := StrSplit(ShortcutNameE.Value, '|')
    ShortcutFile := ShortcutList.GetText(ShortcutRef[1], 3)
    O := FileOpen(AppDir[3] '\' ShortcutFile, 'w')
    O.WriteLine(';' ShortcutCommentE.Value ';')
    O.WriteLine('#Requires AutoHotkey v2')
    O.WriteLine('#SingleInstance Force')
    O.WriteLine("GroupAdd('AOKAOC', 'ahk_exe empires2.exe')")
    O.WriteLine("GroupAdd('AOKAOC', 'ahk_exe age2_x1.exe')")
    O.WriteLine("GroupAdd('AOKAOC', 'ahk_exe age2_x2.exe')")
    O.WriteLine('HotIfWinActive("ahk_group AOKAOC")')
    O.WriteLine("Hotkey('" ShortcutRef[2] "', Action)")
    O.WriteLine("Action(*) {")
    O.WriteLine(ShortcutActionE.Value)
    O.WriteLine('}')
    O.WriteLine('ProcessWaitClose(A_Args[1])')
    O.Write('ExitApp')
    O.Close()
    ShortcutList.Modify(ShortcutRef[1], 'Check', ShortcutRef[2], ShortcutCommentE.Value, ShortcutFile)
    IniWrite(1, Config, 'Hotkey', ShortcutRef[2])
    Run(AppDir[3] '\' ShortcutFile ' ' ProcessExist())
    ShortcutsGE.Hide()
}
ShortcutRemove := ShortcutsG.AddButton('xp+400 yp w100', 'Remove')
ShortcutRemove.OnEvent('Click', (*) => RemoveShortcut())
RemoveShortcut() {
    If !R := ShortcutList.GetNext() {
        Return
    }
    If R <= 2 {
        Msgbox('You can only un-check this shortcut!', 'Unable to delete!', 0x30)
        Return
    }
    Choice := MsgBox('Are you sure want to remove this hotkey?', 'Remove?', 0x4 + 0x20)
    If Choice != 'Yes' {
        Return
    }
    ShortcutFile := ShortcutList.GetText(R, 3)
    If FileExist(AppDir[3] '\' ShortcutFile)
        FileDelete(AppDir[3] '\' ShortcutFile)
    ShortcutName := ShortcutList.GetText(R)
    IniDelete(Config, 'Hotkey', ShortcutName)
    ShortcutList.Delete(R)
}
Shortcuts.OnEvent('Click', (*) => ShowShortcut())
ShowShortcut() {
    ShortcutsG.Show()
}
LoadShortcuts()
LoadShortcuts() {
    Loop Files, AppDir[3] '\*.ahk' {
        Info := ShortcutGetInfo(A_LoopFileFullPath)
        Checked := 0
        If !Info['ShortcutOK']
            Continue
        Item := ShortcutList.Add(, Info['ShortcutName'], Info['ShortcutComment'], A_LoopFileName)
        Checked := IniRead(Config, 'Hotkey', Info['ShortcutName'], 0)
        If Checked {
            Run(A_LoopFileFullPath ' ' ProcessExist())
            ShortcutList.Modify(Item, 'Check')
        }
    }
}
ShortcutList.OnEvent('ItemCheck', RunShortcut)
RunShortcut(Ctrl, Item, Checked) {
    ShortcutName := ShortcutList.GetText(Item)
    IniWrite(Checked, Config, 'Hotkey', ShortcutName)
}
ShortcutGetInfo(FileN) {
    ShortcutAction := ''
    ShortcutComment := ''
    ShortcutName := ''
    ShortcutTable := Map('ShortcutOK', False)
    Loop Read, FileN {
        If InStr(A_LoopReadLine, ';') {
            ShortcutComment := Trim(A_LoopReadLine, ';')
        }
        If InStr(A_LoopReadLine, 'Hotkey(') {
            ShortcutName := SubStr(A_LoopReadLine
                                 , L := InStr(A_LoopReadLine, "'") + 1
                                 , InStr(A_LoopReadLine, "',") - L)
        }
        If (ShortcutName = '') || InStr(A_LoopReadLine, 'Action')
            Continue
        If (A_LoopReadLine != '}')
            ShortcutAction .= (ShortcutAction = '') ? A_LoopReadLine : '`n' A_LoopReadLine
        Else {
            ShortcutTable := Map('ShortcutName', ShortcutName
                               , 'ShortcutComment', ShortcutComment
                               , 'ShortcutAction', ShortcutAction
                               , 'ShortcutOK', (ShortcutName != '') && (ShortcutComment != '') && (ShortcutAction != ''))
        }
    }
    Return ShortcutTable
}

; # VPN
Features['VPN'] := []
_ATools_ := Manager.AddText('xm yp+50 w600 h26 cBlue', 'VPN:')
Features['VPN'].Push(_ATools_)
_ATools_.SetFont('Bold s14')

VPN := Manager.AddButton('xm+10 w56 h56', 'VPN')
VPN.SetFont('Bold s20')
Features['VPN'].Push(VPN)
VPN.OnEvent('Click', (*) => OpenHAI())
OpenHAI() {
    If !FileExist(VPNPath) {
        Choice := MsgBox('Hide All IP does not seem to be installed!`nInstall now?', 'VPN', 0x30 + 0x4 + 0x100)
        If Choice = 'Yes' {
            RunWait('DB\000\hideallipsetup.exe', 'DB\000')
        }
    }
    If ProcessExist(VPNExe) {
        MsgBox('The VPN is already running', 'VPN', 0x40)
        Return
    }
    Run(VPNPath, VPNDir)
}
ClearVPNReg := Manager.AddButton('yp w130', 'Clear Registry')
Features['VPN'].Push(ClearVPNReg)
ClearVPNReg.OnEvent('Click', (*) => ClearRegHAI())
ClearRegHAI() {
    Try {
        Log := ""
        Loop Parse, "HKCU|HKLM", '|' {
            HK := A_LoopField
            Loop Parse, "Software\HideAllIP|Software\Wow6432Node\HideAllIP", '|' {
                Key := A_LoopField
                Loop Reg, HK "\" Key {
                    RegKey := A_LoopRegkey
                    Log .= Log ? "`n" RegKey : RegKey
                    RegDeleteKey(A_LoopRegkey)
                }
            }
        }
    } Catch As E {
        MsgBox('An error occured while trying to clear the registery.', 'Error', 0x10)
        Return
    }
    MsgBox((Log != '') ? "The following key(s) is(are) cleared:`n" Log : "Clear!", 'Clear!', 0x40)
}
VPNCompat := Manager.AddDropDownList('w130')
Features['VPN'].Push(VPNCompat)
For Each, Compat in Compatibilities {
    VPNCompat.Add([Compat[1]])
}
VPNCompat.Choose(1)
VPNCompat.OnEvent('Change', (*) => SetVPNCompat())
SetVPNCompat() {
    If VPNCompat.Value = 1 {
        Try {
            RegDelete(Layers, VPNPath)
        }
        Return
    }
    RegVal := Compatibilities[VPNCompat.Value][2] ' RUNASADMIN'
    RegWrite(RegVal, 'REG_SZ', Layers, VPNPath)
}

; # Game Records
Features['Game Records'] := []
_ATools_ := Manager.AddText('xm yp+50 w600 h26 cBlue', 'GAME RECORDS:')
Features['Game Records'].Push(_ATools_)
_ATools_.SetFont('Bold s14')

RecordFix := Manager.AddButton('xm+10', '(.mgx/.mgl) Records biegleux Fixes')
Features['Game Records'].Push(RecordFix)
RecordFix.OnEvent('Click', (*) => FixRecords())
RecordFixG := Gui(, 'Fix Records')
RecordFixG.OnEvent('Close', (*) => CancelRecordFix())
RecordFixText := RecordFixG.AddText('w300 Center cRed')
RecordFixProgress := RecordFixG.AddProgress('wp -Smooth')
RecordFixCancel := RecordFixG.AddButton('wp', 'Cancel')
RecordFixCancel.OnEvent('Click', (*) => CancelRecordFix()) 
CancelRecordFix() {
   Global Cancel := True
}
FixRecords() {
    Records := FileSelect('M', ChosenFolder.Value '\SaveGame', 'Select a record(s)', 'AoE II record games (*.mgl; *.mgx)')
    If !Records.Length
        Return
    Global Cancel := False
    RecordFixG.Show()
    RecordFixText.Text := 0 ' / ' Records.Length
    RecordFixProgress.Value := 0
    RecordFixProgress.Opt('Range1-' Records.Length)
    For Each, Record in Records {
        If Cancel {
            Break
        }
        If InStr(Record, 'aoeii_aio_fix') {
            RRecord := StrReplace(Record, '_aoeii_aio_fix')
            FileMove(Record, RRecord)
            Record := RRecord
        }
        RecordFixText.Text := Each ' / ' Records.Length
        RecordFixProgress.Value := Each
        SplitPath(Record, &FullName, &Dir, &Ext, &Name)
        If IniRead(MConfig, 'MgxFix', FullName, 0)
            || IniRead(RConfig, 'RevealFix', FullName, 0) {
            Continue
        }
        If (Ext ~= 'i)MGX|MGL') {
            RunWait('DB\000\MgxFix.exe -f "' Record '"',, 'Hide')
            IniWrite(1, MConfig, 'MgxFix', FullName)
            If Ext = 'MGL' {
                RunWait('DB\000\MgxFix.exe -f "' Record '"',, 'Hide')
                IniWrite(1, RConfig, 'RevealFix', FullName)
            }
        }
    }
    RecordFixG.Hide()
    RecordsCheck__________()
}
RecordsCheck__________() {
    Records1 := 0
    RecordCount.Text := '0 Records Found'
    Records2 := 0
    RecordMgxFixed.Text := '0 Records Mgx only Processed ✓'
    Records3 := 0
    RecordRevealFixed.Text := '0 Records Mgx & Reveal Processed ✓'
    Records4 := 0
    RecordUnknown.Text := '0 Records Not Processed X'
    Loop Files, ChosenFolder.Value '\SaveGame\*.mg*' {
        Records1 := CountRecords__________(Records1)
        Records2 := CountMgxFixedRecords__(Records2)
        Records3 := CountRevealFixedRecords(Records3)
        Records4 := CountUnknownRecords___(Records4)
    }
    RecordMgxFixed.Text := Records2 ' Records Mgx only Processed ✓'
    RecordRevealFixed.Text := Records3 ' Records Mgx && Reveal Processed ✓'
    RecordUnknown.Text := Records4 ' Records Not Processed X'
}
RecordCount := Manager.AddText('xp yp+30 w400 BackgroundTrans', '0 Records Found')
RecordCount.SetFont('Bold')
Features['Game Records'].Push(RecordCount)
CountRecords__________(Records) {
    If (A_LoopFileExt ~= 'i)MGX|MGL') {
        Records += 1
        RecordCount.Text := Records ' Records Found'
    }
    Return Records
}
RecordMgxFixed := Manager.AddText('xp yp+20 w400 BackgroundTrans cBlue', '0 Records Mgx only Processed ✓')
RecordMgxFixed.SetFont('Bold')
Features['Game Records'].Push(RecordMgxFixed)
CountMgxFixedRecords__(Records) {
    If (A_LoopFileExt ~= 'i)MGX|MGL') 
        && IniRead(MConfig, 'MgxFix', A_LoopFileName, 0)
        && !IniRead(RConfig, 'RevealFix', A_LoopFileName, 0) {
        Records += 1
    }
    Return Records
}
RecordRevealFixed := Manager.AddText('xp yp+20 w400 BackgroundTrans cGreen', '0 Records Mgx && Reveal Processed ✓')
RecordRevealFixed.SetFont('Bold')
Features['Game Records'].Push(RecordRevealFixed)
CountRevealFixedRecords(Records) {
    If (A_LoopFileExt ~= 'i)MGX|MGL')
        && IniRead(MConfig, 'MgxFix', A_LoopFileName, 0)
        && IniRead(RConfig, 'RevealFix', A_LoopFileName, 0) {
        Records += 1
    }
    Return Records
}
RecordUnknown := Manager.AddText('xp yp+20 w400 BackgroundTrans cRed', '0 Records Not Processed X')
RecordUnknown.SetFont('Bold')
Features['Game Records'].Push(RecordUnknown)
CountUnknownRecords___(Records) {
    If (A_LoopFileExt ~= 'i)MGX|MGL') 
        && !IniRead(MConfig, 'MgxFix', A_LoopFileName, 0)
        && !IniRead(RConfig, 'RevealFix', A_LoopFileName, 0) {
        Records += 1
    }
    Return Records
}

RecordDetailsG := Gui(, 'Records Statistics')
RecordDetailsG.BackColor := 'White'
RecordDetailsG.SetFont('Bold s10', 'Calibri')
RecordsList := RecordDetailsG.AddListView('w300 h283 -E0x200 -Multi', ['Record File Names:'])
RecordsList.OnEvent('ItemFocus', ViewRecordDetails)
ColorId := Map(0, 0x358BE0, 1, 0xFF0000, 2, 0x00FF00, 3, 0xFFFF00, 4, 0x00CBCB, 5, 0xFF00FF, 6, 0xFFFFFF, 7, 0xD86D00)
ReadTime(Time) {
    Seconds := Time // 1000
    Minutes := Seconds // 60
    Hours := Minutes // 60
    Return (Hours ? Hours ' : ' : '') (Minutes ? Minutes ' : ' : '') (Seconds ? Seconds : '')
}
SaveStatistic(FileName) {
    Hash := HashFile(FileName)
    If !DirExist(AppDir[4] '\' Hash) {
        DirCreate(AppDir[4] '\' Hash)
    }
    FileCopy('DB\000\Minimap.bmp', AppDir[4] '\' Hash '\Minimap.bmp')
    FileCopy('DB\000\Record.ini', AppDir[4] '\' Hash '\Record.ini')
}
ViewRecordDetails(GuiCtrlObj, Item) {
    RecordsList.Enabled := False
    Loop RecordsTeamsDetails.GetCount()
        RecordsTeamsDetails.Modify(A_Index,,'','','','','','','','','','','')
    RecordsTeamsDetails.ModifyCol(1,, 'Team')
    RecordsTeamsDetails.ModifyCol(2,, 'Team')
    RecordDetailsMinimap.Value := 'DB\000\Minimapd.png'
    Hash := HashFile(ChosenFolder.Value '\SaveGame\' RecordsList.GetText(Item))
    If !DirExist(AppDir[4] '\' Hash) {
        RunWait('DB\000\recanalyst.exe "' ChosenFolder.Value '\SaveGame\' RecordsList.GetText(Item) '"', 'DB\000', 'Hide')
        SaveStatistic(ChosenFolder.Value '\SaveGame\' RecordsList.GetText(Item))
    }
    RecordDetailsMinimap.Value := AppDir[4] '\' Hash '\Minimap.bmp'
    RecordsTeamsDetails.Modify(1,,,,, IniRead(AppDir[4] '\' Hash '\Record.ini', 'Game Settings' , 'POV Name', '')
                                    , IniRead(AppDir[4] '\' Hash '\Record.ini', 'Game Settings' , 'Game Type', '')
                                    , IniRead(AppDir[4] '\' Hash '\Record.ini', 'Game Settings' , 'Map Style', '')
                                    , IniRead(AppDir[4] '\' Hash '\Record.ini', 'Game Settings' , 'Map', '')
                                    , IniRead(AppDir[4] '\' Hash '\Record.ini', 'Game Settings' , 'Versus', '')
                                    , IniRead(AppDir[4] '\' Hash '\Record.ini', 'Game Settings' , 'Difficulty Level', '')
                                    , IniRead(AppDir[4] '\' Hash '\Record.ini', 'Game Settings' , 'Game Speed', '')
                                    , IniRead(AppDir[4] '\' Hash '\Record.ini', 'Game Settings' , 'Reveal Map', '')
                                    , IniRead(AppDir[4] '\' Hash '\Record.ini', 'Game Settings' , 'Map Size', '')
                                    , IniRead(AppDir[4] '\' Hash '\Record.ini', 'Game Settings' , 'Version', '')
                                    , IniRead(AppDir[4] '\' Hash '\Record.ini', 'Game Settings' , 'Pop Limit', '')
                                    , ReadTime(IniRead(AppDir[4] '\' Hash '\Record.ini', 'Game Settings' , 'Play Time', '')))
    Players := Map()
    Teams := Map()
    Loop {
        Players['Number_' A_Index] := Map()
        Players['Number_' A_Index]['Name'] := IniRead(AppDir[4] '\' Hash '\Record.ini', 'Number_' A_Index, 'Name', '')
        Players['Number_' A_Index]['Team'] := IniRead(AppDir[4] '\' Hash '\Record.ini', 'Number_' A_Index, 'Team', '')
        Players['Number_' A_Index]['Civilization'] := IniRead(AppDir[4] '\' Hash '\Record.ini', 'Number_' A_Index, 'Civilization', '')
        Id := IniRead(AppDir[4] '\' Hash '\Record.ini', 'Number_' A_Index, 'Color Id', 0) + 0
        Players['Number_' A_Index]['Color Id'] := ColorId[Id <= 7 ? Id : 7]
        Players['Number_' A_Index]['Feudal Time'] := ReadTime(IniRead(AppDir[4] '\' Hash '\Record.ini', 'Number_' A_Index, 'Feudal Time', 0))
        Players['Number_' A_Index]['Castle Time'] := ReadTime(IniRead(AppDir[4] '\' Hash '\Record.ini', 'Number_' A_Index, 'Castle Time', 0))
        Players['Number_' A_Index]['Imperial Time'] := ReadTime(IniRead(AppDir[4] '\' Hash '\Record.ini', 'Number_' A_Index, 'Imperial Time', 0))
        Players['Number_' A_Index]['Resign Time'] := ReadTime(IniRead(AppDir[4] '\' Hash '\Record.ini', 'Number_' A_Index, 'Resign Time', 0))
        Players['Number_' A_Index]['Initial Wood'] := IniRead(AppDir[4] '\' Hash '\Record.ini', 'Number_' A_Index, 'Initial Wood', 0)
        Players['Number_' A_Index]['Initial Food'] := IniRead(AppDir[4] '\' Hash '\Record.ini', 'Number_' A_Index, 'Initial Food', 0)
        Players['Number_' A_Index]['Initial Gold'] := IniRead(AppDir[4] '\' Hash '\Record.ini', 'Number_' A_Index, 'Initial Gold', 0)
        Players['Number_' A_Index]['Initial Stone'] := IniRead(AppDir[4] '\' Hash '\Record.ini', 'Number_' A_Index, 'Initial Stone', 0)
        Players['Number_' A_Index]['Initial Stone'] := IniRead(AppDir[4] '\' Hash '\Record.ini', 'Number_' A_Index, 'Initial Stone', 0)
        If !Teams.Has(Players['Number_' A_Index]['Team']) {
            Teams[Players['Number_' A_Index]['Team']] := []
        }
        Teams[Players['Number_' A_Index]['Team']].Push('Number_' A_Index)
    } Until IniRead(AppDir[4] '\' Hash '\Record.ini', 'Number_' A_Index + 1, 'Name', '') = ''
    RecordsTeamsDetails.Redraw()
    IndexCol := 0
    For Each, Team in Teams {
        If Each = 0 {
            For Every, Member in Team {
                RecordsTeamsDetails.Modify(Every,,,, Players[Member]['Name'])
                RecordsTeamsDetailsColors.Cell(Every, 3, 0x000000, Players[Member]['Color Id'])
            }
            Continue
        }
        ++IndexCol
        RecordsTeamsDetails.ModifyCol(IndexCol,, 'Team [' Each ']')
        For Every, Member in Team {
            IndexCol = 1 ? RecordsTeamsDetails.Modify(Every,, Players[Member]['Name']) : RecordsTeamsDetails.Modify(Every,,, Players[Member]['Name'])
            RecordsTeamsDetailsColors.Cell(Every, IndexCol, 0x000000, Players[Member]['Color Id'])
        }
    }
    Loop RecordsTeamsDetails.GetCount('Col') {
        RecordsTeamsDetails.ModifyCol(A_Index, 'AutoHdr')
    }
    RecordsList.Enabled := True
}
RecordDetailsG.AddText('yp w500 h250 h26', 'Minimap:').SetFont('s12')
RecordDetailsMinimap := RecordDetailsG.AddPicture('w500 h250')
RecordsTeamsDetails := RecordDetailsG.AddListView('xm w820 r5 -E0x200 -Multi', ['Team', 'Team', 'Player [No Team]', 'POV Name', 'Game Type', 'Map Style', 'Map', 'Versus', 'Difficulty Level', 'Game Speed', 'Reveal Map', 'Map Size', 'Version', 'Pop Limit', 'Play Time'])
RecordsTeamsDetailsColors := LV_Colors(RecordsTeamsDetails)
Loop 4 {
    RecordsTeamsDetailsColors.Row(RecordsTeamsDetails.Add(), 0x000000, 0xFFFFFF)
}
RecordDetails := Manager.AddButton(, 'View Records Statistics')
RecordDetails.OnEvent('Click', (*) => LoadRecordList())
LoadRecordList() {
    RecordDetailsG.Show()
    RecordsList.Delete()
    Loop Files, ChosenFolder.Value '\SaveGame\*.mg*' {
        If (A_LoopFileExt ~= 'i)MGX|MGL') {
            RecordsList.Add(, A_LoopFileName)
        }
    }
}
;Manager.AddText('yp+40 cBlue w200 BackgroundTrans Center', '3 - Repair Game Files').SetFont('Bold')
;RepairGame := Manager.AddButton('wp', 'Repair')
;RepairGame.OnEvent('Click', (*) => MsgBox('Not Yet Implemented!', 'Hoy!', 0x40))
;
;Manager.AddText('yp+40 cBlue w200 BackgroundTrans Center', '5 - Scenario Files Select').SetFont('Bold')
;FixMgz := Manager.AddButton('wp', 'Select')
;FixMgz.OnEvent('Click', (*) => MsgBox('Not Yet Implemented!', 'Hoy!', 0x40))

FileMenu := Menu()
FileMenu.Add("Age of Empires II: The Age of Kings`tAlt+E", (*) => Run(ChosenFolder.Value '\empires2.exe', ChosenFolder.Value))
FileMenu.Add("Age of Empires II: The Conquerors`tAlt+C", (*) => Run(ChosenFolder.Value '\age2_x1\age2_x1.exe', ChosenFolder.Value '\age2_x1'))
FileMenu.Add("Age of Empires II: Forgotten Empires`tAlt+F", (*) => Run(ChosenFolder.Value '\age2_x1\age2_x2.exe', ChosenFolder.Value '\age2_x1'))
FileMenu.Add("Exit`tEsc", (*) => ExitApp())
FileMenu.SetIcon('Age of Empires II: The Age of Kings`tAlt+E', 'DB\000\aok.png')
FileMenu.SetIcon('Age of Empires II: The Conquerors`tAlt+C', 'DB\000\aoc.png')
FileMenu.SetIcon('Age of Empires II: Forgotten Empires`tAlt+F', 'DB\000\fe.png')
GoToMenu := Menu()
GoToMenu.Add('Go to GAME LOCATION', (*) => SB1.ScrollMsg(100, 0, 0x115, Manager.Hwnd))
GoToMenu.Add('Go to OPTION', (*) => SB1.ScrollMsg(101, 0, 0x115, Manager.Hwnd))
GoToMenu.Add('Go to GAME VERSION', (*) => SB1.ScrollMsg(102, 0, 0x115, Manager.Hwnd))
GoToMenu.Add('Go to GAME LANGUAGES', (*) => SB1.ScrollMsg(103, 0, 0x115, Manager.Hwnd))
GoToMenu.Add('Go to GAME VISUAL MODS', (*) => SB1.ScrollMsg(104, 0, 0x115, Manager.Hwnd))
GoToMenu.Add('Go to GAME DATA MODS', (*) => SB1.ScrollMsg(105, 0, 0x115, Manager.Hwnd))
;GoToMenu.Add('Go to GAME OTHER TOOLS', (*) => SB1.ScrollMsg(7, 0, 0x115, Manager.Hwnd))
AboutMenu := Menu()

AboutMenu.Add('&About this app`tAlt+B', (*) => Msgbox('Nothing but a small AutoHotkey app'
                                                    . '`n`nIn which I tried to include all the helphul things I found on the internet'
                                                    . '`nTo save you the time searching on already done features'
                                                    . '`n`nThe script of this app is already revealed'
                                                    . '`nSo if you feel like it is not safe, you can check it out'
                                                    . "`nI didn't create those stuffs, I just assembled them into one easy app"
                                                    . '`n`nAfter many years GameRanger still here for AoEII'
                                                    . '`nBut the support for this game is not well maintained speciallt when it comes to the mods and the patchs'
                                                    . '`n`nResources:'
                                                    . '`nhttps://www.voobly.com'
                                                    . '`nhttps://aok.heavengames.com'
                                                    . '`n`nMany thanks to my GR friends:'
                                                    . '`nKatsuie - For getting all the useful patches and mods for us here on GameRanger'
                                                    . '`npreo - For helping testing my app'
                                                    . '`nMr.Plow - For helping testing my app'
                                                    . '`n[ Ebul-Feth ] | Cagri the Turk - For helping testing my app'
                                                    . '`n`nAnd to all others who did help make this game better', 'About it', 0x40))
Menus := MenuBar()
Menus.Add("Games", FileMenu)
Menus.Add("Go to", GoToMenu)
Menus.Add("About", AboutMenu)
Manager.MenuBar := Menus
Manager.Show('w520 h300')

RightMenu := Gui('-Caption')
RightMenu.MarginX := RightMenu.MarginY := 10
RightMenu.BackColor := 'White'
RunAOK := RightMenu.AddButton('w36 H36')
Features['The Game'].Push(RunAOK)
CreateImageButton(RunAOK, 0, [['DB\000\aok_normal.png'], ['DB\000\aok_hover.png'], ['DB\000\aok_click.png'], ['DB\000\aok_disable.png']]*)
RunAOK.OnEvent('Click', (*) => Run(ChosenFolder.Value '\empires2.exe', ChosenFolder.Value))
RunAOC := RightMenu.AddButton('wp hp')
Features['The Game'].Push(RunAOC)
CreateImageButton(RunAOC, 0, [['DB\000\aoc_normal.png'], ['DB\000\aoc_hover.png'], ['DB\000\aoc_click.png'], ['DB\000\aoc_disable.png']]*)
RunAOC.OnEvent('Click', (*) => Run(ChosenFolder.Value '\age2_x1\age2_x1.exe', ChosenFolder.Value '\age2_x1'))
RunFOE := RightMenu.AddButton('wp hp')
Features['The Game'].Push(RunFOE)
CreateImageButton(RunFOE, 0, [['DB\000\fe_normal.png'], ['DB\000\fe_hover.png'], ['DB\000\fe_click.png'], ['DB\000\fe_disable.png']]*)
RunFOE.OnEvent('Click', (*) => Run(ChosenFolder.Value '\age2_x1\age2_x2.exe', ChosenFolder.Value '\age2_x1'))
;SetTimer(CheckEach1Second, 1)
;CheckEach1Second() {
;    RightMenu.GetPos(&X1, &Y1)
;    Manager.GetPos(&X2, &Y2, &W2, &H2)
;    CX1 := X2 + W2
;    CY1 := Y2
;    If X1 != CX1 || Y1 != CY1 {
;        RightMenu.Move(CX1, CY1)
;    }
;}
;RightMenu.Show()

ChargeEnableFixes_____()
ChargeEnableFixes_____() {
    ;Loop Files, 'DB\001\*', 'D' {
    ;    Patch.Add([A_LoopFileName])
    ;}
    ;Patch.Choose('Do Not Enable Fixes')
    ;If DirExist('DB\001\' Fix := IniRead(Config, 'Game', 'Fix', 'Do Not Enable Fixes')) {
    ;    Patch.Choose(Fix)
    ;}
}
ChargeSettings________()
CheckForUpdates_______()
CheckForUpdates_______() {
    If A_IsCompiled {
        Return
    }
    ;SB.SetText('Update check is disabled!', 2)
    UpdateChk := IniRead(Config, 'Game', 'UpdateChk', '')
    If UpdateChk = '0' {
        Return
    }
    Try {
        ;SB.SetText('Checking for updates...', 3)
        GetTextFromLink(Link) {
            whr := ComObject("WinHttp.WinHttpRequest.5.1")
            whr.Open("GET", Link, true)
            whr.Send()
            whr.WaitForResponse()
            Return whr.ResponseText
        }
        Hashsums := GetTextFromLink(Server '/' User '/' Repo '/main/Hashsums.ini')
        HashsumsMap := Map()
        For Each, Line in StrSplit(Hashsums, '`n') {
            KeyValue := StrSplit(Line, '=')
            If KeyValue.Length = 2
                HashsumsMap[KeyValue[1]] := KeyValue[2]
        }
        FoundUpdates := []
        If HashFile(A_ScriptName) != HashsumsMap[A_ScriptName] {
            FoundUpdates.Push(A_ScriptName)
        }
        Loop Files, 'DB\*.7z.*', 'R' {
            If HashsumsMap.Has(A_LoopFileDir '\' A_LoopFileName)
                && (HashFile(A_LoopFileDir '\' A_LoopFileName) != HashsumsMap[A_LoopFileDir '\' A_LoopFileName]) {
                    FoundUpdates.Push(A_LoopFileDir '\' A_LoopFileName)
            }
        }
        If FoundUpdates.Length {
            UpdatesList := '`n=======================`n`n'
            For Each, UpdateFile in FoundUpdates {
                UpdatesList .= Each ' - ' UpdateFile '`n'
            }
            UpdatesList .= '`n=======================`n'
            ;SB.SetText('Update found!', 2)
            Choice := MsgBox('The following needs to be updated:`n' UpdatesList '`nUpdate now?', 'New Update!', 0x4 + 0x40 + 0x100)
            If Choice = 'Yes' {
                DoneSteps.Value := 0
                DoneSteps.Opt('Range1-' FoundUpdates.Length + 1)
                DoneStepsText.Text := ''
                Prepare.Show()
                Manager.Hide()
                PrepareTheUnpacker()
                For Each, UpdateFile in FoundUpdates {
                    DownloadLink := Server '/' User '/' Repo '/main/' StrReplace(StrReplace(UpdateFile, ' ', '%20'), '\', '/')
                    Download(DownloadLink, UpdateFile)
                    If !PackageIsValid(PackagePath)
                        Reload
                    Buff := FileRead(UpdateFile, 'RAW')
                    Str := StrGet(Buff, 2, 'cp0')
                    If (Str = '7z') {
                        Name := StrSplit(UpdateFile, '.')
                        DirDelete(Name[1], 1)
                        RunWait('DB\7za.exe x ' Name[1] '.7z.001 -o' Name[1] ' -aoa', , 'Hide')
                    }
                    DoneSteps.Value += 1
                    DoneStepsText.Text := UpdateFile
                }
                Reload
            }
            ;SB.SetText('New update is waiting!', 2)
        } Else {
            ;SB.SetText('Up to date!', 2)
        }
    } Catch As Err {
        ;SB.SetText('Failed to check for updates!', 2)
    }
}

ChargeSettings________(Browse := False) {
    SoundPlay('DB\000\30 wololo.mp3')
    SectionInteract(Features['Versions'], False)
    SectionInteract(Features['Compatibilities'], False)
    SectionInteract(Features['Language'], False)
    SectionInteract(Features['Visual Modes'], False)
    SectionInteract(Features['Data Mods'], False)
    SectionInteract(Features['Game Hotkeys'], False)
    SectionInteract(Features['VPN'], False)
    SectionInteract(Features['Game Records'], False)
    ChosenFolder.Value := IniRead(Config, 'Game', 'Path', '')
    ValidGameLocation(Location) {
        Return FileExist(Location '\empires2.exe')
            && FileExist(Location '\language.dll')
            && FileExist(Location '\Data\graphics.drs')
            && FileExist(Location '\Data\interfac.drs')
            && FileExist(Location '\Data\terrain.drs')
    }
    If !ValidGameLocation(ChosenFolder.Value) {
        SplitPath(ChosenFolder.Value, , &Expected)
        If FileExist(Expected '\empires2.exe') {
            ChosenFolder.Value := Expected
            IniWrite(ChosenFolder.Value, Config, 'Game', 'Path')
        }
    }
    If !ValidGameLocation(ChosenFolder.Value) {
        Loop Files, ChosenFolder.Value '\*', 'D' {
            If FileExist(ChosenFolder.Value '\' A_LoopFileName '\empires2.exe') {
                ChosenFolder.Value := ChosenFolder.Value '\' A_LoopFileName
                IniWrite(ChosenFolder.Value, Config, 'Game', 'Path')
                Break
            }
        }
    }
    If !Browse && !ValidGameLocation(ChosenFolder.Value) {
        SelectAFolder()
    }
    If !ValidGameLocation(ChosenFolder.Value) {
        SplitPath(ChosenFolder.Value, , , , , &Drive)
        If Drive != '' {
            Choice := MsgBox('Game location not found!`n`nWant to launch a search in the current drive?', 'Game', 0x4 + 0x20)
            If Choice = 'Yes' {
                Loop Files, Drive '\empires2.exe', 'R' {
                    If !ValidGameLocation(A_LoopFileDir)
                        Continue
                    ChosenFolder.Value := A_LoopFileDir
                    IniWrite(ChosenFolder.Value, Config, 'Game', 'Path')
                    Break
                }
            }
        }
    }
    If !ValidGameLocation(ChosenFolder.Value) {
        SelectTheGameFromGR()
    }
    If !ValidGameLocation(ChosenFolder.Value) {
        Return
    }
    SectionInteract(Features['Versions'])
    SectionInteract(Features['Compatibilities'])
    SectionInteract(Features['Language'])
    SectionInteract(Features['Visual Modes'])
    SectionInteract(Features['Data Mods'])
    SectionInteract(Features['Game Hotkeys'])
    SectionInteract(Features['VPN'])
    SectionInteract(Features['Game Records'])
    ChargeVersions________()
    UpdateVersionRadio()
    ChargeCompatibilities_()
    ChargeCompatibilities_() {
        AoKCom.Choose(1)
        AoKRun.Value := False
        AoCCom.Choose(1)
        AoCRun.Value := False
        FOECom.Choose(1)
        FOERun.Value := False
        AoKReg := Trim(RegRead(Layers, ChosenFolder.Value '\empires2.exe', ''), ' ')
        If AoKReg {
            AoKReg := StrSplit(AoKReg, ' ')
            For Each, RegVal in AoKReg {
                If (RegVal = 'RUNASADMIN')
                    AoKRun.Value := True
                Else {
                    For Each, Compat in Compatibilities {
                        If (Compat[2] = RegVal) {
                            AoKCom.Choose(Compat[1])
                        }
                    }
                }
            }
        }
        AoCReg := Trim(RegRead(Layers, ChosenFolder.Value '\age2_x1\age2_x1.exe', ''), ' ')
        If AoCReg {
            AoCReg := StrSplit(AoCReg, ' ')
            For Each, RegVal in AoCReg {
                If (RegVal = 'RUNASADMIN')
                    AoCRun.Value := True
                Else {
                    For Each, Compat in Compatibilities {
                        If (Compat[2] = RegVal)
                            AoCCom.Choose(Compat[1])
                    }
                }
            }
        }
        FEReg := Trim(RegRead(Layers, ChosenFolder.Value '\age2_x1\age2_x2.exe', ''), ' ')
        If FEReg {
            FEReg := StrSplit(FEReg, ' ')
            For Each, RegVal in FEReg {
                If (RegVal = 'RUNASADMIN')
                    FOERun.Value := True
                Else {
                    For Each, Compat in Compatibilities {
                        If (Compat[2] = RegVal)
                            FOECom.Choose(Compat[1])
                    }
                }
            }
        }
    }
    ChargeLanguage________()
    BackupDefaultLanguage_()
    ChargeOtherTools______()
    ChargeOtherTools______() {
        ;---
        RegVal := RegRead(Layers, VPNPath, '')
        If RegVal = '' {
            VPNCompat.Value := 1
            Return
        }
        For Each, Compat in Compatibilities {
            If (RegVal = Compat[2] ' RUNASADMIN') {
                VPNCompat.Choose(Compat[1])
                Break
            }
        }
    }
    FixCommonIssues_______() {
        If FileExist(ChosenFolder.Value '\age2_x1.exe') {
            If !DirExist(ChosenFolder.Value '\age2_x1') {
                DirCreate(ChosenFolder.Value '\age2_x1')
            }
            FileMove(ChosenFolder.Value '\age2_x1.exe', ChosenFolder.Value '\age2_x1', 1)
        }
        If FileExist(ChosenFolder.Value '\windmode.ini') {
            FileDelete(ChosenFolder.Value '\windmode.ini')
        }
        If FileExist(ChosenFolder.Value '\age2_x1\windmode.ini') {
            FileDelete(ChosenFolder.Value '\age2_x1\windmode.ini')
        }
        CreateGameShortcuts()
    }
    FixCommonIssues_______()
    RecordsCheck__________()
}
SectionInteract(Items, Default := True) {
    Status := Default ? True : False
    For Each, Item in Items {
        Item.Enabled := Status
        Item.Redraw()
    }
}
; https://autohotkey.com/board/topic/66139-ahk-l-calculating-md5sha-checksum-from-file/
HashFile(FilePath, HashType := 2) {
    Static PROV_RSA_AES := 24
    Static CRYPT_VERIFYCONTEXT := 0xF0000000
    Static BUFF_SIZE := 1024 * 1024 ; 1 MB
    Static HP_HASHVAL := 0x0002
    Static HP_HASHSIZE := 0x0004
    Switch HashType {
        Case 1: Hash_Alg := (CALG_MD2 := 32769)
        Case 2: Hash_Alg := (CALG_MD5 := 32771)
        Case 3: Hash_Alg := (CALG_SHA := 32772)
        Case 4: Hash_Alg := (CALG_SHA_256 := 32780)
        Case 5: Hash_Alg := (CALG_SHA_384 := 32781)
        Case 6: Hash_Alg := (CALG_SHA_512 := 32782)
        Default: throw ValueError('Invalid HashType', -1, HashType)
    }
    F := FileOpen(FilePath, "r")
    F.Pos := 0 ; Rewind in case of BOM.
    HCRYPTPROV() => {
        ptr: 0,
        __delete: this => this.ptr && DllCall("Advapi32\CryptReleaseContext", "Ptr", this, "UInt", 0)
    }
    If !DllCall("Advapi32\CryptAcquireContextW"
        , "Ptr*", hProv := HCRYPTPROV()
        , "Uint", 0
        , "Uint", 0
        , "Uint", PROV_RSA_AES
        , "UInt", CRYPT_VERIFYCONTEXT)
        Throw OSError()
    HCRYPTHASH() => {
        Ptr: 0,
        __Delete: This => This.Ptr && DllCall("Advapi32\CryptDestroyHash", "Ptr", This)
    }
    If !DllCall("Advapi32\CryptCreateHash"
        , "Ptr", hProv
        , "Uint", Hash_Alg
        , "Uint", 0
        , "Uint", 0
        , "Ptr*", hHash := HCRYPTHASH())
        Throw OSError()
    READ_BUF := Buffer(BUFF_SIZE, 0)
    While (cbCount := F.RawRead(READ_BUF, BUFF_SIZE)) {
        if !DllCall("Advapi32\CryptHashData"
            , "Ptr", hHash
            , "Ptr", READ_BUF
            , "Uint", cbCount
            , "Uint", 0)
            Throw OSError()
    }
    If !DllCall("Advapi32\CryptGetHashParam"
        , "Ptr", hHash
        , "Uint", HP_HASHSIZE
        , "Uint*", &HashLen := 0
        , "Uint*", &HashLenSize := 4
        , "UInt", 0)
        Throw OSError()
    bHash := Buffer(HashLen, 0)
    If !DllCall("Advapi32\CryptGetHashParam"
        , "Ptr", hHash
        , "Uint", HP_HASHVAL
        , "Ptr", bHash
        , "Uint*", &HashLen
        , "UInt", 0)
        Throw OSError()
    Loop HashLen
        HashVal .= Format('{:02x}', (NumGet(bHash, A_Index - 1, "UChar")) & 0xff)
    F.Close()
    Return HashVal
}

; ======================================================================================================================
; Name:              CreateImageButton()
; Function:          Create images and assign them to pushbuttons.
; Tested with:       AHK 2.0.11 (U32/U64)
; Tested on:         Win 10 (x64)
; Change history:    1.0.01/2024-01-01/just me   - Use Gui.Backcolor as default for the background if available
;                    1.0.00/2023-02-03/just me   - Initial stable release for AHK v2
; Credits:           THX tic for GDIP.AHK, tkoi for ILBUTTON.AHK
; ======================================================================================================================
; How to use:
;     1. Call UseGDIP() to initialize the Gdiplus.dll before the first call of this function.
;     2. Create a push button (e.g. "MyGui.AddButton("option", "caption").
;     3. If you want to want to use another color than the GUI's current Backcolor for the background of the images
;        - especially for rounded buttons - call CreateImageButton("SetDefGuiColor", NewColor) where NewColor is a RGB
;        integer value (0xRRGGBB) or a HTML color name ("Red"). You can also change the default text color by calling
;        CreateImageButton("SetDefTxtColor", NewColor).
;        To reset the colors to the AHK/system default pass "*DEF*" in NewColor, to reset the background to use
;        Gui.Backcolor pass "*GUI*".
;     4. To create an image button call CreateImageButton() passing two or more parameters:
;        GuiBtn      -  Gui.Button object.
;        Mode        -  The mode used to create the bitmaps:
;                       0  -  unicolored or bitmap
;                       1  -  vertical bicolored
;                       2  -  horizontal bicolored
;                       3  -  vertical gradient
;                       4  -  horizontal gradient
;                       5  -  vertical gradient using StartColor at both borders and TargetColor at the center
;                       6  -  horizontal gradient using StartColor at both borders and TargetColor at the center
;                       7  -  'raised' style
;                       8  -  forward diagonal gradient from the upper-left corner to the lower-right corner
;                       9  -  backward diagonal gradient from the upper-right corner to the lower-left corner
;                      -1  -  reset the button
;        Options*    -  variadic array containing up to 6 option arrays (see below).
;        ---------------------------------------------------------------------------------------------------------------
;        The index of each option object determines the corresponding button state on which the bitmap will be shown.
;        MSDN defines 6 states (http://msdn.microsoft.com/en-us/windows/bb775975):
;           PBS_NORMAL    = 1
;	         PBS_HOT       = 2
;	         PBS_PRESSED   = 3
;	         PBS_DISABLED  = 4
;	         PBS_DEFAULTED = 5
;	         PBS_STYLUSHOT = 6 <- used only on tablet computers (that's false for Windows Vista and 7, see below)
;        If you don't want the button to be 'animated' on themed GUIs, just pass one option object with index 1.
;        On Windows Vista and 7 themed bottons are 'animated' using the images of states 5 and 6 after clicked.
;        ---------------------------------------------------------------------------------------------------------------
;        Each option array may contain the following values:
;           Index Value
;           1     StartColor  mandatory for Option[1], higher indices will inherit the value of Option[1], if omitted:
;                             -  ARGB integer value (0xAARRGGBB) or HTML color name ("Red").
;                             -  Path of an image file or HBITMAP handle for mode 0.
;           2     TargetColor mandatory for Option[1] if Mode > 0. Higher indcices will inherit the color of Option[1],
;                             if omitted:
;                             -  ARGB integer value (0xAARRGGBB) or HTML color name ("Red").
;                             -  String "HICON" if StartColor contains a HICON handle.
;           3     TextColor   optional, if omitted, the default text color will be used for Option[1], higher indices
;                             will inherit the color of Option[1]:
;                             -  ARGB integer value (0xAARRGGBB) or HTML color name ("Red").
;                                Default: 0xFF000000 (black)
;           4     Rounded     optional:
;                             -  Radius of the rounded corners in pixel; the letters 'H' and 'W' may be specified
;                                also to use the half of the button's height or width respectively.
;                                Default: 0 - not rounded
;           5     BorderColor optional, ignored for modes 0 (bitmap) and 7, color of the border:
;                             -  RGB integer value (0xRRGGBB) or HTML color name ("Red").
;           6     BorderWidth optional, ignored for modes 0 (bitmap) and 7, width of the border in pixels:
;                             -  Default: 1
;        ---------------------------------------------------------------------------------------------------------------
;        If the the button has a caption it will be drawn upon the bitmaps.
;     5. Call GdiplusShutDown() to clean up the resources used by GDI+ after the last function call or
;        before the script terminates.
; ======================================================================================================================
; This software is provided 'as-is', without any express or implied warranty.
; In no event will the authors be held liable for any damages arising from the use of this software.
; ======================================================================================================================
; CreateImageButton()
; ======================================================================================================================
CreateImageButton(GuiBtn, Mode, Options*) {
    ; Default colors - COLOR_3DFACE is used by AHK as default Gui background color
    Static DefGuiColor := SetDefGuiColor("*GUI*"),
        DefTxtColor := SetDefTxtColor("*DEF*"),
        GammaCorr := False
    Static HTML := { BLACK: 0x000000, GRAY: 0x808080, SILVER: 0xC0C0C0, WHITE: 0xFFFFFF,
        MAROON: 0x800000, PURPLE: 0x800080, FUCHSIA: 0xFF00FF, RED: 0xFF0000,
        GREEN: 0x008000, OLIVE: 0x808000, YELLOW: 0xFFFF00, LIME: 0x00FF00,
        NAVY: 0x000080, TEAL: 0x008080, AQUA: 0x00FFFF, BLUE: 0x0000FF }
    Static MaxBitmaps := 6, MaxOptions := 6
    Static BitMaps := [], Buttons := Map()
    Static Bitmap := 0, Graphics := 0, Font := 0, StringFormat := 0, HIML := 0
    Static BtnCaption := "", BtnStyle := 0
    Static HWND := 0
    Bitmap := Graphics := Font := StringFormat := HIML := 0
    NumBitmaps := 0
    BtnCaption := ""
    BtnStyle := 0
    BtnW := 0
    BtnH := 0
    GuiColor := ""
    TxtColor := ""
    HWND := 0
    ; -------------------------------------------------------------------------------------------------------------------
    ; Check for 'special calls'
    If !IsObject(GuiBtn) {
        Switch GuiBtn {
            Case "SetDefGuiColor":
                DefGuiColor := SetDefGuiColor(Mode)
                Return True
            Case "SetDefTxtColor":
                DefTxtColor := SetDefTxtColor(Mode)
                Return True
            Case "SetGammaCorrection":
                GammaCorr := !!Mode
                Return True
        }
    }
    ; -------------------------------------------------------------------------------------------------------------------
    ; Check the control object
    If (Type(GuiBtn) != "Gui.Button")
        Return ErrorExit("Invalid parameter GuiBtn!")
    HWND := GuiBtn.Hwnd
    ; -------------------------------------------------------------------------------------------------------------------
    ; Check Mode
    If !IsInteger(Mode) || (Mode < -1) || (Mode > 9)
        Return ErrorExit("Invalid parameter Mode!")
    If (Mode = -1) { ; reset the button
        If Buttons.Has(HWND) {
            Btn := Buttons[HWND]
            BIL := Buffer(20 + A_PtrSize, 0)
            NumPut("Ptr", -1, BIL) ; BCCL_NOGLYPH
            SendMessage(0x1602, 0, BIL.Ptr, HWND) ; BCM_SETIMAGELIST
            IL_Destroy(Btn["HIML"])
            ControlSetStyle(Btn["Style"], HWND)
            Buttons.Delete(HWND)
            Return True
        }
        Return False
    }
    ; -------------------------------------------------------------------------------------------------------------------
    ; Check Options
    If !(Options Is Array) || !Options.Has(1) || (Options.Length > MaxOptions)
        Return ErrorExit("Invalid parameter Options!")
    ; -------------------------------------------------------------------------------------------------------------------
    HBITMAP := HFORMAT := PBITMAP := PBRUSH := PFONT := PGRAPHICS := PPATH := 0
    ; -------------------------------------------------------------------------------------------------------------------
    ; Get control's styles
    BtnStyle := ControlGetStyle(HWND)
    ; -------------------------------------------------------------------------------------------------------------------
    ; Get the button's font
    PFONT := 0
    If (HFONT := SendMessage(0x31, 0, 0, HWND)) { ; WM_GETFONT
        DC := DllCall("GetDC", "Ptr", HWND, "Ptr")
        DllCall("SelectObject", "Ptr", DC, "Ptr", HFONT)
        DllCall("Gdiplus.dll\GdipCreateFontFromDC", "Ptr", DC, "PtrP", &PFONT)
        DllCall("ReleaseDC", "Ptr", HWND, "Ptr", DC)
    }
    If !(Font := PFONT)
        Return ErrorExit("Couldn't get button's font!")
    ; -------------------------------------------------------------------------------------------------------------------
    ; Get the button's width and height
    ControlGetPos(, , &BtnW, &BtnH, HWND)
    ; -------------------------------------------------------------------------------------------------------------------
    ; Get the button's caption
    BtnCaption := GuiBtn.Text
    ; -------------------------------------------------------------------------------------------------------------------
    ; Create a GDI+ bitmap
    PBITMAP := 0
    DllCall("Gdiplus.dll\GdipCreateBitmapFromScan0",
        "Int", BtnW, "Int", BtnH, "Int", 0, "UInt", 0x26200A, "Ptr", 0, "PtrP", &PBITMAP)
    If !(Bitmap := PBITMAP)
        Return ErrorExit("Couldn't create the GDI+ bitmap!")
    ; Get the pointer to its graphics
    PGRAPHICS := 0
    DllCall("Gdiplus.dll\GdipGetImageGraphicsContext", "Ptr", PBITMAP, "PtrP", &PGRAPHICS)
    If !(Graphics := PGRAPHICS)
        Return ErrorExit("Couldn't get the the GDI+ bitmap's graphics!")
    ; Quality settings
    DllCall("Gdiplus.dll\GdipSetSmoothingMode", "Ptr", PGRAPHICS, "UInt", 4)
    DllCall("Gdiplus.dll\GdipSetInterpolationMode", "Ptr", PGRAPHICS, "Int", 7)
    DllCall("Gdiplus.dll\GdipSetCompositingQuality", "Ptr", PGRAPHICS, "UInt", 4)
    DllCall("Gdiplus.dll\GdipSetRenderingOrigin", "Ptr", PGRAPHICS, "Int", 0, "Int", 0)
    DllCall("Gdiplus.dll\GdipSetPixelOffsetMode", "Ptr", PGRAPHICS, "UInt", 4)
    DllCall("Gdiplus.dll\GdipSetTextRenderingHint", "Ptr", PGRAPHICS, "Int", 0)
    ; Create a StringFormat object
    HFORMAT := 0
    DllCall("Gdiplus.dll\GdipStringFormatGetGenericTypographic", "PtrP", &HFORMAT)
    ; Horizontal alignment
    ; BS_LEFT = 0x0100, BS_RIGHT = 0x0200, BS_CENTER = 0x0300, BS_TOP = 0x0400, BS_BOTTOM = 0x0800, BS_VCENTER = 0x0C00
    ; SA_LEFT = 0, SA_CENTER = 1, SA_RIGHT = 2
    HALIGN := (BtnStyle & 0x0300) = 0x0300 ? 1
        : (BtnStyle & 0x0300) = 0x0200 ? 2
            : (BtnStyle & 0x0300) = 0x0100 ? 0
            : 1
    DllCall("Gdiplus.dll\GdipSetStringFormatAlign", "Ptr", HFORMAT, "Int", HALIGN)
    ; Vertical alignment
    VALIGN := (BtnStyle & 0x0C00) = 0x0400 ? 0
        : (BtnStyle & 0x0C00) = 0x0800 ? 2
            : 1
    DllCall("Gdiplus.dll\GdipSetStringFormatLineAlign", "Ptr", HFORMAT, "Int", VALIGN)
    DllCall("Gdiplus.dll\GdipSetStringFormatHotkeyPrefix", "Ptr", HFORMAT, "UInt", 1) ; THX robodesign
    StringFormat := HFORMAT
    ; -------------------------------------------------------------------------------------------------------------------
    ; Create the bitmap(s)
    BitMaps := []
    BitMaps.Length := MaxBitmaps
    Opt1 := Options[1]
    Opt1.Length := MaxOptions
    Loop MaxOptions
        If !Opt1.Has(A_Index)
            Opt1[A_Index] := ""
    If (Opt1[3] = "")
        Opt1[3] := GetARGB(DefTxtColor)
    For Idx, Opt In Options {
        If !IsSet(Opt) || !IsObject(Opt) || !(Opt Is Array)
            Continue
        BkgColor1 := BkgColor2 := TxtColor := Rounded := GuiColor := Image := ""
        ; Replace omitted options with the values of Options.1
        If (Idx > 1) {
            Opt.Length := MaxOptions
            Loop MaxOptions {
                If !Opt.Has(A_Index) || (Opt[A_Index] = "")
                    Opt[A_Index] := Opt1[A_Index]
            }
        }
        ; ----------------------------------------------------------------------------------------------------------------
        ; Check option values
        ; StartColor & TargetColor
        If (Mode = 0) && BitmapOrIcon(Opt[1], Opt[2])
            Image := Opt[1]
        Else {
            If !IsInteger(Opt[1]) && !HTML.HasOwnProp(Opt[1])
                Return ErrorExit("Invalid value for StartColor in Options[" . Idx . "]!")
            BkgColor1 := GetARGB(Opt[1])
            If (Opt[2] = "")
                Opt[2] := Opt[1]
            If !IsInteger(Opt[2]) && !HTML.HasOwnProp(Opt[2])
                Return ErrorExit("Invalid value for TargetColor in Options[" . Idx . "]!")
            BkgColor2 := GetARGB(Opt[2])
        }
        ; TextColor
        If (Opt[3] = "")
            Opt[3] := GetARGB(DefTxtColor)
        If !IsInteger(Opt[3]) && !HTML.HasOwnProp(Opt[3])
            Return ErrorExit("Invalid value for TxtColor in Options[" . Idx . "]!")
        TxtColor := GetARGB(Opt[3])
        ; Rounded
        Rounded := Opt[4]
        If (Rounded = "H")
            Rounded := BtnH * 0.5
        If (Rounded = "W")
            Rounded := BtnW * 0.5
        If !IsNumber(Rounded)
            Rounded := 0
        ; GuiColor
        If DefGuiColor = "*GUI*"
            GuiColor := GetARGB(GuiBtn.Gui.Backcolor != "" ? "0x" GuiBtn.Gui.Backcolor : SetDefGuiColor("*DEF*"))
        Else
            GuiColor := GetARGB(DefGuiColor)
        ; BorderColor
        BorderColor := ""
        If (Opt[5] != "") {
            If !IsInteger(Opt[5]) && !HTML.HasOwnProp(Opt[5])
                Return ErrorExit("Invalid value for BorderColor in Options[" . Idx . "]!")
            BorderColor := 0xFF000000 | GetARGB(Opt[5]) ; BorderColor must be always opaque
        }
        ; BorderWidth
        BorderWidth := Opt[6] ? Opt[6] : 1
        ; ----------------------------------------------------------------------------------------------------------------
        ; Clear the background
        DllCall("Gdiplus.dll\GdipGraphicsClear", "Ptr", PGRAPHICS, "UInt", GuiColor)
        ; Create the image
        If (Image = "") { ; Create a BitMap based on the specified colors
            PathX := PathY := 0, PathW := BtnW, PathH := BtnH
            ; Create a GraphicsPath
            PPATH := 0
            DllCall("Gdiplus.dll\GdipCreatePath", "UInt", 0, "PtrP", &PPATH)
            If (Rounded < 1) ; the path is a rectangular rectangle
                PathAddRectangle(PPATH, PathX, PathY, PathW, PathH)
            Else ; the path is a rounded rectangle
                PathAddRoundedRect(PPATH, PathX, PathY, PathW, PathH, Rounded)
            ; If BorderColor and BorderWidth are specified, 'draw' the border (not for Mode 7)
            If (BorderColor != "") && (BorderWidth > 0) && (Mode != 7) {
                ; Create a SolidBrush
                PBRUSH := 0
                DllCall("Gdiplus.dll\GdipCreateSolidFill", "UInt", BorderColor, "PtrP", &PBRUSH)
                ; Fill the path
                DllCall("Gdiplus.dll\GdipFillPath", "Ptr", PGRAPHICS, "Ptr", PBRUSH, "Ptr", PPATH)
                ; Free the brush
                DllCall("Gdiplus.dll\GdipDeleteBrush", "Ptr", PBRUSH)
                ; Reset the path
                DllCall("Gdiplus.dll\GdipResetPath", "Ptr", PPATH)
                ; Add a new 'inner' path
                PathX := PathY := BorderWidth, PathW -= BorderWidth, PathH -= BorderWidth, Rounded -= BorderWidth
                If (Rounded < 1) ; the path is a rectangular rectangle
                    PathAddRectangle(PPATH, PathX, PathY, PathW - PathX, PathH - PathY)
                Else ; the path is a rounded rectangle
                    PathAddRoundedRect(PPATH, PathX, PathY, PathW, PathH, Rounded)
                ; If a BorderColor has been drawn, BkgColors must be opaque
                BkgColor1 := 0xFF000000 | BkgColor1
                BkgColor2 := 0xFF000000 | BkgColor2
            }
            PathW -= PathX
            PathH -= PathY
            PBRUSH := 0
            RECTF := 0
            Switch Mode {
                Case 0:                    ; the background is unicolored
                    ; Create a SolidBrush
                    DllCall("Gdiplus.dll\GdipCreateSolidFill", "UInt", BkgColor1, "PtrP", &PBRUSH)
                    ; Fill the path
                    DllCall("Gdiplus.dll\GdipFillPath", "Ptr", PGRAPHICS, "Ptr", PBRUSH, "Ptr", PPATH)
                Case 1, 2:                 ; the background is bicolored
                    ; Create a LineGradientBrush
                    SetRectF(&RECTF, PathX, PathY, PathW, PathH)
                    DllCall("Gdiplus.dll\GdipCreateLineBrushFromRect",
                        "Ptr", RECTF, "UInt", BkgColor1, "UInt", BkgColor2, "Int", Mode & 1, "Int", 3, "PtrP", &PBRUSH)
                    DllCall("Gdiplus.dll\GdipSetLineGammaCorrection", "Ptr", PBRUSH, "Int", GammaCorr)
                    ; Set up colors and positions
                    SetRect(&COLORS, BkgColor1, BkgColor1, BkgColor2, BkgColor2) ; sorry for function misuse
                    SetRectF(&POSITIONS, 0, 0.5, 0.5, 1) ; sorry for function misuse
                    DllCall("Gdiplus.dll\GdipSetLinePresetBlend",
                        "Ptr", PBRUSH, "Ptr", COLORS, "Ptr", POSITIONS, "Int", 4)
                    ; Fill the path
                    DllCall("Gdiplus.dll\GdipFillPath", "Ptr", PGRAPHICS, "Ptr", PBRUSH, "Ptr", PPATH)
                Case 3, 4, 5, 6, 8, 9:     ; the background is a gradient
                    ; Determine the brush's width/height
                    W := Mode = 6 ? PathW / 2 : PathW  ; horizontal
                    H := Mode = 5 ? PathH / 2 : PathH  ; vertical
                    ; Create a LineGradientBrush
                    SetRectF(&RECTF, PathX, PathY, W, H)
                    LGM := Mode > 6 ? Mode - 6 : Mode & 1 ; LinearGradientMode
                    DllCall("Gdiplus.dll\GdipCreateLineBrushFromRect",
                        "Ptr", RECTF, "UInt", BkgColor1, "UInt", BkgColor2, "Int", LGM, "Int", 3, "PtrP", &PBRUSH)
                    DllCall("Gdiplus.dll\GdipSetLineGammaCorrection", "Ptr", PBRUSH, "Int", GammaCorr)
                    ; Fill the path
                    DllCall("Gdiplus.dll\GdipFillPath", "Ptr", PGRAPHICS, "Ptr", PBRUSH, "Ptr", PPATH)
                Case 7:                    ; raised mode
                    DllCall("Gdiplus.dll\GdipCreatePathGradientFromPath", "Ptr", PPATH, "PtrP", &PBRUSH)
                    ; Set Gamma Correction
                    DllCall("Gdiplus.dll\GdipSetPathGradientGammaCorrection", "Ptr", PBRUSH, "UInt", GammaCorr)
                    ; Set surround and center colors
                    ColorArray := Buffer(4, 0)
                    NumPut("UInt", BkgColor1, ColorArray)
                    DllCall("Gdiplus.dll\GdipSetPathGradientSurroundColorsWithCount",
                        "Ptr", PBRUSH, "Ptr", ColorArray, "IntP", 1)
                    DllCall("Gdiplus.dll\GdipSetPathGradientCenterColor", "Ptr", PBRUSH, "UInt", BkgColor2)
                    ; Set the FocusScales
                    FS := (BtnH < BtnW ? BtnH : BtnW) / 3
                    XScale := (BtnW - FS) / BtnW
                    YScale := (BtnH - FS) / BtnH
                    DllCall("Gdiplus.dll\GdipSetPathGradientFocusScales", "Ptr", PBRUSH, "Float", XScale, "Float", YScale)
                    ; Fill the path
                    DllCall("Gdiplus.dll\GdipFillPath", "Ptr", PGRAPHICS, "Ptr", PBRUSH, "Ptr", PPATH)
            }
            ; Free resources
            DllCall("Gdiplus.dll\GdipDeleteBrush", "Ptr", PBRUSH)
            DllCall("Gdiplus.dll\GdipDeletePath", "Ptr", PPATH)
        }
        Else { ; Create a bitmap from HBITMAP or file
            PBM := 0
            If IsInteger(Image)
                If (Opt[2] = "HICON")
                    DllCall("Gdiplus.dll\GdipCreateBitmapFromHICON", "Ptr", Image, "PtrP", &PBM)
                Else
                    DllCall("Gdiplus.dll\GdipCreateBitmapFromHBITMAP", "Ptr", Image, "Ptr", 0, "PtrP", &PBM)
            Else
                DllCall("Gdiplus.dll\GdipCreateBitmapFromFile", "WStr", Image, "PtrP", &PBM)
            ; Draw the bitmap
            DllCall("Gdiplus.dll\GdipDrawImageRectI",
                "Ptr", PGRAPHICS, "Ptr", PBM, "Int", 0, "Int", 0, "Int", BtnW, "Int", BtnH)
            ; Free the bitmap
            DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", PBM)
        }
        ; ----------------------------------------------------------------------------------------------------------------
        ; Draw the caption
        If (BtnCaption != "") {
            ; Text color
            DllCall("Gdiplus.dll\GdipCreateSolidFill", "UInt", TxtColor, "PtrP", &PBRUSH)
            ; Set the text's rectangle
            RECT := Buffer(16, 0)
            NumPut("Float", BtnW, "Float", BtnH, RECT, 8)
            ; Draw the text
            DllCall("Gdiplus.dll\GdipDrawString",
                "Ptr", PGRAPHICS, "Str", BtnCaption, "Int", -1,
                "Ptr", PFONT, "Ptr", RECT, "Ptr", HFORMAT, "Ptr", PBRUSH)
        }
        ; ----------------------------------------------------------------------------------------------------------------
        ; Create a HBITMAP handle from the bitmap and add it to the array
        HBITMAP := 0
        DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "Ptr", PBITMAP, "PtrP", &HBITMAP, "UInt", 0X00FFFFFF)
        BitMaps[Idx] := HBITMAP
        NumBitmaps++
        ; Free resources
        DllCall("Gdiplus.dll\GdipDeleteBrush", "Ptr", PBRUSH)
    }
    ; Now free remaining the GDI+ objects
    DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", PBITMAP)
    DllCall("Gdiplus.dll\GdipDeleteGraphics", "Ptr", PGRAPHICS)
    DllCall("Gdiplus.dll\GdipDeleteFont", "Ptr", PFONT)
    DllCall("Gdiplus.dll\GdipDeleteStringFormat", "Ptr", HFORMAT)
    Bitmap := Graphics := Font := StringFormat := 0
    ; -------------------------------------------------------------------------------------------------------------------
    ; Create the ImageList
    ; ILC_COLOR32 = 0x20
    HIL := DllCall("Comctl32.dll\ImageList_Create"
        , "UInt", BtnW, "UInt", BtnH, "UInt", 0x20, "Int", 6, "Int", 0, "Ptr") ; ILC_COLOR32
    Loop (NumBitmaps > 1) ? MaxBitmaps : 1 {
        HBITMAP := BitMaps.Has(A_Index) ? BitMaps[A_Index] : BitMaps[1]
        DllCall("Comctl32.dll\ImageList_Add", "Ptr", HIL, "Ptr", HBITMAP, "Ptr", 0)
    }
    ; Create a BUTTON_IMAGELIST structure
    BIL := Buffer(20 + A_PtrSize, 0)
    ; Get the currently assigned image list
    SendMessage(0x1603, 0, BIL.Ptr, HWND) ; BCM_GETIMAGELIST
    PrevIL := NumGet(BIL, "UPtr")
    ; Remove the previous image list, if any
    BIL := Buffer(20 + A_PtrSize, 0)
    NumPut("Ptr", -1, BIL) ; BCCL_NOGLYPH
    SendMessage(0x1602, 0, BIL.Ptr, HWND) ; BCM_SETIMAGELIST
    ; Create a new BUTTON_IMAGELIST structure
    ; BUTTON_IMAGELIST_ALIGN_LEFT = 0, BUTTON_IMAGELIST_ALIGN_RIGHT = 1, BUTTON_IMAGELIST_ALIGN_CENTER = 4,
    BIL := Buffer(20 + A_PtrSize, 0)
    NumPut("Ptr", HIL, BIL)
    Numput("UInt", 4, BIL, A_PtrSize + 16) ; BUTTON_IMAGELIST_ALIGN_CENTER
    ControlSetStyle(BtnStyle | 0x0080, HWND) ; BS_BITMAP
    ; Remove the currently assigned image list, if any
    If (PrevIL)
        IL_Destroy(PrevIL)
    ; Assign the ImageList to the button
    SendMessage(0x1602, 0, BIL.Ptr, HWND) ; BCM_SETIMAGELIST
    ; Free the bitmaps
    FreeBitmaps()
    NumBitmaps := 0
    ; -------------------------------------------------------------------------------------------------------------------
    ; All done successfully
    Buttons[HWND] := Map("HIML", HIL, "Style", BtnStyle)
    Return True
    ; ===================================================================================================================
    ; Internally used functions
    ; ===================================================================================================================
    ; Set the default GUI color.
    ; GuiColor - RGB integer value (0xRRGGBB) or HTML color name ("Red").
    ;          - "*GUI*" to use Gui.Backcolor (default)
    ;          - "*DEF*" to use AHK's default Gui color.
    SetDefGuiColor(GuiColor) {
        Static DefColor := DllCall("GetSysColor", "Int", 15, "UInt") ; COLOR_3DFACE
        Switch
        {
            Case (GuiColor = "*GUI*"):
                Return GuiColor
            Case (GuiColor = "*DEF*"):
                Return GetRGB(DefColor)
            Case IsInteger(GuiColor):
                Return GuiColor & 0xFFFFFF
            Case HTML.HasOwnProp(GuiColor):
                Return HTML.%GuiColor% &0xFFFFFF
            Default:
                Throw ValueError("Parameter GuiColor invalid", -1, GuiColor)
        }
    }
    ; ===================================================================================================================
    ; Set the default text color.
    ; TxtColor - RGB integer value (0xRRGGBB) or HTML color name ("Red").
    ;          - "*DEF*" to reset to AHK's default text color.
    SetDefTxtColor(TxtColor) {
        Static DefColor := DllCall("GetSysColor", "Int", 18, "UInt") ; COLOR_BTNTEXT
        Switch
        {
            Case (TxtColor = "*DEF*"):
                Return GetRGB(DefColor)
            Case IsInteger(TxtColor):
                Return TxtColor & 0xFFFFFF
            Case HTML.HasOwnProp(TxtColor):
                Return HTML.%TxtColor% &0xFFFFFF
            Default:
                Throw ValueError("Parameter TxtColor invalid", -1, TxtColor)
        }
        Return True
    }
    ; ===================================================================================================================
    ; PRIVATE FUNCTIONS =================================================================================================
    ; ===================================================================================================================
    BitmapOrIcon(O1, O2) {
        ; OBJ_BITMAP = 7
        Return IsInteger(O1) ? (O2 = "HICON") || (DllCall("GetObjectType", "Ptr", O1, "UInt") = 7) : FileExist(O1)
    }
    ; -------------------------------------------------------------------------------------------------------------------
    FreeBitmaps() {
        For HBITMAP In BitMaps
            IsSet(HBITMAP) ? DllCall("DeleteObject", "Ptr", HBITMAP) : 0
        BitMaps := []
    }
    ; -------------------------------------------------------------------------------------------------------------------
    GetARGB(RGB) {
        ARGB := HTML.HasOwnProp(RGB) ? HTML.%RGB% : RGB
        Return (ARGB & 0xFF000000) = 0 ? 0xFF000000 | ARGB : ARGB
    }
    ; -------------------------------------------------------------------------------------------------------------------
    GetRGB(BGR) {
        Return ((BGR & 0xFF0000) >> 16) | (BGR & 0x00FF00) | ((BGR & 0x0000FF) << 16)
    }
    ; -------------------------------------------------------------------------------------------------------------------
    PathAddRectangle(Path, X, Y, W, H) {
        Return DllCall("Gdiplus.dll\GdipAddPathRectangle", "Ptr", Path, "Float", X, "Float", Y, "Float", W, "Float", H)
    }
    ; -------------------------------------------------------------------------------------------------------------------
    PathAddRoundedRect(Path, X1, Y1, X2, Y2, R) {
        D := (R * 2), X2 -= D, Y2 -= D
        DllCall("Gdiplus.dll\GdipAddPathArc",
            "Ptr", Path, "Float", X1, "Float", Y1, "Float", D, "Float", D, "Float", 180, "Float", 90)
        DllCall("Gdiplus.dll\GdipAddPathArc",
            "Ptr", Path, "Float", X2, "Float", Y1, "Float", D, "Float", D, "Float", 270, "Float", 90)
        DllCall("Gdiplus.dll\GdipAddPathArc",
            "Ptr", Path, "Float", X2, "Float", Y2, "Float", D, "Float", D, "Float", 0, "Float", 90)
        DllCall("Gdiplus.dll\GdipAddPathArc",
            "Ptr", Path, "Float", X1, "Float", Y2, "Float", D, "Float", D, "Float", 90, "Float", 90)
        Return DllCall("Gdiplus.dll\GdipClosePathFigure", "Ptr", Path)
    }
    ; -------------------------------------------------------------------------------------------------------------------
    SetRect(&Rect, L := 0, T := 0, R := 0, B := 0) {
        Rect := Buffer(16, 0)
        NumPut("Int", L, "Int", T, "Int", R, "Int", B, Rect)
        Return True
    }
    ; -------------------------------------------------------------------------------------------------------------------
    SetRectF(&Rect, X := 0, Y := 0, W := 0, H := 0) {
        Rect := Buffer(16, 0)
        NumPut("Float", X, "Float", Y, "Float", W, "Float", H, Rect)
        Return True
    }
    ; -------------------------------------------------------------------------------------------------------------------
    ErrorExit(ErrMsg) {
        If (Bitmap)
            DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", Bitmap)
        If (Graphics)
            DllCall("Gdiplus.dll\GdipDeleteGraphics", "Ptr", Graphics)
        If (Font)
            DllCall("Gdiplus.dll\GdipDeleteFont", "Ptr", Font)
        If (StringFormat)
            DllCall("Gdiplus.dll\GdipDeleteStringFormat", "Ptr", StringFormat)
        If (HIML) {
            BIL := Buffer(20 + A_PtrSize, 0)
            NumPut("Ptr", -1, BIL) ; BCCL_NOGLYPH
            DllCall("SendMessage", "Ptr", HWND, "UInt", 0x1602, "Ptr", 0, "Ptr", BIL) ; BCM_SETIMAGELIST
            IL_Destroy(HIML)
        }
        Bitmap := 0
        Graphics := 0
        Font := 0
        StringFormat := 0
        HIML := 0
        FreeBitmaps()
        Throw Error(ErrMsg)
    }
}

; ----------------------------------------------------------------------------------------------------------------------
; Loads and initializes the Gdiplus.dll.
; Must be called once before you use any of the DLL functions.
; ----------------------------------------------------------------------------------------------------------------------
#DllLoad "Gdiplus.dll"
UseGDIP() {
    Static GdipObject := 0
    If !IsObject(GdipObject) {
        GdipToken := 0
        SI := Buffer(24, 0) ; size of 64-bit structure
        NumPut("UInt", 1, SI)
        If DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", &GdipToken, "Ptr", SI, "Ptr", 0, "UInt") {
            MsgBox("GDI+ could not be startet!`n`nThe program will exit!", A_ThisFunc, 262160)
            ExitApp
        }
        GdipObject := { __Delete: UseGdipShutDown }
    }
    UseGdipShutDown(*) {
        DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", GdipToken)
    }
}

; ======================================================================================================================
; Namespace:      LV_Colors
; Function:       Individual row and cell coloring for AHK ListView controls.
; Tested with:    AHK 2.0.2 (U32/U64)
; Tested on:      Win 10 (x64)
; Changelog:      2023-01-04/2.0.0/just me   Initial release of the AHK v2 version
; ======================================================================================================================
; CLASS LV_Colors
;
; The class provides methods to set individual colors for rows and/or cells, to clear all colors, to prevent/allow
; sorting and rezising of columns dynamically, and to deactivate/activate the notification handler for NM_CUSTOMDRAW
; notifications (see below).
;
; A message handler for NM_CUSTOMDRAW notifications will be activated for the specified ListView whenever a new
; instance is created. If you want to temporarily disable coloring call MyInstance.ShowColors(False). This must
; be done also before you try to destroy the instance. To enable it again, call MyInstance.ShowColors().
;
; To avoid the loss of Gui events and messages the message handler is set 'critical'. To prevent 'freezing' of the
; list-view or the whole GUI this script requires AHK v2.0.1+.
; ======================================================================================================================
Class LV_Colors {
    ; ===================================================================================================================
    ; __New()         Constructor - Create a new LV_Colors instance for the given ListView
    ; Parameters:     HWND        -  ListView's HWND.
    ;                 Optional ------------------------------------------------------------------------------------------
    ;                 StaticMode  -  Static color assignment, i.e. the colors will be assigned permanently to the row
    ;                                contents rather than to the row number.
    ;                                Values:  True/False
    ;                                Default: False
    ;                 NoSort      -  Prevent sorting by click on a header item.
    ;                                Values:  True/False
    ;                                Default: False
    ;                 NoSizing    -  Prevent resizing of columns.
    ;                                Values:  True/False
    ;                                Default: False
    ; ===================================================================================================================
    __New(LV, StaticMode := False, NoSort := False, NoSizing := False) {
        If (LV.Type != "ListView")
            Throw TypeError("LV_Colors requires a ListView control!", -1, LV.Type)
        ; ----------------------------------------------------------------------------------------------------------------
        ; Set LVS_EX_DOUBLEBUFFER (0x010000) style to avoid drawing issues.
        LV.Opt("+LV0x010000")
        ; Get the default colors
        BkClr := SendMessage(0x1025, 0, 0, LV) ; LVM_GETTEXTBKCOLOR
        TxClr := SendMessage(0x1023, 0, 0, LV) ; LVM_GETTEXTCOLOR
        ; Get the header control
        Header := SendMessage(0x101F, 0, 0, LV) ; LVM_GETHEADER
        ; Set other properties
        This.LV := LV
        This.HWND := LV.HWND
        This.Header := Header
        This.BkClr := BkCLr
        This.TxClr := Txclr
        This.IsStatic := !!StaticMode
        This.AltCols := False
        This.AltRows := False
        This.SelColors := False
        This.NoSort(!!NoSort)
        This.NoSizing(!!NoSizing)
        This.ShowColors()
        This.RowCount := LV.GetCount()
        This.ColCount := LV.GetCount("Col")
        This.Rows := Map()
        This.Rows.Capacity := This.RowCount
        This.Cells := Map()
        This.Cells.Capacity := This.RowCount
    }
    ; ===================================================================================================================
    ; __Delete()      Destructor
    ; ===================================================================================================================
    __Delete() {
        This.ShowColors(False)
        If WinExist(This.HWND)
            WinRedraw(This.HWND)
    }
    ; ===================================================================================================================
    ; Clear()         Clears all row and cell colors.
    ; Parameters:     AltRows     -  Reset alternate row coloring (True / False)
    ;                                Default: False
    ;                 AltCols     -  Reset alternate column coloring (True / False)
    ;                                Default: False
    ; Return Value:   Always True.
    ; ===================================================================================================================
    Clear(AltRows := False, AltCols := False) {
        If (AltCols)
            This.AltCols := False
        If (AltRows)
            This.AltRows := False
        This.Rows.Clear()
        This.Rows.Capacity := This.RowCount
        This.Cells.Clear()
        This.Cells.Capacity := This.RowCount
        Return True
    }
    ; ===================================================================================================================
    ; UpdateProps()   Updates the RowCount, ColCount, BkClr, and TxClr properties.
    ; Return Value:   True on success, otherwise false.
    ; ===================================================================================================================
    UpdateProps() {
        If !(This.HWND)
            Return False
        This.BkClr := SendMessage(0x1025, 0, 0, This.LV) ; LVM_GETTEXTBKCOLOR
        This.TxClr := SendMessage(0x1023, 0, 0, This.LV) ; LVM_GETTEXTCOLOR
        This.RowCount := This.LV.GetCount()
        This.Colcount := This.LV.GetCount("Col")
        If WinExist(This.HWND)
            WinRedraw(This.HWND)
        Return True
    }
    ; ===================================================================================================================
    ; AlternateRows() Sets background and/or text color for even row numbers.
    ; Parameters:     BkColor     -  Background color as RGB color integer (e.g. 0xFF0000 = red) or HTML color name.
    ;                                Default: Empty -> default background color
    ;                 TxColor     -  Text color as RGB color integer (e.g. 0xFF0000 = red) or HTML color name.
    ;                                Default: Empty -> default text color
    ; Return Value:   True on success, otherwise false.
    ; ===================================================================================================================
    AlternateRows(BkColor := "", TxColor := "") {
        If !(This.HWND)
            Return False
        This.AltRows := False
        If (BkColor = "") && (TxColor = "")
            Return True
        BkBGR := This.BGR(BkColor)
        TxBGR := This.BGR(TxColor)
        If (BkBGR = "") && (TxBGR = "")
            Return False
        This.ARB := (BkBGR != "") ? BkBGR : This.BkClr
        This.ART := (TxBGR != "") ? TxBGR : This.TxClr
        This.AltRows := True
        Return True
    }
    ; ===================================================================================================================
    ; AlternateCols() Sets background and/or text color for even column numbers.
    ; Parameters:     BkColor     -  Background color as RGB color integer (e.g. 0xFF0000 = red) or HTML color name.
    ;                                Default: Empty -> default background color
    ;                 TxColor     -  Text color as RGB color integer (e.g. 0xFF0000 = red) or HTML color name.
    ;                                Default: Empty -> default text color
    ; Return Value:   True on success, otherwise false.
    ; ===================================================================================================================
    AlternateCols(BkColor := "", TxColor := "") {
        If !(This.HWND)
            Return False
        This.AltCols := False
        If (BkColor = "") && (TxColor = "")
            Return True
        BkBGR := This.BGR(BkColor)
        TxBGR := This.BGR(TxColor)
        If (BkBGR = "") && (TxBGR = "")
            Return False
        This.ACB := (BkBGR != "") ? BkBGR : This.BkClr
        This.ACT := (TxBGR != "") ? TxBGR : This.TxClr
        This.AltCols := True
        Return True
    }
    ; ===================================================================================================================
    ; SelectionColors() Sets background and/or text color for selected rows.
    ; Parameters:     BkColor     -  Background color as RGB color integer (e.g. 0xFF0000 = red) or HTML color name.
    ;                                Default: Empty -> default selected background color
    ;                 TxColor     -  Text color as RGB color integer (e.g. 0xFF0000 = red) or HTML color name.
    ;                                Default: Empty -> default selected text color
    ; Return Value:   True on success, otherwise false.
    ; ===================================================================================================================
    SelectionColors(BkColor := "", TxColor := "") {
        If !(This.HWND)
            Return False
        This.SelColors := False
        If (BkColor = "") && (TxColor = "")
            Return True
        BkBGR := This.BGR(BkColor)
        TxBGR := This.BGR(TxColor)
        If (BkBGR = "") && (TxBGR = "")
            Return False
        This.SELB := BkBGR
        This.SELT := TxBGR
        This.SelColors := True
        Return True
    }
    ; ===================================================================================================================
    ; Row()           Sets background and/or text color for the specified row.
    ; Parameters:     Row         -  Row number
    ;                 Optional ------------------------------------------------------------------------------------------
    ;                 BkColor     -  Background color as RGB color integer (e.g. 0xFF0000 = red) or HTML color name.
    ;                                Default: Empty -> default background color
    ;                 TxColor     -  Text color as RGB color integer (e.g. 0xFF0000 = red) or HTML color name.
    ;                                Default: Empty -> default text color
    ; Return Value:   True on success, otherwise false.
    ; ===================================================================================================================
    Row(Row, BkColor := "", TxColor := "") {
        If !(This.HWND)
            Return False
        ;If (Row > This.RowCount)
        ;    Return False
        If This.IsStatic
            Row := This.MapIndexToID(Row)
        If This.Rows.Has(Row)
            This.Rows.Delete(Row)
        If (BkColor = "") && (TxColor = "")
            Return True
        BkBGR := This.BGR(BkColor)
        TxBGR := This.BGR(TxColor)
        If (BkBGR = "") && (TxBGR = "")
            Return False
        ; Colors := {B: (BkBGR != "") ? BkBGR : This.BkClr, T: (TxBGR != "") ? TxBGR : This.TxClr}
        This.Rows[Row] := Map("B", (BkBGR != "") ? BkBGR : This.BkClr, "T", (TxBGR != "") ? TxBGR : This.TxClr)
        Return True
    }
    ; ===================================================================================================================
    ; Cell()          Sets background and/or text color for the specified cell.
    ; Parameters:     Row         -  Row number
    ;                 Col         -  Column number
    ;                 Optional ------------------------------------------------------------------------------------------
    ;                 BkColor     -  Background color as RGB color integer (e.g. 0xFF0000 = red) or HTML color name.
    ;                                Default: Empty -> row's background color
    ;                 TxColor     -  Text color as RGB color integer (e.g. 0xFF0000 = red) or HTML color name.
    ;                                Default: Empty -> row's text color
    ; Return Value:   True on success, otherwise false.
    ; ===================================================================================================================
    Cell(Row, Col, BkColor := "", TxColor := "") {
        If !(This.HWND)
            Return False
        ;If (Row > This.RowCount) || (Col > This.ColCount)
        ;    Return False
        If This.IsStatic
            Row := This.MapIndexToID(Row)
        If This.Cells.Has(Row) && This.Cells[Row].Has(Col)
            This.Cells[Row].Delete(Col)
        If (BkColor = "") && (TxColor = "")
            Return True
        BkBGR := This.BGR(BkColor)
        TxBGR := This.BGR(TxColor)
        If (BkBGR = "") && (TxBGR = "")
            Return False
        If !This.Cells.Has(Row)
            This.Cells[Row] := [], This.Cells[Row].Capacity := This.ColCount
        If (Col > This.Cells[Row].Length)
            This.Cells[Row].Length := Col
        This.Cells[Row][Col] := Map("B", (BkBGR != "") ? BkBGR : This.BkClr, "T", (TxBGR != "") ? TxBGR : This.TxClr)
        Return True
    }
    ; ===================================================================================================================
    ; NoSort()        Prevents/allows sorting by click on a header item for this ListView.
    ; Parameters:     Apply       -  True/False
    ; Return Value:   True on success, otherwise false.
    ; ===================================================================================================================
    NoSort(Apply := True) {
        If !(This.HWND)
            Return False
        This.LV.Opt((Apply ? "+" : "-") . "NoSort")
        Return True
    }
    ; ===================================================================================================================
    ; NoSizing()      Prevents/allows resizing of columns for this ListView.
    ; Parameters:     Apply       -  True/False
    ; Return Value:   True on success, otherwise false.
    ; ===================================================================================================================
    NoSizing(Apply := True) {
        If !(This.Header)
            Return False
        ControlSetStyle((Apply ? "+" : "-") . "0x0800", This.Header) ; HDS_NOSIZING = 0x0800
        Return True
    }
    ; ===================================================================================================================
    ; ShowColors()    Adds/removes a message handler for NM_CUSTOMDRAW notifications of this ListView.
    ; Parameters:     Apply       -  True/False
    ; Return Value:   Always True
    ; ===================================================================================================================
    ShowColors(Apply := True) {
        If (Apply) && !This.HasOwnProp("OnNotifyFunc") {
            This.OnNotifyFunc := ObjBindMethod(This, "NM_CUSTOMDRAW")
            This.LV.OnNotify(-12, This.OnNotifyFunc)
            WinRedraw(This.HWND)
        }
        Else If !(Apply) && This.HasOwnProp("OnNotifyFunc") {
            This.LV.OnNotify(-12, This.OnNotifyFunc, 0)
            This.OnNotifyFunc := ""
            This.DeleteProp("OnNotifyFunc")
            WinRedraw(This.HWND)
        }
        Return True
    }
    ; ===================================================================================================================
    ; Internally used/called Methods
    ; ===================================================================================================================
    NM_CUSTOMDRAW(LV, L) {
        ; Return values: 0x00 (CDRF_DODEFAULT), 0x20 (CDRF_NOTIFYITEMDRAW / CDRF_NOTIFYSUBITEMDRAW)
        Static SizeNMHDR := A_PtrSize * 3                  ; Size of NMHDR structure
        Static SizeNCD := SizeNMHDR + 16 + (A_PtrSize * 5) ; Size of NMCUSTOMDRAW structure
        Static OffItem := SizeNMHDR + 16 + (A_PtrSize * 2) ; Offset of dwItemSpec (NMCUSTOMDRAW)
        Static OffItemState := OffItem + A_PtrSize         ; Offset of uItemState  (NMCUSTOMDRAW)
        Static OffCT := SizeNCD                           ; Offset of clrText (NMLVCUSTOMDRAW)
        Static OffCB := OffCT + 4                          ; Offset of clrTextBk (NMLVCUSTOMDRAW)
        Static OffSubItem := OffCB + 4                     ; Offset of iSubItem (NMLVCUSTOMDRAW)
        Critical -1
        If !(This.HWND) || (NumGet(L, "UPtr") != This.HWND)
            Return
        ; ----------------------------------------------------------------------------------------------------------------
        DrawStage := NumGet(L + SizeNMHDR, "UInt"),
            Row := NumGet(L + OffItem, "UPtr") + 1,
            Col := NumGet(L + OffSubItem, "Int") + 1,
            Item := Row - 1
        If This.IsStatic
            Row := This.MapIndexToID(Row)
        ; CDDS_SUBITEMPREPAINT = 0x030001 --------------------------------------------------------------------------------
        If (DrawStage = 0x030001) {
            UseAltCol := (This.AltCols) && !(Col & 1),
                ColColors := (This.Cells.Has(Row) && This.Cells[Row].Has(Col)) ? This.Cells[Row][Col] : Map("B", "", "T", ""),
                ColB := (ColColors["B"] != "") ? ColColors["B"] : UseAltCol ? This.ACB : This.RowB,
                    ColT := (ColColors["T"] != "") ? ColColors["T"] : UseAltCol ? This.ACT : This.RowT,
                        NumPut("UInt", ColT, L + OffCT), NumPut("UInt", ColB, L + OffCB)
            Return (!This.AltCols && (Col > This.Cells[Row].Length)) ? 0x00 : 0x020
        }
        ; CDDS_ITEMPREPAINT = 0x010001 -----------------------------------------------------------------------------------
        If (DrawStage = 0x010001) {
            ; LVM_GETITEMSTATE = 0x102C, LVIS_SELECTED = 0x0002
            If (This.SelColors) && SendMessage(0x102C, Item, 0x0002, This.HWND) {
                ; Remove the CDIS_SELECTED (0x0001) and CDIS_FOCUS (0x0010) states from uItemState and set the colors.
                NumPut("UInt", NumGet(L + OffItemState, "UInt") & ~0x0011, L + OffItemState)
                If (This.SELB != "")
                    NumPut("UInt", This.SELB, L + OffCB)
                If (This.SELT != "")
                    NumPut("UInt", This.SELT, L + OffCT)
                Return 0x02 ; CDRF_NEWFONT
            }
            UseAltRow := This.AltRows && (Item & 1),
                RowColors := This.Rows.Has(Row) ? This.Rows[Row] : "",
                This.RowB := RowColors ? RowColors["B"] : UseAltRow ? This.ARB : This.BkClr,
                    This.RowT := RowColors ? RowColors["T"] : UseAltRow ? This.ART : This.TxClr
            If (This.AltCols || This.Cells.Has(Row))
                Return 0x20
            NumPut("UInt", This.RowT, L + OffCT), NumPut("UInt", This.RowB, L + OffCB)
            Return 0x00
        }
        ; CDDS_PREPAINT = 0x000001 ---------------------------------------------------------------------------------------
        Return (DrawStage = 0x000001) ? 0x20 : 0x00
    }
    ; -------------------------------------------------------------------------------------------------------------------
    MapIndexToID(Row) { ; provides the unique internal ID of the given row number
        Return SendMessage(0x10B4, Row - 1, 0, This.HWND) ; LVM_MAPINDEXTOID
    }
    ; -------------------------------------------------------------------------------------------------------------------
    BGR(Color, Default := "") { ; converts colors to BGR
        ; HTML Colors (BGR)
        Static HTML := { AQUA: 0xFFFF00, BLACK: 0x000000, BLUE: 0xFF0000, FUCHSIA: 0xFF00FF, GRAY: 0x808080, GREEN: 0x008000
            , LIME: 0x00FF00, MAROON: 0x000080, NAVY: 0x800000, OLIVE: 0x008080, PURPLE: 0x800080, RED: 0x0000FF
            , SILVER: 0xC0C0C0, TEAL: 0x808000, WHITE: 0xFFFFFF, YELLOW: 0x00FFFF }
        If IsInteger(Color)
            Return ((Color >> 16) & 0xFF) | (Color & 0x00FF00) | ((Color & 0xFF) << 16)
        Return (HTML.HasOwnProp(Color) ? HTML.%Color% : Default)
    }
}

/* This class defines the structure below(SCROLLINFO) on Winuser.h:

typedef struct tagSCROLLINFO {
  UINT cbSize;
  UINT fMask;
  int  nMin;
  int  nMax;
  UINT nPage;
  int  nPos;
  int  nTrackPos;
} SCROLLINFO, *LPSCROLLINFO */
class ScrollInfo {
    __New() {
        ; Reserves space in computer memory for scrollInf structure with 28 bytes
        this.scrollInf := Buffer(28, 0)
        ; Set cbSize
        NumPut("uint", this.scrollInf.size, this.scrollInf)
    }

    Ptr => this.scrollInf.Ptr

    ; cbSize: Specifies the size, in bytes, of this structure. The caller must set this to sizeof(SCROLLINFO).
    cbSize => NumGet(this.scrollInf, "uint")

    /*  Specifies the scroll bar parameters to set or retrieve. This member can be a combination of the following values:

    SIF_ALL                     Combination of SIF_PAGE, SIF_POS, SIF_RANGE, and SIF_TRACKPOS.
    SIF_DISABLENOSCROLL         If the scroll bar's new parameters make the scroll bar unnecessary, disable the scroll bar instead of removing it.
    SIF_PAGE                    The nPage member contains the page size for a proportional scroll bar.
    SIF_POS                     The nPos member contains the scroll box position, which is not updated while the user drags the scroll box.
    SIF_RANGE                   The nMin and nMax members contain the minimum and maximum values for the scrolling range.
    SIF_TRACKPOS                The nTrackPos member contains the current position of the scroll box while the user is dragging it.                 */
    fMask {
        get => NumGet(this.scrollInf, 4, "uint")
        set => NumPut("uint", value, this.scrollInf, 4)
    }

    ; Specifies the minimum scrolling position.
    nMin {
        get => NumGet(this.scrollInf, 8, "int")
        set => NumPut("int", value, this.scrollInf, 8)
    }

    ; Specifies the maximum scrolling position.
    nMax {
        get => NumGet(this.scrollInf, 12, "int")
        set => NumPut("int", value, this.scrollInf, 12)
    }

    ; Specifies the page size, in device units. A scroll bar uses this value to determine the appropriate size of the proportional scroll box.
    nPage {
        get => NumGet(this.scrollInf, 16, "uint")
        set => NumPut("uint", value, this.scrollInf, 16)
    }

    ; Specifies the position of the scroll box.
    nPos {
        get => NumGet(this.scrollInf, 20, "int")
        set => NumPut("ptr", value, this.scrollInf, 20)
    }

    ; Specifies the immediate position of a scroll box that the user is dragging. An application can retrieve this value while processing the SB_THUMBTRACK request code. 
    ; An application cannot set the immediate scroll position; the SetScrollInfo function ignores this member.
    nTrackPos {
        get => NumGet(this.scrollInf, 24, "int")
        set => NumPut("ptr", value, this.scrollInf, 24)
    }
}

A_MaxHotkeysPerInterval := 9999

class ScrollBar {
    ; Notification codes for horizontal and vertical scroll
    WM_HSCROLL => 0x114
    WM_VSCROLL => 0x115

    ; type of scroll bar (nBar)
    SB_HORZ => 0
    SB_VERT => 1
    SB_BOTH => 3

    ; Scroll bar parameters to set or retrieve (fMask)
    SIF_RANGE => 1
    SIF_PAGE => 2
    SIF_POS => 4
    SIF_TRACKPOS => 16
    SIF_ALL => this.SIF_RANGE | this.SIF_PAGE | this.SIF_POS | this.SIF_TRACKPOS

    ; Scroll Bar Commands
    ; The user pressed the LEFT ARROW (VK_LEFT) key or clicked the left arrow button on a horizontal scroll bar.
    SB_LINELEFT => 0
    ; The user pressed the UP ARROW (VK_UP) key or clicked the up arrow button on a vertical scroll bar.
    SB_LINEUP => 0
    ; The user pressed the RIGHT ARROW (VK_RIGHT) key or clicked the right arrow button on a horizontal scroll bar.
    SB_LINERIGHT => 1
    ; The user pressed the DOWN ARROW (VK_DOWN) key or clicked the down arrow button on a vertical scroll bar.
    SB_LINEDOWN => 1
    ; The user clicked the channel above the slider on a vertical scroll bar or to the left of the slider on a horizontal scroll bar (VK_PRIOR).
    SB_PAGELEFT => 2
    SB_PAGEUP => 2
    ; The user clicked the channel below the slider on a vertical scroll bar or to the right of the slider on a horizontal scroll bar (VK_NEXT).
    SB_PAGERIGHT => 3
    SB_PAGEDOWN => 3
    ; The scrollbar received WM_LBUTTONUP following a SB_THUMBTRACK notification code.
    SB_THUMBPOSITION => 4
    ; The user dragged the slider.
    SB_THUMBTRACK => 5
    ; The user pressed the HOME key (VK_HOME) or clicked the top arrow button on a vertical scroll bar or left arrow button on a horizontal scroll bar.
    SB_LEFT => 6
    SB_TOP => 6
    ; The user pressed the END key (VK_END) or clicked the bottom arrow button on a vertical scroll bar or right arrow button on a horizontal scroll bar.
    SB_RIGHT => 7
    SB_BOTTOM => 7
    ; The scrollbar received WM_KEYUP, meaning that the user released a key that sent a relevant virtual key code.
    SB_ENDSCROLL => 8

    ; Custom
    GAMELOCATION => 100
    OPTION => 101
    GAMEVERSION => 102
    GAMELANGUAGES => 103
    GAMEVISUALMODS => 104
    GAMEDATAMODS => 105
    GAMEOTHERTOOLS => 106

    ; Constructor for the ScrollBar class
    __New(guiObj, width, height) {
        ; Check if the first parameter is a Gui object
        if (guiObj is Gui) {
            ; Set the guiObj property to the first parameter
            this.guiObj := guiObj
            ; Show both scroll bars
            this.ShowScrollBar(this.SB_BOTH, true)

            ; Create a buffer for the rectangle
            this.Rect := Buffer(16)

            this.FixedControls := []

            ; Bind the ScrollMsg method to this object and set it as the message handler for WM_HSCROLL and WM_VSCROLL messages
            this.ScrollMsgBind := ObjBindMethod(this, 'ScrollMsg')
            OnMessage(this.WM_HSCROLL, this.ScrollMsgBind)
            OnMessage(this.WM_VSCROLL, this.ScrollMsgBind)

            ; Do update of scroll bars when I resize the window
            this.guiObj.OnEvent('Size', (*) => this.UpdateScrollBars())

            ; Create a new SCROLLINFO object
            this.ScrollInf := SCROLLINFO()

            ; Gets left-most, right-most, top-most, bottom-most control positions
            this.GetEdges(&Left, &Right, &Top, &Bottom)

            ; Calculate the scroll height and width
            ScrollHeight := Bottom - Top
            ScrollWidth := Right - Left

            if (IsNumber(width) and IsNumber(height) and width > 0 and height > 0) {
                ; Set the maximum scroll position and page size for the vertical scroll bar
                this.ScrollInf.nMax := ScrollHeight
                this.ScrollInf.nPage := height

                this.ScrollInf.fMask := this.SIF_RANGE | this.SIF_PAGE

                ; Set the scroll info for the vertical scroll bar
                this.SetScrollInfo(this.SB_VERT, true)

                ; Set the maximum scroll position and page size for the horizontal scroll bar
                this.ScrollInf.nMax := ScrollWidth
                this.ScrollInf.nPage := width

                ; Set the scroll info for the horizontal scroll bar
                this.SetScrollInfo(this.SB_HORZ, true)

                ; Set the mask to retrieve all scroll info
                this.ScrollInf.fMask := this.SIF_ALL
            } else throw Error('Width and height must be valid numbers') ; Throw an error if width or height are not valid numbers
        } else throw Error('Parameter is not a Gui object') ; Throw an error if the first parameter is not a Gui object
    }

    ; Updates the position of fixed controls while the user scrolls
    UpdateFixedControlsPosition() {
        ; Iterates over the list of fixed controls
        for control in this.FixedControls {
            ; Sets the new position of the control
            control.Move(control.startX, control.startY)
        }
    }

    ; Add fixed controls...
    AddFixedControls(controls) {
        ; Verifies if the parameter is an array
        if (!(controls is Array)) {
            throw Error('Parameter must be an array of controls')
        }

        ; Adds each control to the list of fixed controls
        for control in controls {
            ; Gets the coordinates of the control
            control.GetPos(&controlX, &controlY)
            control.startX := controlX
            control.startY := controlY
            ; Stores the control in the list of fixed controls
            this.FixedControls.Push(control)
        }
    }

    UpdateScrollBars() {
        ; Gets left-most, right-most, top-most, bottom-most control positions
        this.GetEdges(&Left, &Right, &Top, &Bottom)

        ; Calculate the scroll width and height
        ScrollWidth := Right - Left
        ScrollHeight := Bottom - Top

        ; Set the mask to update the range and page size of the scroll bar
        this.ScrollInf.fMask := this.SIF_RANGE | this.SIF_PAGE

        ; Update the maximum scroll position and page size for the vertical scroll bar
        this.ScrollInf.nMax := ScrollHeight
        this.ScrollInf.nPage := this.GetHeight()

        ; Set the scroll info for the vertical scroll bar
        this.SetScrollInfo(this.SB_VERT, true)

        ; Update the maximum scroll position and page size for the horizontal scroll bar
        this.ScrollInf.nMax := ScrollWidth
        this.ScrollInf.nPage := this.GetWidth()

        ; Set the scroll info for the horizontal scroll bar
        this.SetScrollInfo(this.SB_HORZ, true)

        /*
        The code below checks if the left or top position of the content is less than 0 and if
        the right or bottom position of the content is less than the width or height of the window. If
        both conditions are true for either axis, it calculates how much to scroll in that axis to bring
        the content back into view. It then calls the ScrollWindow function to scroll the content by that
        amount in both axes.
        */

        x := 0, y := 0

        if (Left < 0 && Right < this.GetWidth()) {
            x := Abs(Left) > this.GetWidth() - Right ? this.GetWidth() - Right : Abs(Left)
        }
        if (Top < 0 && Bottom < this.GetHeight()) {
            y := Abs(Top) > this.GetHeight() - Bottom ? this.GetHeight() - Bottom : Abs(Top)
        }
        if (x || y) {
            DllCall("ScrollWindow", "ptr", this.guiObj.Hwnd, "int", x, "int", y, "uint", 0, "uint", 0)
        }

        ; Set the mask to retrieve all scroll info
        this.ScrollInf.fMask := this.SIF_ALL
    }

    HiWord(wParam) {
        Return (wParam >> 16)
    }

    LoWord(wParam) {
        Return (wParam & 0xFFFF)
    }

    ; The ScrollMsg function is called when the window receives a WM_HSCROLL or WM_VSCROLL message.
    ; It calls the ScrollAction function to update the scroll bar position and then calls the ScrollWindow function to scroll the content.
    ScrollMsg(wParam, lParam, msg, hwnd) {
        switch msg {
            ; If the message is WM_HSCROLL, update the horizontal scroll bar
            case this.WM_HSCROLL:
                this.ScrollAction(this.SB_HORZ, wParam)
                this.ScrollWindow(this.oldPos - this.ScrollInf.nPos, 0)
                this.UpdateFixedControlsPosition()
            ; If the message is WM_VSCROLL, update the vertical scroll bar
            case this.WM_VSCROLL:
                this.ScrollAction(this.SB_VERT, wParam)
                this.ScrollWindow(0, this.oldPos - this.ScrollInf.nPos)
                this.UpdateFixedControlsPosition()
        }
    }

    ; The ScrollAction function updates the scroll bar position based on the scroll action specified in wParam.
    ; It first gets the current scroll info and position for the specified scroll bar and then calculates the new position based on the scroll action.
    ScrollAction(typeOfScrollBar, wParam) {
        ; Get current attributes of scroll bar
        this.GetScrollInfo(typeOfScrollBar)

        ; Store current position of scroll bar
        this.oldPos := this.ScrollInf.nPos

        ; Get current scroll range
        this.GetScrollRange(typeOfScrollBar, &minPos, &maxPos)

        ; Calculates max position of scroll bar's thumb (scroll box)
        maxThumbPos := this.ScrollInf.nMax - this.ScrollInf.nMin + 1 - this.ScrollInf.nPage

        ; Updates scroll bar position based on command received
        switch this.LoWord(wParam) {
            case this.SB_LINELEFT, this.SB_LINEUP:
                this.ScrollInf.nPos := max(this.ScrollInf.nPos - 40, minPos)
            case this.SB_PAGELEFT, this.SB_PAGEUP:
                this.ScrollInf.nPos := max(this.ScrollInf.nPos - this.ScrollInf.nPage, minPos)
            case this.SB_LINERIGHT, this.SB_LINEDOWN:
                this.ScrollInf.nPos := min(this.ScrollInf.nPos + 40, maxThumbPos)
            case this.SB_PAGERIGHT, this.SB_PAGEDOWN:
                this.ScrollInf.nPos := min(this.ScrollInf.nPos + this.ScrollInf.nPage, maxThumbPos)
            case this.SB_THUMBTRACK:
                this.ScrollInf.nPos := this.HiWord(wParam)
            case this.SB_LEFT, this.SB_TOP:
                this.ScrollInf.nPos := minPos
            case this.SB_RIGHT, this.SB_BOTTOM:
                this.ScrollInf.nPos := maxThumbPos
            case this.GAMELOCATION:
                this.ScrollInf.nPos := 285
            case this.OPTION:
                this.ScrollInf.nPos := 530
            case this.GAMEVERSION:
                this.ScrollInf.nPos := 700
            case this.GAMELANGUAGES:
                this.ScrollInf.nPos := 1170
            case this.GAMEVISUALMODS:
                this.ScrollInf.nPos := 1649
            case this.GAMEDATAMODS:
                this.ScrollInf.nPos := 5941
            default:
                return
        }
        this.SetScrollInfo(typeOfScrollBar, true)
    }

    GetClientRect() {
        return DllCall("GetClientRect", "uint", this.guiObj.Hwnd, "ptr", this.Rect.Ptr)
    }

    ; Gets current visible height
    GetHeight() {
        this.GetClientRect()
        return NumGet(this.Rect, 12, "int")
    }

    ; Gets current visible height
    GetWidth() {
        this.GetClientRect()
        return NumGet(this.Rect, 8, "int")
    }

    ; Gets left-most, right-most, top-most, bottom-most control positions
    GetEdges(&Left?, &Right?, &Top?, &Bottom?) {
        ; Calculate scrolling area.
        Left := Top := 9999
        Right := Bottom := 0
        ; Get a list of all controls in guiObj
        ControlList := WinGetControls(this.guiObj.Hwnd)
        ; Loops through all controls and finds the farthest sides
        For i in ControlList {
            ; Gets all positions of current control
            this.guiObj[i].GetPos(&cX, &cY, &cW, &cH)
            ; If it's position is farther than the last one, saves it
            if (cX < Left) {
                Left := cX
            }
            if (cY < Top) {
                Top := cY
            }
            if (cX + cW > Right) {
                Right := cX + cW
            }
            if (cY + cH > Bottom) {
                Bottom := cY + cH
            }
        }

        ; Gives a little more space for the edges
        Left -= 8
        Top -= 8
        Right += 8
        Bottom += 8
    }

    ; The ShowScrollBar function shows or hides the specified scroll bar.
    ; f the function succeeds, the return value is nonzero.
    ShowScrollBar(typeOfScrollBar, bool) {
        return DllCall("ShowScrollBar", "ptr", this.guiObj.Hwnd, "int", typeOfScrollBar, "int", bool)
    }

    ; The GetScrollInfo function retrieves the parameters of a scroll bar, including the minimum and maximum scrolling positions,
    ; the page size, and the position of the scroll box (thumb).
    ; Before calling GetScrollInfo, set the cbSize member to sizeof(SCROLLINFO), and set the fMask member to specify the scroll bar parameters to retrieve.
    ; If the function retrieved any values, the return value is nonzero.
    GetScrollInfo(typeOfScrollBar) {
        return DllCall("GetScrollInfo", "ptr", this.guiObj.Hwnd, "int", typeOfScrollBar, "ptr", this.ScrollInf.Ptr)
    }

    ; The SetScrollInfo function sets the parameters of a scroll bar, including the minimum and maximum scrolling positions,
    ; the page size, and the position of the scroll box (thumb). The function also redraws the scroll bar, if requested.
    ; The return value is the current position of the scroll box.
    SetScrollInfo(typeOfScrollBar, redraw) {
        return DllCall("SetScrollInfo", "ptr", this.guiObj.Hwnd, "int", typeOfScrollBar, "ptr", this.ScrollInf.Ptr, "int", redraw)
    }

    ; The GetScrollRange function retrieves the current minimum and maximum scroll box (thumb) positions for the specified scroll bar.
    ; If the function succeeds, the return value is nonzero.
    GetScrollRange(typeOfScrollBar, &minPos, &maxPos) {
        minnn := Buffer(4)
        maxxx := Buffer(4)
        r := DllCall("GetScrollRange", "ptr", this.guiObj.Hwnd, "int", typeOfScrollBar, "ptr", minnn.Ptr, "ptr", maxxx.Ptr)
        minPos := NumGet(minnn, "int"), maxPos := NumGet(maxxx, "int")
        return r
    }

    ; The ScrollWindow function scrolls the contents of the specified window's client area.
    ; If the function succeeds, the return value is nonzero.
    ScrollWindow(xamount, yamount) {
        return DllCall("ScrollWindow", "ptr", this.guiObj.Hwnd, "int", xamount, "int", yamount, "ptr", 0, "ptr", 0, "int")
    }
}

GuiButtonIcon(Handle, File, Index := 1, Options := '') {
	RegExMatch(Options, 'i)w\K\d+', &W) ? W := W.0 : W := 16
	RegExMatch(Options, 'i)h\K\d+', &H) ? H := H.0 : H := 16
	RegExMatch(Options, 'i)s\K\d+', &S) ? W := H := S.0 : ''
	RegExMatch(Options, 'i)l\K\d+', &L) ? L := L.0 : L := 0
	RegExMatch(Options, 'i)t\K\d+', &T) ? T := T.0 : T := 0
	RegExMatch(Options, 'i)r\K\d+', &R) ? R := R.0 : R := 0
	RegExMatch(Options, 'i)b\K\d+', &B) ? B := B.0 : B := 0
	RegExMatch(Options, 'i)a\K\d+', &A) ? A := A.0 : A := 4
	W *= A_ScreenDPI / 96, H *= A_ScreenDPI / 96
	button_il := Buffer(20 + A_PtrSize)
	normal_il := DllCall('ImageList_Create', 'Int', W, 'Int', H, 'UInt', 0x21, 'Int', 1, 'Int', 1)
	NumPut('Ptr', normal_il, button_il, 0)			; Width & Height
	NumPut('UInt', L, button_il, 0 + A_PtrSize)		; Left Margin
	NumPut('UInt', T, button_il, 4 + A_PtrSize)		; Top Margin
	NumPut('UInt', R, button_il, 8 + A_PtrSize)		; Right Margin
	NumPut('UInt', B, button_il, 12 + A_PtrSize)	; Bottom Margin
	NumPut('UInt', A, button_il, 16 + A_PtrSize)	; Alignment
	SendMessage(BCM_SETIMAGELIST := 5634, 0, button_il, Handle)
	Return IL_Add(normal_il, File, Index)
}