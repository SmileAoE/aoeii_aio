#Requires AutoHotkey v2
#SingleInstance Force

If !A_IsAdmin {
    MsgBox("- This application is not being ran as administrator`n"
        . "- This can cause an unexpected behaviour on using any of it's options`n"
        . "- It is highly recommended that you run it as administrator"
        , 'Warning'
        , 0x30)
}

; Initialization
Server                          := 'https://raw.githubusercontent.com'
User                            := 'SmileAoE'
Repo                            := 'aoeii_aio'
Version                         := '1.4'
Layers                          := 'HKLM\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers'
Config                          := A_AppData '\aoeii_aio\config.ini'
AppDir                          := ['DB', A_AppData '\aoeii_aio', A_AppData '\aoeii_aio\Hotkeys']
GRSetting                       := A_AppData '\GameRanger\GameRanger Prefs\Settings'
DrsTypes                        := Map('gra', 'graphics.drs'
                                     , 'int', 'interfac.drs'
                                     , 'ter', 'terrain.drs')
DrsRange                        := Map('gra', [2, 5312]
                                     , 'int', [50100, 53211]
                                     , 'ter', [15000, 15031])
IDL                             := 5
VCodedSlp                       := '3713EFBE'
NormalSlp                       := '322E304E'
General                         := Map()
General['AOK']                  := Map()
General['AOK']['VersionsN']     := Map()
General['AOK']['Combine']       := Map('2.0b CD'
                                    , ['2.0a No CD'])
General['AOC']                  := Map()
General['AOC']['VersionsN']     := Map()
General['AOC']['Combine']       := Map('1.0e No CD', ['1.0c No CD']
                                     , '1.0e No CD', ['1.0c No CD']
                                     , '1.1  No CD', ['1.0c No CD']
                                     , '1.5  CD'   , ['1.0c No CD'])
General['FOE']                  := Map()
General['FOE']['VersionsN']     := Map()
General['FOE']['Combine']       := Map()
General['LNG']                  := Map()
Compatibilities                 := Map(1 , ["_____Not Set_____" , ""]
                                     , 2 , ["Windows 8"         , "WIN8RTM"]
                                     , 3 , ["Windows 7"         , "WIN7RTM"]
                                     , 4 , ["Windows Vista Sp2" , "VISTASP2"]
                                     , 5 , ["Windows Vista Sp1" , "VISTASP1"]
                                     , 6 , ["Windows Vista"     , "VISTARTM"]
                                     , 7 , ["Windows XP Sp3"    , "WINXPSP3"]
                                     , 8 , ["Windows XP Sp2"    , "WINXPSP2"]
                                     , 9 , ["Windows 98"        , "WIN98"]
                                     , 10, ["Windows 95"        , "WIN95"])
BasePackages                    := ['DB/000.7z.001'
                                  , 'DB/001.7z.001'
                                  , 'DB/002.7z.001'
                                  , 'DB/006.7z.001'
                                  , 'DB/007.7z.001'
                                  , 'DB/008.7z.001']
GamePackages                    := ['DB/003.7z.001'
                                  , 'DB/003.7z.002'
                                  , 'DB/003.7z.003'
                                  , 'DB/003.7z.004'
                                  , 'DB/004.7z.001'
                                  , 'DB/004.7z.002'
                                  , 'DB/004.7z.003'
                                  , 'DB/005.7z.001']
Dots                            := 0
Task                            := 1
TaskNumber                      := BasePackages.Length

ProgramFiles86                  := EnvGet(A_Is64bitOS ? "ProgramFiles(x86)" : "ProgramFiles")
VPNDir                          := ProgramFiles86 '\Hide ALL IP'
VPNExe                          := 'HideALLIP.exe'
VPNPath                         := VPNDir '\' VPNExe
Shortcut1                       := '
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
If !FileExist(AppDir[3] '\001.ahk') || FileRead(AppDir[3] '\001.ahk') != Shortcut1 {
    O := FileOpen(AppDir[3] '\001.ahk', 'w')
    O.Write(Shortcut1)
    O.Close()
}
Shortcut2                       := '
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
If !FileExist(AppDir[3] '\002.ahk') || FileRead(AppDir[3] '\002.ahk') != Shortcut2 {
    O := FileOpen(AppDir[3] '\002.ahk', 'w')
    O.Write(Shortcut2)
    O.Close()
}
; Preparation

; Create app folders
For _, Item in AppDir {
    If !DirExist(Item) {
        DirCreate(Item)
    }
}

; Create Default Shortcuts


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
        MsgBox('An error occured while preparing the unpacker!', 'Oops!', '48')
        ; Run installation help page
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
    MsgBox('An error occured while preparing the necessary files!', 'Oops!', '48')
    ; Run installation help page
    ExitApp
}
Prepare.Hide()

; Main window
Manager := Gui(, 'AoE II Manager AIO')
Manager.OnEvent('Close', (*) => ExitApp())

; Features
Features := Map()

; # The Game
Features['The Game'] := []
_Game_ := Manager.AddText('xm ym w220 h260 Center c800000 BackgroundFFFFFF Border', '# The Game')
Features['The Game'].Push(_Game_)
_Game_.SetFont('Bold')
GetTheGame := Manager.AddButton('xm+10 ym+25 w200 h21', 'Download AoE II')
Features['The Game'].Push(GetTheGame)
GetTheGame.OnEvent('Click', (*) => DownloadInstallGame())
CreateImageButton(GetTheGame, 0, [['DB\000\download_aoeii_normal.png' ]
                                , ['DB\000\download_aoeii_hover.png'  ]
                                , ['DB\000\download_aoeii_click.png'  ]
                                , ['DB\000\download_aoeii_disable.png']]*)
ProgressBar := Manager.AddProgress('xp yp wp h20 Hidden cFF9427 Background804000', 0)
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
    If FileExist(ChosenFolder.Value '\empires2.exe')
        FileCreateShortcut(ChosenFolder.Value '\empires2.exe', A_Desktop '\Age of Empires II.lnk')
    If FileExist(ChosenFolder.Value '\age2_x1\age2_x1.exe')
        FileCreateShortcut(ChosenFolder.Value '\age2_x1\age2_x1.exe', A_Desktop '\The Conquerors.lnk')
    If FileExist(ChosenFolder.Value '\age2_x1\age2_x2.exe')
        FileCreateShortcut(ChosenFolder.Value '\age2_x1\age2_x2.exe', A_Desktop '\Forgotten Empires.lnk')
}
RunAOK := Manager.AddButton('xm+50 yp+20 w36 H36')
Features['The Game'].Push(RunAOK)
CreateImageButton(RunAOK, 0, [['DB\000\aok_normal.png'  ]
                            , ['DB\000\aok_hover.png'   ]
                            , ['DB\000\aok_click.png'   ]
                            , ['DB\000\aok_disable.png' ]]*)
RunAOK.OnEvent('Click', (*) => Run(ChosenFolder.Value '\empires2.exe', ChosenFolder.Value))

RunAOC := Manager.AddButton('yp wp hp')
Features['The Game'].Push(RunAOC)
CreateImageButton(RunAOC, 0, [['DB\000\aoc_normal.png'  ]
                            , ['DB\000\aoc_hover.png'   ]
                            , ['DB\000\aoc_click.png'   ]
                            , ['DB\000\aoc_disable.png' ]]*)
RunAOC.OnEvent('Click', (*) => Run(ChosenFolder.Value '\age2_x1\age2_x1.exe', ChosenFolder.Value '\age2_x1'))

RunFOE := Manager.AddButton('yp wp hp')
Features['The Game'].Push(RunFOE)
CreateImageButton(RunFOE, 0, [['DB\000\fe_normal.png'  ]
                            , ['DB\000\fe_hover.png'   ]
                            , ['DB\000\fe_click.png'   ]
                            , ['DB\000\fe_disable.png' ]]*)
RunFOE.OnEvent('Click', (*) => Run(ChosenFolder.Value '\age2_x1\age2_x2.exe', ChosenFolder.Value '\age2_x1'))

ChooseFolder := Manager.AddButton('xm+10 yp+50 w100 h21', 'Choose')
Features['The Game'].Push(ChooseFolder)
ChooseFolder.OnEvent('Click', (*) => SelectTheGame())
CreateImageButton(ChooseFolder, 0, [['DB\000\pick_folder_normal.png'  ]
                                  , ['DB\000\pick_folder_hover.png'   ]
                                  , ['DB\000\pick_folder_click.png'   ]
                                  , ['DB\000\pick_folder_disable.png' ]]*)

LoadGRFolder := Manager.AddButton('xm+180 yp w30 h21')
Features['The Game'].Push(LoadGRFolder)
LoadGRFolder.OnEvent('Click', (*) => SelectTheGameFromGR())
CreateImageButton(LoadGRFolder, 0, [['DB\000\gr_get_normal.png'  ]
                                  , ['DB\000\gr_get_hover.png'   ]
                                  , ['DB\000\gr_get_click.png'   ]
                                  , ['DB\000\gr_get_disable.png' ]]*)
SelectTheGameFromGR() {
    TextFound := LoadGRSettingText()
    LoadGRSettingText() {
        Setting := FileRead(GRSetting, 'RAW')
        TextFound := ''
        Loop Setting.Size {
            Byte := NumGet(Setting, A_Index - 1, 'UChar')
            If (32 <= Byte && Byte <= 126) || (Byte = 10) || (Byte = 13)
                TextFound .= Chr(Byte)
        }
        Return TextFound
    }
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
ChosenFolder := Manager.AddEdit('xm+10 yp+30 w200 Center ReadOnly r4 -VScroll Border -E0x200 BackgroundWhite cBlue')
Features['The Game'].Push(ChosenFolder)
ChosenFolder.SetFont('Bold')
OpenTheGameFolder := Manager.AddButton('w200 h21', 'Open')
Features['The Game'].Push(OpenTheGameFolder)
CreateImageButton(OpenTheGameFolder, 0, [['DB\000\open_aoeii_normal.png' ]
                                       , ['DB\000\open_aoeii_hover.png'  ]
                                       , ['DB\000\open_aoeii_click.png'  ]
                                       , ['DB\000\open_aoeii_disable.png']]*)
OpenTheGameFolder.OnEvent('Click', (*) => Run(ChosenFolder.Value))
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

; # App
Features['App'] := []
_App_ := Manager.AddText('xm+230 ym w40 h280 Center c800000 BackgroundFFFFFF Border', '[ A ]')
Features['App'].Push(_App_)
_App_.SetFont('Bold')
OpenDB := Manager.AddButton('xp+5 yp+25 w30 h21')
Features['App'].Push(OpenDB)
OpenDB.SetFont('Bold')
CreateImageButton(OpenDB, 0, [['DB\000\open_db_normal.png' ]
                            , ['DB\000\open_db_hover.png'  ]
                            , ['DB\000\open_db_click.png'  ]
                            , ['DB\000\open_db_disable.png']]*)
OpenDB.OnEvent('Click', (*) => Run(AppDir[1]))
OpenSetting := Manager.AddButton('xp yp+25 w30 h21')
Features['App'].Push(OpenSetting)
OpenSetting.SetFont('Bold')
CreateImageButton(OpenSetting, 0, [['DB\000\open_setting_normal.png' ]
                                 , ['DB\000\open_setting_hover.png'  ]
                                 , ['DB\000\open_setting_click.png'  ]
                                 , ['DB\000\open_setting_disable.png']]*)
OpenSetting.OnEvent('Click', (*) => Run(AppDir[2]))
EditSetting := Manager.AddButton('xp yp+25 w30 h21')
Features['App'].Push(EditSetting)
EditSetting.SetFont('Bold')
CreateImageButton(EditSetting, 0, [['DB\000\open_opt_normal.png' ]
                                 , ['DB\000\open_opt_hover.png'  ]
                                 , ['DB\000\open_opt_click.png'  ]
                                 , ['DB\000\open_opt_disable.png']]*)
EditSetting.OnEvent('Click', (*) => ShowOption())
AppOption := Gui(, 'Options')
UpdateChk := AppOption.AddCheckbox('w300', '  Check for updates when the app starts')
Check := IniRead(Config, 'Game', 'UpdateChk', 0)
If Check {
    UpdateChk.Value := 1
}
UpdateChk.OnEvent('Click', (*) => IniWrite(UpdateChk.Value, Config, 'Game', 'UpdateChk'))
ShowOption() {
    AppOption.Show()
}

; # Versions
Features['Versions'] := []
; # Compatibilities
Features['Compatibilities'] := []
_Version_ := Manager.AddText('xm+280 ym w400 h280 Center c800000 BackgroundFFFFFF Border', '# Versions')
Features['Versions'].Push(_Version_)
_Version_.SetFont('Bold')
H := Manager.AddPicture('xp+54 ym+25 BackgroundTrans', 'DB\000\aok.png')
Features['Versions'].Push(H)
H := Manager.AddText('xp-44 yp+40 cRed w120 Center BackgroundTrans', 'The Age of Kings')
Features['Versions'].Push(H)
H.SetFont('Bold')
AoKCom := Manager.AddDropDownList('w120')
Features['Compatibilities'].Push(AoKCom)
For Each, Compat in Compatibilities {
    AoKCom.Add([Compat[1]])
}
AoKCom.Choose(1)
AoKCom.OnEvent("Change", (*) => AoKComReg())
AoKRun := Manager.AddCheckbox('xp yp+30 wp hp BackgroundFFFFFF', 'Run as administrator')
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
    Handle := Manager.AddRadio('w30 w100 BackgroundFFFFFF', A_LoopFileName)
    Features['Versions'].Push(Handle)
    Handle.SetFont('s10', 'Consolas')
    Handle.OnEvent('Click', ApplyVersion)
    General['AOK']['VersionsN'][A_LoopFileName] := Handle
}
ApplyVersion(Ctrl, Info) {
    SectionInteract(Features['Versions'], False)
    SectionInteract(Features['The Game'], False)
    If GameIsRunning() {
        ChargeSettings________()
        Return
    }
    Try {
        CleanUp(Ctrl.Text)
        SetVersion(Ctrl.Text)
    } Catch As Err {
        MsgBox('An error occured while trying to set v' Ctrl.Text, 'Version apply error!', 0x20)
    }
    SectionInteract(Features['The Game'])
    SectionInteract(Features['Versions'])
    SoundPlay('DB\000\30 wololo.mp3')
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
    If (Patch.Value = 1) {
        Return
    }
    DirCopy('DB\001\' Patch.Text '\Static', ChosenFolder.Value, 1)
    If DirExist('DB\001\' Patch.Text '\' Version) {
        DirCopy('DB\001\' Patch.Text '\' Version, ChosenFolder.Value, 1)
        If InStr(Patch.Text, 'v2') {
            RegWrite(2, 'REG_DWORD', 'HKEY_CURRENT_USER\SOFTWARE\Microsoft\Microsoft Games\Age of Empires', 'Aoe2Patch')
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
H := Manager.AddPicture('xp+174 ym+25 BackgroundTrans', 'DB\000\aoc.png')
Features['Versions'].Push(H)
H := Manager.AddText('xp-44 yp+40 cBlue w120 Center BackgroundTrans', 'The Conquerors')
Features['Versions'].Push(H)
H.SetFont('Bold')
AoCCom := Manager.AddDropDownList('w120')
Features['Compatibilities'].Push(AoCCom)
For Each, Compat in Compatibilities {
    AoCCom.Add([Compat[1]])
}
AoCCom.Choose(1)
AoCCom.OnEvent("Change", (*) => AoCComReg())
AoCRun := Manager.AddCheckbox('xp yp+30 wp hp BackgroundFFFFFF', 'Run as administrator')
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
    Handle := Manager.AddRadio('w30 w100 BackgroundFFFFFF', A_LoopFileName)
    Features['Versions'].Push(Handle)
    Handle.SetFont('s10', 'Consolas')
    Handle.OnEvent('Click', ApplyVersion)
    General['AOC']['VersionsN'][A_LoopFileName] := Handle
}
H := Manager.AddPicture('xp+174 ym+25 BackgroundTrans', 'DB\000\fe.png')
Features['Versions'].Push(H)
H := Manager.AddText('xp-44 yp+40 cGreen w120 Center BackgroundTrans', 'Forgotten Empires')
Features['Versions'].Push(H)
H.SetFont('Bold')
FOECom := Manager.AddDropDownList('w120')
Features['Compatibilities'].Push(FOECom)
For Each, Compat in Compatibilities {
    FOECom.Add([Compat[1]])
}
FOECom.Choose(1)
FOECom.OnEvent("Change", (*) => FOEComReg())
FOERun := Manager.AddCheckbox('xp yp+30 wp hp BackgroundFFFFFF', 'Run as administrator')
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
Handle := Manager.AddRadio('w30 w100 Checked BackgroundFFFFFF', '2.2  CD')
Features['Versions'].Push(Handle)
Handle.SetFont('s10', 'Consolas')
General['FOE']['VersionsN']['2.2  CD'] := Handle
Patch := Manager.AddDropDownList('xm+290 ym+250 w380', ['Do Not Enable Fixes'])
Features['Versions'].Push(Patch)
Patch.OnEvent('Change', (*) => IniWrite(Patch.Text, Config, 'Game', 'Fix'))
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

; # Language
Features['Language'] := []
_Language_ := Manager.AddText('xm ym+270 w220 h385 Center c800000 BackgroundFFFFFF Border', '# Languages')
Features['Language'].Push(_Language_)
_Language_.SetFont('Bold')
H := Manager.AddText('xp+10 yp w200 BackgroundTrans')
Features['Language'].Push(H)
Loop Files, 'DB\006\*', 'D' {
    Handle := Manager.AddRadio('wp Center BackgroundFFFFFF', A_LoopFileName)
    Features['Language'].Push(Handle)
    Handle.SetFont('Underline Bold')
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

; # Visual Modes
Features['Visual Modes'] := []
_VisualMods_ := Manager.AddText('xm+230 ym+290 w220 h365 Center c800000 BackgroundFFFFFF Border', '# Visual Mods')
Features['Visual Modes'].Push(_VisualMods_)
_VisualMods_.SetFont('Bold')
H := Manager.AddText('xp+10 yp+10 w200 BackgroundTrans')
Features['Visual Modes'].Push(H)
VMList := Manager.AddListView('w200 h320 -E0x200 -Hdr Checked', ['Mode Name'])
Features['Visual Modes'].Push(VMList)
VMList.SetFont('Bold')
Loop Files, 'DB\007\*', 'D' {
    VMList.Add(, A_LoopFileName)
}
VMList.OnEvent('ItemCheck', ApplyVM)
ApplyVM(Ctrl, Item, Checked) {
    SectionInteract(Features['Visual Modes'], False)
    VMName := VMList.GetText(Item)
    SlpDir := Checked ? 'DB\007\' VMName : 'DB\007\' VMName '\U'
    RunWait('DB\000\DrsBuild.exe /a "' ChosenFolder.Value '\Data\' DrsTypes['gra'] '" "' SlpDir '\gra*.slp"',, 'Hide')
    RunWait('DB\000\DrsBuild.exe /a "' ChosenFolder.Value '\Data\' DrsTypes['int'] '" "' SlpDir '\int*.slp"',, 'Hide')
    RunWait('DB\000\DrsBuild.exe /a "' ChosenFolder.Value '\Data\' DrsTypes['ter'] '" "' SlpDir '\ter*.slp"',, 'Hide')
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
    Hash := HashFile(ChosenFolder.Value '\Data\' DrsTypes['gra'])
          . HashFile(ChosenFolder.Value '\Data\' DrsTypes['int'])
          . HashFile(ChosenFolder.Value '\Data\' DrsTypes['ter'])
    CheckedRows := Map(), NCR := 0
    While (NCR := VMList.GetNext(NCR, 'Checked')) {
        CheckedRows[NCR] := True
    }
    Stat := ''
    Loop VMList.GetCount() {
        IniWrite(CheckedRows.Has(A_Index), Config, Hash, VMList.GetText(A_Index))
    }
    SectionInteract(Features['Visual Modes'])
    SoundPlay('DB\000\30 wololo.mp3')
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
        VMList.Add(, ModeName)
        MsgBox(ModeName ' should be added to the list by now!', 'Info', 0x40)
        SectionInteract(Features['Visual Modes'])
    }
}

; # Data Modes
Features['Data Modes'] := []
_VisualMods_.GetPos(, &Y)
_DataModes_ := Manager.AddText('xm+460 y' Y ' w220 h365 Center c800000 BackgroundFFFFFF Border', '# Data Mods')
Features['Data Modes'].Push(_DataModes_)
_DataModes_.SetFont('Bold')
H := Manager.AddText('xp+10 yp+10 w200 BackgroundTrans')
Features['Data Modes'].Push(H)
DMList := Manager.AddListView('w200 h240 -E0x200 -Hdr Checked BackgroundFFFFFF', ['Mode Name'])
Features['Data Modes'].Push(DMList)
DMList.SetFont('Bold')
For Each, Mode in StrSplit(IniRead('DB\008\DataMode.ini', 'DataMode', , ''), '`n') {
    DMList.Add(, StrSplit(Mode, '=')[1])
}
DMList.OnEvent('ItemCheck', ApplyDM)
ApplyDM(Ctrl, Item, Checked) {
    DMName := DMList.GetText(Item)
    If GameIsRunning() {
        DMList.Modify(Item, '-Check')
        Return
    }
    If (Checked) {
        ModeDir := IniRead('DB\008\DataMode.ini', 'DataMode', DMName, '')
        ModeDir := StrSplit(ModeDir, '|')
        Parts := StrSplit(ModeDir[2], ',')
        If !PrepareTheDataMode() {
            DMList.Modify(Item, '-Check')
            Return
        }
        PrepareTheDataMode() {
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
        CleanUp('1.5  CD')
        SetVersion('1.5  CD', 'AOC')
        SectionInteract(Features['Data Modes'], False)
        If !DirExist(ChosenFolder.Value '\Games') {
            DirCreate(ChosenFolder.Value '\Games')
        }
        If FileExist(ChosenFolder.Value '\Games\age2_x1.xml')
            FileDelete(ChosenFolder.Value '\Games\age2_x1.xml')
        Loop Files, 'DB\' ModeDir[1] '\*.*', 'R' {
            GameFileDir := ChosenFolder.Value SubStr(A_LoopFileDir, StrLen('DB\' ModeDir[1]) + 1)
            GameFile := GameFileDir '\' A_LoopFileName
            If FileExist(GameFile) {
                GameFileMD5 := HashFile(GameFile)
                CurrentMD5 := HashFile(A_LoopFileFullPath)
                If (GameFileMD5 != CurrentMD5) {
                    FileCopy(A_LoopFileFullPath, GameFile, 1)
                }
            } Else {
                If !DirExist(GameFileDir) {
                    DirCreate(GameFileDir)
                }
                FileCopy(A_LoopFileFullPath, GameFile)
            }
        }
        IniWrite(DMName, Config, 'Game', 'CurrDM')
    } Else {
        CleanUp('1.5  CD')
        SetVersion('1.5  CD', 'AOC')
        If FileExist(ChosenFolder.Value '\Games\age2_x1.xml') {
            FileDelete(ChosenFolder.Value '\Games\age2_x1.xml')
        }
        If DMName = 'Sheep vs Wolf 2' {
            DirCopy('DB\008\Sound\stream', ChosenFolder.Value '\Sound\stream', 1)
        }
        IniDelete(Config, 'Game', 'CurrDM')
    }
    SectionInteract(Features['Data Modes'])
    SoundPlay('DB\000\30 wololo.mp3')
}
VMDM := Gui(, 'Customize')
VMDM.BackColor := 'White'
VMDMTitle := VMDM.AddText('w220 h280 Center c800000 BackgroundFFFFFF', '# Mode Name')
VMDMTitle.SetFont('Bold')
VMDM.AddText('xp+10 yp+20 w200 Center cBlue', '1 - Visual modes').SetFont('Bold')
VMDMList := VMDM.AddListView('w200 h240 -E0x200 -Hdr Checked BackgroundFFFFFF', ['Mode Name'])
VMDMList.SetFont('Bold')
Loop Files, 'DB\007\*', 'D' {
    VMDMList.Add(, A_LoopFileName)
}
DMList.OnEvent('DoubleClick', ShowVMDMList)
ShowVMDMList(Ctrl, Item) {
    DMName := DMList.GetText(Item)
    If !DirExist(ChosenFolder.Value '\Games\' DMName) {
        Return
    }
    VMDMTitle.Text := DMName
    VMDM.Show()
}
VMDMList.OnEvent('ItemCheck', ApplyVMDM)
ApplyVMDM(Ctrl, Item, Checked) {
    SectionInteract(Features['Data Modes'], False)
    VMName := VMList.GetText(Item)
    SlpDir := Checked ? 'DB\007\' VMName : 'DB\007\' VMName '\U'
    RunWait(A_Clipboard := 'DB\000\DrsBuild.exe /a "' ChosenFolder.Value '\Games\' VMDMTitle.Text '\Data\gamedata_x1_p1.drs" "' SlpDir '\*.slp"', , 'Hide')
    SectionInteract(Features['Data Modes'])
    SoundPlay('DB\000\30 wololo.mp3')
}
;ImportDM := Manager.AddButton('wp Disabled', 'Import')
;ImportDM.SetFont('Bold')
;ImportDM.OnEvent('Click', (*) => ImportDataMode())
ImportDataMode() {
    If Selected := FileSelect('D') {
        SplitPath(Selected, &ModeName)
        If IniRead('DB\008\DataMode.ini', 'DataMode', ModeName, '') {
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
        SectionInteract(Features['Data Modes'], False)
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
        IniWrite(DID, 'DB\008\DataMode.ini', 'DataMode', ModeName)
        SectionInteract(Features['Data Modes'])
    }
}

; # Other tools
Features['Other Tools'] := []
_ATools_ := Manager.AddText('xm+690 ym w220 h650 Center c800000 BackgroundFFFFFF Border', '# Other Tools`n`n')
Features['Other Tools'].Push(_ATools_)
_ATools_.SetFont('Bold')
_ATools_.GetPos(&X, &Y, &Width, &Height)

H := Manager.AddText('xp+10 yp+10 w200 BackgroundTrans')
Features['Other Tools'].Push(H)

; # Shortcuts
Shortcuts := Manager.AddButton('xp yp+20 w200 Center', 'Hotkeys')
;Macro1 := Manager.AddCheckBox('xp yp+20 w200 BackgroundWhite Center', '{Left Alt + Right Mouse Button}`n[Send && un-select one unit]')
Features['Other Tools'].Push(Shortcuts)
CreateImageButton(Shortcuts, 0, [['DB\000\hotkey_normal.png' ]
                               , ['DB\000\hotkey_hover.png'  ]
                               , ['DB\000\hotkey_click.png'  ]
                               , ['DB\000\hotkey_disable.png']]*)
Shortcuts.SetFont('Bold')
ShortcutsG := Gui(, 'Defined Hotkeys')
ShortcutList := ShortcutsG.AddListView('w305 r10 Checked', ['Hotkey', 'Comment', 'ID'])
ShortcutList.ModifyCol(3, 0)
ShortcutList.ModifyCol(4, 0)
ShortcutList.ModifyCol(1, 100)
ShortcutList.ModifyCol(2, 200)
ShortcutList.SetFont('Bold')
ShortcutAdd := ShortcutsG.AddButton('w150', 'Add')
ShortcutAdd.OnEvent('Click', (*) => AddShortcut())
ShortcutsGA := Gui(, 'Add A Shortcut')
ShortcutsGA.AddText(, '1 - Hotkey Name')
ShortcutName := ShortcutsGA.AddEdit('w300 cRed')
ShortcutName.SetFont('Bold')
ShortcutsGA.AddText(, '2 - Hotkey Comment')
ShortcutComment := ShortcutsGA.AddEdit('w300 cGreen')
ShortcutComment.SetFont('Bold')
ShortcutsGA.AddText(, '3 - Hotkey Action')
ShortcutAction := ShortcutsGA.AddEdit('w300 r10 cBlue HScroll')
ShortcutAction.SetFont('Bold')
ShortcutAddOK := ShortcutsGA.AddButton('w300', 'Submit')
ShortcutAddOK.OnEvent('Click', (*) => SaveShortcut())
SaveShortcut() {
    If ShortcutName.Value = '' || ShortcutAction.Value = '' || ShortcutComment.Value = '' {
        MsgBox('One of the inputs is not being filled!', 'Fill!', 0x30)
        Return
    }
    Loop {
        Index := Format('{:03}', A_Index) '.ahk'
    } Until !FileExist(AppDir[3] '\' Index)
    O := FileOpen(AppDir[3] '\' Index, 'w')
    O.WriteLine(';' ShortcutComment.Value ';')
    O.WriteLine('#Requires AutoHotkey v2')
    O.WriteLine('#SingleInstance Force')
    O.WriteLine("GroupAdd('AOKAOC', 'ahk_exe empires2.exe')")
    O.WriteLine("GroupAdd('AOKAOC', 'ahk_exe age2_x1.exe')")
    O.WriteLine("GroupAdd('AOKAOC', 'ahk_exe age2_x2.exe')")
    O.WriteLine('HotIfWinActive("ahk_group AOKAOC")')
    O.WriteLine("Hotkey('" ShortcutName.Value "', Action)")
    O.WriteLine("Action(*) {")
    O.WriteLine(ShortcutAction.Value)
    O.WriteLine('}')
    O.WriteLine('ProcessWaitClose(A_Args[1])')
    O.Write('ExitApp')
    O.Close()
    ShortcutList.Add('Check', ShortcutName.Value, ShortcutComment.Value, Index)
    IniWrite(1, Config, 'Hotkey', ShortcutName.Value)
    ShortcutName.Value := ''
    ShortcutAction.Value := ''
    ShortcutsGA.Hide()
}
AddShortcut() {
    ShortcutsGA.Show()
}
ShortcutRemove := ShortcutsG.AddButton('yp w150', 'Remove')
ShortcutRemove.OnEvent('Click', (*) => RemoveShortcut())
RemoveShortcut() {
    If !R := ShortcutList.GetNext() {
        Return
    }
    If R = 1 {
        Msgbox('You can only un-check this shortcut!', 'Unable to delete!', 0x30)
        Return
    }
    ShortcutFile := ShortcutList.GetText(R, 3)
    If FileExist(AppDir[3] '\' ShortcutFile)
        FileDelete(AppDir[3] '\' ShortcutFile)
    ShortcutList.Delete(R)
}
Shortcuts.OnEvent('Click', (*) => ShowShortcut())
ShowShortcut() {
    ShortcutsG.Show()
}
LoadShortcuts()
LoadShortcuts() {
    Loop Files, AppDir[3] '\*.ahk' {
        SScript := FileRead(A_LoopFileFullPath)
        If !RegExMatch(SScript, ";(.*);", &ShortcutComment)
            Continue
        ShortcutComment := ShortcutComment.1
        If !RegExMatch(SScript, "Hotkey\('(.*)'\, Action\)", &ShortcutName)
            Continue
        ShortcutName := ShortcutName.1
        Item := ShortcutList.Add(, ShortcutName, ShortcutComment, A_LoopFileName)
        Checked := IniRead(Config, 'Hotkey', ShortcutName, 0)
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
H := Manager.AddText('x' (X + 1) ' yp+35 w220 0x10')
VPN := Manager.AddButton('xp+10 yp+10 w56 h56')
Features['Other Tools'].Push(VPN)
CreateImageButton(VPN, 0, [['DB\000\vpn_normal.png'  ]
                         , ['DB\000\vpn_hover.png'   ]
                         , ['DB\000\vpn_click.png'   ]
                         , ['DB\000\vpn_disable.png' ]]*)
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
ClearVPNReg := Manager.AddButton('xp+65 yp+2 w130', 'Clear Registry')
CreateImageButton(ClearVPNReg, 0, [['DB\000\clear_vpn_normal.png'  ]
                                 , ['DB\000\clear_vpn_hover.png'   ]
                                 , ['DB\000\clear_vpn_click.png'   ]
                                 , ['DB\000\clear_vpn_disable.png' ]]*)
Features['Other Tools'].Push(ClearVPNReg)
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
Features['Other Tools'].Push(VPNCompat)
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
H := Manager.AddText('x' (X + 1) ' yp+35 w220 0x10')
RecordFix := Manager.AddButton('xp+10 yp+10 w200 Center', '(.mgx/.mgl) Records biegleux Fixes')
Features['Other Tools'].Push(RecordFix)
CreateImageButton(RecordFix, 0, [['DB\000\hotkey_normal.png' ]
                               , ['DB\000\hotkey_hover.png'  ]
                               , ['DB\000\hotkey_click.png'  ]
                               , ['DB\000\hotkey_disable.png']]*)
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
    Global Cancel := False
    RecordFixG.Show()
    Records := []
    Loop Files, ChosenFolder.Value '\SaveGame\*.mg*' {
        If (A_LoopFileExt = 'MGX' || A_LoopFileExt = 'MGL') && !InStr(A_LoopFileName, 'aoeii_aio_fix')
            Records.Push(A_LoopFileFullPath)
    }
    RecordFixText.Text := 0 ' / ' Records.Length
    RecordFixProgress.Value := 0
    RecordFixProgress.Opt('Range1-' Records.Length)
    For Each, Record in Records {
        If Cancel {
            Break
        }
        RunWait('DB\000\MgxFix.exe -f "' Record '"',, 'Hide')
        RunWait('DB\000\RevealFix.exe "' Record '"',, 'Hide')
        SplitPath(Record,, &Dir, &Ext, &Name)
        FileMove(Record, Dir '\' Name '_aoeii_aio_fix.' Ext)
        RecordFixText.Text := Each ' / ' Records.Length
        RecordFixProgress.Value += 1
    }
    RecordsCheck__________()
    RecordFixG.Hide()
}
RecordCount := Manager.AddText('xp yp+30 w200 BackgroundTrans', '0 Records Found')
RecordCount.SetFont('Bold')
Features['Other Tools'].Push(RecordCount)
CountRecords__________() {
    Records := 0
    Loop Files, ChosenFolder.Value '\SaveGame\*.mg*' {
        If (A_LoopFileExt = 'MGX' || A_LoopFileExt = 'MGL') {
            Records += 1
            RecordCount.Text := Records ' Records Found'
        }
    }
}
RecordFixed := Manager.AddText('xp yp+20 w200 BackgroundTrans cGreen', '0 Records Processed ✓')
RecordFixed.SetFont('Bold')
Features['Other Tools'].Push(RecordFixed)
CountFixedRecords_____() {
    Records := 0
    Loop Files, ChosenFolder.Value '\SaveGame\*.mg*' {
        If (A_LoopFileExt = 'MGX' || A_LoopFileExt = 'MGL') && InStr(A_LoopFileName, 'aoeii_aio_fix') {
            Records += 1
            RecordFixed.Text := Records ' Records Processed ✓'
        }
    }
}
RecordUnknown := Manager.AddText('xp yp+20 w200 BackgroundTrans cRed', '0 Records Not Processed X')
RecordUnknown.SetFont('Bold')
Features['Other Tools'].Push(RecordUnknown)
CountUnknownRecords___() {
    Records := 0
    Loop Files, ChosenFolder.Value '\SaveGame\*.mg*' {
        If (A_LoopFileExt = 'MGX' || A_LoopFileExt = 'MGL') && !InStr(A_LoopFileName, 'aoeii_aio_fix') {
            Records += 1
            RecordUnknown.Text := Records ' Records Not Processed X'
        }
    }
}
H := Manager.AddText('x' (X + 1) ' yp+25 w220 0x10')

;
;Manager.AddText('yp+40 cBlue w200 BackgroundTrans Center', '3 - Repair Game Files').SetFont('Bold')
;RepairGame := Manager.AddButton('wp', 'Repair')
;RepairGame.OnEvent('Click', (*) => MsgBox('Not Yet Implemented!', 'Hoy!', 0x40))
;
;Manager.AddText('yp+40 cBlue w200 BackgroundTrans Center', '5 - Scenario Files Select').SetFont('Bold')
;FixMgz := Manager.AddButton('wp', 'Select')
;FixMgz.OnEvent('Click', (*) => MsgBox('Not Yet Implemented!', 'Hoy!', 0x40))

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

; '| AGE OF EMPIRES II MANAGER ALL IN ONE, '
; 'AUTOHOTKEY BASED APP, '
; 'CREATED BY SMILE, '
; 'TOTALLY SECURE, '
; 'TESTED MANY TIMES, '
; 'BUT USE ON YOUR OWN RISK, '
; 'ANY FEEDBACK WILL BE HELPFUL, '
; 'MY EMAIL, '
; 'CHANDOUL.MOHAMED26@GMAIL.COM, '
; 'WEBSITE FOR THIS APP, '
; 'HTTPS://SMILEAOE.GITHUB.IO |'

SB := Manager.AddStatusBar()
SB.SetFont('Bold', 'Calibri')
SB.SetParts(10, 50, 200)
SB.SetText('v' Version, 2)
SB.SetText('Loading...', 3)
SB.SetText(A_Tab A_Tab 'A Collective App From The Internet On What I Found Useful About AoE II!    ', 4)
Manager.Show('x10 y10 w930 h690')
ChargeEnableFixes_____() {
    Loop Files, 'DB\001\*', 'D' {
        Patch.Add([A_LoopFileName])
    }
    Patch.Choose('Do Not Enable Fixes')
    If DirExist('DB\001\' Fix := IniRead(Config, 'Game', 'Fix', 'Do Not Enable Fixes')) {
        Patch.Choose(Fix)
    }
}
ChargeEnableFixes_____()
ChargeSettings________()
CheckForUpdates_______()
Return

ChargeSettings________(Browse := False) {
    SoundPlay('DB\000\30 wololo.mp3')
    SectionInteract(Features['Versions'], False)
    SectionInteract(Features['Compatibilities'], False)
    SectionInteract(Features['Language'], False)
    SectionInteract(Features['Visual Modes'], False)
    SectionInteract(Features['Data Modes'], False)
    SectionInteract(Features['Other Tools'], False)
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
    SectionInteract(Features['Data Modes'])
    SectionInteract(Features['Other Tools'])
    ChargeVersions________()
    ChargeCompatibilities_()
    ChargeLanguage________()
    ChargeVModes__________()
    ChargeDModes__________()
    BackupDefaultLanguage_()
    ChargeOtherTools______()
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

RecordsCheck__________() {
    CountRecords__________()
    CountFixedRecords_____()
    CountUnknownRecords___()
}

ChargeDModes__________() {
    Loop DMList.GetCount() {
        DMList.Modify(A_Index, '-Check')
    }
    CurrDMDir := IniRead(Config, 'Game', 'CurrDMDir', '')
    If ChosenFolder.Value != CurrDMDir {
        Return
    }
    Name := IniRead(Config, 'Game', 'CurrDM', '')
    Loop DMList.GetCount() {
        If DMList.GetText(A_Index) = Name {
            DMList.Modify(A_Index, 'Check')
            Break
        }
    }
}
ChargeVModes__________() {
    VMList.Enabled := True
    Hash := HashFile(ChosenFolder.Value '\Data\' DrsTypes['gra'])
        . HashFile(ChosenFolder.Value '\Data\' DrsTypes['int'])
        . HashFile(ChosenFolder.Value '\Data\' DrsTypes['ter'])
    Loop VMList.GetCount() {
        VMList.Modify(A_Index, '-Check')
    }
    If Values := IniRead(Config, Hash, , '') {
        For Each, Value in StrSplit(Values, '`n') {
            If StrSplit(Value, '=')[2]
                VMList.Modify(Each, '+Check')
        }
    }
}
CheckForUpdates_______() {
    If A_IsCompiled {
        Return
    }
    SB.SetText('Update check is disabled!', 3)
    UpdateChk := IniRead(Config, 'Game', 'UpdateChk', 0)
    If !UpdateChk {
        Return
    }
    Try {
        SB.SetText('Checking for updates...', 3)
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
            SB.SetText('Update found!', 3)
            Choice := MsgBox('The following needs to be updated:`n' UpdatesList '`nUpdate now?', 'New Update!', 0x4 + 0x40 + 0x100)
            If Choice = 'Yes' {
                DoneSteps.Value := 0
                DoneSteps.Opt('Range1-' FoundUpdates.Length)
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
                    Str := StrGet(Buff, 2, '')
                    If (Str != '7z') {
                        Continue
                    }
                    Name := StrSplit(UpdateFile, '.')
                    DirDelete(Name[1], 1)
                    RunWait('DB\7za.exe x ' Name[1] '.7z.001 -o' Name[1] ' -aoa', , 'Hide')
                    DoneSteps.Value += 1
                    DoneStepsText.Text := UpdateFile
                }
                Reload
            }
            SB.SetText('New update is waiting!', 3)
        } Else {
            SB.SetText('Up to date!', 3)
        }
    } Catch As Err {
        SB.SetText('Failed to check for updates!', 3)
    }
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