#Requires AutoHotkey v2
#SingleInstance Force
CoordMode('Mouse', 'Screen')

If !A_IsAdmin {
    MsgBox("- This application is not being ran as administrator`n"
         . "- This can cause an unexpected behaviour on using any of it's options`n"
         . "- It is highly recommended that you run it as administrator"
         , 'Warning', 0x30)
}

Server := 'https://raw.githubusercontent.com'
User := 'SmileAoE'
Repo := 'aoeii_aio'
Version := '1.7'
Layers := 'HKLM\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers'
Config := A_AppData '\aoeii_aio\config.ini'
AppDir := ['DB', A_AppData '\aoeii_aio']
GRSetting := A_AppData '\GameRanger\GameRanger Prefs\Settings'
DrsTypes := Map('gra', 'graphics.drs', 'int', 'interfac.drs', 'ter', 'terrain.drs')
DrsRange := Map('gra', [2, 5312], 'int', [50100, 53211], 'ter', [15000, 15031])
IDL := 5
VCodedSlp := '3713EFBE'
NormalSlp := '322E304E'

SysDrive := EnvGet('SystemDrive')
;ProgramFilesDir := EnvGet(A_Is64bitOS ? "ProgramW6432" : "ProgramFiles")

General := Map()

General['AOK'] := Map()
General['AOK']['VersionsN'] := Map()
General['AOK']['Combine'] := Map('2.0b CD', ['2.0a No CD'])

General['AOC'] := Map()
General['AOC']['VersionsN'] := Map()
General['AOC']['Combine'] := Map( '1.0e No CD', ['1.0c No CD']
                                , '1.0e No CD', ['1.0c No CD']
                                , '1.1  No CD', ['1.0c No CD']
                                , '1.5  CD'   , ['1.0c No CD'])

General['FOE'] := Map()
General['FOE']['VersionsN'] := Map()
General['FOE']['Combine'] := Map()

General['LNG'] := Map()

Compatibilities := Map(1, ["_____Not Set_____", ""]
    , 2, ["Windows 8", "WIN8RTM"]
    , 3, ["Windows 7", "WIN7RTM"]
    , 4, ["Windows Vista Sp2", "VISTASP2"]
    , 5, ["Windows Vista Sp1", "VISTASP1"]
    , 6, ["Windows Vista", "VISTARTM"]
    , 7, ["Windows XP Sp2", "WINXPSP2"]
    , 8, ["Windows 98", "WIN98"]
    , 9, ["Windows 95", "WIN95"])

BasePackages := ['DB/7za.exe'
               , 'DB/000.7z.001'
               , 'DB/001.7z.001'
               , 'DB/002.7z.001'
               , 'DB/006.7z.001'
               , 'DB/007.7z.001']
Dots := 0
Task := 1
TaskNumber := BasePackages.Length

Try {
    Prepare := Gui('-MinimizeBox', 'Preparing...')
    Prepare.OnEvent('Close', (*) => ExitApp())
    HoldOn := Prepare.AddText('Center w400 h40', PW := 'Please Wait')
    HoldOn.SetFont('s12 Bold')
    DoneSteps := Prepare.AddText('Center w400 h30 cRed')
    DoneSteps.SetFont('Bold')
    Prepare.Show()

    ShowInfo() {
        Global Dots, Task
        HoldOn.Text := PW '`n'  ((Mod(++Dots, 4) = 0) ? '●'
                                : ((Mod(Dots, 4) = 1) ? '●●'
                                : ((Mod(Dots, 4) = 2) ? '●●●'
                                :                       '●●●●')))
        DoneSteps.Text := Task ' / ' TaskNumber ' of prepare steps (is/are) done'
    }
    SetTimer(ShowInfo, 500)
    ; Create app folders
    For Each, Folder in AppDir {
        If !DirExist(Folder) {
            DirCreate(Folder)
        }
    }
    ; Download base files
    For Each, Package in BasePackages {
        PackagePath := StrReplace(Package, '/', '\')
        If !FileExist(PackagePath) {
            Download(Server '/' User '/' Repo '/main/' Package, PackagePath)
        }
        If InStr(Package, '.7z.') {
            PackageFolder := StrSplit(PackagePath, '.')[1]
            If !DirExist(PackageFolder)
                RunWait('DB\7za.exe x ' PackagePath ' -o' PackageFolder, , 'Hide')
        }
        ++Task
    }
    SetTimer(ShowInfo, 0)
    Prepare.Hide()
} Catch As Err {
    MsgBox('There was an error while preparing the necessary files!', 'Oops!', '48')
    ; Run installation help page
    ExitApp
}
; Main window
Manager := Gui(, 'AoE II Manager AIO')
Manager.OnEvent('Close', (*) => ExitApp())
; First section: # The Game
_Game_ := Manager.AddText('w220 h260 Center c800000 BackgroundFFFFFF Border', '# The Game')
_Game_.SetFont('Bold')
GetTheGame := Manager.AddButton('xm+10 ym+25 w200', 'Download AoE II')
GetTheGame.OnEvent('Click', (*) => DownloadInstallGame())
GuiButtonIcon(GetTheGame, 'DB\000\Down.png', , 'W16 H16 T2 A1')
ProgressBar := Manager.AddProgress('xp yp wp h20 Hidden', 0)
ProgressInfo := Manager.AddText('xp yp+25 wp Hidden Center')
DownloadInstallGame() {
    S_7z := '-aoa'
    DownloadRange := 7
    ExportRange := 4
    ProgressInfo.Value := ''
    ProgressBar.Value := 0
    ProgressBar.Opt('Range0-' DownloadRange)
    GameSectionInstallView()
    If !ExportDir := FileSelect('D') {
        GameSectionNormalView()
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
        ; The Age Of Kings
        Loop 4 {
            If !FileExist('DB\003.7z.00' A_Index)
                Download(Server '/' User '/' Repo '/main/DB/003.7z.00' A_Index, 'DB\003.7z.00' A_Index)
            ProgressInfo.Value := 'Downloaded [ ' Round((++ProgressBar.Value / DownloadRange) * 100) ' % ]'
        }
        ; The Conquerors
        Loop 2 {
            If !FileExist('DB\004.7z.00' A_Index)
                Download(Server '/' User '/' Repo '/main/DB/004.7z.00' A_Index, 'DB\004.7z.00' A_Index)
            ProgressInfo.Value := 'Downloaded [ ' Round((++ProgressBar.Value / DownloadRange) * 100) ' % ]'
        }
        ; Forgotten Empires
        Loop 1 {
            If !FileExist('DB\005.7z.00' A_Index)
                Download(Server '/' User '/' Repo '/main/DB/005.7z.00' A_Index, 'DB\005.7z.00' A_Index)
            ProgressInfo.Value := 'Downloaded [ ' Round((++ProgressBar.Value / DownloadRange) * 100) ' % ]'
        }
        ProgressBar.Value := 0
        ProgressBar.Opt('Range0-' ExportRange)
        ++ProgressBar.Value
        ; Export The Age Of Kings
        ProgressInfo.Value := 'Exporting The Age Of Kings...'
        RunWait('DB\7za.exe x DB\003.7z.001 -o"' ExportDir '\Age of Empires II" ' S_7z, , 'Hide')
        ++ProgressBar.Value
        ; Export The Conquerors
        ProgressInfo.Value := 'Exporting The Conquerors...'
        RunWait('DB\7za.exe x DB\004.7z.001 -o"' ExportDir '\Age of Empires II" ' S_7z, , 'Hide')
        ++ProgressBar.Value
        ; Export Forgotten Empires
        ProgressInfo.Value := 'Exporting Forgotten Empires...'
        RunWait('DB\7za.exe x DB\005.7z.001 -o"' ExportDir '\Age of Empires II" ' S_7z, , 'Hide')
        ++ProgressBar.Value
    } Catch As Err {
        GameSectionNormalView()
        MsgBox('Unable to get the game!', 'Oops!', '48')
        Return
    }
    GameSectionNormalView()
    Choice := MsgBox('Done!`n`nGame located at: "' ExportDir '\Age of Empires II"`n`nWanna select this game location?', 'Question', 0x20 + 0x4)
    If Choice = 'Yes' {
        ChosenFolder.Value := ExportDir '\Age of Empires II'
        IniWrite(ChosenFolder.Value, Config, 'Game', 'Path')
        LoadCurrentSettings()
    }
}

RunAOK := Manager.AddButton('xm+30 yp+20 w48 H48')
GuiButtonIcon(RunAOK, 'DB\000\aok.png', , 'W32 H32')
RunAOK.OnEvent('Click', (*) => Run(ChosenFolder.Value '\empires2.exe', ChosenFolder.Value))

RunAOC := Manager.AddButton('yp wp hp')
GuiButtonIcon(RunAOC, 'DB\000\aoc.png', , 'W32 H32')
RunAOC.OnEvent('Click', (*) => Run(ChosenFolder.Value '\age2_x1\age2_x1.exe', ChosenFolder.Value '\age2_x1'))

RunFOE := Manager.AddButton('yp wp hp')
GuiButtonIcon(RunFOE, 'DB\000\fe.png', , 'W32 H32')
RunFOE.OnEvent('Click', (*) => Run(ChosenFolder.Value '\age2_x1\age2_x2.exe', ChosenFolder.Value '\age2_x1'))

ChooseFolder := Manager.AddButton('xm+10 yp+60 w30 w100', 'Choose')
ChooseFolder.OnEvent('Click', (*) => SelectTheGame())
GuiButtonIcon(ChooseFolder, 'DB\000\Folder.png', , 'W16 H16 T1 A1')

LoadGRFolder := Manager.AddButton('xm+180 yp w30 w30')
LoadGRFolder.OnEvent('Click', (*) => SelectTheGameFromGR())
GuiButtonIcon(LoadGRFolder, 'DB\000\GR.png', , 'W16 H16')

SelectTheGameFromGR() {
    TextFound := LoadGRSettingText()
    FoundLocations := []
    AOKDir := ''
    AOCDir := ''
    FOEDir := ''
    ChosenDir := ''
    If AOKFile := GRGamePath(TextFound, 'empires2.exe') {
        SplitPath(AOKFile,, &AOKDir)
        FoundLocations.Push(ChosenDir := AOKDir)
    }
    If AOCFile := GRGamePath(TextFound, 'age2_x1.exe') {
        SplitPath(AOCFile,, &AOCDir)
        SplitPath(AOCDir,, &AOCDir)
        FoundLocations.Push(ChosenDir := AOCDir)
    }
    If FOEFile := GRGamePath(TextFound, 'age2_x2.exe') {
        SplitPath(FOEFile,, &FOEDir)
        SplitPath(FOEDir,, &FOEDir)
        FoundLocations.Push(ChosenDir := FOEDir)
    }
    RemoveDuplications(Arr) {
        E := ''
        For Each, Value in Arr {
            If !InStr(E, Value) {
                E .= E = '' ? Value : ',' Value
            }
        }
        Return StrSplit(E, ',')
    }
    FoundLocations := RemoveDuplications(FoundLocations)
    If FoundLocations.Length > 1 {
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
            LoadCurrentSettings()
            Return
        }
        Location.Show()
        Return
    }
    If !ChosenDir
        Return
    IniWrite(ChosenDir, Config, 'Game', 'Path')
    ChosenFolder.Value := ChosenDir
    LoadCurrentSettings()
}

ChosenFolder := Manager.AddEdit('xm+10 yp+30 w200 Center ReadOnly r4 -VScroll cBlue BackgroundWhite')

OpenTheGameFolder := Manager.AddButton('w200', 'Open')
GuiButtonIcon(OpenTheGameFolder, 'DB\000\Folder.png', , 'W16 H16 T1 A1')
OpenTheGameFolder.OnEvent('Click', (*) => Run(ChosenFolder.Value))
SelectTheGame() {
    SelectAFolder()
    LoadCurrentSettings(True)
}
; Second Section: # Versions
_Version_ := Manager.AddText('ym w450 h220 Center c800000 BackgroundFFFFFF Border', '# Versions')
_Version_.SetFont('Bold')

Manager.AddPicture('xp+54 ym+25 BackgroundTrans', 'DB\000\aok.png')
Manager.AddText('xp-44 yp+40 cRed w120 Center BackgroundTrans', 'The Age of Kings').SetFont('Bold')
Manager.AddText('xp+20 yp+20 w1 h1 BackgroundTrans')
Loop Files, 'DB\002\2*', 'D' {
    Handle := Manager.AddRadio('w30 w100 BackgroundFFFFFF', A_LoopFileName)
    Handle.SetFont('s10', 'Consolas')
    Handle.OnEvent('Click', ApplyVersion)
    General['AOK']['VersionsN'][A_LoopFileName] := Handle
}
ApplyVersion(Ctrl, Info) {
    DisableVersions()
    DisableGameRun()
    If GameIsRunning() {
        LoadCurrentSettings()
        Return
    }
    Try {
        CleanUp(Ctrl.Text)
        SetVersion(Ctrl.Text)
    } Catch As Err {
        Try {
            GRPath := ''
            Try
                GRPath := ProcessGetPath('GameRanger.exe')
            ProcessClose('GameRanger.exe')
            CleanUp(Ctrl.Text)
            SetVersion(Ctrl.Text)
            If GRPath
                Run(GRPath)
        } Catch As Err {
            MsgBox('An error occured while trying to set v' Ctrl.Text, 'Version apply error!', 0x20)
        }
    }
    EnableGameRun()
    EnableVersions()
    SoundPlay('DB\000\30 wololo.mp3')
}
Manager.AddPicture('xp+174 ym+25 BackgroundTrans', 'DB\000\aoc.png')
Manager.AddText('xp-44 yp+40 cBlue w120 Center BackgroundTrans', 'The Conquerors').SetFont('Bold')
Manager.AddText('xp+20 yp+20 w1 h1 BackgroundTrans')
Loop Files, 'DB\002\1*', 'D' {
    Handle := Manager.AddRadio('w30 w100 BackgroundFFFFFF', A_LoopFileName)
    Handle.SetFont('s10', 'Consolas')
    Handle.OnEvent('Click', ApplyVersion)
    General['AOC']['VersionsN'][A_LoopFileName] := Handle
}

Manager.AddPicture('xp+174 ym+25 BackgroundTrans', 'DB\000\fe.png')
Manager.AddText('xp-44 yp+40 cGreen w120 Center BackgroundTrans', 'Forgotten Empires').SetFont('Bold')
Manager.AddText('xp+20 yp+20 w1 h1 BackgroundTrans')
Handle := Manager.AddRadio('w30 w100 Checked BackgroundFFFFFF', '2.2  CD')
Handle.SetFont('s10', 'Consolas')
General['FOE']['VersionsN']['2.2  CD'] := Handle

Patch := Manager.AddDropDownList('xm+240 ym+195 w430', ['Do Not Enable Fixes'])
Patch.OnEvent('Change', (*) => IniWrite(Patch.Text, Config, 'Game', 'Fix'))

_Compatibility_ := Manager.AddText('xm+230 ym+225 w450 h140 Center c800000 BackgroundFFFFFF Border', '# Compatibilities')
_Compatibility_.SetFont('Bold')
Manager.AddPicture('xp+54 yp+25 BackgroundTrans', 'DB\000\aok.png')
Manager.AddText('xp-44 yp+40 cRed w120 Center BackgroundTrans', 'The Age of Kings').SetFont('Bold')
AoKCom := Manager.AddDropDownList('xp yp+20 w120')
For Each, Compat in Compatibilities {
    AoKCom.Add([Compat[1]])
}
AoKCom.Choose(1)
AoKCom.OnEvent("Change", (*) => AoKComReg())
AoKRun := Manager.AddCheckbox('yp+30 wp hp BackgroundFFFFFF', 'Run as administrator')
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

Manager.AddPicture('xp+194 yp-90 BackgroundTrans', 'DB\000\aoc.png')
Manager.AddText('xp-44 yp+40 cBlue w120 Center BackgroundTrans', 'The Conquerors').SetFont('Bold')
AoCCom := Manager.AddDropDownList('xp yp+20 w120')
For Each, Compat in Compatibilities {
    AoCCom.Add([Compat[1]])
}
AoCCom.Choose(1)
AoCCom.OnEvent("Change", (*) => AoCComReg())
AoCRun := Manager.AddCheckbox('yp+30 wp hp BackgroundFFFFFF', 'Run as administrator')
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

Manager.AddPicture('xp+194 yp-90 BackgroundTrans', 'DB\000\fe.png')
Manager.AddText('xp-44 yp+40 cGreen w120 Center BackgroundTrans', 'Forgotten Empires').SetFont('Bold')
FOECom := Manager.AddDropDownList('xp yp+20 w120')
For Each, Compat in Compatibilities {
    FOECom.Add([Compat[1]])
}
FOECom.Choose(1)
FOECom.OnEvent("Change", (*) => FOEComReg())
FOERun := Manager.AddCheckbox('yp+30 wp hp BackgroundFFFFFF', 'Run as administrator')
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

_Language_ := Manager.AddText('xm yp-75 w220 h385 Center c800000 BackgroundFFFFFF Border', '# Languages')
_Language_.SetFont('Bold')
Manager.AddText('xp+10 yp w200 BackgroundTrans')
Loop Files, 'DB\006\*', 'D' {
    Handle := Manager.AddRadio('wp Center BackgroundFFFFFF', A_LoopFileName)
    Handle.SetFont('Underline Bold')
    Handle.OnEvent('Click', ApplyLanguage)
    General['LNG'][A_LoopFileName] := Handle
}
ApplyLanguage(Ctrl, Info) {
    DisableLanguage(), Sleep(500)
    If !((Time := FoundDefaultLanguage()) && ((Ctrl.Text = '___Default___'))) {
        DirCopy('DB\006\' Ctrl.Text, ChosenFolder.Value, 1)
    } Else {
        DirCopy(AppDir[2] '\' Time, ChosenFolder.Value, 1)
    }
    EnableLanguage()
    SoundPlay('DB\000\30 wololo.mp3')
}
ModePic := Gui('-Caption AlwaysOnTop')
ModePic.MarginX := 3
ModePic.MarginY := 3
ModePic.BackColor := 0x000000
ModeThumb := ModePic.AddPicture('w150 h113')
_VisualMods_ := Manager.AddText('xm+230 yp-255 w220 h280 Center c800000 BackgroundFFFFFF Border', '# Visual Mods')
_VisualMods_.SetFont('Bold')
Manager.AddText('xp+10 yp+10 w200 BackgroundTrans')
VMList := Manager.AddListView('w200 h210 -E0x200 -Hdr Checked BackgroundFFFFFF', ['Mod Name'])
VMList.SetFont('Bold')
Loop Files, 'DB\007\*', 'D' {
    VMList.Add(, A_LoopFileName)
}
VMList.OnEvent('ItemSelect', ShowModePic)
ShowModePic(Ctrl, Item, Selected) {

    SetTimer(Follow, 0)
    VMName := VMList.GetText(Item)
    Try {
        Size := imgSize('DB\007\' VMName '\Mode.pic')
        W := Size.w
        H := Size.h
    } Catch As Err {
        Dummy := Gui()
        Pic := Dummy.AddPicture(, 'DB\007\' VMName '\Mode.pic')
        Pic.GetPos(,, &W, &H)
        Dummy.Destroy()
    }
    MouseGetPos(&X, &Y)

    ModeThumb.Move(,, W, H)
    ModeThumb.Value := 'DB\007\' VMName '\Mode.pic'
    ModePic.Show('NA x' X + 10 ' y' Y + 5 ' w' W + 6 ' h' H + 6)
    Show := True
    SetTimer(Follow, 50)
    Follow() {
        MouseGetPos(&X, &Y,, &Ctrl)
        If Ctrl != 'SysListView321' {
            SetTimer(Follow, 0)
            ModePic.Hide()
            Show := False
        }
        If Show
            ModePic.Show('NA x' X + 10 ' y' Y + 5)
    }
}
VMList.OnEvent('ItemCheck', ApplyVM)
ApplyVM(Ctrl, Item, Checked) {
    DisableVisualMod()
    VMName := VMList.GetText(Item)
    SlpDir := Checked ? 'DB\007\' VMName : 'DB\007\' VMName '\U'
    RunWait('DB\000\DrsBuild.exe /a "' ChosenFolder.Value '\Data\' DrsTypes['gra'] '" "' SlpDir '\gra*.slp"', , 'Hide')
    RunWait('DB\000\DrsBuild.exe /a "' ChosenFolder.Value '\Data\' DrsTypes['int'] '" "' SlpDir '\int*.slp"', , 'Hide')
    RunWait('DB\000\DrsBuild.exe /a "' ChosenFolder.Value '\Data\' DrsTypes['ter'] '" "' SlpDir '\ter*.slp"', , 'Hide')
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
    EnableVisualMod()
    SoundPlay('DB\000\30 wololo.mp3')
}
LoadVM := Manager.AddButton('wp', 'Load')
LoadVM.SetFont('Bold')
LoadVM.OnEvent('Click', (*) => LoadVisualMod())
LoadVisualMod() {
    If Selected := FileSelect('D') {
        SplitPath(Selected, &ModeName)
        If DirExist('DB\007\' ModeName) {
            MsgBox(ModeName ' is already loaded!', ModeName, 0x30)
            Return
        }
        DisableVisualMod()
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
        EnableVisualMod()
    }
    LoadVM.Enabled := True
}

_VisualMods_.GetPos(, &Y)
_DataMods_ := Manager.AddText('xm+460 y' Y ' w220 h280 Center c800000 BackgroundFFFFFF Border', '# Data Mods`n`n(Comming Soon)')
_DataMods_.SetFont('Bold')

_ATools_ := Manager.AddText('ym w220 h650 Center c800000 BackgroundFFFFFF Border', '# Other Tools`n`n(Comming Soon)')
_ATools_.SetFont('Bold')

;Manager.AddText('xp+10 yp+20 cBlue w200 BackgroundTrans Center', '1 - Hide All IP [VPN]').SetFont('Bold')
;VPN := Manager.AddButton('w56 h56')
;VPN.OnEvent('Click', (*) => MsgBox('Not Yet Implemented!', 'Hoy!', 0x40))
;GuiButtonIcon(VPN, 'DB\000\vpn.png',, 'W48 H48')
;ClearVPNReg := Manager.AddButton('xp+65 yp+2 w130', 'Reset Trial')
;ClearVPNReg.OnEvent('Click', (*) => MsgBox('Not Yet Implemented!', 'Hoy!', 0x40))
;VPNCompat := Manager.AddDropDownList('w130')
;
;Manager.AddText('xp-65 yp+40 cBlue w200 BackgroundTrans Center', '2 - Shortcuts/Keys Remapper').SetFont('Bold')
;KRemap := Manager.AddButton('wp', 'Create/Modify')
;KRemap.OnEvent('Click', (*) => MsgBox('Not Yet Implemented!', 'Hoy!', 0x40))
;
;Manager.AddText('yp+40 cBlue w200 BackgroundTrans Center', '3 - Repair Game Files').SetFont('Bold')
;RepairGame := Manager.AddButton('wp', 'Repair')
;RepairGame.OnEvent('Click', (*) => MsgBox('Not Yet Implemented!', 'Hoy!', 0x40))
;
;Manager.AddText('yp+40 cBlue w200 BackgroundTrans Center', '4 - Repair Record Files (.mgz)').SetFont('Bold')
;FixMgz := Manager.AddButton('wp', 'Repair')
;FixMgz.OnEvent('Click', (*) => MsgBox('Not Yet Implemented!', 'Hoy!', 0x40))
;
;Manager.AddText('yp+40 cBlue w200 BackgroundTrans Center', '5 - Scenario Files Select').SetFont('Bold')
;FixMgz := Manager.AddButton('wp', 'Select')
;FixMgz.OnEvent('Click', (*) => MsgBox('Not Yet Implemented!', 'Hoy!', 0x40))

AboutText := ''
           . '| AGE OF EMPIRES II MANAGER ALL IN ONE, '
           . 'AUTOHOTKEY BASED APP, '
           . 'CREATED BY SMILE, '
           . 'TOTALLY SECURE, '
           . 'TESTED MANY TIMES, '
           . 'BUT USE ON YOUR OWN RISK, '
           . 'ANY FEEDBACK WILL BE HELPFUL, '
           . 'MY EMAIL, '
           . 'CHANDOUL.MOHAMED26@GMAIL.COM, '
           . 'WEBSITE FOR THIS APP, '
           . 'HTTPS://SMILEAOE.GITHUB.IO |'

SB := Manager.AddStatusBar()
SB.SetFont('Bold', 'Calibri')
SB.SetParts(10, 50, 200)
SB.SetText('v' Version, 2)
SB.SetText('Loading...', 3)
SB.SetText(A_Tab A_Tab 'A Collective App From The Internet On What I Found Useful About AoE II!    ', 4)
Manager.Show()
LoadCurrentSettings()
__CheckForUpdates__()
Return

LoadCurrentSettings(Browse := False) {
    SoundPlay('DB\000\30 wololo.mp3')
    DisableGameRun()
    DisableVersions()
    DisableCompatibilitys()
    DisableLanguage()
    DisableVisualMod()
    ChosenFolder.Value := IniRead(Config, 'Game', 'Path', '')
    If !ValidGameLocation(ChosenFolder.Value) {
        SplitPath(ChosenFolder.Value,, &Expected)
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
        SplitPath(ChosenFolder.Value,,,,, &Drive)
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
    EnableAOKVersion()
    EnableAOKRun()
    EnableAOKCompatibility()
    FixCommonIssues()
    If FileExist(ChosenFolder.Value '\age2_x1\age2_x1.exe') {
        EnableAOCVersion()
        EnableAOCRun()
        EnableAOCCompatibility()
    }
    If FileExist(ChosenFolder.Value '\age2_x1\age2_x2.exe') {
        EnableFOEVersion()
        EnableFOERun()
        EnableFOECompatibility()
    }
    VersionsLoad()
    LanguageLoad()
    LoadAppliedVM()
    CompatibilityCheck()
    CopyDefaultLanguage()
    EnableLanguage()
    EnableVisualMod()
    LoadEnableFixes()
}
imgSize(img) { ; https://www.autohotkey.com/boards/viewtopic.php?f=76&t=81665
    ; Returns an array indicating the image's width (w) and height (h), obtained from the file's properties
    SplitPath img, &fn, &dir
    (dir = '' && dir := A_WorkingDir)
    objShell := ComObject("Shell.Application")
    objFolder := objShell.NameSpace(dir), objFolderItem := objFolder.ParseName(fn)
    scale := StrSplit(RegExReplace(objFolder.GetDetailsOf(objFolderItem, 31), ".(.+).", "$1"), " x ")
    Return { w: scale[1], h: scale[2] }
}
LoadAppliedVM() {
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
CompatibilityCheck() {
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
CopyDefaultLanguage() {
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
GameIsRunning() {
    For Each, App in ['empires2.exe', 'age2_x1.exe', 'age2_x2.exe'] {
        If ProcessExist(App) {
            ProcessClose(App)
        }
        ProcessWaitClose(App, 5)
        If ProcessExist(App) {
            Return True
        }
    }
    Return False
}
SetVersion(Version) {
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
    If Patch.Value = 1 {
        Return
    }
    DirCopy('DB\001\' Patch.Text '\Static' , ChosenFolder.Value, 1)
    If DirExist('DB\001\' Patch.Text '\' Version) {
        DirCopy('DB\001\' Patch.Text '\' Version, ChosenFolder.Value, 1)
        If InStr(Patch.Text, 'v2') {
            RegWrite(2, 'REG_DWORD', 'HKEY_CURRENT_USER\SOFTWARE\Microsoft\Microsoft Games\Age of Empires', 'Aoe2Patch')
        }
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
LanguageLoad() {
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
VersionsLoad() {
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
; https://www.autohotkey.com/boards/viewtopic.php?f=83&t=115871
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
; https://autohotkey.com/board/topic/66139-ahk-l-calculating-md5sha-checksum-from-file/
HashFile(filePath, hashType := 2) {
    static PROV_RSA_AES := 24
    static CRYPT_VERIFYCONTEXT := 0xF0000000
    static BUFF_SIZE := 1024 * 1024 ; 1 MB
    static HP_HASHVAL := 0x0002
    static HP_HASHSIZE := 0x0004

    switch hashType {
        case 1: hash_alg := (CALG_MD2 := 32769)
        case 2: hash_alg := (CALG_MD5 := 32771)
        case 3: hash_alg := (CALG_SHA := 32772)
        case 4: hash_alg := (CALG_SHA_256 := 32780)
        case 5: hash_alg := (CALG_SHA_384 := 32781)
        case 6: hash_alg := (CALG_SHA_512 := 32782)
        default: throw ValueError('Invalid hashType', -1, hashType)
    }

    If !FileExist(filePath) {
        Return
    }

    f := FileOpen(filePath, "r")
    f.Pos := 0 ; Rewind in case of BOM.

    HCRYPTPROV() => {
        ptr: 0,
        __delete: this => this.ptr && DllCall("Advapi32\CryptReleaseContext", "Ptr", this, "UInt", 0)
    }

    if !DllCall("Advapi32\CryptAcquireContextW"
        , "Ptr*", hProv := HCRYPTPROV()
        , "Uint", 0
        , "Uint", 0
        , "Uint", PROV_RSA_AES
        , "UInt", CRYPT_VERIFYCONTEXT)
        throw OSError()

    HCRYPTHASH() => {
        ptr: 0,
        __delete: this => this.ptr && DllCall("Advapi32\CryptDestroyHash", "Ptr", this)
    }

    if !DllCall("Advapi32\CryptCreateHash"
        , "Ptr", hProv
        , "Uint", hash_alg
        , "Uint", 0
        , "Uint", 0
        , "Ptr*", hHash := HCRYPTHASH())
        throw OSError()

    read_buf := Buffer(BUFF_SIZE, 0)

    While (cbCount := f.RawRead(read_buf, BUFF_SIZE))
    {
        if !DllCall("Advapi32\CryptHashData"
            , "Ptr", hHash
            , "Ptr", read_buf
            , "Uint", cbCount
            , "Uint", 0)
            throw OSError()
    }

    if !DllCall("Advapi32\CryptGetHashParam"
        , "Ptr", hHash
        , "Uint", HP_HASHSIZE
        , "Uint*", &HashLen := 0
        , "Uint*", &HashLenSize := 4
        , "UInt", 0)
        throw OSError()

    bHash := Buffer(HashLen, 0)
    if !DllCall("Advapi32\CryptGetHashParam"
        , "Ptr", hHash
        , "Uint", HP_HASHVAL
        , "Ptr", bHash
        , "Uint*", &HashLen
        , "UInt", 0)
        throw OSError()

    loop HashLen
        HashVal .= Format('{:02x}', (NumGet(bHash, A_Index - 1, "UChar")) & 0xff)

    return HashVal
}
ConnectedToInternet(Flag := 0x40) {
    Return DllCall("Wininet.dll\InternetGetConnectedState", "Str", Flag, "Int", 0)
}
GetTextFromLink(Link) {
    whr := ComObject("WinHttp.WinHttpRequest.5.1")
    whr.Open("GET", Link, true)
    whr.Send()
    whr.WaitForResponse()
    Return whr.ResponseText
}
BufferToBase64(BufferObj) {
	if !DllCall('Crypt32.dll\CryptBinaryToString', 'Ptr', BufferObj, 'UInt', BufferObj.Size, 'UInt', 1, 'Ptr', 0, 'UInt*', &numChars := 0)
		Throw 'Cant compute the destination buffer size, error: ' A_LastError
    if !DllCall('Crypt32.dll\CryptBinaryToString', 'Ptr', BufferObj, 'UInt', BufferObj.Size, 'UInt', 1, 'Ptr', BufferString := Buffer(numChars * 2), 'UInt*', numChars * 2)
		Throw 'Cant convert source buffer to base64, error: ' A_LastError
	return StrGet(BufferString)
}
Base64ToBuffer(Base64) {
    If !DllCall("Crypt32.dll\CryptStringToBinary", "Str", Base64, "UInt", BLen := StrLen(Base64), "UInt", 1, "UInt", 0, "UIntP", &Rqd := 0, "Int", 0, "Int", 0)
        Throw 'Cant compute the destination buffer size, error: ' A_LastError
    If !DllCall("Crypt32.dll\CryptStringToBinary", "Str", Base64, "UInt", BLen, "UInt", 1, "Ptr", BufferObj := Buffer(Rqd, 0), "UIntP", Rqd, "Int", 0, "Int", 0)
        Throw 'Cant convert source base64 to buffer, error: ' A_LastError
    Return BufferObj
}
FileToBase64(FileName) {
    Return BufferToBase64(FileRead(FileName, 'RAW'))
}
Base64ToFile(Base64, FileName) {
    BufferObj := Base64ToBuffer(Base64)
    O := FileOpen(FileName, "w")
    O.RawWrite(BufferObj, BufferObj.Size)
    O.Close()
}
__CheckForUpdates__() {
    Global Version, TaskNumber, Task
    If A_IsCompiled {
        Return
    }
    Try {
        SB.SetText('Checking for updates...', 3)
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
        Loop Files, 'DB\*7z*', 'R' {
            If HashsumsMap.Has(A_LoopFileDir '\' A_LoopFileName) 
                && (HashFile(A_LoopFileDir '\' A_LoopFileName) != HashsumsMap[A_LoopFileDir '\' A_LoopFileName]) {
                FoundUpdates.Push(A_LoopFileDir '\' A_LoopFileName)
            }
        }
        If FoundUpdates.Length {
            UpdatesList := '`n'
            For Each, UpdateFile in FoundUpdates{
                UpdatesList .= '- ' UpdateFile '`n'
            }
            SB.SetText('Update found!', 3)
            Choice := MsgBox('The following needs to be updated`n' UpdatesList '`nUpdate now?', 'Update', 0x4 + 0x20)
            If Choice = 'Yes' {
                Task := 1
                TaskNumber := FoundUpdates.Length
                SetTimer(ShowInfo, 500)
                Prepare.Show()
                Manager.Hide()
                For Each, UpdateFile in FoundUpdates {
                    Task := Each
                    DownloadLink := Server '/' User '/' Repo '/main/' StrReplace(StrReplace(UpdateFile, ' ', '%20'), '\', '/')
                    Download(DownloadLink, UpdateFile)
                    If InStr(UpdateFile, '7z') {
                        Name := StrSplit(UpdateFile, '.')
                        DirDelete(Name[1], 1)
                        RunWait('DB\7za.exe x ' Name[1] '.7z.001 -o' Name[1] ' -aoa', , 'Hide')
                    }
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
ValidGameLocation(Location) {
    Return FileExist(Location '\empires2.exe')
        && FileExist(Location '\language.dll')
        && FileExist(Location '\Data\graphics.drs')
        && FileExist(Location '\Data\interfac.drs')
        && FileExist(Location '\Data\terrain.drs')
}
FixCommonIssues() {
    If FileExist(ChosenFolder.Value '\age2_x1.exe') {
        If !DirExist(ChosenFolder.Value '\age2_x1') {
            DirCreate(ChosenFolder.Value '\age2_x1')
        }
        FileMove(ChosenFolder.Value '\age2_x1.exe', ChosenFolder.Value '\age2_x1', 1)
    }
}
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
            Byte    := NumGet(F, (A_Index - 1) + 4, 'UChar')
            Val     := (Byte - 17) ^ 0x23
            UChar   := Val & 0xFF
            NByte   := (0x20 * (Val) | (UChar >> 3))
            NumPut('UChar', NByte, NF, A_Index - 1)
        }
        FileOpen(FileName, 'w').RawWrite(NF, NF.Size)
        Return True
    }
    Return True
}
LoadEnableFixes() {
    Loop Files, 'DB\001\*', 'D' {
        Patch.Add([A_loopFileName])
    }
    Patch.Choose('Do Not Enable Fixes')
    If DirExist('DB\001\' Fix := IniRead(Config, 'Game', 'Fix', 'Do Not Enable Fixes')) {
        Patch.Choose(Fix)
    }
}
SelectAFolder() {
    ChosenDir := FileSelect('D', 'C:\' (A_Is64bitOS ? 'Program Files (x86)' : 'Program Files') '\Microsoft Games')
    If !ChosenDir
        Return
    IniWrite(ChosenDir, Config, 'Game', 'Path')
    ChosenFolder.Value := ChosenDir
}
GRGamePath(TextFound, AppName) {
    P := InStr(TextFound, LFE := AppName,, -1)
    Loop {
        Char := SubStr(TextFound, P - (I := A_Index), 1)
        LFE := Char LFE
    } Until (Char = ':' || Ord(Char) = 10 || Ord(Char) = 13)
    Result := SubStr(TextFound, P - (I + 1), 1) LFE
    Return (FileExist(Result) ? Result : '')
}
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
GameSectionInstallView() {
    GetTheGame.Visible := False
    ProgressBar.Visible := True
    ProgressInfo.Visible := True
}
GameSectionNormalView() {
    GetTheGame.Visible := True
    ProgressBar.Visible := False
    ProgressInfo.Visible := False
}
DisableVisualMod() {
    VMList.Enabled := False
    LoadVM.Enabled := False
}
EnableVisualMod() {
    VMList.Enabled := True
    VMList.Redraw()
    LoadVM.Enabled := True
}
DisableLanguage() {
    _Language_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > X && CX < (X + (Width / 3))) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := False
        }
    }
}
EnableLanguage() {
    _Language_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > X && CX < (X + Width)) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := True
        }
    }
}
DisableCompatibilitys() {
    DisableAOKCompatibility()
    DisableAOCCompatibility()
    DisableFOECompatibility()
}
EnableCompatibilitys() {
    EnableAOKCompatibility()
    EnableAOCCompatibility()
    EnableFOECompatibility()
}
EnableAOKCompatibility() {
    _Compatibility_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > X && CX < (X + (Width / 3))) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := True
        }
    }
}
EnableAOCCompatibility() {
    _Compatibility_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > (X + (Width / 3)) && CX < (X + (Width * 2 / 3))) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := True
        }
    }
}
EnableFOECompatibility() {
    _Compatibility_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > (X + (Width * 2 / 3)) && CX < (X + Width)) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := True
        }
    }
}
DisableAOKCompatibility() {
    _Compatibility_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > X && CX < (X + (Width / 3))) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := False
            Control.Redraw()
        }
    }
}
DisableAOCCompatibility() {
    _Compatibility_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > (X + (Width / 3)) && CX < (X + (Width * 2 / 3))) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := False
            Control.Redraw()
        }
    }
}
DisableFOECompatibility() {
    _Compatibility_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > (X + (Width * 2 / 3)) && CX < (X + Width)) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := False
            Control.Redraw()
        }
    }
}
DisableVersions() {
    DisableAOKVersion()
    DisableAOCVersion()
    DisableFOEVersion()
}
EnableVersions() {
    EnableAOKVersion()
    EnableAOCVersion()
    EnableFOEVersion()
}
EnableAOKVersion() {
    _Version_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > X && CX < (X + (Width / 3))) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := True
        }
    }
}
EnableAOCVersion() {
    _Version_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > (X + (Width / 3)) && CX < (X + (Width * 2 / 3))) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := True
        }
    }
}
EnableFOEVersion() {
    _Version_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > (X + (Width * 2 / 3)) && CX < (X + Width)) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := True
        }
    }
}
DisableAOKVersion() {
    _Version_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > X && CX < (X + (Width / 3))) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := False
            Control.Redraw()
        }
    }
}
DisableAOCVersion() {
    _Version_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > (X + (Width / 3)) && CX < (X + (Width * 2 / 3))) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := False
            Control.Redraw()
        }
    }
}
DisableFOEVersion() {
    _Version_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > (X + (Width * 2 / 3)) && CX < (X + Width)) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := False
            Control.Redraw()
        }
    }
}
DisableGameRun() {
    DisableAOKRun()
    DisableAOCRun()
    DisableFOERun()
}
EnableGameRun() {
    EnableAOKRun()
    EnableAOCRun()
    EnableFOERun()
}
EnableAOKRun() {
    RunAOK.Enabled := True
}
EnableAOCRun() {
    RunAOC.Enabled := True
}
EnableFOERun() {
    RunFOE.Enabled := True
}
DisableAOKRun() {
    RunAOK.Enabled := False
}
DisableAOCRun() {
    RunAOC.Enabled := False
}
DisableFOERun() {
    RunFOE.Enabled := False
}