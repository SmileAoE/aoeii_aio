#Requires AutoHotkey v2.0
#SingleInstance Force

Server := 'https://raw.githubusercontent.com'
User := 'SmileAoE'
Repo := 'aoeii_aio'
Version := '1.1'
Layers := 'HKLM\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers'
Config := A_AppData '\aoeii_aio\config.ini'
AppDir := ['DB', A_AppData '\aoeii_aio']
GRSetting := A_AppData '\GameRanger\GameRanger Prefs\Settings'
Dots := 0
Task := 1
TaskNumber := 12

DrsTypes := Map('gra', 'graphics.drs', 'int', 'interfac.drs', 'ter', 'terrain.drs')

General := Map()

General['AOK'] := Map()
General['AOK']['VersionsN'] := Map()
General['AOK']['Combine'] := Map('2.0b CD', ['2.0a No CD'])

General['AOC'] := Map()
General['AOC']['VersionsN'] := Map()
General['AOC']['Combine'] := Map( '1.0e No CD', ['1.0c No CD']
                                , '1.0e No CD', ['1.0c No CD']
                                , '1.1  No CD', ['1.0c No CD']
                                , '1.5  CD', ['1.0c No CD'])

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
        HoldOn.Text := PW '`n' ((Mod(++Dots, 4) = 0) ? '?'
            : ((Mod(Dots, 4) = 1) ? '??'
                : ((Mod(Dots, 4) = 2) ? '???'
                    : '????')))
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
    If !FileExist('DB\7za.exe') {
        Download(Server '/' User '/' Repo '/main/DB/7za.exe', 'DB\7za.exe')
    }
    ++Task
    If !FileExist('DB\000.7z.001') {
        Download(Server '/' User '/' Repo '/main/DB/000.7z.001', 'DB\000.7z.001')
    }
    ++Task
    If !FileExist('DB\001.7z.001') {
        Download(Server '/' User '/' Repo '/main/DB/001.7z.001', 'DB\001.7z.001')
    }
    ++Task
    If !FileExist('DB\002.7z.001') {
        Download(Server '/' User '/' Repo '/main/DB/002.7z.001', 'DB\002.7z.001')
    }
    ++Task
    If !FileExist('DB\006.7z.001') {
        Download(Server '/' User '/' Repo '/main/DB/006.7z.001', 'DB\006.7z.001')
    }
    ++Task
    If !FileExist('DB\007.7z.001') {
        Download(Server '/' User '/' Repo '/main/DB/007.7z.001', 'DB\007.7z.001')
    }
    ++Task
    ; Export base files
    If !DirExist('DB\000') {
        RunWait('DB\7za.exe x DB\000.7z.001 -oDB\000', , 'Hide')
    }
    ++Task
    If !DirExist('DB\001') {
        RunWait('DB\7za.exe x DB\001.7z.001 -oDB\001', , 'Hide')
    }
    ++Task
    If !DirExist('DB\002') {
        RunWait('DB\7za.exe x DB\002.7z.001 -oDB\002', , 'Hide')
    }
    ++Task
    If !DirExist('DB\006') {
        RunWait('DB\7za.exe x DB\006.7z.001 -oDB\006', , 'Hide')
    }
    ++Task
    If !DirExist('DB\007') {
        RunWait('DB\7za.exe x DB\007.7z.001 -oDB\007', , 'Hide')
    }
    ++Task
    SetTimer(ShowInfo, 0)
    Prepare.Destroy()
} Catch As Err {
    MsgBox('There was an error while preparing the necessary files!', 'Oops!', '48')
    ; Run installation help page
    ExitApp
}
; Main window
Manager := Gui(, 'AoE II Manager AIO')
Manager.OnEvent('Close', (*) => ExitApp())
Manager.BackColor := 0xFFFFFF
;Manager.MarginX := Manager.MarginY := 5
; First section: # The Game
_Game_ := Manager.AddText('w220 h260 Center c800000 BackgroundFFF2E4 Border', '# The Game')
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

ChosenFolder := Manager.AddEdit('xm+10 yp+30 w200 Center ReadOnly r4 -VScroll cBlue BackgroundWhite')

OpenTheGameFolder := Manager.AddButton('w200', 'Open')
GuiButtonIcon(OpenTheGameFolder, 'DB\000\Folder.png', , 'W16 H16 T1 A1')
OpenTheGameFolder.OnEvent('Click', (*) => Run(ChosenFolder.Value))
SelectTheGame() {
    ChosenDir := FileSelect('D')
    If !ChosenDir
        Return
    IniWrite(ChosenDir, Config, 'Game', 'Path')
    ChosenFolder.Value := ChosenDir
    LoadCurrentSettings()
}
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
    FoundLocations := RemoveDuplications(FoundLocations)
    If FoundLocations.Length > 1 {
        Location := Gui('ToolWindow', 'Pick one location')
        Location.OnEvent('Close', (*) => DoNothing())
        DoNothing() {
            Location.Destroy()
            Return
        }
        Location.AddText('Center r3 w350 cRed', 'Different locations were found for Age of Empires II game`nPlease pick only one`n').SetFont('Bold')
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
; Second Section: # Versions
_Version_ := Manager.AddText('ym w450 h220 Center c800000 BackgroundFFFFE0 Border', '# Versions')
_Version_.SetFont('Bold')

Manager.AddPicture('xp+54 ym+25 BackgroundTrans', 'DB\000\aok.png')
Manager.AddText('xp-44 yp+40 cRed w120 Center BackgroundTrans', 'The Age of Kings').SetFont('Bold')
Manager.AddText('xp+20 yp+20 w1 h1 BackgroundTrans')
Loop Files, 'DB\002\2*', 'D' {
    Handle := Manager.AddRadio('w30 w100 BackgroundFFFFE0', A_LoopFileName)
    Handle.SetFont('s10', 'Consolas')
    Handle.OnEvent('Click', ApplyVersion)
    General['AOK']['VersionsN'][A_LoopFileName] := Handle
}

Manager.AddPicture('xp+174 ym+25 BackgroundTrans', 'DB\000\aoc.png')
Manager.AddText('xp-44 yp+40 cBlue w120 Center BackgroundTrans', 'The Conquerors').SetFont('Bold')
Manager.AddText('xp+20 yp+20 w1 h1 BackgroundTrans')
Loop Files, 'DB\002\1*', 'D' {
    Handle := Manager.AddRadio('w30 w100 BackgroundFFFFE0', A_LoopFileName)
    Handle.SetFont('s10', 'Consolas')
    Handle.OnEvent('Click', ApplyVersion)
    General['AOC']['VersionsN'][A_LoopFileName] := Handle
}

Manager.AddPicture('xp+174 ym+25 BackgroundTrans', 'DB\000\fe.png')
Manager.AddText('xp-44 yp+40 cGreen w120 Center BackgroundTrans', 'Forgotten Empires').SetFont('Bold')
Manager.AddText('xp+20 yp+20 w1 h1 BackgroundTrans')
Handle := Manager.AddRadio('w30 w100 Checked BackgroundFFFFE0', '2.2  CD')
Handle.SetFont('s10', 'Consolas')
General['FOE']['VersionsN']['2.2  CD'] := Handle
ApplyVersion(Ctrl, Info) {
    DisableGameRun()
    DisableVersions()
    If GameIsRunning() {
        LoadCurrentSettings()
        Return
    }
    CleanUp(Ctrl.Text)
    SetVersion(Ctrl.Text)
    EnableVersions()
    EnableGameRun()
    SoundPlay('DB\000\30 wololo.mp3')
}

Patch := Manager.AddCheckbox('BackgroundFFFFE0 xm+240 ym+200' (IniRead(Config, 'Game', 'Fix', 1) ? ' Checked' : ''), 'Enable fixs after each patching if available')
Patch.OnEvent('Click', (*) => IniWrite(Patch.Value, Config, 'Game', 'Fix'))

_Compatibility_ := Manager.AddText('xm+230 ym+225 w450 h140 Center c800000 BackgroundECFFDA Border', '# Compatibilities')
_Compatibility_.SetFont('Bold')
Manager.AddPicture('xp+54 yp+25 BackgroundTrans', 'DB\000\aok.png')
Manager.AddText('xp-44 yp+40 cRed w120 Center BackgroundTrans', 'The Age of Kings').SetFont('Bold')
AoKCom := Manager.AddDropDownList('xp yp+20 w120')
For Each, Compat in Compatibilities {
    AoKCom.Add([Compat[1]])
}
AoKCom.Choose(1)
AoKCom.OnEvent("Change", (*) => AoKComReg())
AoKRun := Manager.AddCheckbox('yp+30 wp hp BackgroundECFFDA', 'Run as administrator')
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
AoCRun := Manager.AddCheckbox('yp+30 wp hp BackgroundECFFDA', 'Run as administrator')
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
FOERun := Manager.AddCheckbox('yp+30 wp hp BackgroundECFFDA', 'Run as administrator')
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

_Language_ := Manager.AddText('xm yp-75 w220 h385 Center c800000 BackgroundE8FFF4 Border', '# Languages')
_Language_.SetFont('Bold')
Manager.AddText('xp+10 yp w200 BackgroundTrans')
Loop Files, 'DB\006\*', 'D' {
    Handle := Manager.AddRadio('wp Center BackgroundE8FFF4', A_LoopFileName)
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
    SoundPlay('DB\000\30 wololo.mp3')
    EnableLanguage()
}
_VisualMods_ := Manager.AddText('xm+230 yp-255 w220 h280 Center c800000 BackgroundDCDCFF Border', '# Visual Mods')
_VisualMods_.SetFont('Bold')
Manager.AddText('xp+10 yp+10 w200 BackgroundTrans')
VMList := Manager.AddListView('w200 h210 -E0x200 -Hdr Checked BackgroundDCDCFF', ['Mod Name'])
VMList.SetFont('Bold')
Loop Files, 'DB\007\*', 'D' {
    VMList.Add(, A_LoopFileName)
}
VMList.OnEvent('ItemCheck', ApplyVM)
ApplyVM(Ctrl, Item, Checked) {
    VMName := VMList.GetText(Item)
    SlpDir := Checked ? 'DB\007\' VMName : 'DB\007\' VMName '\U'
    VMList.Enabled := False
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
    SoundPlay('DB\000\30 wololo.mp3')
    VMList.Enabled := True
}
LoadVM := Manager.AddButton('wp', 'Load')
LoadVM.SetFont('Bold')
LoadVM.OnEvent('Click', (*) => LoadVisualMod())
LoadVisualMod() {
    LoadVM.Enabled := False
    If Selected := FileSelect('D') {
        SplitPath(Selected, &OutFileName)
        DirCreate('DB\007\' OutFileName '\U')
        Loop Files, Selected '\*.slp', 'R' {
            ID := SubStr(A_LoopFileName, 1, -4)
            If !IsDigit(ID) {
                Continue
            }
            If (ID < 6000) || ((ID >= 15000) && (ID <= 16000)) || ((ID >= 50000) && (ID <= 54000)) {
                FileCopy(A_LoopFileFullPath, 'DB\007\' OutFileName '\' A_LoopFileName, 1)
                RunWait('DB\000\vooblyslpdecode.exe "DB\007\' OutFileName '\' A_LoopFileName '" "DB\007\' OutFileName '\' A_LoopFileName '"', , 'Hide')
                AddPrefix(['DB\007\' OutFileName '\' A_LoopFileName], 5)
            }
        }
        Loop Files, 'DB\007\*.slp', 'R' {
            If !InStr(A_LoopFileDir, '\U')
                || InStr(A_LoopFileDir, OutFileName)
                Continue
            If FileExist('DB\007\' OutFileName '\' A_LoopFileName) {
                FileCopy(A_LoopFileFullPath, 'DB\007\' OutFileName '\U\' A_LoopFileName, 1)
            }
        }
        BackUpOrgSlp(OutFileName)
    }
    LoadVM.Enabled := True
}

_VisualMods_.GetPos(, &Y)
_DataMods_ := Manager.AddText('xm+460 y' Y ' w220 h280 Center c800000 BackgroundF6EEFF Border', '# Data Mods')
_DataMods_.SetFont('Bold')

AboutText := '| AGE OF EMPIRES II MANAGER ALL IN ONE, '
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
;About1 := Manager.AddText('x0 cYellow BackgroundBlack', AboutText)
;About1.SetFont('Bold', 'Calibri')
;About1.GetPos(&AX1,, &AW1)
;About2 := Manager.AddText('xp-' AW1 ' yp cYellow BackgroundBlack', AboutText)
;About2.SetFont('Bold', 'Calibri')
;About2.GetPos(&AX2)
SB := Manager.AddStatusBar('cRed')
SB.SetFont('Bold', 'Calibri')
SB.SetParts(10, 50, 200)
SB.SetText('v' Version, 2)
SB.SetText('Loading...', 3)
SB.SetText(A_Tab A_Tab 'A Collective App From The Internet On What I Found Useful About AoE II!    ', 4)
Manager.Show('w700')
LoadCurrentSettings()
__CheckForUpdates__()
;AX := 0
;SetTimer(AnimateAbout, 10)
;AnimateAbout() {
;    Global AX
;    About1.Move(AX1 + (++AX))
;    About2.Move(AX2 + AX)
;}
Return

LoadCurrentSettings() {
    If !DirExist(GameFolder := IniRead(Config, 'Game', 'Path', '')) {
        Return
    }
    ChosenFolder.Value := GameFolder
    DisableGameRun()
    DisableVersions()
    DisableCompatibilitys()
    DisableLanguage()
    DisableVisualMod()
    If !FileExist(ChosenFolder.Value '\empires2.exe') {
        Expected := ChosenFolder.Value
        Loop Files, Expected '\empires2.exe', 'R' {
            ChosenFolder.Value := A_LoopFileDir
            Break
        }
        If !FileExist(ChosenFolder.Value '\empires2.exe') {
            SplitPath(Expected,, &Expected)
            If !FileExist(Expected '\empires2.exe') {
                Return
            }
            ChosenFolder.Value := Expected
        }
    }
    EnableAOKVersion()
    EnableAOKRun()
    EnableAOKCompatibility()
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

BackUpOrgSlp(VMName) {
    Slps := []
    Loop Files, 'DB\007\' VMName '\*.slp' {
        Slps.Push(A_LoopFileFullPath)
    }
    Loop Files, 'DB\007\' VMName '\*.slp' {
        Flag := SubStr(A_LoopFileName, 1, 3)
        If !FileExist('DB\007\' VMName '\U\' A_LoopFileName) {
            RunWait('DB\000\DrsBuild.exe /e "' ChosenFolder.Value '\Data\' DrsTypes[Flag] '" ' A_LoopFileName ' /o "DB\007\' VMName '\U"', , 'Hide')
        }
    }
}

AddPrefix(Slps, NL) {
    For Each, Slp in Slps {
        SplitPath(Slp, &OutFileName, &OutDir, , &OutNameNoExt)
        ID := OutNameNoExt
        If ID < 6000 {
            Prefix := 'gra'
        } Else If (ID >= 15000) && (ID <= 16000) {
            Prefix := 'ter'
        } Else If (ID >= 50000) && (ID <= 54000) {
            Prefix := 'int'
        } Else {
            FileDelete(Slp)
            Continue
        }
        If (Flag := SubStr(OutFileName, 1, 3)) != Prefix {
            Loop (NL - StrLen(OutNameNoExt))
                OutFileName := '0' OutFileName
            FileMove(Slp, OutDir '\' Prefix OutFileName, 1)
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
    If (Patch.Value) {
        If DirExist('DB\001\' Version) {
            DirCopy('DB\001\' Version, ChosenFolder.Value, 1)
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
    Loop Files, 'DB\001\' Edition '*', 'D' {
        Version := A_LoopFileName
        Loop Files, 'DB\001\' Version '\*.*', 'R' {
            PatchFile := A_LoopFileDir '\' A_LoopFileName
            GameFile := ChosenFolder.Value StrReplace(PatchFile, 'DB\001\' Version)
            If FileExist(GameFile) {
                FileDelete(GameFile)
            }
        }
    }
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
    Loop Files, 'DB\002\*', 'D' {
        Version := A_LoopFileName
        Found := True
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
                Case 1: General['AOC']['VersionsN'][Version].Value := 1
                Case 2: General['AOK']['VersionsN'][Version].Value := 1
            }
        }
    }
}

GameSectionNormalView() {
    GetTheGame.Visible := True
    ProgressBar.Visible := False
    ProgressInfo.Visible := False
}
GameSectionInstallView() {
    GetTheGame.Visible := False
    ProgressBar.Visible := True
    ProgressInfo.Visible := True
}

GameIsRunning() {
    For Each, Game in ['empires2.exe', 'age2_x1.exe', 'age2_x2.exe'] {
        If ProcessExist(Game) {
            ProcessClose(Game)
        }
        ProcessWaitClose(Game, 5)
        If ProcessExist(Game) {
            Return True
        }
    }
    Return False
}

EnableVisualMod() {
    _VisualMods_.Enabled := True
    VMList.Enabled := True
    VMList.Redraw()
    LoadVM.Enabled := True
}

DisableVisualMod() {
    _VisualMods_.Enabled := False
    VMList.Enabled := False
    LoadVM.Enabled := False
}

EnableLanguage() {
    _Language_.Enabled := True
    _Language_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > X && CX < (X + Width)) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := True
        }
    }
}
DisableLanguage() {
    _Language_.Enabled := False
    _Language_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > X && CX < (X + (Width / 3))) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := False
        }
    }
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

DisableGameRun() {
    DisableAOKRun()
    DisableAOCRun()
    DisableFOERun()
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

EnableCompatibilitys() {
    EnableAOKCompatibility()
    EnableAOCCompatibility()
    EnableFOECompatibility()
}
EnableAOKCompatibility() {
    _Compatibility_.Enabled := True
    _Compatibility_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > X && CX < (X + (Width / 3))) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := True
        }
    }
}
EnableAOCCompatibility() {
    _Compatibility_.Enabled := True
    _Compatibility_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > (X + (Width / 3)) && CX < (X + (Width * 2 / 3))) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := True
        }
    }
}
EnableFOECompatibility() {
    _Compatibility_.Enabled := True
    _Compatibility_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > (X + (Width * 2 / 3)) && CX < (X + Width)) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := True
        }
    }
}

DisableCompatibilitys() {
    DisableAOKCompatibility()
    DisableAOCCompatibility()
    DisableFOECompatibility()
}
DisableAOKCompatibility() {
    _Compatibility_.Enabled := False
    _Compatibility_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > X && CX < (X + (Width / 3))) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := False
        }
    }
}
DisableAOCCompatibility() {
    _Compatibility_.Enabled := False
    _Compatibility_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > (X + (Width / 3)) && CX < (X + (Width * 2 / 3))) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := False
        }
    }
}
DisableFOECompatibility() {
    _Compatibility_.Enabled := False
    _Compatibility_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > (X + (Width * 2 / 3)) && CX < (X + Width)) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := False
        }
    }
}

EnableVersions() {
    EnableAOKVersion()
    EnableAOCVersion()
    EnableFOEVersion()
}
EnableAOKVersion() {
    _Version_.Enabled := True
    _Version_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > X && CX < (X + (Width / 3))) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := True
        }
    }
}
EnableAOCVersion() {
    _Version_.Enabled := True
    _Version_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > (X + (Width / 3)) && CX < (X + (Width * 2 / 3))) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := True
        }
    }
}
EnableFOEVersion() {
    _Version_.Enabled := True
    _Version_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > (X + (Width * 2 / 3)) && CX < (X + Width)) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := True
        }
    }
}
DisableVersions() {
    DisableAOKVersion()
    DisableAOCVersion()
    DisableFOEVersion()
}
DisableAOKVersion() {
    _Version_.Enabled := False
    _Version_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > X && CX < (X + (Width / 3))) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := False
        }
    }
}
DisableAOCVersion() {
    _Version_.Enabled := False
    _Version_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > (X + (Width / 3)) && CX < (X + (Width * 2 / 3))) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := False
        }
    }
}
DisableFOEVersion() {
    _Version_.Enabled := False
    _Version_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > (X + (Width * 2 / 3)) && CX < (X + Width)) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := False
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

GRGamePath(TextFound, AppName) {
    P := InStr(TextFound, LFE := AppName,, -1)
    Loop {
        Char := SubStr(TextFound, P - (I := A_Index), 1)
        LFE := Char LFE
    } Until (Char = ':' || Ord(Char) = 10 || Ord(Char) = 13)
    Result := SubStr(TextFound, P - (I + 1), 1) LFE
    Return (FileExist(Result) ? Result : '')
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

__CheckForUpdates__() {
    Global Version
    If A_IsCompiled {
        Return
    }
    Try {
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
            If HashFile(A_LoopFileDir '\' A_LoopFileName) != HashsumsMap[A_LoopFileDir '\' A_LoopFileName] {
                FoundUpdates.Push(A_LoopFileDir '\' A_LoopFileName)
            }
        }
        If FoundUpdates.Length {
            UpdatesList := '`n'
            For Each, UpdateFile in FoundUpdates{
                UpdatesList .= '- ' UpdateFile '`n'
            }
            Choice := MsgBox('The following needs to be updated`n' UpdatesList '`nUpdate now?', 'Update', 0x4 + 0x20)
            If Choice = 'Yes' {
                For Each, UpdateFile in FoundUpdates {
                    DownloadLink := Server '/' User '/' Repo '/main/' StrReplace(StrReplace(UpdateFile, ' ', '%20'), '\', '/')
                    Download(DownloadLink, UpdateFile)
                }
                Reload
            }
        }
        SB.SetText('Up to date!', 3)
    } Catch As Err {
        SB.SetText('Failed to check for updates!', 3)
    }
}