#Requires AutoHotkey v2.0
#SingleInstance Force
Server := 'https://raw.githubusercontent.com'
User := 'SmileAoE'
Repo := 'aoeii_aio'
Layers := 'HKLM\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers'
Config := A_AppData '\aoeii_aio\config.ini'
AppDir := ['DB', A_AppData '\aoeii_aio']
Dots := 0
Task := 1
TaskNumber := 8

General := Map()

General['AOK'] := Map()
General['AOK']['VersionsN'] := Map()
General['AOK']['Combine'] := Map('2.0b CD' , ['2.0a No CD'])

General['AOC'] := Map()
General['AOC']['VersionsN'] := Map()
General['AOC']['Combine'] := Map('1.0e No CD' , ['1.0c No CD']
                               , '1.0e No CD' , ['1.0c No CD']
                               , '1.1  No CD' , ['1.0c No CD']
                               , '1.5  CD'    , ['1.0c No CD'])

General['FOE'] := Map()
General['FOE']['VersionsN'] := Map()
General['FOE']['Combine'] := Map()

Compatibilities := Map(1 , [ "_____Not Set_____" , ""         ]
                     , 2 , [ "Windows 8"         , "WIN8RTM"  ]
                     , 3 , [ "Windows 7"         , "WIN7RTM"  ]
                     , 4 , [ "Windows Vista Sp2" , "VISTASP2" ]
                     , 5 , [ "Windows Vista Sp1" , "VISTASP1" ]
                     , 6 , [ "Windows Vista"     , "VISTARTM" ]
                     , 7 , [ "Windows XP Sp2"    , "WINXPSP2" ]
                     , 8 , [ "Windows 98"        , "WIN98"    ]
                     , 9 , [ "Windows 95"        , "WIN95"    ])
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
        HoldOn.Text := PW '`n' ((Mod(++Dots, 4) = 0) ? '●' 
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

    SetTimer(ShowInfo, 0)
    Prepare.Destroy()
} Catch As Err {
    MsgBox('There was an error while preparing the necessary files!', 'Oops!', '48 T5')
    ExitApp
}
; Main window
Manager := Gui(, 'AoE II Manager AIO')
Manager.OnEvent('Close', (*) => ExitApp())
Manager.BackColor := 0xFFFFFF
; First section: # The Game
_Game_ := Manager.AddGroupBox('w220 h260 Right', '# The Game')
_Game_.SetFont('Bold')
GetTheGame := Manager.AddButton('xm+10 ym+25 w200', 'Download AoE II')
GetTheGame.OnEvent('Click', (*) => DownloadInstallGame())
GuiButtonIcon(GetTheGame, 'DB\000\Down.ico', , 'W16 H16 T2 A1')
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

ChooseFolder := Manager.AddButton('xm+10 yp+60 w30 w200', 'Choose')
ChooseFolder.OnEvent('Click', (*) => SelectTheGame())
GuiButtonIcon(ChooseFolder, 'DB\000\Folder.ico', , 'W16 H16 T2 A1')

ChosenFolder := Manager.AddEdit('xm+10 yp+30 w200 Center ReadOnly r4 -VScroll cBlue BackgroundWhite')

OpenTheGameFolder := Manager.AddButton('w200', 'Open')
GuiButtonIcon(OpenTheGameFolder, 'DB\000\Folder.ico', , 'W16 H16 T2 A1')
OpenTheGameFolder.OnEvent('Click', (*) => Run(ChosenFolder.Value))

SelectTheGame() {
    ChosenDir := FileSelect('D')
    If !ChosenDir
        Return
    IniWrite(ChosenDir, Config, 'Game', 'Path')
    ChosenFolder.Value := ChosenDir
    LoadCurrentSettings()
}
; Second Section: # Versions
_Version_ := Manager.AddGroupBox('ym w450 h220 Right', '# Versions')
_Version_.SetFont('Bold')

Manager.AddPicture('xp+54 ym+25', 'DB\000\aok.png')
Manager.AddText('xp-44 yp+40 cRed w120 Center', 'The Age of Kings').SetFont('Bold')
Manager.AddText('xp+20 yp+20 w1 h1')
Loop Files, 'DB\002\2*', 'D' {
    Handle := Manager.AddRadio('w30 w100', A_LoopFileName)
    Handle.SetFont('s10', 'Consolas')
    Handle.OnEvent('Click', ApplyVersion)
    General['AOK']['VersionsN'][A_LoopFileName] := Handle
}

Manager.AddPicture('xp+174 ym+25', 'DB\000\aoc.png')
Manager.AddText('xp-44 yp+40 cBlue w120 Center', 'The Conquerors').SetFont('Bold')
Manager.AddText('xp+20 yp+20 w1 h1')
Loop Files, 'DB\002\1*', 'D' {
    Handle := Manager.AddRadio('w30 w100', A_LoopFileName)
    Handle.SetFont('s10', 'Consolas')
    Handle.OnEvent('Click', ApplyVersion)
    General['AOC']['VersionsN'][A_LoopFileName] := Handle
}

Manager.AddPicture('xp+174 ym+25', 'DB\000\fe.png')
Manager.AddText('xp-44 yp+40 cGreen w120 Center', 'Forgotten Empires').SetFont('Bold')
Manager.AddText('xp+20 yp+20 w1 h1')
Handle := Manager.AddRadio('w30 w100 Checked', '2.2  CD')
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

Patch := Manager.AddCheckbox('xm+240 ym+200' (IniRead(Config, 'Game', 'Fix', 1) ? ' Checked' : ''), 'Enable fixs after each patching if available')
Patch.OnEvent('Click', (*) => IniWrite(Patch.Value, Config, 'Game', 'Fix'))

_Compatibility_ := Manager.AddGroupBox('xm+230 ym+225 w450 h140 Right', '# Compatibilities')
_Compatibility_.SetFont('Bold')
Manager.AddPicture('xp+54 yp+25', 'DB\000\aok.png')
Manager.AddText('xp-44 yp+40 cRed w120 Center', 'The Age of Kings').SetFont('Bold')
AoKCom := Manager.AddDropDownList('xp yp+20 w120')
For Each, Compat in Compatibilities {
    AoKCom.Add([Compat[1]])
}
AoKCom.Choose(1)
AoKCom.OnEvent("Change", (*) => AoKComReg())
AoKRun := Manager.AddCheckbox('yp+30 wp hp', 'Run as administrator')
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

Manager.AddPicture('xp+194 yp-90', 'DB\000\aoc.png')
Manager.AddText('xp-44 yp+40 cBlue w120 Center', 'The Conquerors').SetFont('Bold')
AoCCom := Manager.AddDropDownList('xp yp+20 w120')
For Each, Compat in Compatibilities {
    AoCCom.Add([Compat[1]])
}
AoCCom.Choose(1)
AoCCom.OnEvent("Change", (*) => AoCComReg())
AoCRun := Manager.AddCheckbox('yp+30 wp hp', 'Run as administrator')
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

Manager.AddPicture('xp+194 yp-90', 'DB\000\fe.png')
Manager.AddText('xp-44 yp+40 cGreen w120 Center', 'Forgotten Empires').SetFont('Bold')
FECom := Manager.AddDropDownList('xp yp+20 w120')
For Each, Compat in Compatibilities {
    FECom.Add([Compat[1]])
}
FECom.Choose(1)
FECom.OnEvent("Change", (*) => FEComReg())
FERun := Manager.AddCheckbox('yp+30 wp hp', 'Run as administrator')
FERun.OnEvent("Click", (*) => FEComReg())
FEComReg() {
    RegVal := Compatibilities[FECom.Value][2] (Compatibilities[FECom.Value][2] ? ' ' : '') (FERun.Value ? 'RUNASADMIN' : '')
    If !RegVal {
        Try {
            RegDelete(Layers, ChosenFolder.Value '\age2_x1\age2_x2.exe')
        }
        Return
    }
    RegWrite(RegVal, 'REG_SZ', Layers, ChosenFolder.Value '\age2_x1\age2_x2.exe')
}

_Language_ := Manager.AddGroupBox('xm yp-75 w220 h100 Right', '# Language')
_Language_.SetFont('Bold')
Language := Manager.AddDropDownList('xp+10 yp+40 w200')

Manager.AddStatusBar(, 'v1.0')
Manager.Show()
LoadCurrentSettings()
Return

CompatibilityCheck() {
    AoKReg := RegRead(Layers, ChosenFolder.Value '\empires2.exe', '')
    If AoKReg {
        AoKReg := StrSplit(AoKReg, ' ')
        For Each, RegVal in AoKReg {
            If (RegVal = 'RUNASADMIN')
                AoKRun.Value := 1
            Else {
                For Each, Compat in Compatibilities {
                    If (Compat[2] = RegVal)
                        AoKCom.Choose(Compat[1])
                }
            }
        }
    }
    AoCReg := RegRead(Layers, ChosenFolder.Value '\age2_x1\age2_x1.exe', '')
    If AoCReg {
        AoCReg := StrSplit(AoCReg, ' ')
        For Each, RegVal in AoCReg {
            If (RegVal = 'RUNASADMIN')
                AoCRun.Value := 1
            Else {
                For Each, Compat in Compatibilities {
                    If (Compat[2] = RegVal)
                        AoCCom.Choose(Compat[1])
                }
            }
        }
    }
    FEReg := RegRead(Layers, ChosenFolder.Value '\age2_x1\age2_x2.exe', '')
    If FEReg {
        FEReg := StrSplit(FEReg, ' ')
        For Each, RegVal in FEReg {
            If (RegVal = 'RUNASADMIN')
                FERun.Value := 1
            Else {
                For Each, Compat in Compatibilities {
                    If (Compat[2] = RegVal)
                        FECom.Choose(Compat[1])
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

LoadCurrentSettings() {
    If DirExist(GameFolder := IniRead(Config, 'Game', 'Path', '')) {
        ChosenFolder.Value := GameFolder
    }
    DisableGameRun()
    DisableVersions()
    DisableCompatibilitys()
    If FileExist(ChosenFolder.Value '\empires2.exe') {
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
            If FileExist(GameFile) && HashFile(PatchFile) = HashFile(GameFile) {
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

DisableCompatibilitys() {
    DisableAOKCompatibility()
    DisableAOCCompatibility()
    DisableFOECompatibility()
}
DisableAOKCompatibility() {
    _Compatibility_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > X && CX < (X + (Width / 3))) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := False
        }
    }
}
DisableAOCCompatibility() {
    _Compatibility_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > (X + (Width / 3)) && CX < (X + (Width * 2 / 3))) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := False
        }
    }
}
DisableFOECompatibility() {
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
DisableVersions() {
    DisableAOKVersion()
    DisableAOCVersion()
    DisableFOEVersion()
}
DisableAOKVersion() {
    _Version_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > X && CX < (X + (Width / 3))) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := False
        }
    }
}
DisableAOCVersion() {
    _Version_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > (X + (Width / 3)) && CX < (X + (Width * 2 / 3))) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := False
        }
    }
}
DisableFOEVersion() {
    _Version_.GetPos(&X, &Y, &Width, &Height)
    For Each, Control in Manager {
        Control.GetPos(&CX, &CY)
        If (CX > (X + (Width * 2 / 3)) && CX < (X + Width)) && (CY > Y && CY < (Y + Height)) {
            Control.Enabled := False
        }
    }
}

#Include DB\000\lib\ButtonIcon.ahk
#Include DB\000\lib\HashFile.ahk
#Include DB\000\lib\ConnectedToInternet.ahk