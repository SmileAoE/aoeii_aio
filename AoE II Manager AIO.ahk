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
Server                      := 'https://raw.githubusercontent.com'
User                        := 'SmileAoE'
Repo                        := 'aoeii_aio'
Version                     := '1.2'
Layers                      := 'HKLM\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers'
Config                      := A_AppData '\aoeii_aio\config.ini'
AppDir                      :=    ['DB'
                                 , A_AppData '\aoeii_aio']
GRSetting                   := A_AppData '\GameRanger\GameRanger Prefs\Settings'
DrsTypes                    := Map('gra'        , 'graphics.drs'
                                 , 'int'        , 'interfac.drs'
                                 , 'ter'        , 'terrain.drs')
DrsRange                    := Map('gra'        , [2, 5312]
                                 , 'int'        , [50100, 53211]
                                 , 'ter'        , [15000, 15031])
IDL                         := 5
VCodedSlp                   := '3713EFBE'
NormalSlp                   := '322E304E'
General                     := Map()
General['AOK']              := Map()
General['AOK']['VersionsN'] := Map()
General['AOK']['Combine']   := Map('2.0b CD'    , ['2.0a No CD'])
General['AOC']              := Map()
General['AOC']['VersionsN'] := Map()
General['AOC']['Combine']   := Map('1.0e No CD' , ['1.0c No CD']
                                  ,'1.0e No CD' , ['1.0c No CD']
                                  ,'1.1  No CD' , ['1.0c No CD']
                                  ,'1.5  CD'    , ['1.0c No CD'])
General['FOE']              := Map()
General['FOE']['VersionsN'] := Map()
General['FOE']['Combine']   := Map()
General['LNG']              := Map()
Compatibilities             := Map(1            , ["_____Not Set_____"  , ""]
                                 , 2            , ["Windows 8"          , "WIN8RTM"]
                                 , 3            , ["Windows 7"          , "WIN7RTM"]
                                 , 4            , ["Windows Vista Sp2"  , "VISTASP2"]
                                 , 5            , ["Windows Vista Sp1"  , "VISTASP1"]
                                 , 6            , ["Windows Vista"      , "VISTARTM"]
                                 , 7            , ["Windows XP Sp3"     , "WINXPSP3"]
                                 , 8            , ["Windows XP Sp2"     , "WINXPSP2"]
                                 , 9            , ["Windows 98"         , "WIN98"]
                                 , 10           , ["Windows 95"         , "WIN95"])
BasePackages                :=   ['DB/000.7z.001'
                                 ,'DB/001.7z.001'
                                 ,'DB/002.7z.001'
                                 ,'DB/006.7z.001'
                                 ,'DB/007.7z.001'
                                 ,'DB/008.7z.001']
GamePackages                :=   ['DB/003.7z.001'
                                 ,'DB/003.7z.002'
                                 ,'DB/003.7z.003'
                                 ,'DB/003.7z.004'
                                 ,'DB/004.7z.001'
                                 ,'DB/004.7z.002'
                                 ,'DB/004.7z.003'
                                 ,'DB/005.7z.001']
Dots                        := 0
Task                        := 1
TaskNumber                  := BasePackages.Length
Features                    := Map()
;SysDrive                    := EnvGet('SystemDrive')
ProgramFiles86              := EnvGet(A_Is64bitOS ? "ProgramFiles(x86)" : "ProgramFiles")
VPNDir                      := ProgramFiles86 '\Hide ALL IP'
VPNExe                      := 'HideALLIP.exe'
VPNPath                     := VPNDir '\' VPNExe

; Preparation
CoordMode('Mouse', 'Screen')
GroupAdd('AOKAOC', 'ahk_exe empires2.exe')
GroupAdd('AOKAOC', 'ahk_exe age2_x1.exe')

; Create app folders
For _, Item in AppDir {
    If !DirExist(Item) {
        DirCreate(Item)
    }
}

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

; # The Game
Features['The Game'] := []
_Game_ := Manager.AddText('w220 h260 Center c800000 BackgroundFFFFFF Border', '# The Game')
Features['The Game'].Push(_Game_)
_Game_.SetFont('Bold')
GetTheGame := Manager.AddButton('xm+10 ym+25 w200', 'Download AoE II')
Features['The Game'].Push(GetTheGame)
GetTheGame.OnEvent('Click', (*) => DownloadInstallGame())
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
GuiButtonIcon(GetTheGame, 'DB\000\Down.png', , 'W16 H16 T2 A1')
ProgressBar := Manager.AddProgress('xp yp wp h20 Hidden', 0)
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
        GetTheGame.Visible      := False
        ProgressBar.Visible     := True
        ProgressInfo.Visible    := True
    }
    If !ExportDir := FileSelect('D') {
        GameSectionNormalView()
        GameSectionNormalView() {
            GetTheGame.Visible      := True
            ProgressBar.Visible     := False
            ProgressInfo.Visible    := False
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
            RunWait('DB\7za.exe x DB\00' (2 + A_Index) '.7z.001 -o"' ExportDir '\Age of Empires II" -aoa',, 'Hide')
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
RunAOK := Manager.AddButton('xm+30 yp+20 w48 H48')
Features['The Game'].Push(RunAOK)
GuiButtonIcon(RunAOK, 'DB\000\aok.png', , 'W32 H32')
RunAOK.OnEvent('Click', (*) => Run(ChosenFolder.Value '\empires2.exe', ChosenFolder.Value))

RunAOC := Manager.AddButton('yp wp hp')
Features['The Game'].Push(RunAOC)
GuiButtonIcon(RunAOC, 'DB\000\aoc.png', , 'W32 H32')
RunAOC.OnEvent('Click', (*) => Run(ChosenFolder.Value '\age2_x1\age2_x1.exe', ChosenFolder.Value '\age2_x1'))

RunFOE := Manager.AddButton('yp wp hp')
Features['The Game'].Push(RunFOE)
GuiButtonIcon(RunFOE, 'DB\000\fe.png', , 'W32 H32')
RunFOE.OnEvent('Click', (*) => Run(ChosenFolder.Value '\age2_x1\age2_x2.exe', ChosenFolder.Value '\age2_x1'))

ChooseFolder := Manager.AddButton('xm+10 yp+60 w30 w100', 'Choose')
Features['The Game'].Push(ChooseFolder)
ChooseFolder.OnEvent('Click', (*) => SelectTheGame())
GuiButtonIcon(ChooseFolder, 'DB\000\Folder.png', , 'W16 H16 T1 A1')

LoadGRFolder := Manager.AddButton('xm+180 yp w30 w30')
Features['The Game'].Push(LoadGRFolder)
LoadGRFolder.OnEvent('Click', (*) => SelectTheGameFromGR())
GuiButtonIcon(LoadGRFolder, 'DB\000\GR.png', , 'W16 H16')
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
    AOKDir      := ''
    AOCDir      := ''
    FOEDir      := ''
    ChosenDir   := ''
    GRGamePath(TextFound, AppName) {
        P := InStr(TextFound, LFE := AppName,, -1)
        Loop {
            Char := SubStr(TextFound, P - (I := A_Index), 1)
            LFE := Char LFE
        } Until (Char = ':' || Ord(Char) = 10 || Ord(Char) = 13)
        Result := SubStr(TextFound, P - (I + 1), 1) LFE
        Return (FileExist(Result) ? Result : '')
    }
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
ChosenFolder := Manager.AddEdit('xm+10 yp+30 w200 Center ReadOnly r4 -VScroll cBlue BackgroundWhite')
Features['The Game'].Push(ChosenFolder)
OpenTheGameFolder := Manager.AddButton('w200', 'Open')
Features['The Game'].Push(OpenTheGameFolder)
GuiButtonIcon(OpenTheGameFolder, 'DB\000\Folder.png', , 'W16 H16 T1 A1')
OpenTheGameFolder.OnEvent('Click', (*) => Run(ChosenFolder.Value))
SelectTheGame() {
    SelectAFolder()
    ChargeSettings________(True)
}
SelectAFolder() {
    ChosenDir := FileSelect('D', 'C:\' (A_Is64bitOS ? 'Program Files (x86)' : 'Program Files') '\Microsoft Games')
    If !ChosenDir
        Return
    IniWrite(ChosenDir, Config, 'Game', 'Path')
    ChosenFolder.Value := ChosenDir
}

; # Versions
Features['Versions'] := []
_Version_ := Manager.AddText('ym w450 h220 Center c800000 BackgroundFFFFFF Border', '# Versions')
Features['Versions'].Push(_Version_)
_Version_.SetFont('Bold')
H := Manager.AddPicture('xp+54 ym+25 BackgroundTrans', 'DB\000\aok.png')
Features['Versions'].Push(H)
H := Manager.AddText('xp-44 yp+40 cRed w120 Center BackgroundTrans', 'The Age of Kings')
Features['Versions'].Push(H)
H.SetFont('Bold')
H := Manager.AddText('xp+20 yp+20 w1 h1 BackgroundTrans')
Features['Versions'].Push(H)
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
    DirCopy('DB\001\' Patch.Text '\Static' , ChosenFolder.Value, 1)
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
H := Manager.AddText('xp+20 yp+20 w1 h1 BackgroundTrans')
Features['Versions'].Push(H)
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
H := Manager.AddText('xp+20 yp+20 w1 h1 BackgroundTrans')
Features['Versions'].Push(H)
Handle := Manager.AddRadio('w30 w100 Checked BackgroundFFFFFF', '2.2  CD')
Features['Versions'].Push(Handle)
Handle.SetFont('s10', 'Consolas')
General['FOE']['VersionsN']['2.2  CD'] := Handle
Patch := Manager.AddDropDownList('xm+240 ym+195 w430', ['Do Not Enable Fixes'])
Features['Versions'].Push(Patch)
Patch.OnEvent('Change', (*) => IniWrite(Patch.Text, Config, 'Game', 'Fix'))

; # Compatibilities
Features['Compatibilities'] := []
_Compatibility_ := Manager.AddText('xm+230 ym+225 w450 h140 Center c800000 BackgroundFFFFFF Border', '# Compatibilities')
Features['Compatibilities'].Push(_Compatibility_)
_Compatibility_.SetFont('Bold')
H := Manager.AddPicture('xp+54 yp+25 BackgroundTrans', 'DB\000\aok.png')
Features['Compatibilities'].Push(H)
H := Manager.AddText('xp-44 yp+40 cRed w120 Center BackgroundTrans', 'The Age of Kings')
Features['Compatibilities'].Push(H)
H.SetFont('Bold')
AoKCom := Manager.AddDropDownList('xp yp+20 w120')
Features['Compatibilities'].Push(AoKCom)
For Each, Compat in Compatibilities {
    AoKCom.Add([Compat[1]])
}
AoKCom.Choose(1)
AoKCom.OnEvent("Change", (*) => AoKComReg())
AoKRun := Manager.AddCheckbox('yp+30 wp hp BackgroundFFFFFF', 'Run as administrator')
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
H := Manager.AddPicture('xp+194 yp-90 BackgroundTrans', 'DB\000\aoc.png')
Features['Compatibilities'].Push(H)
H := Manager.AddText('xp-44 yp+40 cBlue w120 Center BackgroundTrans', 'The Conquerors')
Features['Compatibilities'].Push(H)
H.SetFont('Bold')
AoCCom := Manager.AddDropDownList('xp yp+20 w120')
Features['Compatibilities'].Push(AoCCom)
For Each, Compat in Compatibilities {
    AoCCom.Add([Compat[1]])
}
AoCCom.Choose(1)
AoCCom.OnEvent("Change", (*) => AoCComReg())
AoCRun := Manager.AddCheckbox('yp+30 wp hp BackgroundFFFFFF', 'Run as administrator')
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
H := Manager.AddPicture('xp+194 yp-90 BackgroundTrans', 'DB\000\fe.png')
Features['Compatibilities'].Push(H)
H := Manager.AddText('xp-44 yp+40 cGreen w120 Center BackgroundTrans', 'Forgotten Empires')
Features['Compatibilities'].Push(H)
H.SetFont('Bold')
FOECom := Manager.AddDropDownList('xp yp+20 w120')
Features['Compatibilities'].Push(FOECom)
For Each, Compat in Compatibilities {
    FOECom.Add([Compat[1]])
}
FOECom.Choose(1)
FOECom.OnEvent("Change", (*) => FOEComReg())
FOERun := Manager.AddCheckbox('yp+30 wp hp BackgroundFFFFFF', 'Run as administrator')
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
_Language_ := Manager.AddText('xm yp-75 w220 h385 Center c800000 BackgroundFFFFFF Border', '# Languages')
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
_VisualMods_ := Manager.AddText('xm+230 yp-255 w220 h280 Center c800000 BackgroundFFFFFF Border', '# Visual Mods')
Features['Visual Modes'].Push(_VisualMods_)
_VisualMods_.SetFont('Bold')
H := Manager.AddText('xp+10 yp+10 w200 BackgroundTrans')
Features['Visual Modes'].Push(H)
VMList := Manager.AddListView('w200 h210 -E0x200 -Hdr Checked BackgroundFFFFFF', ['Mode Name'])
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
    SectionInteract(Features['Visual Modes'])
    SoundPlay('DB\000\30 wololo.mp3')
}
LoadVM := Manager.AddButton('wp Disabled', 'Import')
LoadVM.SetFont('Bold')
LoadVM.OnEvent('Click', (*) => ImportVisualMod())
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
    LoadVM.Enabled := True
}

; # Data Modes
Features['Data Modes'] := []
_VisualMods_.GetPos(, &Y)
_DataModes_ := Manager.AddText('xm+460 y' Y ' w220 h280 Center c800000 BackgroundFFFFFF Border', '# Data Mods')
Features['Data Modes'].Push(_DataModes_)
_DataModes_.SetFont('Bold')
H := Manager.AddText('xp+10 yp+10 w200 BackgroundTrans')
Features['Data Modes'].Push(H)
DMList := Manager.AddListView('w200 h210 -E0x200 -Hdr Checked BackgroundFFFFFF', ['Mode Name'])
Features['Data Modes'].Push(DMList)
DMList.SetFont('Bold')
For Each, Mode in StrSplit(IniRead('DB\008\DataMode.ini', 'DataMode',, ''), '`n') {
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
        Parts   := StrSplit(ModeDir[2], ',')
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
ImportDM := Manager.AddButton('wp Disabled', 'Import')
ImportDM.SetFont('Bold')
ImportDM.OnEvent('Click', (*) => ImportDataMode())
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
                FileMove(A_LoopFileFullPath,  Dir '\' ModeName '\Drs\' LZID '.' A_LoopFileExt)
            }
            RunWait('DB\000\DrsBuild.exe /a "' Dir '\' ModeName '\Data\gamedata_x1_p1.drs" "'  Dir '\' ModeName '\Drs\gam*.*"',, 'Hide')
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
_ATools_ := Manager.AddText('ym w220 h590 Center c800000 BackgroundFFFFFF Border', '# Other Tools`n`n')
Features['Other Tools'].Push(_ATools_)
_ATools_.SetFont('Bold')
_ATools_.GetPos(&X, &Y, &Width, &Height)

H := Manager.AddText('xp+10 yp+10 w200 BackgroundTrans')
Features['Other Tools'].Push(H)

; # Shortcut send & un-select one unit
Macro1 := Manager.AddCheckBox('xp yp+20 w200 BackgroundWhite Center' , '{Left Alt + Right Mouse Button}`n[Send && un-select one unit]')
Features['Other Tools'].Push(Macro1)
Macro1.SetFont('Bold')
H := Manager.AddText('x' (X + 1) ' yp+35 w220 0x10')
Features['Other Tools'].Push(H)
#HotIf WinActive('ahk_group AOKAOC')
Hotkey('!RButton', Macro1Action, 'Off')
Macro1Action(*) {
    MouseClick('Right', , , , 0)
    MouseGetPos(&X, &Y)
    SendInput('{LCtrl Down}')
    MouseClick('Left', 315, A_ScreenHeight - 130, , 0)
    SendInput('{Ctrl Up}')
    MouseMove(X, Y, 0)
}
Macro1.OnEvent('Click', (*) => EDMacro1())
EDMacro1() {
    IniWrite(Macro1.Value, Config, 'Game', 'Macro1')
    Hotkey('!RButton', Macro1Action, Macro1.Value ? 'On' : 'Off')
}

VPN := Manager.AddButton('xp+10 yp+10 w56 h56')
Features['Other Tools'].Push(VPN)
GuiButtonIcon(VPN, 'DB\000\vpn.png',, 'W48 H48')
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

ChargeOtherTools______() {
    Macro1On := IniRead(Config, 'Game', 'Macro1', 0)
    Macro1.Value := Macro1On
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

_ATools_.GetPos(&X, &Y,, &Height)
ProgressText := Manager.AddText('x' X ' y' (Y + Height + 10) ' w220 h20 Center BackgroundTrans cBlue', '...')
ProgressText.SetFont('Bold')
Progress := Manager.AddProgress('xp yp+20 wp hp -Smooth')

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
ChargeSettings________()
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
CheckForUpdates_______()
Return

ChargeSettings________(Browse := False) {
    SoundPlay('DB\000\30 wololo.mp3')
    SectionInteract(Features['Versions']        , False)
    SectionInteract(Features['Compatibilities'] , False)
    SectionInteract(Features['Language']        , False)
    SectionInteract(Features['Visual Modes']    , False)
    SectionInteract(Features['Data Modes']      , False)
    SectionInteract(Features['Other Tools']     , False)
    ChosenFolder.Value := IniRead(Config, 'Game', 'Path', '')
    ValidGameLocation(Location) {
        Return FileExist(Location '\empires2.exe')
            && FileExist(Location '\language.dll')
            && FileExist(Location '\Data\graphics.drs')
            && FileExist(Location '\Data\interfac.drs')
            && FileExist(Location '\Data\terrain.drs')
    }
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
            For Each, UpdateFile in FoundUpdates{
                UpdatesList .= Each ' - ' UpdateFile '`n'
            }
            UpdatesList .= '`n=======================`n'
            SB.SetText('Update found!', 3)
            Choice := MsgBox('The following needs to be updated:`n' UpdatesList '`nUpdate now?', 'New Update!', 0x4 + 0x40)
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
                    Buff    := FileRead(UpdateFile, 'RAW')
                    Str     := StrGet(Buff, 2, '')
                    If (Str != '7z') {
                        Continue
                    }
                    Name := StrSplit(UpdateFile, '.')
                    DirDelete(Name[1], 1)
                    RunWait('DB\7za.exe x ' Name[1] '.7z.001 -o' Name[1] ' -aoa',, 'Hide')
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