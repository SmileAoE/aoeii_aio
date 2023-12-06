#Requires AutoHotkey v2.0
#SingleInstance Force
;--------------------------------------------
Server := 'https://raw.githubusercontent.com'
User := 'SmileAoE'
Repo := 'aoeiigrdb'
Layers := 'HKLM\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers'
Versions := Map()
Compatibilities := Map(), C := 0
Compats := Map( 1 , [ "_____Not Set_____" , ""         ]
              , 2 , [ "Windows 8"         , "WIN8RTM"  ]
              , 3 , [ "Windows 7"         , "WIN7RTM"  ]
              , 4 , [ "Windows Vista Sp2" , "VISTASP2" ]
              , 5 , [ "Windows Vista Sp1" , "VISTASP1" ]
              , 6 , [ "Windows Vista"     , "VISTARTM" ]
              , 7 , [ "Windows XP Sp2"    , "WINXPSP2" ]
              , 8 , [ "Windows 98"        , "WIN98"    ]
              , 9 , [ "Windows 95"        , "WIN95"    ] )
Combine := Map('2.0b CD'    , ['2.0a No CD']
             , '1.0e No CD' , ['1.0c No CD']
             , '1.1  No CD' , ['1.0c No CD']
             , '1.5  CD'    , ['1.0c No CD'])
;---------------------------------------------------
AppDir := ['DB', A_AppData '\aoeiigrdb']
Try {
    Info := Gui('-MinimizeBox', 'Preparing...'), Info.OnEvent('Close', (*) => ExitApp())
    InfoText := Info.AddText('Center w400 h40', PW := 'Please Wait')
    InfoText.SetFont('s12 Bold')
    ProgText := Info.AddText('Center w400 h30 cRed')
    ProgText.SetFont('Bold')
    Info.Show()
    SetTimer(ShowInfo, 500), T := 0, S := 1, SF := 8
    ShowInfo() {
        Global T, S
        InfoText.Text := PW '`n' ((Mod(T, 4) = 0) ? '●' : ((Mod(T, 4) = 1) ? '●●' : ((Mod(T, 4) = 2) ? '●●●' : '●●●●')))
        ProgText.Text := S ' / ' SF ' of prepare steps (is/are) done'
        ++T
    }
    For Each, Folder in AppDir {
        If !DirExist(Folder) {
            DirCreate(Folder)
        }
    }
    If !FileExist('DB\7za.exe') {
        Download(Server '/' User '/' Repo '/main/DB/7za.exe', 'DB\7za.exe')
    }
    ++S
    If !FileExist('DB\000.7z.001') {
        Download(Server '/' User '/' Repo '/main/DB/000.7z.001', 'DB\000.7z.001')
    }
    ++S
    If !FileExist('DB\001.7z.001') {
        Download(Server '/' User '/' Repo '/main/DB/001.7z.001', 'DB\001.7z.001')
    }
    ++S
    If !FileExist('DB\002.7z.001') {
        Download(Server '/' User '/' Repo '/main/DB/002.7z.001', 'DB\002.7z.001')
    }
    ++S
    ;-----------------------
    If !DirExist('DB\000') {
        RunWait('DB\7za.exe x DB\000.7z.001 -oDB\000', , 'Hide')
    }
    ++S
    If !DirExist('DB\001') {
        RunWait('DB\7za.exe x DB\001.7z.001 -oDB\001', , 'Hide')
    }
    ++S
    If !DirExist('DB\002') {
        RunWait('DB\7za.exe x DB\002.7z.001 -oDB\002', , 'Hide')
    }
    ++S
    SetTimer(ShowInfo, 0)
    Info.Destroy()
} Catch As Err {
    MsgBox('There was an error while preparing the necessary files!', 'Oops!', '48 T5')
    ExitApp
}
Manager := Gui(, 'AoE II Manager AIO'), Manager.OnEvent('Close', (*) => ExitApp())
Manager.BackColor := 0xFFFFFF
Manager.AddGroupBox('w220 h260 Right', '# The Game').SetFont('Bold')
;------------------------------------------------------------------
GetGame := Manager.AddButton('xm+10 ym+25 w200', 'Download AoE II')
GuiButtonIcon(GetGame, 'DB\000\Down.ico', , 'W16 H16 T2 A1')
GetGame.OnEvent('Click', (*) => DownloadInstallGame())
DownloadInstallGame() {
    If !OutGame := FileSelect('D') {
        Return
    }
    OutGame := RTrim(OutGame, '\')
    Sw := ''
    If DirExist(OutGame '\Age of Empires II') {
        Rsd := MsgBox('Game seems to be already exported at this location!`n`nOverwrite?', 'Game Exist', 0x4 + 0x30)
        If Rsd != 'Yes' {
            Return
        }
        Sw := '-aoa'
    }
    GetGame.Visible := False
    DIGame.Visible := True
    DIGameText.Opt('cDefault')
    DIGameText.Visible := True
    DIGameText.Value := ''
    DIGame.Value := 0
    DIGame.Opt('Range0-' R := 7)
    Try {
        ; The Age Of Kings
        Loop 4 {
            If !FileExist('DB\003.7z.00' A_Index)
                Download(Server '/' User '/' Repo '/main/DB/003.7z.00' A_Index, 'DB\003.7z.00' A_Index)
            DIGameText.Value := 'Downloaded [ ' Round((++DIGame.Value / R) * 100) ' % ]'
        }
        ; The Conquerors
        Loop 2 {
            If !FileExist('DB\004.7z.00' A_Index)
                Download(Server '/' User '/' Repo '/main/DB/004.7z.00' A_Index, 'DB\004.7z.00' A_Index)
            DIGameText.Value := 'Downloaded [ ' Round((++DIGame.Value / R) * 100) ' % ]'
        }
        ; Forgotten Empires
        Loop 1 {
            If !FileExist('DB\005.7z.00' A_Index)
                Download(Server '/' User '/' Repo '/main/DB/005.7z.00' A_Index, 'DB\005.7z.00' A_Index)
            DIGameText.Value := 'Downloaded [ ' Round((++DIGame.Value / R) * 100) ' % ]'
        }
        DIGameText.Opt('cGreen')
        DIGame.Opt('Range0-4')
        DIGame.Value := 0
        ++DIGame.Value
        ; Export The Age Of Kings
        DIGameText.Value := 'Exporting The Age Of Kings...'
        RunWait('DB\7za.exe x DB\003.7z.001 -o"' OutGame '\Age of Empires II" ' Sw, , 'Hide')
        ++DIGame.Value
        ; Export The Conquerors
        DIGameText.Value := 'Exporting The Conquerors...'
        RunWait('DB\7za.exe x DB\004.7z.001 -o"' OutGame '\Age of Empires II" ' Sw, , 'Hide')
        ++DIGame.Value
        ; Export Forgotten Empires
        DIGameText.Value := 'Exporting Forgotten Empires...'
        RunWait('DB\7za.exe x DB\005.7z.001 -o"' OutGame '\Age of Empires II" ' Sw, , 'Hide')
        ++DIGame.Value
    } Catch As Err {
        DownloadDefaultView()
        MsgBox('Unable to get the game!', 'Oops!', '48')
        Return
    }
    DownloadDefaultView()
    Rsd := MsgBox('Done!`n`nGame located at: "' OutGame '\Age of Empires II"`n`nWanna select this game location?', 'Question', 0x20 + 0x4)
    If Rsd = 'Yes' {
        ChosenFolder.Value := OutGame '\Age of Empires II'
        IniWrite(ChosenFolder.Value, A_AppData '\aoeiigrdb\config.ini', 'Game', 'Path')
        GameFolderValid()
    }
}
DownloadDefaultView() {
    GetGame.Visible := True
    DIGame.Visible := False
    DIGameText.Visible := False
}
;------------------------------------------------------
DIGame := Manager.AddProgress('xp yp wp h20 Hidden', 0)
DIGameText := Manager.AddText('xp yp+25 wp Hidden Center')
AoKLogo := Manager.AddButton('xm+30 yp+20 w48 H48')
GuiButtonIcon(AoKLogo, 'DB\000\aok.png', , 'W32 H32')
AoKLogo.OnEvent('Click', (*) => Run(ChosenFolder.Value '\empires2.exe', ChosenFolder.Value))
AoCLogo := Manager.AddButton('yp wp hp')
GuiButtonIcon(AoCLogo, 'DB\000\aoc.png', , 'W32 H32')
AoCLogo.OnEvent('Click', (*) => Run(ChosenFolder.Value '\age2_x1\age2_x1.exe', ChosenFolder.Value '\age2_x1'))
FELogo := Manager.AddButton('yp wp hp')
GuiButtonIcon(FELogo, 'DB\000\fe.png', , 'W32 H32')
FELogo.OnEvent('Click', (*) => Run(ChosenFolder.Value '\age2_x1\age2_x2.exe', ChosenFolder.Value '\age2_x1'))
ChooseFolder := Manager.AddButton('xm+10 yp+60 w30 w200', 'Choose')
GuiButtonIcon(ChooseFolder, 'DB\000\Folder.ico', , 'W16 H16 T2 A1')
ChooseFolder.OnEvent('Click', (*) => SelectTheGame())
SelectTheGame() {
    Selected := FileSelect('D')
    If !Selected
        Return
    IniWrite(Selected, A_AppData '\aoeiigrdb\config.ini', 'Game', 'Path')
    ChosenFolder.Value := Selected
    GameFolderValid()
    CurrentVersions()
}
ChosenFolder := Manager.AddEdit('xm+10 yp+30 w200 Center ReadOnly r4 -VScroll cBlue')
OpenFolder := Manager.AddButton('w200', 'Open')
GuiButtonIcon(OpenFolder, 'DB\000\Folder.ico', , 'W16 H16 T2 A1')
OpenFolder.OnEvent('Click', (*) => Run(ChosenFolder.Value))

Manager.AddGroupBox('ym w450 h220 Right', '# Versions').SetFont('Bold')
Manager.AddPicture('xp+54 ym+25', 'DB\000\aok.png')
Manager.AddText('xp-44 yp+40 cRed w120 Center', 'The Age of Kings').SetFont('Bold')
Manager.AddText('xp+20 yp+20 w1 h1')
Loop Files, 'DB\002\2*', 'D' {
    Versions[A_LoopFileName] := Manager.AddRadio('w30 w100', A_LoopFileName)
    Versions[A_LoopFileName].SetFont('s10', 'Consolas')
    Versions[A_LoopFileName].OnEvent('Click', ApplyVersion)
}

Manager.AddPicture('xp+174 ym+25', 'DB\000\aoc.png')
Manager.AddText('xp-44 yp+40 cBlue w120 Center', 'The Conquerors').SetFont('Bold')
Manager.AddText('xp+20 yp+20 w1 h1')
Loop Files, 'DB\002\1*', 'D' {
    Versions[A_LoopFileName] := Manager.AddRadio('w30 w100', A_LoopFileName)
    Versions[A_LoopFileName].SetFont('s10', 'Consolas')
    Versions[A_LoopFileName].OnEvent('Click', ApplyVersion)
}

ApplyVersion(Ctrl, Info) {
    HoldRadios()
    If !CloseGameCheck() {
        CurrentVersions()
        Return
    }
    CleanUp(Ctrl.Text)
    If Combine.Has(Ctrl.Text) {
        For Each, Version in Combine[Ctrl.Text] {
            DirCopy('DB\002\' Version, ChosenFolder.Value, 1)
        }
    }
    DirCopy('DB\002\' Ctrl.Text, ChosenFolder.Value, 1)
    If Patch.Value {
        If DirExist('DB\001\' Ctrl.Text) {
            DirCopy('DB\001\' Ctrl.Text, ChosenFolder.Value, 1)
        }
    }
    HoldRadios(False)
    SoundPlay('DB\000\30 wololo.mp3')
}
Manager.AddPicture('xp+174 ym+25', 'DB\000\fe.png')
Manager.AddText('xp-44 yp+40 cGreen w120 Center', 'Forgotten Empires').SetFont('Bold')
Manager.AddText('xp+20 yp+20 w1 h1')
Versions['2.2  CD'] := Manager.AddRadio('w30 w100 Checked', '2.2  CD')
Versions['2.2  CD'].SetFont('s10', 'Consolas')

Patch := Manager.AddCheckbox('xm+240 ym+200' (IniRead(A_AppData '\aoeiigrdb\config.ini', 'Game', 'Fix', 1) ? ' Checked' : ''), 'Enable fixs after each patching if available')
Patch.OnEvent('Click', (*) => IniWrite(Patch.Value, A_AppData '\aoeiigrdb\config.ini', 'Game', 'Fix'))

Manager.AddGroupBox('xm+230 ym+225 w450 h120 Right', '# Compatibilities').SetFont('Bold')
Manager.AddPicture('xp+54 yp+25', 'DB\000\aok.png')
Manager.AddText('xp-44 yp+40 cRed w120 Center', 'The Age of Kings').SetFont('Bold')
AoKCom := Manager.AddDropDownList('xp yp+20 w120')
For Each, Compat in Compats {
    AoKCom.Add([Compat[1]])
}
AoKCom.Choose(1)
AoKCom.OnEvent("Change", (*) => AoKComReg())
AoKRun := Manager.AddCheckbox('yp+30 wp hp', 'Run as administrator')
AoKRun.OnEvent("Click", (*) => AoKComReg())
AoKComReg() {
    RegVal := Compats[AoKCom.Value][2] (Compats[AoKCom.Value][2] ? ' ' : '') (AoKRun.Value ? 'RUNASADMIN' : '')
    If !RegVal {
        Try {
            RegDelete(Layers, ChosenFolder.Value '\empires2.exe')
        }
        Return
    }
    RegWrite(RegVal, 'REG_SZ', Layers, ChosenFolder.Value '\empires2.exe')
}
Compatibilities[++C] := AoKCom
Compatibilities[++C] := AoKRun

Manager.AddPicture('xp+194 yp-90', 'DB\000\aoc.png')
Manager.AddText('xp-44 yp+40 cBlue w120 Center', 'The Conquerors').SetFont('Bold')
AoCCom := Manager.AddDropDownList('xp yp+20 w120')
For Each, Compat in Compats {
    AoCCom.Add([Compat[1]])
}
AoCCom.Choose(1)
AoCCom.OnEvent("Change", (*) => AoCComReg())
AoCRun := Manager.AddCheckbox('yp+30 wp hp', 'Run as administrator')
AoCRun.OnEvent("Click", (*) => AoCComReg())
AoCComReg() {
    RegVal := Compats[AoCCom.Value][2] (Compats[AoCCom.Value][2] ? ' ' : '') (AoCRun.Value ? 'RUNASADMIN' : '')
    If !RegVal {
        Try {
            RegDelete(Layers, ChosenFolder.Value '\age2_x1\age2_x1.exe')
        }
        Return
    }
    RegWrite(RegVal, 'REG_SZ', Layers, ChosenFolder.Value '\age2_x1\age2_x1.exe')
}
Compatibilities[++C] := AoCCom
Compatibilities[++C] := AoCRun

Manager.AddPicture('xp+194 yp-90', 'DB\000\fe.png')
Manager.AddText('xp-44 yp+40 cGreen w120 Center', 'Forgotten Empires').SetFont('Bold')
FECom := Manager.AddDropDownList('xp yp+20 w120')
For Each, Compat in Compats {
    FECom.Add([Compat[1]])
}
FECom.Choose(1)
FECom.OnEvent("Change", (*) => FEComReg())
FERun := Manager.AddCheckbox('yp+30 wp hp', 'Run as administrator')
FERun.OnEvent("Click", (*) => FEComReg())
FEComReg() {
    RegVal := Compats[FECom.Value][2] (Compats[FECom.Value][2] ? ' ' : '') (FERun.Value ? 'RUNASADMIN' : '')
    If !RegVal {
        Try {
            RegDelete(Layers, ChosenFolder.Value '\age2_x1\age2_x2.exe')
        }
        Return
    }
    RegWrite(RegVal, 'REG_SZ', Layers, ChosenFolder.Value '\age2_x1\age2_x2.exe')
}
Compatibilities[++C] := FECom
Compatibilities[++C] := FERun

Manager.AddStatusBar(, 'v1.0')
Manager.Show()
GameFolderValid()
CurrentVersions()
CompatibilityCheck()
Return

CompatibilityCheck() {
    AoKReg := RegRead(Layers, ChosenFolder.Value '\empires2.exe', '')
    If AoKReg {
        AoKReg := StrSplit(AoKReg, ' ')
        For Each, RegVal in AoKReg {
            If (RegVal = 'RUNASADMIN')
                AoKRun.Value := 1
            Else {
                For Each, Compat in Compats {
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
                For Each, Compat in Compats {
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
                For Each, Compat in Compats {
                    If (Compat[2] = RegVal)
                        FECom.Choose(Compat[1])
                }
            }
        }
    }
}

HoldRadios(H := True) {
    If H {
        For Version, Radio in Versions {
            Radio.Enabled := False
        }
        For Version, Radio in Compatibilities {
            Radio.Enabled := False
        }
    } Else {
        For Version, Radio in Versions {
            Radio.Enabled := True
        }
        For Version, Radio in Compatibilities {
            Radio.Enabled := True
        }
    }
}

CloseGameCheck() {
    If ProcessExist('empires2.exe')
        ProcessClose('empires2.exe')
    ProcessWaitClose('empires2.exe', 5)
    If ProcessExist('empires2.exe')
        Return False
    If ProcessExist('age2_x1.exe')
        ProcessClose('age2_x1.exe')
    ProcessWaitClose('age2_x1.exe', 5)
    If ProcessExist('age2_x1.exe')
        Return False
    If ProcessExist('age2_x2.exe')
        ProcessClose('age2_x2.exe')
    ProcessWaitClose('age2_x2.exe', 5)
    If ProcessExist('age2_x2.exe')
        Return False
    Return True
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

CurrentVersions() {
    FoundVersion := ''
    Loop Files, 'DB\002\*', 'D' {
        Version := A_LoopFileName
        Found := True
        Loop Files, 'DB\002\' Version '\*.*', 'R' {
            PatchFile := A_LoopFileDir '\' A_LoopFileName
            GameFile := ChosenFolder.Value StrReplace(PatchFile, 'DB\002\' Version)
            If !FileExist(GameFile) || HashFile(PatchFile) != HashFile(GameFile) {
                Found := False
                Break
            }
        }
        If Found {
            Versions[Version].Value := 1
        }
    }
}

GameFolderValid() {
    If DirExist(GameFolder := IniRead(A_AppData '\aoeiigrdb\config.ini', 'Game', 'Path', '')) {
        ChosenFolder.Value := GameFolder
    }
    AoKLogo.Enabled := False
    AoCLogo.Enabled := False
    FELogo.Enabled := False
    If FileExist(ChosenFolder.Value '\empires2.exe') {
        AoKLogo.Enabled := True
    }
    If FileExist(ChosenFolder.Value '\age2_x1\age2_x1.exe') {
        AoCLogo.Enabled := True
    }
    If FileExist(ChosenFolder.Value '\age2_x1\age2_x2.exe') {
        FELogo.Enabled := True
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

RetrieveMap() {
    whr := ComObject("WinHttp.WinHttpRequest.5.1")
    whr.Open("GET", "https://raw.githubusercontent.com/FreeP4lestine/aoeiigrdb/main/map.txt", true)
    whr.Send()
    whr.WaitForResponse()
    return Trim(whr.ResponseText, '`n')
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