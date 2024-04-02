; Require AutoHotkey v2
#Requires AutoHotkey v2

; Force the script instance to be overwritten
#SingleInstance Force

; Must be called once before you use any of the DLL functions.
#DllLoad 'Gdiplus.dll'

; Start the app
Try {
    AoE_II_Manager_AIO()
} Catch Error As Err {
    MsgBox('Unexpected error occured!'
         . '`n`nMessage: `n-> [' Err.Message ']'
         . '`n`nFunction: `n-> [' Err.What ']'
         . '`n`nExtra: `n-> [' Err.Extra ']'
         . '`n`nFile: `n-> [' Err.File ']'
         . '`n`nLine: `n-> [' Err.Line ']'
         . '`n`nYou can contact this app provider!'
         . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
}

; AoE II Manager AIO class
Class AoE_II_Manager_AIO {
    ; Version
    Version         := '2.0'
    ; App server
    Server          := 'https://raw.githubusercontent.com'
    User            := 'SmileAoE'
    Repositry       := 'aoeii_aio'
    DownloadDB      := This.Server '/' This.User '/' This.Repositry '/main'
    ; Packages
    BasePackages    := ['DB/000.7z.001', 'DB/001.7z.001', 'DB/002.7z.001', 'DB/006.7z.001', 'DB/007.7z.001', 'DB/008.7z.001']
    GamePackages    := ['DB/003.7z.001', 'DB/003.7z.002', 'DB/003.7z.003', 'DB/003.7z.004', 'DB/004.7z.001', 'DB/004.7z.002', 'DB/004.7z.003', 'DB/005.7z.001']
    RestPackages    := ['DB/009.7z.001', 'DB/009.7z.002', 'DB/010.7z.001', 'DB/010.7z.002', 'DB/010.7z.003', 'DB/010.7z.004', 'DB/010.7z.005', 'DB/011.7z.001', 'DB/012.7z.001', 'DB/013.7z.001', 'DB/014.7z.001', 'DB/014.7z.002']
    ; Reg key
    InstallRegKey   := "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Age of Empires II AIO"
    UninstallScript := '
    (
     Please do not use this, unless you know what you are doing!
    #Requires AutoHotkey v2.0
    #SingleInstance Force
    If !A_Args.Length || A_Args[1] != 'aoeii_aio_uninstall_game_request' {
        ExitApp()
    }
    Name := StrSplit(A_ScriptDir, '\')
    Name := Name[Name.Length]
    Unio := FileOpen(A_Temp '\Uninstall.ahk', 'w')
    Unio.WriteLine('#Requires AutoHotkey v2.0')
    Unio.WriteLine('#SingleInstance Force')
    Unio.WriteLine("If !A_Args.Length || A_Args[1] != 'aoeii_aio_uninstall_game_request' {")
    Unio.WriteLine("    ExitApp()")
    Unio.WriteLine("}")
    Unio.WriteLine("If 'Yes' = MsgBox('Are you sure want to uninstall " Name " ?', 'Uninstall', 0x4 + 0x40) {")
    Unio.WriteLine("DirDelete('" A_ScriptDir "', 1)")
    Unio.WriteLine("Msgbox('Uninstall completed!')")
    Unio.WriteLine("}")
    Run(A_Temp '\Uninstall.ahk aoeii_aio_uninstall_game_request')
    )'
    ; Base package hashs
    LinkHashs       := This.DownloadDB '/DB/Hashsums.ini'
    ; Configuration
    Config          := 'Config.ini'
    Update          := IniRead(This.Config, 'Settings', 'Update', 0)
    ; Default folders
    AppDir          := ['DB', 'Hotkeys', 'Records']
    ; Default shortucts
    Shortcut1       := '
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
    Shortcut2       := '
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
    Shortcutslist   := [This.Shortcut1, This.Shortcut2]
    ; 7zip
    Location7z      := 'DB\7za.exe'
    Link7z          := This.DownloadDB '/7za.exe'
    Hash7z          := '80014d2b38a815f1a6ea220e679111c6'
    7zPID           := 0
    ; 
    Features        := Map('Version', [], 'Game', [])
    ; Versions maps
    GameVersion     := Map('AOK'            , ['2.0  CD', '2.0a No CD', '2.0b CD']
                         , 'AOKCombine'     , Map('2.0b CD', ['2.0a No CD'])
                         , 'AOKHandle'      , Map()
                         , 'AOC'            , ['1.0  CD', '1.0c No CD', '1.0e No CD', '1.1  No CD', '1.5  CD']
                         , 'AOCCombine'     , Map('1.0e No CD', ['1.0c No CD'], '1.1  No CD', ['1.0c No CD'], '1.5  CD', ['1.0c No CD'])
                         , 'AOCHandle'      , Map()
                         , 'FE'             , ['2.2  CD'])
    ; GameRanger
    GRSetting       := A_AppData '\GameRanger\GameRanger Prefs\Settings'
    GRApp           := A_AppData '\GameRanger\GameRanger\GameRanger.exe'
    ; Start the app
    __New() {
        This.IsAdmin()
        This.CreateDirectory()
        This.CreateShortcuts()
        This.UseGDIP()
        This.GUILoading()
        If This.GetConnectedState() && This.Update
            This.UpdatedHashs := This.UpdatedPackagesHashs()
        PackagesNumber := This.BasePackages.Length + This.GamePackages.Length + This.RestPackages.Length + 1
        This.DoneSteps.Opt('Range1-' PackagesNumber)
        For Each, Package in This.BasePackages {
            ; Update the progress GUI
            This.DoneSteps.Value += 1
            This.Prepare.Title := 'Loading -> (' This.DoneSteps.Value ' / ' PackagesNumber ')...'
            ; Set the needed parameters
            PackagePath := StrReplace(Package, '/', '\')
            PackageFolder := StrSplit(PackagePath, '.')[1]
            If This.GetConnectedState() && This.Update
                This.DownloadPackage(Package, PackagePath, PackageFolder)
            This.ExtractPackage(PackagePath, PackageFolder)
        }
        For Each, Package in This.GamePackages {
            ; Update the progress GUI
            This.DoneSteps.Value += 1
            This.Prepare.Title := 'Loading -> (' This.DoneSteps.Value ' / ' PackagesNumber ')...'
            ; Set the needed parameters
            PackagePath := StrReplace(Package, '/', '\')
            PackageFolder := StrSplit(PackagePath, '.')[1]
            If FileExist(PackagePath) {
                If This.GetConnectedState() && This.Update
                    This.DownloadPackage(Package, PackagePath, PackageFolder)
            }
        }
        For Each, Package in This.RestPackages {
            ; Update the progress GUI
            This.DoneSteps.Value += 1
            This.Prepare.Title := 'Loading -> (' This.DoneSteps.Value ' / ' PackagesNumber ')...'
            ; Set the needed parameters
            PackagePath := StrReplace(Package, '/', '\')
            PackageFolder := StrSplit(PackagePath, '.')[1]
            If FileExist(PackagePath) {
                If This.GetConnectedState() && This.Update
                    This.DownloadPackage(Package, PackagePath, PackageFolder)
                This.ExtractPackage(PackagePath, PackageFolder)
            }
        }
        This.Prepare.Destroy()
        This.GUIManager()
        This.SectionTitle()
        This.GameLocation()
        This.VersionSection()
        ; Finally display the window
        This.HMGUI.Show('x' (A_ScreenWidth - 640) / 2 ' y' (A_ScreenHeight - 500) / 2)
        This.Loader()
    }
    ; Quit
    Quit(HG) {
        If ProcessExist(This.7zPID) {
            ProcessClose(This.7zPID)
        }
        ExitApp()
    }
    ; Shows Loading GUI
    GUILoading(Quit := 1) {
        This.Prepare := Gui(, 'Loading...')
        If Quit {
            This.Prepare.OnEvent('Close', ObjBindMethod(This, 'Quit'))
        }
        This.Prepare.AddText('Center w400 h25', 'Please Wait...').SetFont('s12 Bold')
        This.DoneSteps := This.Prepare.AddProgress('Center w400 h20 -Smooth')
        This.DoneStepsText := This.Prepare.AddText('Center wp cBlue')
        This.Prepare.Show()
    }
    ; Checks if the script run as admin
    IsAdmin() {
        If !A_IsAdmin {
            MsgBox('Script must run as administrator!', 'Warn', 0x30)
            ExitApp
        }
    }
    ; Creates the app default dirs
    CreateDirectory() {
        Try {
            For Every, Directory in This.AppDir {
                If !DirExist(Directory) {
                    DirCreate(Directory)
                }
            }
        } Catch Error As Err {
            MsgBox('Unexpected error occured!'
                 . '`n`nMessage: `n-> [' Err.Message ']'
                 . '`n`nFunction: `n-> [' Err.What ']'
                 . '`n`nExtra: `n-> [' Err.Extra ']'
                 . '`n`nFile: `n-> [' Err.File ']'
                 . '`n`nLine: `n-> [' Err.Line ']'
                 . '`n`nYou can contact this app provider!'
                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
            ExitApp
        }
    }
    ; Creates and run the default game shortcuts
    CreateShortcuts()
    {
        Try {
            For Every, Shortcut in This.Shortcutslist {
                If !FileExist(This.AppDir[2] '\00' Every '.ahk') || FileRead(This.AppDir[2] '\00' Every '.ahk') != Shortcut
                {
                    O := FileOpen(This.AppDir[2] '\00' Every '.ahk', 'w')
                    O.Write(Shortcut)
                    O.Close()
                }
                Run(This.AppDir[2] '\00' Every '.ahk ' ProcessExist())
            }
        } Catch Error As Err {
            MsgBox('Unexpected error occured!'
                 . '`n`nMessage: `n-> [' Err.Message ']'
                 . '`n`nFunction: `n-> [' Err.What ']'
                 . '`n`nExtra: `n-> [' Err.Extra ']'
                 . '`n`nFile: `n-> [' Err.File ']'
                 . '`n`nLine: `n-> [' Err.Line ']'
                 . '`n`nYou can contact this app provider!'
                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
            ExitApp
        }
    }
    ; Loads and initializes the Gdiplus.dll.
    UseGDIP() {
        Static GdipObject := 0
        If !IsObject(GdipObject) {
            GdipToken := 0
            SI := Buffer(24, 0) ; size of 64-bit structure
            NumPut("UInt", 1, SI)
            If DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", &GdipToken, "Ptr", SI, "Ptr", 0, "UInt") {
                Return 4
            }
            GdipObject := { __Delete: UseGdipShutDown }
        }
        UseGdipShutDown(*) {
            DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", GdipToken)
        }
    }
    ; Checks the internet connection
    GetConnectedState() {
        Return DllCall("Wininet.dll\InternetGetConnectedState", "Str", Flag := 0x40, "Int", 0)
    }
    ; Checks the unpacker
    PrepareTheUnpacker() {
        Try {
            If !FileExist(This.Location7z) || This.HashFile(This.Location7z) != This.Hash7z {
                Download(This.Link7z, This.Location7z)
            }
        } Catch Error As Err {
            MsgBox('Unexpected error occured!'
                 . '`n`nMessage: `n-> [' Err.Message ']'
                 . '`n`nFunction: `n-> [' Err.What ']'
                 . '`n`nExtra: `n-> [' Err.Extra ']'
                 . '`n`nFile: `n-> [' Err.File ']'
                 . '`n`nLine: `n-> [' Err.Line ']'
                 . '`n`nYou can contact this app provider!'
                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
            ExitApp
        }
    }
    ; Returns a file hash
    HashFile(FilePath, HashType := 2) {
        Try {
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
        } Catch Error As Err {
            MsgBox('Unexpected error occured!'
                 . '`n`nMessage: `n-> [' Err.Message ']'
                 . '`n`nFunction: `n-> [' Err.What ']'
                 . '`n`nExtra: `n-> [' Err.Extra ']'
                 . '`n`nFile: `n-> [' Err.File ']'
                 . '`n`nLine: `n-> [' Err.Line ']'
                 . '`n`nYou can contact this app provider!'
                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
            ExitApp
        }
    }
    ; Returns the updated packages hashs
    UpdatedPackagesHashs() {
        Try {
            whr := ComObject("WinHttp.WinHttpRequest.5.1")
            whr.Open("GET", This.LinkHashs, true)
            whr.Send()
            whr.WaitForResponse()
            Data := Trim(whr.ResponseText, '`n`r')
            Hashs := Map()
            Loop Parse, Data, '`n`r' {
                Pair := StrSplit(A_LoopField, '=')
                Hashs[Pair[1]] := Pair[2]
            }
            Return Hashs
        } Catch Error As Err {
            MsgBox('Unexpected error occured!'
                 . '`n`nMessage: `n-> [' Err.Message ']'
                 . '`n`nFunction: `n-> [' Err.What ']'
                 . '`n`nExtra: `n-> [' Err.Extra ']'
                 . '`n`nFile: `n-> [' Err.File ']'
                 . '`n`nLine: `n-> [' Err.Line ']'
                 . '`n`nYou can contact this app provider!'
                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
            ExitApp
        }
    }
    ; Downloads a given package
    DownloadPackage(Package, PackagePath, PackageFolder) {
        Try {
            This.DoneStepsText.Text := 'Check for updates -> ' PackagePath '...'
            If !FileExist(PackagePath) || !This.UpdatedHashs.Has(PackagePath) || (This.HashFile(PackagePath) != This.UpdatedHashs[PackagePath]) {
                This.DoneStepsText.Text := 'Updating -> ' PackagePath '...'
                Download(This.DownloadDB '/' Package, PackagePath)
                DirDelete(PackageFolder, 1)
            }
        } Catch Error As Err {
            MsgBox('Unexpected error occured!'
                 . '`n`nMessage: `n-> [' Err.Message ']'
                 . '`n`nFunction: `n-> [' Err.What ']'
                 . '`n`nExtra: `n-> [' Err.Extra ']'
                 . '`n`nFile: `n-> [' Err.File ']'
                 . '`n`nLine: `n-> [' Err.Line ']'
                 . '`n`nYou can contact this app provider!'
                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
            ExitApp
        }
    }
    ; Extracts a given package
    ExtractPackage(PackagePath, PackageFolder, Overwrite := 0) {
        Try {
            Overwrite := !Overwrite ? !DirExist(PackageFolder) : Overwrite
            If Overwrite && FileExist(PackagePath) {
                This.DoneStepsText.Text := 'Exporting -> ' PackagePath '...'
                Run(This.Location7z ' x ' PackagePath ' -o"' PackageFolder '" -aoa',, 'Hide', &PID)
                ProcessWaitClose(This.7zPID := PID)
            }
        } Catch Error As Err {
            MsgBox('Unexpected error occured!'
                 . '`n`nMessage: `n-> [' Err.Message ']'
                 . '`n`nFunction: `n-> [' Err.What ']'
                 . '`n`nExtra: `n-> [' Err.Extra ']'
                 . '`n`nFile: `n-> [' Err.File ']'
                 . '`n`nLine: `n-> [' Err.Line ']'
                 . '`n`nYou can contact this app provider!'
                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
            ExitApp
        }
    }
    ; Set the manager GUI
    GUIManager() {
        This.HMGUI := Gui('-DPIScale Resize MinSize' 640 'x' 500 ' MaxSize' 640 'x' 500, 'Age of Empire II Easy Manager AIO v' This.Version)
        This.HMGUI.SetFont('s10 Bold', 'Calibri')
        This.HMGUI.BackColor := 'White'
        This.HMGUI.OnEvent('Close', ObjBindMethod(This, 'Quit'))
        ; Set the scroll bars
        SB := ScrollBar(This.HMGUI, [640, 500]*)
        HotIfWinActive("ahk_id " This.HMGUI.Hwnd)
        Hotkey("WheelUp"    , Scrolls)
        Hotkey("WheelDown"  , Scrolls)
        Hotkey("+WheelUp"   , Scrolls)
        Hotkey("+WheelDown" , Scrolls)
        Hotkey("Up"         , Scrolls)
        Hotkey("Down"       , Scrolls)
        Hotkey("+Up"        , Scrolls)
        Hotkey("+Down"      , Scrolls)
        Hotkey("PgUp"       , Scrolls)
        Hotkey("PgDn"       , Scrolls)
        Hotkey("+PgUp"      , Scrolls)
        Hotkey("+PgDn"      , Scrolls)
        Hotkey("Home"       , Scrolls)
        Hotkey("End"        , Scrolls)
        Scrolls(H) {
            Switch H {
                Case 'WheelUp'      : SB.ScrollMsg((InStr(H,"Down") || InStr(H,"Dn")) ? 1 : 0, 0, GetKeyState("Shift") ? 0x114 : 0x115, This.HMGUI.Hwnd)
                Case 'WheelDown'    : SB.ScrollMsg((InStr(H,"Down") || InStr(H,"Dn")) ? 1 : 0, 0, GetKeyState("Shift") ? 0x114 : 0x115, This.HMGUI.Hwnd)
                Case '+WheelUp'     : SB.ScrollMsg((InStr(H,"Down") || InStr(H,"Dn")) ? 1 : 0, 0, GetKeyState("Shift") ? 0x114 : 0x115, This.HMGUI.Hwnd)
                Case '+WheelDown'   : SB.ScrollMsg((InStr(H,"Down") || InStr(H,"Dn")) ? 1 : 0, 0, GetKeyState("Shift") ? 0x114 : 0x115, This.HMGUI.Hwnd)
                Case 'Up'           : SB.ScrollMsg((InStr(H,"Down") || InStr(H,"Dn")) ? 1 : 0, 0, GetKeyState("Shift") ? 0x114 : 0x115, This.HMGUI.Hwnd)
                Case 'Down'         : SB.ScrollMsg((InStr(H,"Down") || InStr(H,"Dn")) ? 1 : 0, 0, GetKeyState("Shift") ? 0x114 : 0x115, This.HMGUI.Hwnd)
                Case '+Up'          : SB.ScrollMsg((InStr(H,"Down") || InStr(H,"Dn")) ? 1 : 0, 0, GetKeyState("Shift") ? 0x114 : 0x115, This.HMGUI.Hwnd)
                Case '+Down'        : SB.ScrollMsg((InStr(H,"Down") || InStr(H,"Dn")) ? 1 : 0, 0, GetKeyState("Shift") ? 0x114 : 0x115, This.HMGUI.Hwnd)
                Case 'PgUp'         : SB.ScrollMsg((InStr(H,"Down") || InStr(H,"Dn")) ? 3 : 2, 0, GetKeyState("Shift") ? 0x114 : 0x115, This.HMGUI.Hwnd)
                Case 'PgDn'         : SB.ScrollMsg((InStr(H,"Down") || InStr(H,"Dn")) ? 3 : 2, 0, GetKeyState("Shift") ? 0x114 : 0x115, This.HMGUI.Hwnd)
                Case '+PgUp'        : SB.ScrollMsg((InStr(H,"Down") || InStr(H,"Dn")) ? 3 : 2, 0, GetKeyState("Shift") ? 0x114 : 0x115, This.HMGUI.Hwnd)
                Case '+PgDn'        : SB.ScrollMsg((InStr(H,"Down") || InStr(H,"Dn")) ? 3 : 2, 0, GetKeyState("Shift") ? 0x114 : 0x115, This.HMGUI.Hwnd)
                Case 'Home'         : SB.ScrollMsg(6, 0, GetKeyState("Shift") ? 0x114 : 0x115, This.HMGUI.Hwnd)
                Case 'End'          : SB.ScrollMsg(7, 0, GetKeyState("Shift") ? 0x114 : 0x115, This.HMGUI.Hwnd)
            }
        }
    }
    ; Adds the title
    SectionTitle() {
        This.Thumb := This.HMGUI.AddPicture('xm+118', 'DB\000\gameoff.png')
        This.Thumb.Focus()
        This.Title := This.HMGUI.AddText('xm yp+200 Center w600 h50 cGray', 'Age of Empire II Easy Manager AIO v' This.Version)
        This.Title.SetFont('s16')
        This.Log := This.HMGUI.AddListview('-Hdr xm+150 w300 r8 -E0x200', ['S', 'L'])
        This.Log.ModifyCol(1, '30 Center')
        This.Log.OnEvent('ItemSelect', ObjBindMethod(This, 'LoggerSelectColor'))
        This.LogCLV := LV_Colors(This.Log)
        This.Reloader := This.HMGUI.AddButton('wp Disabled', 'RELOAD')
        This.Reloader.OnEvent('Click', ObjBindMethod(This, 'ReloadGame'))
    }
    ; Reloads for a game selection
    ReloadGame(Ctrl, Info) {
        Ctrl.Enabled := False
        This.Loader()
        Ctrl.Enabled := True
    }
    ; Updates the log list
    Logger(Text, OK := 0, R := 0) {
        Switch R {
            Case 0 :
                R := This.Log.Add(, '→', Text)
            Default :
                Switch OK {
                    Case 0 :
                        Color := 'Green'
                        TextStatus := 'OK'
                    Case 1 :
                        Color := 'Red'
                        TextStatus := '!OK'
                    Case 3 :
                        Color := '0xDC6F00'
                        TextStatus := '!FIX'
                }
                This.Log.Modify(R,, TextStatus)
                This.LogCLV.Cell(R, 1, Color, 'White')
                This.Log.Modify(R,,, Text)
                This.LogCLV.Cell(R, 2,, Color)
        }
        This.Log.ModifyCol(2, 'AutoHdr')
        Return R
    }
    ; Updates the log list select color
    LoggerSelectColor(Ctrl, Item, Selected) {
        If !Selected {
            Return
        }
        Status := Ctrl.GetText(Item, 1)
        Switch Status {
            Case 'OK' : This.LogCLV.SelectionColors(0x008000, 0xFFFFFF)
            Case '!FIX' : This.LogCLV.SelectionColors(0xDC6F00, 0xFFFFFF)
            Case '!OK' : This.LogCLV.SelectionColors(0xFF0000, 0xFFFFFF)
        }
    }
    ; Adds the game location
    GameLocation() {
        H := This.HMGUI.AddText('xm cBlue w600 h30', 'GAME LOCATION:')
        H.SetFont('s16')
        This.Features['Game'].Push(H)
        This.GameDirectory := This.HMGUI.AddEdit('ReadOnly xm+20 w580 -E0x200 Border')
        This.Features['Game'].Push(This.GameDirectory)
        H := This.HMGUI.AddButton('w100', 'Select')
        This.Features['Game'].Push(H)
        This.GuiButtonIcon(H, 'DB\000\folder.png',, 'A1')
        H.OnEvent('Click', ObjBindMethod(This, 'SelectDirectory'))
        H := This.HMGUI.AddButton('w170 yp', 'Select from GameRanger')
        This.GuiButtonIcon(H, 'DB\000\gr.png',, 'A1')
        H.OnEvent('Click', ObjBindMethod(This, 'SelectDirectoryGR'))
        This.Features['Game'].Push(H)
        H := This.HMGUI.AddButton('w140 yp', 'Open the selected')
        This.GuiButtonIcon(H, 'DB\000\sfolder.png',, 'A1')
        This.Features['Game'].Push(H)
        H.OnEvent('Click', (*) => This.GameDirectory.Value ? Run(This.GameDirectory.Value '\') : 0)
        H := This.HMGUI.AddCheckBox('xm+20 Checked', 'Perform a common issues fix on each load')
        This.Features['Game'].Push(H)
        H.OnEvent('Click', ObjBindMethod(this, 'CommonIssueFix'))
        H := This.HMGUI.AddButton('xm+20 w200', 'Download and install the game')
        This.GuiButtonIcon(H, 'DB\000\download.png',, 'A1')
        This.Features['Game'].Push(H)
        H.OnEvent('Click', ObjBindMethod(this, 'DownloadGame'))
    }
    ; Selects a location from GR
    SelectDirectoryGR(Ctrl, Info) {
        Ctrl.Enabled := False
        Text := This.BinGrabText(This.GRSetting)
        Locations := This.TextGrabPath(Text, ['empires2.exe', 'age2_x1.exe', 'age2_x2.exe'])
        For Location in Locations {
            If RC := This.ValidGameDirectory(Location) {
                Choice := MsgBox('Want to select this location?`n`n' Location, 'Game Location', 0x4 + 0x40)
                If Choice = 'Yes' {
                    This.GameDirectory.Value := Location
                    IniWrite(Location, This.Config, 'Settings', 'GameDirectory')
                    This.Loader(Location)
                    MsgBox('Game selected sucessfully!', 'Game Select', 0x40 ' T5')
                    Break
                }
            }
        }
        Ctrl.Enabled := True
    }
    ; Selects a location
    SelectDirectory(Ctrl, Info) {
        Ctrl.Enabled := False
        If SelectedDirectory := FileSelect('D', 'C:\' (A_Is64bitOS ? 'Program Files (x86)' : 'Program Files') '\Microsoft Games') {
            If !Valid := This.ValidGameDirectory(SelectedDirectory) {
                SelectedDirectoryEx := SelectedDirectory
                SelectedDirectory := ''
                SplitPath(SelectedDirectoryEx, &_, &ParentSelectedDirectory)
                If Valid := This.ValidGameDirectory(ParentSelectedDirectory) {
                    Choice := MsgBox('Want to select this location?`n`n' ParentSelectedDirectory, 'Game Location', 0x4 + 0x40)
                    If Choice = 'Yes' {
                        SelectedDirectory := ParentSelectedDirectory
                    }
                }
            }
            If !Valid {
                Loop Files, SelectedDirectoryEx '\*', 'D' {
                    If This.ValidGameDirectory(A_LoopFileFullPath) {
                        Choice := MsgBox('Want to select this location?`n`n' A_LoopFileFullPath, 'Game Location', 0x4 + 0x40)
                        If Choice = 'Yes' {
                            SelectedDirectory := A_LoopFileFullPath
                            Break
                        }
                    }
                }
            }
            If SelectedDirectory != '' {
                SelectedDirectory := StrUpper(SelectedDirectory)
                This.GameDirectory.Value := SelectedDirectory
                IniWrite(SelectedDirectory, This.Config, 'Settings', 'GameDirectory')
                This.Loader(SelectedDirectory)
                MsgBox('Game selected sucessfully!', 'Game Select', 0x40 ' T5')
            }
            Else {
                MsgBox('Invalid game location!', 'Game Select', 0x30)
            }
        }
        Ctrl.Enabled := True
    }
    ; Downloads and installs the game
    DownloadGame(Ctrl, Info) {
        Try {
            If !This.GetConnectedState() {
                MsgBox('Make sure you are connected to the internet!', "Can't download!", 0x30)
                Return
            }
            If (GameDirectory := FileSelect('D',, 'Game install location')) && 'Yes' = MsgBox('Are you sure want to install at this location?`n' GameDirectory, 'Game install location', 0x40 + 0x4) {
                GameDirectory := RegExReplace(GameDirectory, "\\$")
                GameDirectory := GameDirectory '\Age of Empires II'
                If !DirExist(GameDirectory) {
                    DirCreate(GameDirectory)
                }
                If This.ValidGameDirectory(GameDirectory) && 'Yes' != MsgBox('It seems like the game already installed at this location!`nWant continue?', 'Game location install', 0x30 + 0x4) {
                    Return
                }
                ; Check for packages updates
                Ctrl.Enabled := False
                This.UpdatedHashs := This.UpdatedPackagesHashs()
                This.GUILoading(0)
                This.DoneSteps.Opt('Range0-' (PackagesNumber := This.GamePackages.Length + 3))
                This.DoneSteps.Value := 0
                For Each, Package in This.GamePackages {
                    ; Update the progress GUI
                    This.DoneSteps.Value += 1
                    This.Prepare.Title := 'Downloading -> (' This.DoneSteps.Value ' / ' PackagesNumber ')...'
                    ; Set the needed parameters
                    PackagePath := StrReplace(Package, '/', '\')
                    PackageFolder := StrSplit(PackagePath, '.')[1]
                    This.DownloadPackage(Package, PackagePath, PackageFolder)
                }
                ; AOK extract
                This.Prepare.Title := 'Exporting -> (' This.DoneSteps.Value ' / ' PackagesNumber ')...'
                This.ExtractPackage('DB\003.7z.001', GameDirectory, 1)
                This.DoneSteps.Value += 1
                ; AOC extract
                This.Prepare.Title := 'Exporting -> (' This.DoneSteps.Value ' / ' PackagesNumber ')...'
                This.ExtractPackage('DB\004.7z.001', GameDirectory, 1)
                This.DoneSteps.Value += 1
                ; FE extract
                This.Prepare.Title := 'Exporting -> (' This.DoneSteps.Value ' / ' PackagesNumber ')...'
                This.ExtractPackage('DB\005.7z.001', GameDirectory, 1)
                This.DoneSteps.Value += 1
                ; Add the reg keys
                This.UpdateGameReg(GameDirectory)
                If 'Yes' = MsgBox('Game installation complete!`nWanna select this game?', 'Game install location', 0x4 + 0x40) {
                    This.Loader(GameDirectory)
                }
                This.Prepare.Destroy()
            }
            Ctrl.Enabled := True
        } Catch Error As Err {
            MsgBox('Unexpected error occured!'
                 . '`n`nMessage: `n-> [' Err.Message ']'
                 . '`n`nFunction: `n-> [' Err.What ']'
                 . '`n`nExtra: `n-> [' Err.Extra ']'
                 . '`n`nFile: `n-> [' Err.File ']'
                 . '`n`nLine: `n-> [' Err.Line ']'
                 . '`n`nYou can contact this app provider!'
                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
            ExitApp
        }
    }
    ; Updates game install registery settings
    UpdateGameReg(GameDirectory) {
        Try {
            RegWrite('Age of Empires II AIO', 'REG_SZ', This.InstallRegKey, 'DisplayName')
            RegWrite('AOK (2.0) / AOC (1.0) / FE (2.1)', 'REG_SZ', This.InstallRegKey, 'DisplayVersion')
            RegWrite(GameDirectory '\age2_x1\age2_x1.exe', 'REG_SZ', This.InstallRegKey, 'DisplayIcon')
            RegWrite(GameDirectory, 'REG_SZ', This.InstallRegKey, 'InstallLocation')
            RegWrite(1, 'REG_DWORD', This.InstallRegKey, 'NoModify')
            RegWrite(1, 'REG_DWORD', This.InstallRegKey, 'NoRepair')
            RegWrite(This.FolderGetSize(GameDirectory), 'REG_DWORD', This.InstallRegKey, 'EstimatedSize')
            RegWrite('Microsoft Corporation', 'REG_SZ', This.InstallRegKey, 'Publisher')
            ;RegWrite(GameDirectory '\Uninstall.ahk "aoeii_aio_uninstall_game_request"', 'REG_SZ', This.InstallRegKey, 'UninstallString')
        } Catch Error As Err {
            MsgBox('Unexpected error occured!'
                 . '`n`nMessage: `n-> [' Err.Message ']'
                 . '`n`nFunction: `n-> [' Err.What ']'
                 . '`n`nExtra: `n-> [' Err.Extra ']'
                 . '`n`nFile: `n-> [' Err.File ']'
                 . '`n`nLine: `n-> [' Err.Line ']'
                 . '`n`nYou can contact this app provider!'
                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
            ExitApp
        } 
    }
    ; Returns a folder size in KB
    FolderGetSize(Location) {
        Size := 0
        Loop Files, Location '\*.*', 'R' {
            Size += FileGetSize(A_LoopFileFullPath, 'K')
        }
        Return Size
    }
    ; Checks if a directory contains the game
    ValidGameDirectory(Location) {
        Return   FileExist(Location '\empires2.exe')
              && FileExist(Location '\language.dll')
              && FileExist(Location '\Data\graphics.drs')
              && FileExist(Location '\Data\interfac.drs')
              && FileExist(Location '\Data\terrain.drs') ? 1 : 0
    }
    ; Grabs the readable text from binary file
    BinGrabText(Filepath) {
        Try {
            Text := ''
            BufferObj := FileRead(Filepath, 'RAW')
            Loop BufferObj.Size {
                Address := A_Index - 1
                Byte := NumGet(BufferObj, Address, 'UChar')
                If (C := Chr(Byte)) != '' {
                    Text .= C
                }
            }
            Return Text
        } Catch Error As Err {
            MsgBox('Unexpected error occured!'
                 . '`n`nMessage: `n-> [' Err.Message ']'
                 . '`n`nFunction: `n-> [' Err.What ']'
                 . '`n`nExtra: `n-> [' Err.Extra ']'
                 . '`n`nFile: `n-> [' Err.File ']'
                 . '`n`nLine: `n-> [' Err.Line ']'
                 . '`n`nYou can contact this app provider!'
                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
            ExitApp
        }
    }
    ; Parses the game locations out of a text
    TextGrabPath(TextFound, Excutables) {
        Try {
            ResultMap := Map()
            For Each, Excutable in Excutables {
                P := InStr(TextFound, LFE := Excutable,, -1)
                Loop {
                    Char := SubStr(TextFound, P - (I := A_Index), 1)
                    LFE := Char LFE
                } Until (Char = ':' || Ord(Char) = 10 || Ord(Char) = 13)
                FoundPath := SubStr(TextFound, P - (I + 1), 1) LFE
                FoundPath := StrReplace(FoundPath, '\' Excutables[1])
                FoundPath := StrReplace(FoundPath, '\age2_x1\' Excutables[2])
                FoundPath := StrReplace(FoundPath, '\age2_x1\' Excutables[3])
                ResultMap[StrUpper(FoundPath)] := True
            }
            Return ResultMap
        } Catch Error As Err {
            MsgBox('Unexpected error occured!'
                 . '`n`nMessage: `n-> [' Err.Message ']'
                 . '`n`nFunction: `n-> [' Err.What ']'
                 . '`n`nExtra: `n-> [' Err.Extra ']'
                 . '`n`nFile: `n-> [' Err.File ']'
                 . '`n`nLine: `n-> [' Err.Line ']'
                 . '`n`nYou can contact this app provider!'
                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
            ExitApp
        }
    }
    ; Common issues fix setting
    CommonIssueFix(Ctrl, Info) {
        IniWrite(Ctrl.Value, This.Config, 'Settings', 'CommonFix')
    }
    ; Adds the versions
    VersionSection() {
        H := This.HMGUI.AddText('xm cBlue w600 h30', 'GAME VERSION:')
        H.GetPos(, &Y)
        Y += 30
        H.SetFont('s16')
        This.Features['Version'].Push(H)
        H := This.HMGUI.AddPicture('xm+20 y' Y, 'DB\000\aok.png')
        This.Features['Version'].Push(H)
        H := This.HMGUI.AddText('cRed', 'The Age of Kings')
        This.Features['Version'].Push(H)
        For Each, AOK in This.GameVersion['AOK'] {
            H := This.HMGUI.AddRadio('w150', AOK)
            H.SetFont(, 'Consolas')
            This.Features['Version'].Push(H)
            H.OnEvent('Click', ObjBindMethod(this, 'ApplyVersion'))
            This.GameVersion['AOKHandle'][AOK] := H
        }
        H := This.HMGUI.AddPicture('xm+220 y' Y, 'DB\000\aoc.png')
        This.Features['Version'].Push(H)
        H := This.HMGUI.AddText('cBlue', 'The Conquerors')
        This.Features['Version'].Push(H)
        For Each, AOC in This.GameVersion['AOC'] {
            H := This.HMGUI.AddRadio('w150', AOC)
            H.SetFont(, 'Consolas')
            This.Features['Version'].Push(H)
            H.OnEvent('Click', ObjBindMethod(this, 'ApplyVersion'))
            This.GameVersion['AOCHandle'][AOC] := H
        }
        H := This.HMGUI.AddPicture('xm+440 y' Y, 'DB\000\fe.png')
        This.Features['Version'].Push(H)
        H := This.HMGUI.AddText('cGreen', 'Forgotten Empires')
        This.Features['Version'].Push(H)
        For Each, FE in This.GameVersion['FE'] {
            H := This.HMGUI.AddRadio('w150 Checked', FE)
            H.SetFont(, 'Consolas')
            This.Features['Version'].Push(H)
        }
        This.OLV := This.HMGUI.AddListView('xm+20 r4 -Hdr Checked -E0x200 -Multi', ['Option'])
        This.Features['Version'].Push(This.OLV)
        This.OLVC := LV_Colors(This.OLV)
        This.OLVC.SelectionColors
        This.OLV.Add(, 'Advanced interface')
        This.OLV.Add(, 'Advanced interface + Overlay')
        This.OLV.Add(, 'Widescreen')
        This.OLV.Add(, 'Centered widescreen')
        This.OLV.OnEvent('ItemCheck', ApplyFix)
        This.OLV.OnEvent('ItemSelect', ApplyFix)
        ApplyFix(Ctrl, Item, CheckedSelected) {
            If !CheckedSelected {
                RegWrite(0, 'REG_DWORD', 'HKEY_CURRENT_USER\SOFTWARE\Microsoft\Microsoft Games\Age of Empires', 'Aoe2Patch')
                IniWrite(0, This.Config, 'Settings', 'Fix')
                Return
            }
            RegWrite(Item, 'REG_DWORD', 'HKEY_CURRENT_USER\SOFTWARE\Microsoft\Microsoft Games\Age of Empires', 'Aoe2Patch')
            IniWrite(Item, This.Config, 'Settings', 'Fix')
            Loop Ctrl.GetCount() {
                If A_Index != Item {
                    Ctrl.Modify(A_Index, '-Check')
                }
            }
            Ctrl.Modify(Item, 'Select')
            Ctrl.Modify(Item, 'Check')
        }
    }
    ; Applys the version
    ApplyVersion(Radio, Info) {
        Try {
            ; Checks the game directory
            If !This.ValidGameDirectory(This.GameDirectory.Value) {
                MsgBox('Invalid game location!', 'Game Select', 0x30)
                Radio.Value := 0
                Return
            }
            This.EnableControls(This.Features['Version'], 0)
            ; Close the game if running
            For Each, Process in ['empires2.exe', 'age2_x1.exe', 'age2_x2.exe'] {
                If ProcessExist(Process) {
                    ProcessClose(Process)
                }
            }
            ; Cleans up previous versions files
            TargetVersion := SubStr(Radio.Text, 1, 1)
            Loop Files, 'DB\002\*', 'D' {
                If TargetVersion != SubStr(Version := A_LoopFileName, 1, 1) {
                    Continue
                }
                Loop Files, 'DB\002\' Version '\*.*', 'R' {
                    PathFile := StrReplace(A_LoopFileDir '\' A_LoopFileName, 'DB\002\' Version '\')
                    If FileExist(This.GameDirectory.Value '\' PathFile) {
                        FileDelete(This.GameDirectory.Value '\' PathFile)
                    }
                }
            }
            ; Cleans up previous fix files
            Loop Files, 'DB\001\*', 'D' {
                Fix := A_LoopFileName
                Loop Files, 'DB\001\' Fix '\*', 'D' {
                    If TargetVersion != SubStr(Version := A_LoopFileName, 1, 1) {
                        Continue
                    }
                    Loop Files, 'DB\001\' Fix '\' Version '\*.*', 'R' {
                        PathFile := StrReplace(A_LoopFileDir '\' A_LoopFileName, 'DB\001\' Fix '\' Version '\')
                        If FileExist(This.GameDirectory.Value '\' PathFile) {
                            FileDelete(This.GameDirectory.Value '\' PathFile)
                        }
                    }
                }
            }
            ; Copy the selected version files
            Key := TargetVersion = '1' ? 'AOCCombine' : 'AOKCombine'
            If This.GameVersion[Key].Has(Radio.Text) {
                For Each, Version in This.GameVersion[Key][Radio.Text] {
                    If DirExist('DB\002\' Version) {
                        DirCopy('DB\002\' Version, This.GameDirectory.Value, 1)
                    }
                }
            }
            If DirExist('DB\002\' Radio.Text) {
                DirCopy('DB\002\' Radio.Text, This.GameDirectory.Value, 1)
            }
            ; Copy fixs
            If Fix := This.OLV.GetNext(0, 'C') {
                DirCopy('DB\001\Enable Fix v2\Static', This.GameDirectory.Value, 1)
                If DirExist('DB\001\Enable Fix v2\' Radio.Text) {
                    DirCopy('DB\001\Enable Fix v2\' Radio.Text, This.GameDirectory.Value, 1)
                }
            }
            This.EnableControls(This.Features['Version'])
            SoundPlay('DB\000\30 Wololo.mp3')
        } Catch Error As Err {
            MsgBox('Unexpected error occured!'
                 . '`n`nMessage: `n-> [' Err.Message ']'
                 . '`n`nFunction: `n-> [' Err.What ']'
                 . '`n`nExtra: `n-> [' Err.Extra ']'
                 . '`n`nFile: `n-> [' Err.File ']'
                 . '`n`nLine: `n-> [' Err.Line ']'
                 . '`n`nYou can contact this app provider!'
                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
            ExitApp
        }
    }
    ; Sets a push button icon
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
    ; Updates the desktop shortcuts
    UpdateShortcuts(FileLocation, Name := '') {
        Try {
            SplitPath(FileLocation, &_, &OutDir, &_, &OutNameNoExt)
            FileShortcut := A_Desktop '\' (Name != '' ? Name : OutNameNoExt) '.lnk'
            If !FileExist(FileShortcut) && FileExist(FileLocation) {
                FileCreateShortcut(FileLocation, FileShortcut, OutDir)
            }
            If FileExist(FileShortcut) && !FileExist(FileLocation) {
                FileRecycle(FileShortcut)
            }
            If FileExist(FileShortcut) && FileExist(FileLocation) {
                FileGetShortcut(FileShortcut, &OutTarget)
                If OutTarget != FileLocation {
                    FileCreateShortcut(FileLocation, FileShortcut, OutDir)
                }
            }
        } Catch Error As Err {
            MsgBox('Unexpected error occured!'
                 . '`n`nMessage: `n-> [' Err.Message ']'
                 . '`n`nExtra: `n-> [' Err.Extra ']'
                 . '`n`nFile: `n-> [' Err.File ']'
                 . '`n`nLine: `n-> [' Err.Line ']'
                 . '`n`nYou can contact this app provider!'
                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
            ExitApp
        }
    }
    ; Clear unwanted files
    ClearUnwanted(FileLocation) {
        Try {
            If FileExist(FileLocation) {
            FileRecycle(FileLocation)
            }
        } Catch Error As Err {
            MsgBox('Unexpected error occured!'
                 . '`n`nMessage: `n-> [' Err.Message ']'
                 . '`n`nFunction: `n-> [' Err.What ']'
                 . '`n`nExtra: `n-> [' Err.Extra ']'
                 . '`n`nFile: `n-> [' Err.File ']'
                 . '`n`nLine: `n-> [' Err.Line ']'
                 . '`n`nYou can contact this app provider!'
                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
            ExitApp
        }
    }
    ; Controls enable or disable
    EnableControls(Controls, Enable := 1) {
        If Enable {
            For Each, Control in Controls {
                Control.Enabled := True
            }
        } Else {
            For Each, Control in Controls {
                Control.Enabled := False
            }
        }
    }
    ; Loads a game
    Loader(GameDirectoryLoad := '') {
        ; Disable features
        This.EnableControls(This.Features['Version'], 0)
        ; Game location section loads
        If !This.GameLocationLoads(GameDirectoryLoad) {
            Return
        }
        ; Enable title
        This.Thumb.Value := 'DB\000\game.png'
        This.Title.Opt('cBlack')
        ; Version section loads
        This.VersionLoads()
        ; Other updates
        This.OLV.Redraw()
        This.Reloader.Enabled := True
    }
    ; Game location section loads
    GameLocationLoads(GameDirectoryLoad) {
        ;1
        This.Log.Delete()
        R := This.Logger('Looking for the game folder...')
        GameDirectory := GameDirectoryLoad != '' ? GameDirectoryLoad : IniRead(This.Config, 'Settings', 'GameDirectory', '')
        If !This.ValidGameDirectory(GameDirectory) {
            This.Logger('No game folder is selected', 1, R)
            Return False
        }
        This.GameDirectory.Value := GameDirectory
        IniWrite(This.GameDirectory.Value, This.Config, 'Settings', 'GameDirectory')
        This.Logger('Game is located at: ' This.GameDirectory.Value,, R)
        ;2
        If IniRead(This.Config, 'Settings', 'CommonFix', 0) {
            ;1
            R := This.Logger('Performing a common issue fix...')
            If FileExist(This.GameDirectory.Value '\age2_x1.exe') {
                If !DirExist(This.GameDirectory.Value '\age2_x1') {
                    DirCreate(This.GameDirectory.Value '\age2_x1')
                }
                FileMove(This.GameDirectory.Value '\age2_x1.exe', This.GameDirectory.Value '\age2_x1', 1)
            }
            ;2
            This.ClearUnwanted(This.GameDirectory.Value '\windmode.ini')
            This.ClearUnwanted(This.GameDirectory.Value '\age2_x1\windmode.ini')
            ;3
            This.UpdateShortcuts(This.GameDirectory.Value '\empires2.exe', 'The Age of Kings')
            This.UpdateShortcuts(This.GameDirectory.Value '\age2_x1\age2_x1.exe', 'The Conquerors')
            This.UpdateShortcuts(This.GameDirectory.Value '\age2_x1\age2_x2.exe', 'Forgotten Empires')
            This.Logger('Perform common issue fix',, R)
        }
        Return True
    }
    ; Version section loads
    VersionLoads() {
        For Each, Control in This.Features['Version'] {
            If Type(Control) = 'Gui.Radio' {
                Control.Value := 0
            }
        }
        R := This.Logger('Scanning the game versions')
        Loop Files, 'DB\002\2.*', 'D' {
            Version := A_LoopFileName
            VersionIsSet := Version
            Loop Files, 'DB\002\' Version '\*.*', 'R' {
                PathFile := StrReplace(A_LoopFileDir '\' A_LoopFileName, 'DB\002\' Version '\')
                If !FileExist(This.GameDirectory.Value '\' PathFile) && VersionIsSet {
                    VersionIsSet := ''
                    Break
                }
                CurrentHash := This.HashFile(A_LoopFileFullPath)
                FoundHash := This.HashFile(This.GameDirectory.Value '\' PathFile)
                If (CurrentHash != FoundHash) && VersionIsSet {
                    VersionIsSet := ''
                    Break
                }
            }
            If VersionIsSet {
                This.GameVersion['AOKHandle'][VersionIsSet].Value := 1
                This.Logger('AOK ' VersionIsSet ' found',, R)
            }
            If This.Log.GetText(R, 1) != 'OK' {
                This.Logger('AOK version not found', 3, R)
            }
        }
        R := This.Logger('Scanning the game versions')
        Loop Files, 'DB\002\1.*', 'D' {
            Version := A_LoopFileName
            VersionIsSet := Version
            Loop Files, 'DB\002\' Version '\*.*', 'R' {
                PathFile := StrReplace(A_LoopFileDir '\' A_LoopFileName, 'DB\002\' Version '\')
                If !FileExist(This.GameDirectory.Value '\' PathFile) {
                    VersionIsSet := ''
                    Break
                }
                CurrentHash := This.HashFile(A_LoopFileFullPath)
                FoundHash := This.HashFile(This.GameDirectory.Value '\' PathFile)
                If (CurrentHash != FoundHash) && VersionIsSet != '' {
                    VersionIsSet := ''
                    Break
                }
            }
            If VersionIsSet != '' {
                This.GameVersion['AOCHandle'][VersionIsSet].Value := 1
                This.Logger('AOC ' VersionIsSet ' found',, R)
            }
            If This.Log.GetText(R, 1) != 'OK' {
                This.Logger('AOC version not found', 3, R)
            }
        }
        R := This.Logger('Scanning the game fixes')
        FixIsSetAOC := True
        FixIsSetAOK := True
        Loop Files, 'DB\001\Enable Fix v2\Static\*.*', 'R' {
            PathFile := StrReplace(A_LoopFileDir '\' A_LoopFileName, 'DB\001\Enable Fix v2\Static\')
            If !FileExist(This.GameDirectory.Value '\' PathFile) {
                FixIsSetAOC := False
                FixIsSetAOK := False
                Break
            }
            CurrentHash := This.HashFile(A_LoopFileFullPath)
            FoundHash := This.HashFile(This.GameDirectory.Value '\' PathFile)
            If (CurrentHash != FoundHash) && FixIsSetAOC {
                FixIsSetAOC := False
                FixIsSetAOK := False
                Break
            }
        }
        If FixIsSetAOK && (!FileExist(This.GameDirectory.Value '\dsound.dll') || This.HashFile(This.GameDirectory.Value '\dsound.dll') != This.HashFile('DB\001\Enable Fix v2\2.0  CD\dsound.dll')) {
            FixIsSetAOK := False
        }
        If FixIsSetAOC && (!FileExist(This.GameDirectory.Value '\age2_x1\dsound.dll') || This.HashFile(This.GameDirectory.Value '\age2_x1\dsound.dll') != This.HashFile('DB\001\Enable Fix v2\1.0  CD\age2_x1\dsound.dll')) {
            FixIsSetAOC := False
        }
        Fix := IniRead(This.Config, 'Settings', 'Fix', 0)
        If Fix && FixIsSetAOK {
            This.OLV.Modify(Fix, 'Select')
            This.OLV.Modify(Fix, 'Check')
            This.Logger('AOK ' This.OLV.GetText(Fix, 1) ' mod found',, R)
        } Else {
            This.Logger('AOK No fix is applied', 3, R)
        }
        R := This.Logger('Scanning the game fixes')
        If Fix && FixIsSetAOC {
            This.OLV.Modify(Fix, 'Select')
            This.OLV.Modify(Fix, 'Check')
            This.Logger('AOC ' This.OLV.GetText(Fix, 1) ' mod found',, R)
        } Else {
            This.Logger('AOC No fix is applied', 3, R)
        }
        This.EnableControls(This.Features['Version'])
    }
}
; Class ScrollBar
Class ScrollBar {
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
        this.ScrollInf.nPos := this.ScrollInf.nPos < 0 ? 0 : this.ScrollInf.nPos
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
Class ScrollInfo {
    ;/* This class defines the structure below(SCROLLINFO) on Winuser.h:
    ;
    ;typedef struct tagSCROLLINFO {
    ;  UINT cbSize;
    ;  UINT fMask;
    ;  int  nMin;
    ;  int  nMax;
    ;  UINT nPage;
    ;  int  nPos;
    ;  int  nTrackPos;
    ;} SCROLLINFO, *LPSCROLLINFO */
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