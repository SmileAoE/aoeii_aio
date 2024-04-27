#Include SharedLib.ahk
Features['Main'] := []
WD := AoEIIAIO.AddButton('x0 y0', '...')
AoEIIAIO.SetFont('Bold s20')
T := AoEIIAIO.AddText('xm cGreen Center', 'Age of Empires II Easy Manager v' Version)
P := AoEIIAIO.AddPicture('xm+90', 'DB\000\game.png')
R := AoEIIAIO.AddButton('xm ym+30', 'RELOAD')
R.SetFont('Bold s12')
CreateImageButton(R, 0, [[0xFFFFFF,, 0xFF0000, 4, 0xFF0000, 2], [0xE6E6E6], [0xCCCCCC], [0xFFFFFF,, 0xCCCCCC]]*)
R.OnEvent('Click', (*) => Reload())
H := AoEIIAIO.AddButton('xm', 'GAME')
H.SetFont('Bold s12')
CreateImageButton(H, 0, [[0xFFFFFF,,, 4, 0x000000, 2], [0xE6E6E6], [0xCCCCCC], [0xFFFFFF,, 0xCCCCCC]]*)
H.OnEvent('Click', LaunchGame)
LaunchGame(Ctrl, Info) {
    Try {
        Run('Game.ahk ' ProcessExist())
    } Catch Error As Err {
        MsgBox("Launch failed!`n`n" Err.Message '`n' Err.Line '`n' Err.File, 'Game', 0x10)
    }
}
H := AoEIIAIO.AddButton('yp', 'VERSION')
H.SetFont('Bold s12')
CreateImageButton(H, 0, [[0xFFFFFF,, 0x0000FF, 4, 0x0000FF, 2], [0xE6E6E6], [0xCCCCCC], [0xFFFFFF,, 0xCCCCCC]]*)
H.OnEvent('Click', LaunchVersion)
Features['Main'].Push(H)
LaunchVersion(Ctrl, Info) {
    Try {
        Run('Version.ahk')
    } Catch Error As Err {
        MsgBox("Launch failed!`n`n" Err.Message '`n' Err.Line '`n' Err.File, 'Version', 0x10)
    }
}
H := AoEIIAIO.AddButton('yp', 'FIX')
H.SetFont('Bold s12')
CreateImageButton(H, 0, [[0xFFFFFF,, 0x0000FF, 4, 0x0000FF, 2], [0xE6E6E6], [0xCCCCCC], [0xFFFFFF,, 0xCCCCCC]]*)
H.OnEvent('Click', LaunchFixes)
Features['Main'].Push(H)
LaunchFixes(Ctrl, Info) {
    Try {
        Run('Fixes.ahk')
    } Catch Error As Err {
        MsgBox("Launch failed!`n`n" Err.Message '`n' Err.Line '`n' Err.File, 'Fix', 0x10)
    }
}
H := AoEIIAIO.AddButton('yp', 'LANGUAGE')
H.SetFont('Bold s12')
CreateImageButton(H, 0, [[0xFFFFFF,,, 4, 0x000000, 2], [0xE6E6E6], [0xCCCCCC], [0xFFFFFF,, 0xCCCCCC]]*)
H.OnEvent('Click', LaunchLanguage)
Features['Main'].Push(H)
LaunchLanguage(Ctrl, Info) {
    Try {
        Run('Language.ahk')
    } Catch Error As Err {
        MsgBox("Launch failed!`n`n" Err.Message '`n' Err.Line '`n' Err.File, 'Language', 0x10)
    }
}
H := AoEIIAIO.AddButton('yp', 'VISUAL MODS')
H.SetFont('Bold s12')
CreateImageButton(H, 0, [[0xFFFFFF,, 0x008000, 4, 0x008000, 2], [0xE6E6E6], [0xCCCCCC], [0xFFFFFF,, 0xCCCCCC]]*)
H.OnEvent('Click', LaunchVM)
Features['Main'].Push(H)
LaunchVM(Ctrl, Info) {
    Try {
        Run('VM.ahk')
    } Catch Error As Err {
        MsgBox("Launch failed!`n`n" Err.Message '`n' Err.Line '`n' Err.File, 'Language', 0x10)
    }
}
H := AoEIIAIO.AddButton('yp', 'DATA MODS')
H.SetFont('Bold s12')
CreateImageButton(H, 0, [[0xFFFFFF,, 0x008000, 4, 0x008000, 2], [0xE6E6E6], [0xCCCCCC], [0xFFFFFF,, 0xCCCCCC]]*)
H.OnEvent('Click', LaunchDM)
Features['Main'].Push(H)
LaunchDM(Ctrl, Info) {
    Try {
        Run('DM.ahk')
    } Catch Error As Err {
        MsgBox("Launch failed!`n`n" Err.Message '`n' Err.Line '`n' Err.File, 'Language', 0x10)
    }
}
H := AoEIIAIO.AddButton('xm', 'HIDE ALL IP')
H.SetFont('Bold s12')
CreateImageButton(H, 0, [[0xFFFFFF,, 0x804000, 4, 0x804000, 2], [0xE6E6E6], [0xCCCCCC], [0xFFFFFF,, 0xCCCCCC]]*)
H.OnEvent('Click', LaunchVPN)
Features['Main'].Push(H)
LaunchVPN(Ctrl, Info) {
    Try {
        Run('VPN.ahk')
    } Catch Error As Err {
        MsgBox("Launch failed!`n`n" Err.Message '`n' Err.Line '`n' Err.File, 'Language', 0x10)
    }
}
H := AoEIIAIO.AddButton('YP', 'SHORTCUTS')
H.SetFont('Bold s12')
CreateImageButton(H, 0, [[0xFFFFFF,, 0x804000, 4, 0x804000, 2], [0xE6E6E6], [0xCCCCCC], [0xFFFFFF,, 0xCCCCCC]]*)
H.OnEvent('Click', LaunchAHK)
Features['Main'].Push(H)
LaunchAHK(Ctrl, Info) {
    Try {
        Run('AHK.ahk')
    } Catch Error As Err {
        MsgBox("Launch failed!`n`n" Err.Message '`n' Err.Line '`n' Err.File, 'Language', 0x10)
    }
}
AoEIIAIO.Show()
R.Redraw()
; Graphics updates
AoEIIAIO.GetPos(,, &W, &H)
T.Move(0,, W)
T.Redraw()
P.Move((W - 373) / 2)
WD.Move(,, W)
WD.SetFont('Bold')
WD.OnEvent('Click', (*) => OpenGameFolder())
GameDirectory := IniRead(Config, 'Settings', 'GameDirectory', '')
If !ValidGameDirectory(GameDirectory) {
    P.Value := 'DB\000\gameoff.png'
    For Each, Version in Features['Main'] {
        Version.Enabled := False
    }
    If 'Yes' = MsgBox('Game is not yet located!, want to select now?', 'Game', 0x4 + 0x40) {
        Run('Game.ahk')
    }
    Return
}
WD.Text := 'Game: ' GameDirectory
CreateImageButton(WD, 0, [[0xCCCCCC], [0xB2B2B2], [0x999999], [0xCCCCCC,, 0xCCCCCC]]*)
; Stay up to date with the new selections
OnMessage(0x1001, GameUpdate)
GameUpdate(wParam, LParam, Msg, Hwnd) {
    If Msg = 0x1001 {
        Apps := IniRead(Config, 'PIDs',, '')
        Loop Parse, Apps, '`n', '`r' {
            PID := StrSplit(A_LoopField, '=')
            If ProcessExist(PID[2]) && PID[2] != ProcessExist() {
                Run(PID[1])
            }
        }
        Reload()
    }
}
; Opens the game folder
OpenGameFolder() {
    GameDirectory := IniRead(Config, 'Settings', 'GameDirectory', '')
    If ValidGameDirectory(GameDirectory) {
        Run(GameDirectory '\')
    }
}
;
;; AoE II Manager AIO class
;Class AoE_II_Manager_AIO {
;    ; Version
;    Version         := '2.0'
;    ; App server
;    Server          := 'https://raw.githubusercontent.com'
;    User            := 'SmileAoE'
;    Repositry       := 'aoeii_aio'
;    DownloadDB      := This.Server '/' This.User '/' This.Repositry '/main'
;    ; Packages
;    BasePackages    := ['DB/000.7z.001', 'DB/001.7z.001', 'DB/002.7z.001', 'DB/006.7z.001', 'DB/007.7z.001', 'DB/008.7z.001']
;    GamePackages    := ['DB/003.7z.001', 'DB/003.7z.002', 'DB/003.7z.003', 'DB/003.7z.004', 'DB/004.7z.001', 'DB/004.7z.002', 'DB/004.7z.003', 'DB/005.7z.001']
;    RestPackages    := ['DB/009.7z.001', 'DB/009.7z.002', 'DB/010.7z.001', 'DB/010.7z.002', 'DB/010.7z.003', 'DB/010.7z.004', 'DB/010.7z.005', 'DB/011.7z.001', 'DB/012.7z.001', 'DB/013.7z.001', 'DB/014.7z.001', 'DB/014.7z.002']
;    ; Reg key
;    InstallRegKey   := "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Age of Empires II AIO"
;    UninstallScript := '
;    (
;     Please do not use this, unless you know what you are doing!
;    #Requires AutoHotkey v2.0
;    #SingleInstance Force
;    If !A_Args.Length || A_Args[1] != 'aoeii_aio_uninstall_game_request' {
;        ExitApp()
;    }
;    Name := StrSplit(A_ScriptDir, '\')
;    Name := Name[Name.Length]
;    Unio := FileOpen(A_Temp '\Uninstall.ahk', 'w')
;    Unio.WriteLine('#Requires AutoHotkey v2.0')
;    Unio.WriteLine('#SingleInstance Force')
;    Unio.WriteLine("If !A_Args.Length || A_Args[1] != 'aoeii_aio_uninstall_game_request' {")
;    Unio.WriteLine("    ExitApp()")
;    Unio.WriteLine("}")
;    Unio.WriteLine("If 'Yes' = MsgBox('Are you sure want to uninstall " Name " ?', 'Uninstall', 0x4 + 0x40) {")
;    Unio.WriteLine("DirDelete('" A_ScriptDir "', 1)")
;    Unio.WriteLine("Msgbox('Uninstall completed!')")
;    Unio.WriteLine("}")
;    Run(A_Temp '\Uninstall.ahk aoeii_aio_uninstall_game_request')
;    )'
;    ; Base package hashs
;    LinkHashs       := This.DownloadDB '/DB/Hashsums.ini'
;    ; Configuration
;    Config          := 'Config.ini'
;    Update          := IniRead(This.Config, 'Settings', 'Update', 0)
;    ; Default folders
;    AppDir          := ['DB', 'Hotkeys', 'Records']
;    ; Default shortucts
;    Shortcut1       := '
;    (
;    ;Fast One Unit Un-Select;
;    #Requires AutoHotkey v2
;    #SingleInstance Force
;    GroupAdd('AOKAOC', 'ahk_exe empires2.exe')
;    GroupAdd('AOKAOC', 'ahk_exe age2_x1.exe')
;    GroupAdd('AOKAOC', 'ahk_exe age2_x2.exe')
;    HotIfWinActive("ahk_group AOKAOC")
;    Hotkey('!RButton', Action)
;    Action(*) {
;    WinGetPos(,, &W, &H, 'ahk_group AOKAOC')
;    If W != A_ScreenWidth || H != A_ScreenHeight
;    Return
;    MouseClick('Right', , , , 0)
;    MouseGetPos(&X, &Y)
;    SendInput('{LCtrl Down}')
;    MouseClick('Left', 315, A_ScreenHeight - 130, , 0)
;    SendInput('{Ctrl Up}')
;    MouseMove(X, Y, 0)
;    }
;    ProcessWaitClose(A_Args[1])
;    ExitApp
;    )'
;    Shortcut2       := '
;    (
;    ;Terminates The Game;
;    #Requires AutoHotkey v2
;    #SingleInstance Force
;    GroupAdd('AOKAOC', 'ahk_exe empires2.exe')
;    GroupAdd('AOKAOC', 'ahk_exe age2_x1.exe')
;    GroupAdd('AOKAOC', 'ahk_exe age2_x2.exe')
;    HotIfWinActive("ahk_group AOKAOC")
;    Hotkey('#q', Action)
;    Action(*) {
;    If GameIsRunning()
;    Msgbox('Game termination failure!', 'Game Terminate', 0x30)
;    }
;    GameIsRunning() {
;    Processes := ['empires2.exe', 'age2_x1.exe', 'age2_x2.exe']
;    For Each, Process in Processes {
;    If ProcessExist(Process) {
;    ProcessClose(Process)
;    }
;    ProcessWaitClose(Process, 5)
;    If ProcessExist(Process) {
;    Return True
;    }
;    }
;    Return False
;    }
;    ProcessWaitClose(A_Args[1])
;    ExitApp
;    )'
;    Shortcutslist   := [This.Shortcut1, This.Shortcut2]
;    ; 7zip
;    Location7z      := 'DB\7za.exe'
;    Link7z          := This.DownloadDB '/7za.exe'
;    Hash7z          := '80014d2b38a815f1a6ea220e679111c6'
;    7zPID           := 0
;    ; 
;    Features        := Map('Version', [], 'Game', [])
;    ; Versions maps
;    GameVersion     := Map('AOK'            , ['2.0  CD', '2.0a No CD', '2.0b CD']
;                         , 'AOKCombine'     , Map('2.0b CD', ['2.0a No CD'])
;                         , 'AOKHandle'      , Map()
;                         , 'AOC'            , ['1.0  CD', '1.0c No CD', '1.0e No CD', '1.1  No CD', '1.5  CD']
;                         , 'AOCCombine'     , Map('1.0e No CD', ['1.0c No CD'], '1.1  No CD', ['1.0c No CD'], '1.5  CD', ['1.0c No CD'])
;                         , 'AOCHandle'      , Map()
;                         , 'FE'             , ['2.2  CD'])
;    ; GameRanger
;    GRSetting       := A_AppData '\GameRanger\GameRanger Prefs\Settings'
;    GRApp           := A_AppData '\GameRanger\GameRanger\GameRanger.exe'
;    ; Start the app
;    __New() {
;        This.IsAdmin()
;        This.CreateDirectory()
;        This.CreateShortcuts()
;        This.UseGDIP()
;        This.GUILoading()
;        If This.GetConnectedState() && This.Update
;            This.UpdatedHashs := This.UpdatedPackagesHashs()
;        PackagesNumber := This.BasePackages.Length + This.GamePackages.Length + This.RestPackages.Length + 1
;        This.DoneSteps.Opt('Range1-' PackagesNumber)
;        For Each, Package in This.BasePackages {
;            ; Update the progress GUI
;            This.DoneSteps.Value += 1
;            This.Prepare.Title := 'Loading -> (' This.DoneSteps.Value ' / ' PackagesNumber ')...'
;            ; Set the needed parameters
;            PackagePath := StrReplace(Package, '/', '\')
;            PackageFolder := StrSplit(PackagePath, '.')[1]
;            If This.GetConnectedState() && This.Update
;                This.DownloadPackage(Package, PackagePath, PackageFolder)
;            This.ExtractPackage(PackagePath, PackageFolder)
;        }
;        For Each, Package in This.GamePackages {
;            ; Update the progress GUI
;            This.DoneSteps.Value += 1
;            This.Prepare.Title := 'Loading -> (' This.DoneSteps.Value ' / ' PackagesNumber ')...'
;            ; Set the needed parameters
;            PackagePath := StrReplace(Package, '/', '\')
;            PackageFolder := StrSplit(PackagePath, '.')[1]
;            If FileExist(PackagePath) {
;                If This.GetConnectedState() && This.Update
;                    This.DownloadPackage(Package, PackagePath, PackageFolder)
;            }
;        }
;        For Each, Package in This.RestPackages {
;            ; Update the progress GUI
;            This.DoneSteps.Value += 1
;            This.Prepare.Title := 'Loading -> (' This.DoneSteps.Value ' / ' PackagesNumber ')...'
;            ; Set the needed parameters
;            PackagePath := StrReplace(Package, '/', '\')
;            PackageFolder := StrSplit(PackagePath, '.')[1]
;            If FileExist(PackagePath) {
;                If This.GetConnectedState() && This.Update
;                    This.DownloadPackage(Package, PackagePath, PackageFolder)
;                This.ExtractPackage(PackagePath, PackageFolder)
;            }
;        }
;        This.Prepare.Destroy()
;        This.GUIManager()
;        This.SectionTitle()
;        This.GameLocation()
;        This.VersionSection()
;        ; Finally display the window
;        This.HMGUI.Show('x' (A_ScreenWidth - 640) / 2 ' y' (A_ScreenHeight - 500) / 2)
;        This.Loader()
;    }
;    ; Checks if the script run as admin
;    IsAdmin() {
;        If !A_IsAdmin {
;            MsgBox('Script must run as administrator!', 'Warn', 0x30)
;            ExitApp
;        }
;    }
;    ; Creates the app default dirs
;    CreateDirectory() {
;        Try {
;            For Every, Directory in This.AppDir {
;                If !DirExist(Directory) {
;                    DirCreate(Directory)
;                }
;            }
;        } Catch Error As Err {
;            MsgBox('Unexpected error occured!'
;                 . '`n`nMessage: `n-> [' Err.Message ']'
;                 . '`n`nFunction: `n-> [' Err.What ']'
;                 . '`n`nExtra: `n-> [' Err.Extra ']'
;                 . '`n`nFile: `n-> [' Err.File ']'
;                 . '`n`nLine: `n-> [' Err.Line ']'
;                 . '`n`nYou can contact this app provider!'
;                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
;            ExitApp
;        }
;    }
;    ; Creates and run the default game shortcuts
;    CreateShortcuts() {
;        Try {
;            For Every, Shortcut in This.Shortcutslist {
;                If !FileExist(This.AppDir[2] '\00' Every '.ahk') || FileRead(This.AppDir[2] '\00' Every '.ahk') != Shortcut
;                {
;                    O := FileOpen(This.AppDir[2] '\00' Every '.ahk', 'w')
;                    O.Write(Shortcut)
;                    O.Close()
;                }
;                Run(This.AppDir[2] '\00' Every '.ahk ' ProcessExist())
;            }
;        } Catch Error As Err {
;            MsgBox('Unexpected error occured!'
;                 . '`n`nMessage: `n-> [' Err.Message ']'
;                 . '`n`nFunction: `n-> [' Err.What ']'
;                 . '`n`nExtra: `n-> [' Err.Extra ']'
;                 . '`n`nFile: `n-> [' Err.File ']'
;                 . '`n`nLine: `n-> [' Err.Line ']'
;                 . '`n`nYou can contact this app provider!'
;                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
;            ExitApp
;        }
;    }
;    ; Loads and initializes the Gdiplus.dll.
;    UseGDIP() {
;        Static GdipObject := 0
;        If !IsObject(GdipObject) {
;            GdipToken := 0
;            SI := Buffer(24, 0) ; size of 64-bit structure
;            NumPut("UInt", 1, SI)
;            If DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", &GdipToken, "Ptr", SI, "Ptr", 0, "UInt") {
;                Return 4
;            }
;            GdipObject := { __Delete: UseGdipShutDown }
;        }
;        UseGdipShutDown(*) {
;            DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", GdipToken)
;        }
;    }
;    ; Checks the internet connection
;    GetConnectedState() {
;        Return DllCall("Wininet.dll\InternetGetConnectedState", "Str", Flag := 0x40, "Int", 0)
;    }
;    ; Checks the unpacker
;    PrepareTheUnpacker() {
;        Try {
;            If !FileExist(This.Location7z) || This.HashFile(This.Location7z) != This.Hash7z {
;                Download(This.Link7z, This.Location7z)
;            }
;        } Catch Error As Err {
;            MsgBox('Unexpected error occured!'
;                 . '`n`nMessage: `n-> [' Err.Message ']'
;                 . '`n`nFunction: `n-> [' Err.What ']'
;                 . '`n`nExtra: `n-> [' Err.Extra ']'
;                 . '`n`nFile: `n-> [' Err.File ']'
;                 . '`n`nLine: `n-> [' Err.Line ']'
;                 . '`n`nYou can contact this app provider!'
;                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
;            ExitApp
;        }
;    }
;    ; Adds the title
;    SectionTitle() {
;        This.Thumb := This.HMGUI.AddPicture('xm+118', 'DB\000\gameoff.png')
;        This.Thumb.Focus()
;        This.Title := This.HMGUI.AddText('xm yp+200 Center w600 h50 cGray', 'Age of Empire II Easy Manager AIO v' This.Version)
;        This.Title.SetFont('s16')
;        This.Log := This.HMGUI.AddListview('-Hdr xm+150 w300 r8 -E0x200', ['S', 'L'])
;        This.Log.ModifyCol(1, '30 Center')
;        This.Log.OnEvent('ItemSelect', ObjBindMethod(This, 'LoggerSelectColor'))
;        This.LogCLV := LV_Colors(This.Log)
;        This.Reloader := This.HMGUI.AddButton('wp Disabled', 'RELOAD')
;        This.Reloader.OnEvent('Click', ObjBindMethod(This, 'ReloadGame'))
;    }
;    ; Reloads for a game selection
;    ReloadGame(Ctrl, Info) {
;        Ctrl.Enabled := False
;        This.Loader()
;        Ctrl.Enabled := True
;    }
;    ; Updates the log list
;    Logger(Text, OK := 0, R := 0) {
;        Switch R {
;            Case 0 :
;                R := This.Log.Add(, '→', Text)
;            Default :
;                Switch OK {
;                    Case 0 :
;                        Color := 'Green'
;                        TextStatus := 'OK'
;                    Case 1 :
;                        Color := 'Red'
;                        TextStatus := '!OK'
;                    Case 3 :
;                        Color := '0xDC6F00'
;                        TextStatus := '!FIX'
;                }
;                This.Log.Modify(R,, TextStatus)
;                This.LogCLV.Cell(R, 1, Color, 'White')
;                This.Log.Modify(R,,, Text)
;                This.LogCLV.Cell(R, 2,, Color)
;        }
;        This.Log.ModifyCol(2, 'AutoHdr')
;        Return R
;    }
;    ; Updates the log list select color
;    LoggerSelectColor(Ctrl, Item, Selected) {
;        If !Selected {
;            Return
;        }
;        Status := Ctrl.GetText(Item, 1)
;        Switch Status {
;            Case 'OK' : This.LogCLV.SelectionColors(0x008000, 0xFFFFFF)
;            Case '!FIX' : This.LogCLV.SelectionColors(0xDC6F00, 0xFFFFFF)
;            Case '!OK' : This.LogCLV.SelectionColors(0xFF0000, 0xFFFFFF)
;        }
;    }
;    ; Adds the game location
;    GameLocation() {
;        H := This.HMGUI.AddText('xm cBlue w600 h30', 'GAME LOCATION:')
;        H.SetFont('s16')
;        This.Features['Game'].Push(H)
;        This.GameDirectory := This.HMGUI.AddEdit('ReadOnly xm+20 w580 -E0x200 Border')
;        This.Features['Game'].Push(This.GameDirectory)
;        H := This.HMGUI.AddButton('w100', 'Select')
;        This.Features['Game'].Push(H)
;        This.GuiButtonIcon(H, 'DB\000\folder.png',, 'A1')
;        H.OnEvent('Click', ObjBindMethod(This, 'SelectDirectory'))
;        H := This.HMGUI.AddButton('w170 yp', 'Select from GameRanger')
;        This.GuiButtonIcon(H, 'DB\000\gr.png',, 'A1')
;        H.OnEvent('Click', ObjBindMethod(This, 'SelectDirectoryGR'))
;        This.Features['Game'].Push(H)
;        H := This.HMGUI.AddButton('w140 yp', 'Open the selected')
;        This.GuiButtonIcon(H, 'DB\000\sfolder.png',, 'A1')
;        This.Features['Game'].Push(H)
;        H.OnEvent('Click', (*) => This.GameDirectory.Value ? Run(This.GameDirectory.Value '\') : 0)
;        H := This.HMGUI.AddCheckBox('xm+20 Checked', 'Perform a common issues fix on each load')
;        This.Features['Game'].Push(H)
;        H.OnEvent('Click', ObjBindMethod(this, 'CommonIssueFix'))
;        H := This.HMGUI.AddButton('xm+20 w200', 'Download and install the game')
;        This.GuiButtonIcon(H, 'DB\000\download.png',, 'A1')
;        This.Features['Game'].Push(H)
;        H.OnEvent('Click', ObjBindMethod(this, 'DownloadGame'))
;    }
;    ; Selects a location from GR
;    SelectDirectoryGR(Ctrl, Info) {
;        Ctrl.Enabled := False
;        Text := This.BinGrabText(This.GRSetting)
;        Locations := This.TextGrabPath(Text, ['empires2.exe', 'age2_x1.exe', 'age2_x2.exe'])
;        For Location in Locations {
;            If RC := This.ValidGameDirectory(Location) {
;                Choice := MsgBox('Want to select this location?`n`n' Location, 'Game Location', 0x4 + 0x40)
;                If Choice = 'Yes' {
;                    This.GameDirectory.Value := Location
;                    IniWrite(Location, This.Config, 'Settings', 'GameDirectory')
;                    This.Loader(Location)
;                    MsgBox('Game selected sucessfully!', 'Game Select', 0x40 ' T5')
;                    Break
;                }
;            }
;        }
;        Ctrl.Enabled := True
;    }
;    ; Selects a location
;    SelectDirectory(Ctrl, Info) {
;        Ctrl.Enabled := False
;        If SelectedDirectory := FileSelect('D', 'C:\' (A_Is64bitOS ? 'Program Files (x86)' : 'Program Files') '\Microsoft Games') {
;            If !Valid := This.ValidGameDirectory(SelectedDirectory) {
;                SelectedDirectoryEx := SelectedDirectory
;                SelectedDirectory := ''
;                SplitPath(SelectedDirectoryEx, &_, &ParentSelectedDirectory)
;                If Valid := This.ValidGameDirectory(ParentSelectedDirectory) {
;                    Choice := MsgBox('Want to select this location?`n`n' ParentSelectedDirectory, 'Game Location', 0x4 + 0x40)
;                    If Choice = 'Yes' {
;                        SelectedDirectory := ParentSelectedDirectory
;                    }
;                }
;            }
;            If !Valid {
;                Loop Files, SelectedDirectoryEx '\*', 'D' {
;                    If This.ValidGameDirectory(A_LoopFileFullPath) {
;                        Choice := MsgBox('Want to select this location?`n`n' A_LoopFileFullPath, 'Game Location', 0x4 + 0x40)
;                        If Choice = 'Yes' {
;                            SelectedDirectory := A_LoopFileFullPath
;                            Break
;                        }
;                    }
;                }
;            }
;            If SelectedDirectory != '' {
;                SelectedDirectory := StrUpper(SelectedDirectory)
;                This.GameDirectory.Value := SelectedDirectory
;                IniWrite(SelectedDirectory, This.Config, 'Settings', 'GameDirectory')
;                This.Loader(SelectedDirectory)
;                MsgBox('Game selected sucessfully!', 'Game Select', 0x40 ' T5')
;            }
;            Else {
;                MsgBox('Invalid game location!', 'Game Select', 0x30)
;            }
;        }
;        Ctrl.Enabled := True
;    }
;    ; Downloads and installs the game
;    DownloadGame(Ctrl, Info) {
;        Try {
;            If !This.GetConnectedState() {
;                MsgBox('Make sure you are connected to the internet!', "Can't download!", 0x30)
;                Return
;            }
;            If (GameDirectory := FileSelect('D',, 'Game install location')) && 'Yes' = MsgBox('Are you sure want to install at this location?`n' GameDirectory, 'Game install location', 0x40 + 0x4) {
;                GameDirectory := RegExReplace(GameDirectory, "\\$")
;                GameDirectory := GameDirectory '\Age of Empires II'
;                If !DirExist(GameDirectory) {
;                    DirCreate(GameDirectory)
;                }
;                If This.ValidGameDirectory(GameDirectory) && 'Yes' != MsgBox('It seems like the game already installed at this location!`nWant continue?', 'Game location install', 0x30 + 0x4) {
;                    Return
;                }
;                ; Check for packages updates
;                Ctrl.Enabled := False
;                This.UpdatedHashs := This.UpdatedPackagesHashs()
;                This.GUILoading(0)
;                This.DoneSteps.Opt('Range0-' (PackagesNumber := This.GamePackages.Length + 3))
;                This.DoneSteps.Value := 0
;                For Each, Package in This.GamePackages {
;                    ; Update the progress GUI
;                    This.DoneSteps.Value += 1
;                    This.Prepare.Title := 'Downloading -> (' This.DoneSteps.Value ' / ' PackagesNumber ')...'
;                    ; Set the needed parameters
;                    PackagePath := StrReplace(Package, '/', '\')
;                    PackageFolder := StrSplit(PackagePath, '.')[1]
;                    This.DownloadPackage(Package, PackagePath, PackageFolder)
;                }
;                ; AOK extract
;                This.Prepare.Title := 'Exporting -> (' This.DoneSteps.Value ' / ' PackagesNumber ')...'
;                This.ExtractPackage('DB\003.7z.001', GameDirectory, 1)
;                This.DoneSteps.Value += 1
;                ; AOC extract
;                This.Prepare.Title := 'Exporting -> (' This.DoneSteps.Value ' / ' PackagesNumber ')...'
;                This.ExtractPackage('DB\004.7z.001', GameDirectory, 1)
;                This.DoneSteps.Value += 1
;                ; FE extract
;                This.Prepare.Title := 'Exporting -> (' This.DoneSteps.Value ' / ' PackagesNumber ')...'
;                This.ExtractPackage('DB\005.7z.001', GameDirectory, 1)
;                This.DoneSteps.Value += 1
;                ; Add the reg keys
;                This.UpdateGameReg(GameDirectory)
;                If 'Yes' = MsgBox('Game installation complete!`nWanna select this game?', 'Game install location', 0x4 + 0x40) {
;                    This.Loader(GameDirectory)
;                }
;                This.Prepare.Destroy()
;            }
;            Ctrl.Enabled := True
;        } Catch Error As Err {
;            MsgBox('Unexpected error occured!'
;                 . '`n`nMessage: `n-> [' Err.Message ']'
;                 . '`n`nFunction: `n-> [' Err.What ']'
;                 . '`n`nExtra: `n-> [' Err.Extra ']'
;                 . '`n`nFile: `n-> [' Err.File ']'
;                 . '`n`nLine: `n-> [' Err.Line ']'
;                 . '`n`nYou can contact this app provider!'
;                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
;            ExitApp
;        }
;    }
;    ; Updates game install registery settings
;    UpdateGameReg(GameDirectory) {
;        Try {
;            RegWrite('Age of Empires II AIO', 'REG_SZ', This.InstallRegKey, 'DisplayName')
;            RegWrite('AOK (2.0) / AOC (1.0) / FE (2.1)', 'REG_SZ', This.InstallRegKey, 'DisplayVersion')
;            RegWrite(GameDirectory '\age2_x1\age2_x1.exe', 'REG_SZ', This.InstallRegKey, 'DisplayIcon')
;            RegWrite(GameDirectory, 'REG_SZ', This.InstallRegKey, 'InstallLocation')
;            RegWrite(1, 'REG_DWORD', This.InstallRegKey, 'NoModify')
;            RegWrite(1, 'REG_DWORD', This.InstallRegKey, 'NoRepair')
;            RegWrite(This.FolderGetSize(GameDirectory), 'REG_DWORD', This.InstallRegKey, 'EstimatedSize')
;            RegWrite('Microsoft Corporation', 'REG_SZ', This.InstallRegKey, 'Publisher')
;            ;RegWrite(GameDirectory '\Uninstall.ahk "aoeii_aio_uninstall_game_request"', 'REG_SZ', This.InstallRegKey, 'UninstallString')
;        } Catch Error As Err {
;            MsgBox('Unexpected error occured!'
;                 . '`n`nMessage: `n-> [' Err.Message ']'
;                 . '`n`nFunction: `n-> [' Err.What ']'
;                 . '`n`nExtra: `n-> [' Err.Extra ']'
;                 . '`n`nFile: `n-> [' Err.File ']'
;                 . '`n`nLine: `n-> [' Err.Line ']'
;                 . '`n`nYou can contact this app provider!'
;                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
;            ExitApp
;        } 
;    }
;    ; Returns a folder size in KB
;    FolderGetSize(Location) {
;        Size := 0
;        Loop Files, Location '\*.*', 'R' {
;            Size += FileGetSize(A_LoopFileFullPath, 'K')
;        }
;        Return Size
;    }
;    ; Checks if a directory contains the game
;    ValidGameDirectory(Location) {
;        Return   FileExist(Location '\empires2.exe')
;              && FileExist(Location '\language.dll')
;              && FileExist(Location '\Data\graphics.drs')
;              && FileExist(Location '\Data\interfac.drs')
;              && FileExist(Location '\Data\terrain.drs') ? 1 : 0
;    }
;    ; Grabs the readable text from binary file
;    BinGrabText(Filepath) {
;        Try {
;            Text := ''
;            BufferObj := FileRead(Filepath, 'RAW')
;            Loop BufferObj.Size {
;                Address := A_Index - 1
;                Byte := NumGet(BufferObj, Address, 'UChar')
;                If (C := Chr(Byte)) != '' {
;                    Text .= C
;                }
;            }
;            Return Text
;        } Catch Error As Err {
;            MsgBox('Unexpected error occured!'
;                 . '`n`nMessage: `n-> [' Err.Message ']'
;                 . '`n`nFunction: `n-> [' Err.What ']'
;                 . '`n`nExtra: `n-> [' Err.Extra ']'
;                 . '`n`nFile: `n-> [' Err.File ']'
;                 . '`n`nLine: `n-> [' Err.Line ']'
;                 . '`n`nYou can contact this app provider!'
;                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
;            ExitApp
;        }
;    }
;    ; Parses the game locations out of a text
;    TextGrabPath(TextFound, Excutables) {
;        Try {
;            ResultMap := Map()
;            For Each, Excutable in Excutables {
;                P := InStr(TextFound, LFE := Excutable,, -1)
;                Loop {
;                    Char := SubStr(TextFound, P - (I := A_Index), 1)
;                    LFE := Char LFE
;                } Until (Char = ':' || Ord(Char) = 10 || Ord(Char) = 13)
;                FoundPath := SubStr(TextFound, P - (I + 1), 1) LFE
;                FoundPath := StrReplace(FoundPath, '\' Excutables[1])
;                FoundPath := StrReplace(FoundPath, '\age2_x1\' Excutables[2])
;                FoundPath := StrReplace(FoundPath, '\age2_x1\' Excutables[3])
;                ResultMap[StrUpper(FoundPath)] := True
;            }
;            Return ResultMap
;        } Catch Error As Err {
;            MsgBox('Unexpected error occured!'
;                 . '`n`nMessage: `n-> [' Err.Message ']'
;                 . '`n`nFunction: `n-> [' Err.What ']'
;                 . '`n`nExtra: `n-> [' Err.Extra ']'
;                 . '`n`nFile: `n-> [' Err.File ']'
;                 . '`n`nLine: `n-> [' Err.Line ']'
;                 . '`n`nYou can contact this app provider!'
;                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
;            ExitApp
;        }
;    }
;    ; Common issues fix setting
;    CommonIssueFix(Ctrl, Info) {
;        IniWrite(Ctrl.Value, This.Config, 'Settings', 'CommonFix')
;    }
;    ; Adds the versions
;    VersionSection() {
;        H := This.HMGUI.AddText('xm cBlue w600 h30', 'GAME VERSION:')
;        H.GetPos(, &Y)
;        Y += 30
;        H.SetFont('s16')
;        This.Features['Version'].Push(H)
;        H := This.HMGUI.AddPicture('xm+20 y' Y, 'DB\000\aok.png')
;        This.Features['Version'].Push(H)
;        H := This.HMGUI.AddText('cRed', 'The Age of Kings')
;        This.Features['Version'].Push(H)
;        For Each, AOK in This.GameVersion['AOK'] {
;            H := This.HMGUI.AddRadio('w150', AOK)
;            H.SetFont(, 'Consolas')
;            This.Features['Version'].Push(H)
;            H.OnEvent('Click', ObjBindMethod(this, 'ApplyVersion'))
;            This.GameVersion['AOKHandle'][AOK] := H
;        }
;        H := This.HMGUI.AddPicture('xm+220 y' Y, 'DB\000\aoc.png')
;        This.Features['Version'].Push(H)
;        H := This.HMGUI.AddText('cBlue', 'The Conquerors')
;        This.Features['Version'].Push(H)
;        For Each, AOC in This.GameVersion['AOC'] {
;            H := This.HMGUI.AddRadio('w150', AOC)
;            H.SetFont(, 'Consolas')
;            This.Features['Version'].Push(H)
;            H.OnEvent('Click', ObjBindMethod(this, 'ApplyVersion'))
;            This.GameVersion['AOCHandle'][AOC] := H
;        }
;        H := This.HMGUI.AddPicture('xm+440 y' Y, 'DB\000\fe.png')
;        This.Features['Version'].Push(H)
;        H := This.HMGUI.AddText('cGreen', 'Forgotten Empires')
;        This.Features['Version'].Push(H)
;        For Each, FE in This.GameVersion['FE'] {
;            H := This.HMGUI.AddRadio('w150 Checked', FE)
;            H.SetFont(, 'Consolas')
;            This.Features['Version'].Push(H)
;        }
;        This.OLV := This.HMGUI.AddListView('xm+20 r4 -Hdr Checked -E0x200 -Multi', ['Option'])
;        This.Features['Version'].Push(This.OLV)
;        This.OLVC := LV_Colors(This.OLV)
;        This.OLVC.SelectionColors
;        This.OLV.Add(, 'Advanced interface')
;        This.OLV.Add(, 'Advanced interface + Overlay')
;        This.OLV.Add(, 'Widescreen')
;        This.OLV.Add(, 'Centered widescreen')
;        This.OLV.OnEvent('ItemCheck', ApplyFix)
;        This.OLV.OnEvent('ItemSelect', ApplyFix)
;        ApplyFix(Ctrl, Item, CheckedSelected) {
;            If !CheckedSelected {
;                RegWrite(0, 'REG_DWORD', 'HKEY_CURRENT_USER\SOFTWARE\Microsoft\Microsoft Games\Age of Empires', 'Aoe2Patch')
;                IniWrite(0, This.Config, 'Settings', 'Fix')
;                Return
;            }
;            RegWrite(Item, 'REG_DWORD', 'HKEY_CURRENT_USER\SOFTWARE\Microsoft\Microsoft Games\Age of Empires', 'Aoe2Patch')
;            IniWrite(Item, This.Config, 'Settings', 'Fix')
;            Loop Ctrl.GetCount() {
;                If A_Index != Item {
;                    Ctrl.Modify(A_Index, '-Check')
;                }
;            }
;            Ctrl.Modify(Item, 'Select')
;            Ctrl.Modify(Item, 'Check')
;        }
;    }
;    ; Applys the version
;    ApplyVersion(Radio, Info) {
;        Try {
;            ; Checks the game directory
;            If !This.ValidGameDirectory(This.GameDirectory.Value) {
;                MsgBox('Invalid game location!', 'Game Select', 0x30)
;                Radio.Value := 0
;                Return
;            }
;            This.EnableControls(This.Features['Version'], 0)
;            ; Close the game if running
;            For Each, Process in ['empires2.exe', 'age2_x1.exe', 'age2_x2.exe'] {
;                If ProcessExist(Process) {
;                    ProcessClose(Process)
;                }
;            }
;            ; Cleans up previous versions files
;            TargetVersion := SubStr(Radio.Text, 1, 1)
;            Loop Files, 'DB\002\*', 'D' {
;                If TargetVersion != SubStr(Version := A_LoopFileName, 1, 1) {
;                    Continue
;                }
;                Loop Files, 'DB\002\' Version '\*.*', 'R' {
;                    PathFile := StrReplace(A_LoopFileDir '\' A_LoopFileName, 'DB\002\' Version '\')
;                    If FileExist(This.GameDirectory.Value '\' PathFile) {
;                        FileDelete(This.GameDirectory.Value '\' PathFile)
;                    }
;                }
;            }
;            ; Cleans up previous fix files
;            Loop Files, 'DB\001\*', 'D' {
;                Fix := A_LoopFileName
;                Loop Files, 'DB\001\' Fix '\*', 'D' {
;                    If TargetVersion != SubStr(Version := A_LoopFileName, 1, 1) {
;                        Continue
;                    }
;                    Loop Files, 'DB\001\' Fix '\' Version '\*.*', 'R' {
;                        PathFile := StrReplace(A_LoopFileDir '\' A_LoopFileName, 'DB\001\' Fix '\' Version '\')
;                        If FileExist(This.GameDirectory.Value '\' PathFile) {
;                            FileDelete(This.GameDirectory.Value '\' PathFile)
;                        }
;                    }
;                }
;            }
;            ; Copy the selected version files
;            Key := TargetVersion = '1' ? 'AOCCombine' : 'AOKCombine'
;            If This.GameVersion[Key].Has(Radio.Text) {
;                For Each, Version in This.GameVersion[Key][Radio.Text] {
;                    If DirExist('DB\002\' Version) {
;                        DirCopy('DB\002\' Version, This.GameDirectory.Value, 1)
;                    }
;                }
;            }
;            If DirExist('DB\002\' Radio.Text) {
;                DirCopy('DB\002\' Radio.Text, This.GameDirectory.Value, 1)
;            }
;            ; Copy fixs
;            If Fix := This.OLV.GetNext(0, 'C') {
;                DirCopy('DB\001\Enable Fix v2\Static', This.GameDirectory.Value, 1)
;                If DirExist('DB\001\Enable Fix v2\' Radio.Text) {
;                    DirCopy('DB\001\Enable Fix v2\' Radio.Text, This.GameDirectory.Value, 1)
;                }
;            }
;            This.EnableControls(This.Features['Version'])
;            SoundPlay('DB\000\30 Wololo.mp3')
;        } Catch Error As Err {
;            MsgBox('Unexpected error occured!'
;                 . '`n`nMessage: `n-> [' Err.Message ']'
;                 . '`n`nFunction: `n-> [' Err.What ']'
;                 . '`n`nExtra: `n-> [' Err.Extra ']'
;                 . '`n`nFile: `n-> [' Err.File ']'
;                 . '`n`nLine: `n-> [' Err.Line ']'
;                 . '`n`nYou can contact this app provider!'
;                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
;            ExitApp
;        }
;    }
;    ; Updates the desktop shortcuts
;    UpdateShortcuts(FileLocation, Name := '') {
;        Try {
;            SplitPath(FileLocation, &_, &OutDir, &_, &OutNameNoExt)
;            FileShortcut := A_Desktop '\' (Name != '' ? Name : OutNameNoExt) '.lnk'
;            If !FileExist(FileShortcut) && FileExist(FileLocation) {
;                FileCreateShortcut(FileLocation, FileShortcut, OutDir)
;            }
;            If FileExist(FileShortcut) && !FileExist(FileLocation) {
;                FileRecycle(FileShortcut)
;            }
;            If FileExist(FileShortcut) && FileExist(FileLocation) {
;                FileGetShortcut(FileShortcut, &OutTarget)
;                If OutTarget != FileLocation {
;                    FileCreateShortcut(FileLocation, FileShortcut, OutDir)
;                }
;            }
;        } Catch Error As Err {
;            MsgBox('Unexpected error occured!'
;                 . '`n`nMessage: `n-> [' Err.Message ']'
;                 . '`n`nExtra: `n-> [' Err.Extra ']'
;                 . '`n`nFile: `n-> [' Err.File ']'
;                 . '`n`nLine: `n-> [' Err.Line ']'
;                 . '`n`nYou can contact this app provider!'
;                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
;            ExitApp
;        }
;    }
;    ; Clear unwanted files
;    ClearUnwanted(FileLocation) {
;        Try {
;            If FileExist(FileLocation) {
;            FileRecycle(FileLocation)
;            }
;        } Catch Error As Err {
;            MsgBox('Unexpected error occured!'
;                 . '`n`nMessage: `n-> [' Err.Message ']'
;                 . '`n`nFunction: `n-> [' Err.What ']'
;                 . '`n`nExtra: `n-> [' Err.Extra ']'
;                 . '`n`nFile: `n-> [' Err.File ']'
;                 . '`n`nLine: `n-> [' Err.Line ']'
;                 . '`n`nYou can contact this app provider!'
;                 . '`nEmail: chandoul.mohamed26@gmail.com', 'Failed', 0x10)
;            ExitApp
;        }
;    }
;    ; Controls enable or disable
;    EnableControls(Controls, Enable := 1) {
;        If Enable {
;            For Each, Control in Controls {
;                Control.Enabled := True
;            }
;        } Else {
;            For Each, Control in Controls {
;                Control.Enabled := False
;            }
;        }
;    }
;    ; Loads a game
;    Loader(GameDirectoryLoad := '') {
;        ; Disable features
;        This.EnableControls(This.Features['Version'], 0)
;        ; Game location section loads
;        If !This.GameLocationLoads(GameDirectoryLoad) {
;            Return
;        }
;        ; Enable title
;        This.Thumb.Value := 'DB\000\game.png'
;        This.Title.Opt('cBlack')
;        ; Version section loads
;        This.VersionLoads()
;        ; Other updates
;        This.OLV.Redraw()
;        This.Reloader.Enabled := True
;    }
;    ; Game location section loads
;    GameLocationLoads(GameDirectoryLoad) {
;        ;1
;        This.Log.Delete()
;        R := This.Logger('Looking for the game folder...')
;        GameDirectory := GameDirectoryLoad != '' ? GameDirectoryLoad : IniRead(This.Config, 'Settings', 'GameDirectory', '')
;        If !This.ValidGameDirectory(GameDirectory) {
;            This.Logger('No game folder is selected', 1, R)
;            Return False
;        }
;        This.GameDirectory.Value := GameDirectory
;        IniWrite(This.GameDirectory.Value, This.Config, 'Settings', 'GameDirectory')
;        This.Logger('Game is located at: ' This.GameDirectory.Value,, R)
;        ;2
;        If IniRead(This.Config, 'Settings', 'CommonFix', 0) {
;            ;1
;            R := This.Logger('Performing a common issue fix...')
;            If FileExist(This.GameDirectory.Value '\age2_x1.exe') {
;                If !DirExist(This.GameDirectory.Value '\age2_x1') {
;                    DirCreate(This.GameDirectory.Value '\age2_x1')
;                }
;                FileMove(This.GameDirectory.Value '\age2_x1.exe', This.GameDirectory.Value '\age2_x1', 1)
;            }
;            ;2
;            This.ClearUnwanted(This.GameDirectory.Value '\windmode.ini')
;            This.ClearUnwanted(This.GameDirectory.Value '\age2_x1\windmode.ini')
;            ;3
;            This.UpdateShortcuts(This.GameDirectory.Value '\empires2.exe', 'The Age of Kings')
;            This.UpdateShortcuts(This.GameDirectory.Value '\age2_x1\age2_x1.exe', 'The Conquerors')
;            This.UpdateShortcuts(This.GameDirectory.Value '\age2_x1\age2_x2.exe', 'Forgotten Empires')
;            This.Logger('Perform common issue fix',, R)
;        }
;        Return True
;    }
;    ; Version section loads
;    VersionLoads() {
;        For Each, Control in This.Features['Version'] {
;            If Type(Control) = 'Gui.Radio' {
;                Control.Value := 0
;            }
;        }
;        R := This.Logger('Scanning the game versions')
;        Loop Files, 'DB\002\2.*', 'D' {
;            Version := A_LoopFileName
;            VersionIsSet := Version
;            Loop Files, 'DB\002\' Version '\*.*', 'R' {
;                PathFile := StrReplace(A_LoopFileDir '\' A_LoopFileName, 'DB\002\' Version '\')
;                If !FileExist(This.GameDirectory.Value '\' PathFile) && VersionIsSet {
;                    VersionIsSet := ''
;                    Break
;                }
;                CurrentHash := This.HashFile(A_LoopFileFullPath)
;                FoundHash := This.HashFile(This.GameDirectory.Value '\' PathFile)
;                If (CurrentHash != FoundHash) && VersionIsSet {
;                    VersionIsSet := ''
;                    Break
;                }
;            }
;            If VersionIsSet {
;                This.GameVersion['AOKHandle'][VersionIsSet].Value := 1
;                This.Logger('AOK ' VersionIsSet ' found',, R)
;            }
;            If This.Log.GetText(R, 1) != 'OK' {
;                This.Logger('AOK version not found', 3, R)
;            }
;        }
;        R := This.Logger('Scanning the game versions')
;        Loop Files, 'DB\002\1.*', 'D' {
;            Version := A_LoopFileName
;            VersionIsSet := Version
;            Loop Files, 'DB\002\' Version '\*.*', 'R' {
;                PathFile := StrReplace(A_LoopFileDir '\' A_LoopFileName, 'DB\002\' Version '\')
;                If !FileExist(This.GameDirectory.Value '\' PathFile) {
;                    VersionIsSet := ''
;                    Break
;                }
;                CurrentHash := This.HashFile(A_LoopFileFullPath)
;                FoundHash := This.HashFile(This.GameDirectory.Value '\' PathFile)
;                If (CurrentHash != FoundHash) && VersionIsSet != '' {
;                    VersionIsSet := ''
;                    Break
;                }
;            }
;            If VersionIsSet != '' {
;                This.GameVersion['AOCHandle'][VersionIsSet].Value := 1
;                This.Logger('AOC ' VersionIsSet ' found',, R)
;            }
;            If This.Log.GetText(R, 1) != 'OK' {
;                This.Logger('AOC version not found', 3, R)
;            }
;        }
;        R := This.Logger('Scanning the game fixes')
;        FixIsSetAOC := True
;        FixIsSetAOK := True
;        Loop Files, 'DB\001\Enable Fix v2\Static\*.*', 'R' {
;            PathFile := StrReplace(A_LoopFileDir '\' A_LoopFileName, 'DB\001\Enable Fix v2\Static\')
;            If !FileExist(This.GameDirectory.Value '\' PathFile) {
;                FixIsSetAOC := False
;                FixIsSetAOK := False
;                Break
;            }
;            CurrentHash := This.HashFile(A_LoopFileFullPath)
;            FoundHash := This.HashFile(This.GameDirectory.Value '\' PathFile)
;            If (CurrentHash != FoundHash) && FixIsSetAOC {
;                FixIsSetAOC := False
;                FixIsSetAOK := False
;                Break
;            }
;        }
;        If FixIsSetAOK && (!FileExist(This.GameDirectory.Value '\dsound.dll') || This.HashFile(This.GameDirectory.Value '\dsound.dll') != This.HashFile('DB\001\Enable Fix v2\2.0  CD\dsound.dll')) {
;            FixIsSetAOK := False
;        }
;        If FixIsSetAOC && (!FileExist(This.GameDirectory.Value '\age2_x1\dsound.dll') || This.HashFile(This.GameDirectory.Value '\age2_x1\dsound.dll') != This.HashFile('DB\001\Enable Fix v2\1.0  CD\age2_x1\dsound.dll')) {
;            FixIsSetAOC := False
;        }
;        Fix := IniRead(This.Config, 'Settings', 'Fix', 0)
;        If Fix && FixIsSetAOK {
;            This.OLV.Modify(Fix, 'Select')
;            This.OLV.Modify(Fix, 'Check')
;            This.Logger('AOK ' This.OLV.GetText(Fix, 1) ' mod found',, R)
;        } Else {
;            This.Logger('AOK No fix is applied', 3, R)
;        }
;        R := This.Logger('Scanning the game fixes')
;        If Fix && FixIsSetAOC {
;            This.OLV.Modify(Fix, 'Select')
;            This.OLV.Modify(Fix, 'Check')
;            This.Logger('AOC ' This.OLV.GetText(Fix, 1) ' mod found',, R)
;        } Else {
;            This.Logger('AOC No fix is applied', 3, R)
;        }
;        This.EnableControls(This.Features['Version'])
;    }
;}
