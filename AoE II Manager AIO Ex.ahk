#Include SharedLib.ahk
Features['Main'] := []
WD := AoEIIAIO.AddButton('x0 y0', '...')
AoEIIAIO.SetFont('Bold s20')
T := AoEIIAIO.AddText('xm cGreen Center', 'Age of Empires II Easy Manager v' Version)
P := AoEIIAIO.AddPicture('xm+90', 'DB\000\game.png')
R := AoEIIAIO.AddButton('xm ym+30', 'RELOAD')
R.SetFont('Bold s12')
CreateImageButton(R, 0, IBRed*)
R.OnEvent('Click', (*) => Reload())
U := AoEIIAIO.AddButton(, 'UPDATE')
U.SetFont('Bold s12')
CreateImageButton(U, 0, IBBlue*)
U.OnEvent('Click', Check4Updates)
H := AoEIIAIO.AddButton('xm', 'GAME')
H.SetFont('Bold s12')
CreateImageButton(H, 0, IBBlack*)
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
CreateImageButton(H, 0, IBBlack*)
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
CreateImageButton(H, 0, IBBlack*)
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
CreateImageButton(H, 0, IBBlack*)
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
CreateImageButton(H, 0, IBBlack*)
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
CreateImageButton(H, 0, IBBlack*)
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
CreateImageButton(H, 0, IBBlack*)
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
CreateImageButton(H, 0, IBBlack*)
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
R.GetPos(, &Y)
U.GetPos(,, &WU)
U.Move(W - WU - 25, Y)
U.Redraw()
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
; Check for the updates
Check4Updates(Ctrl, Info) {
    Try {
        Hashsums := UpdatedPackagesHashs()
        UpdateList := [[], '']
        For Each, Packages in AllPackagesC {
            For Every, Package in Packages {
                If !FileExist(Package) {
                    Continue
                }
                If Hashsums.Has(Package) && Hashsums[Package] != HashFile(Package) {
                    UpdateList[1].Push(Package)
                    UpdateList[2] .= UpdateList[2] = '' ? '+ ' Package : '`n+ ' Package
                }
            }
        }
        If UpdateList[1].Length > 0 {
            If 'Yes' != MsgBox('Files to be updated!`n`n' UpdateList[2] '`n`nWant to update now?', 'Update list', 0x4 + 0x40) {
                Return
            }
            Prepare.Show()
            Prepare.OnEvent('Close', ExitScript, False)
            ProgressBar.Value := 0
            ProgressBar.Opt('Range1-' UpdateList[1].Length)
            ; Apply updates
            For Each, Package in UpdateList[1] {
                ProgressBar.Value += 1
                ProgressText.Text := 'Preparing [ ' Package ' ]'
                PackagePath := StrReplace(Package, '/', '\')
                PackageFolder := ''
                PackHead := StrGet(FileRead(PackagePath, 'RAW m2'), 2, 'CP0')
                If PackHead = '7z' {
                    SplitPath(PackagePath, &OutFileName, &OutDir)
                    PackageFolder := (OutDir ? OutDir '\' : '') StrSplit(OutFileName, '.')[1]
                }
                FileDelete(PackagePath)
                DownloadPackage(Package, PackagePath, PackageFolder)
                If PackHead = '7z' {
                    ExtractPackage(PackagePath, PackageFolder, True)
                }
            }
            Reload()
        }
    } Catch Error As Err {
        MsgBox("Update check failed!`n`n" Err.Message '`n' Err.Line '`n' Err.File, 'Fix', 0x10)
        Ctrl.Enabled := True
    }
}