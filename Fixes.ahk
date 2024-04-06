#Include SharedLib.ahk
GameFix := Map('FIX'        , ['Fix v1', 'Fix v2']
             , 'FIXHandle'  , Map())
Features['Fixes'] := []
AoEIIAIO.Title := 'GAME FIXES'
H := AoEIIAIO.AddText('w350 Center h25', 'Select one of the fixes below')
H.SetFont('Bold s12')
For Each, FIX in GameFix['FIX'] {
    H := AoEIIAIO.AddButton('w350', FIX)
    H.SetFont('Bold')
    CreateImageButton(H, 0, [[0xFFFFFF,,, 4, 0xCCCCCC, 2], [0xE6E6E6], [0xCCCCCC], [0xFFFFFF,, 0xCCCCCC]]*)
    Features['Fixes'].Push(H)
    H.OnEvent('Click', ApplyFix)
    GameFix['FIXHandle'][FIX] := H
}
H := AoEIIAIO.AddLink('w350', 'Help Links:`n<a href="https://aok.heavengames.com/blacksmith/showfile.php?fileid=13275">Aoe II Wide Screen all version</a>'
                                 . '`n<a href="https://aok.heavengames.com/blacksmith/showfile.php?fileid=13710">Age of Empire II the Age of king version 2.0c patch into 2.0</a>'
                                 . '`n<a href="https://aok.heavengames.com/blacksmith/showfile.php?fileid=13730">Ao2 patch:1.0 ,1.0c,2.0,2.0a,2.0c Widescreen + windowed</a>'
                                 . '`n<a href="https://aok.heavengames.com/blacksmith/showfile.php?fileid=13673">Aok 2.0 Generate Record To Ignore Player Who Leave</a>')
H.SetFont('Bold')
AoEIIAIO.Show()
GameDirectory := IniRead(Config, 'Settings', 'GameDirectory', '')
If !ValidGameDirectory(GameDirectory) {
    For Each, Fix in Features['Fixes'] {
        Fix.Enabled := False
    }
    If 'Yes' = MsgBox('Game is not yet located!, want to select now?', 'Game', 0x4 + 0x40) {
        Run('Game.ahk')
    }
    ExitApp()
}
AnalyzeFix()
ApplyFix(Ctrl, Info) {
    Try {
        CloseGame()
        DefaultPB(Features['Fixes'])
        EnableControls(Features['Fixes'], 0)
        DirCopy('DB\001\' Ctrl.Text, GameDirectory, 1)
        If InStr(Ctrl.Text, 'v2') {
            RegWrite(2, 'REG_DWORD', 'HKEY_CURRENT_USER\SOFTWARE\Microsoft\Microsoft Games\Age of Empires', 'Aoe2Patch')
        }
        EnableControls(Features['Fixes'])
        AnalyzeFix()
        SoundPlay('DB\000\30 Wololo.mp3')
    } Catch Error As Err {
        MsgBox("Patch failed!`n`n" Err.Message '`n' Err.Line '`n' Err.File, 'Fix', 0x10)
    }
}
AnalyzeFix() {
    MatchFix := ''
    Loop Files, 'DB\001\*', 'D' {
        Fix := A_LoopFileName
        Match := True
        Loop Files, 'DB\001\' Fix '\*.*', 'R' {
            PathFile := StrReplace(A_LoopFileDir '\' A_LoopFileName, 'DB\001\' Fix '\')
            If !FileExist(GameDirectory '\' PathFile) && Match {
                Match := False
                Break
            }
            CurrentHash := HashFile(A_LoopFileFullPath)
            FoundHash := HashFile(GameDirectory '\' PathFile)
            If (CurrentHash != FoundHash) && Match {
                Match := False
                Break
            }
        }
        If Match {
            MatchFix := Fix
        }
    }
    If MatchFix {
        CreateImageButton(GameFix['FIXHandle'][MatchFix], 0, [[0x008000,, 0xFFFFFF, 4, 0x008000, 2],,, [0xFFFFFF,, 0xCCCCCC]]*)
        GameFix['FIXHandle'][MatchFix].Redraw()
    }
}