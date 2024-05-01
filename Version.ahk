#Include SharedLib.ahk
GameVersion := Map('aokCombine' , Map('2.0b', '2.0a')
                 , 'aocCombine' , Map('1.0e', '1.0c', '1.1', '1.0c', '1.5', '1.0c') )

AoEIIAIO.Title := 'GAME VERSION'
H := AoEIIAIO.AddText('cRed w150 Center h30', 'The Age of Kings')
H.SetFont('Bold s12')
H := AoEIIAIO.AddPicture('xp+59 yp+30', 'DB\000\aok.png')
AoEIIAIO.AddText('xp-59 yp+35 w1 h1')
GameVersion['aok'] := Map()
Loop Files, 'DB\002\aok\*', 'D' {
    H := AoEIIAIO.AddButton('w150', AOK := A_LoopFileName)
    H.SetFont('Bold')
    CreateImageButton(H, 0, [[0xFFFFFF,, 0xFF0000, 4, 0xCCCCCC, 2], [0xE6E6E6], [0xCCCCCC], [0xFFFFFF,, 0xCCCCCC]]*)
    H.OnEvent('Click', ApplyVersion)
    GameVersion['aok'][H] := 1
}
H := AoEIIAIO.AddText('cBlue ym w150 Center h30', 'The Conquerors')
H.SetFont('Bold s12')
H := AoEIIAIO.AddPicture('xp+59 yp+30', 'DB\000\aoc.png')
AoEIIAIO.AddText('xp-59 yp+35 w1 h1')
GameVersion['aoc'] := Map()
Loop Files, 'DB\002\aoc\*', 'D' {
    H := AoEIIAIO.AddButton('w150', AOC := A_LoopFileName)
    H.SetFont('Bold')
    CreateImageButton(H, 0, [[0xFFFFFF,, 0x0000FF, 4, 0xCCCCCC, 2], [0xE6E6E6], [0xCCCCCC], [0xFFFFFF,, 0xCCCCCC]]*)
    H.OnEvent('Click', ApplyVersion)
    GameVersion['aoc'][H] := 2
}
H := AoEIIAIO.AddText('cGreen ym w150 Center h30', 'Forgotten Empires')
H.SetFont('Bold s12')
H := AoEIIAIO.AddPicture('xp+59 yp+30', 'DB\000\fe.png')
AoEIIAIO.AddText('xp-59 yp+35 w1 h1')
GameVersion['fe'] := Map()
Loop Files, 'DB\002\fe\*', 'D' {
    H := AoEIIAIO.AddButton('w150', FE := A_LoopFileName)
    H.SetFont('Bold')
    CreateImageButton(H, 0, [[0xFFFFFF,, 0x008000, 4, 0xCCCCCC, 2], [0xE6E6E6], [0xCCCCCC], [0xFFFFFF,, 0xCCCCCC]]*)
    H.OnEvent('Click', ApplyVersion)
    GameVersion['fe'][H] := 3
}
AoEIIAIO.Show()
GameDirectory := IniRead(Config, 'Settings', 'GameDirectory', '')
If !ValidGameDirectory(GameDirectory) {
    For Each, Version in Features['Version'] {
        Version.Enabled := False
    }
    If 'Yes' = MsgBox('Game is not yet located!, want to select now?', 'Game', 0x4 + 0x40) {
        Run('Game.ahk')
    }
    ExitApp()
}
AnalyzeVersion()
; Finds which game to patch
FindGame(Ctrl) {
    FGame := ''
    If GameVersion['aok'].Has(Ctrl) {
        FGame := 'aok'
    }
    If GameVersion['aoc'].Has(Ctrl) {
        FGame := 'aoc'
    }
    If GameVersion['fe'].Has(Ctrl) {
        FGame := 'fe'
    }
    Return FGame
}
; Cleans up
CleansUp(FGame) {
    Loop Files, 'DB\002\' FGame '\*', 'D' {
        Version := A_LoopFileName
        Loop Files, 'DB\002\' FGame '\' Version '\*.*', 'R' {
            PathFile := StrReplace(A_LoopFileDir '\' A_LoopFileName, 'DB\002\' FGame '\' Version '\')
            If FileExist(GameDirectory '\' PathFile) {
                FileDelete(GameDirectory '\' PathFile)
            }
        }
    }
}
; Applys the selected version files
ApplyReqVersion(Ctrl, FGame) {
    If GameVersion.Has(FGame 'Combine') 
    && GameVersion[FGame 'Combine'].Has(Ctrl.Text) {
        If DirExist('DB\002\' FGame '\' GameVersion[FGame 'Combine'][Ctrl.Text]) {
            DirCopy('DB\002\' FGame '\' GameVersion[FGame 'Combine'][Ctrl.Text], GameDirectory, 1)
        }
    }
    If DirExist('DB\002\' FGame '\' Ctrl.Text) {
        DirCopy('DB\002\' FGame '\' Ctrl.Text, GameDirectory, 1)
    }
}
ApplyVersion(Ctrl, Info) {
    Try {
        ; Disables game buttons
        DisableOptions(FGame := FindGame(Ctrl))
        ; Closes the game if running
        CloseGame()
        ; Cleans up previous applied version
        CleansUp(FGame)
        ; Applys the selected version files
        ApplyReqVersion(Ctrl, FGame)
        ; Analyzes the game folder
        AnalyzeVersion()
        ; Finish
        SoundPlay('DB\000\30 Wololo.mp3')
    } Catch Error As Err {
        MsgBox("Version set failed!`n`n" Err.Message '`n' Err.Line '`n' Err.File, 'Version', 0x10)
    }
}
; Enables a versions list
EnableOptions(List) {
    For Item in List {
        Item.Enabled := True
    }
}
; Disables a versions list
DisableOptions(Game) {
    Switch Game {
        Case 'aok' : IB := [[0xFFFFFF,, 0xFF0000, 4, 0xCCCCCC, 2], [0xE6E6E6], [0xCCCCCC], [0xFFFFFF,, 0xCCCCCC]]
        Case 'aoc' : IB := [[0xFFFFFF,, 0x0000FF, 4, 0xCCCCCC, 2], [0xE6E6E6], [0xCCCCCC], [0xFFFFFF,, 0xCCCCCC]]
        Case 'fe' : IB := [[0xFFFFFF,, 0x008000, 4, 0xCCCCCC, 2], [0xE6E6E6], [0xCCCCCC], [0xFFFFFF,, 0xCCCCCC]]
    }
    For Item in GameVersion[Game] {
        Item.Enabled := False
        CreateImageButton(Item, 0, IB*)
    }
}
; Return a game version based on the available versions
AppliedVersionLookUp(Location) {
    MatchVersion := ''
    Loop Files, 'DB\002\' Location '\*', 'D' {
        Version := A_LoopFileName
        Match := True
        Loop Files, 'DB\002\' Location '\' Version '\*.*', 'R' {
            PathFile := StrReplace(A_LoopFileDir '\' A_LoopFileName, 'DB\002\' Location '\' Version '\')
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
            MatchVersion := Version
        }
    }
    If MatchVersion {
        For Control in GameVersion[Location] {
            If Control.Text = MatchVersion {
                Return [MatchVersion, Control]
            }
        }
    }
    Return ''
}
; Analyzes game versions
AnalyzeVersion() {
    If FileExist(GameDirectory '\empires2.exe') {
        Version := AppliedVersionLookUp('aok')
        If Type(Version) = 'Array' {
            For Game, Version in GameVersion['aok'] {
                Game.Enabled := True
            }
            CreateImageButton(Version[2], 0, [[0xFF5151,, 0xFFFFFF, 4, 0xFF5151, 2],,, [0xFFFFFF,, 0xCCCCCC,, 0xCCCCCC]]*)
        }
    }
    If FileExist(GameDirectory '\age2_x1\age2_x1.exe') {
        Version := AppliedVersionLookUp('aoc')
        If Type(Version) = 'Array' {
            For Game, Version in GameVersion['aoc'] {
                Game.Enabled := True
            }
            CreateImageButton(Version[2], 0, [[0x0080FF,, 0xFFFFFF, 4, 0x0080FF, 2],,, [0xFFFFFF,, 0xCCCCCC,, 0xCCCCCC]]*)
        }
    }
    If FileExist(GameDirectory '\age2_x1\age2_x2.exe') {
        Version := AppliedVersionLookUp('fe')
        If Type(Version) = 'Array' {
            For Game, Version in GameVersion['fe'] {
                Game.Enabled := True
            }
            CreateImageButton(Version[2], 0, [[0x00A800,, 0xFFFFFF, 4, 0x00A800, 2],,, [0xFFFFFF,, 0xCCCCCC,, 0xCCCCCC]]*)
        }
    }
}