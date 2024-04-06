#Include SharedLib.ahk
GameVersion := Map('AOK'        , ['2.0 CD', '2.0a No CD', '2.0b CD']
                 , 'AOKCombine' , Map('2.0b CD', ['2.0a No CD'])
                 , 'AOKHandle'  , Map()
                 , 'AOC'        , ['1.0 CD', '1.0c No CD', '1.0e No CD', '1.1 No CD', '1.5 CD']
                 , 'AOCCombine' , Map('1.0e No CD', ['1.0c No CD'], '1.1 No CD', ['1.0c No CD'], '1.5 CD', ['1.0c No CD'])
                 , 'AOCHandle'  , Map()
                 , 'FE'         , ['2.2 CD']
                 , 'FECombine'  , Map()
                 , 'FEHandle'   , Map())
AoEIIAIO.Title := 'GAME VERSION'
Features['Version'] := []
H := AoEIIAIO.AddText('cRed w150 Center h30', 'The Age of Kings')
H.SetFont('Bold s12')
H := AoEIIAIO.AddPicture('xp+59 yp+30', 'DB\000\aok.png')
AoEIIAIO.AddText('xp-59 yp+35 w1 h1')
For Each, AOK in GameVersion['AOK'] {
    H := AoEIIAIO.AddButton('w150', AOK)
    H.SetFont('Bold')
    CreateImageButton(H, 0, [[0xFFFFFF,, 0xFF0000, 4, 0xCCCCCC, 2], [0xE6E6E6], [0xCCCCCC], [0xFFFFFF,, 0xCCCCCC]]*)
    Features['Version'].Push(H)
    H.OnEvent('Click', ApplyVersion)
    GameVersion['AOKHandle'][AOK] := H
}
H := AoEIIAIO.AddText('cBlue ym w150 Center h30', 'The Conquerors')
H.SetFont('Bold s12')
H := AoEIIAIO.AddPicture('xp+59 yp+30', 'DB\000\aoc.png')
AoEIIAIO.AddText('xp-59 yp+35 w1 h1')
For Each, AOC in GameVersion['AOC'] {
    H := AoEIIAIO.AddButton('w150', AOC)
    H.SetFont('Bold')
    CreateImageButton(H, 0, [[0xFFFFFF,, 0x0000FF, 4, 0xCCCCCC, 2], [0xE6E6E6], [0xCCCCCC], [0xFFFFFF,, 0xCCCCCC]]*)
    Features['Version'].Push(H)
    H.OnEvent('Click', ApplyVersion)
    GameVersion['AOCHandle'][AOC] := H
}
H := AoEIIAIO.AddText('cGreen ym w150 Center h30', 'Forgotten Empires')
H.SetFont('Bold s12')
H := AoEIIAIO.AddPicture('xp+59 yp+30', 'DB\000\fe.png')
AoEIIAIO.AddText('xp-59 yp+35 w1 h1')
For Each, FE in GameVersion['FE'] {
    H := AoEIIAIO.AddButton('w150', FE)
    H.SetFont('Bold')
    CreateImageButton(H, 0, [[0xFFFFFF,, 0x008000, 4, 0xCCCCCC, 2], [0xE6E6E6], [0xCCCCCC], [0xFFFFFF,, 0xCCCCCC]]*)
    Features['Version'].Push(H)
    ;H.OnEvent('Click', ApplyVersion)
    GameVersion['FEHandle'][FE] := H
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
; Cleans up the game folder
CleanUp(TargetVersion) {
    ; Cleans up previous versions files
    Loop Files, 'DB\002\*', 'D' {
        If TargetVersion != SubStr(Version := A_LoopFileName, 1, 1) {
            Continue
        }
        Loop Files, 'DB\002\' Version '\*.*', 'R' {
            PathFile := StrReplace(A_LoopFileDir '\' A_LoopFileName, 'DB\002\' Version '\')
            If FileExist(GameDirectory '\' PathFile) {
                FileDelete(GameDirectory '\' PathFile)
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
                If FileExist(GameDirectory '\' PathFile) {
                    FileDelete(GameDirectory '\' PathFile)
                }
            }
        }
    }
}
; Sets a version
ApplyVersion(Ctrl, Info) {
    Try {
        TargetVersion := SubStr(Ctrl.Text, 1, 1)
        Key2 := TargetVersion = '1' ? 'AOCHandle' : 'AOKHandle'
        DefaultPB(GameVersion[Key2])
        EnableControls(GameVersion[Key2], 0)
        CloseGame()
        CleanUp(TargetVersion)
        ; Copy the selected version files
        Key1 := TargetVersion = '1' ? 'AOCCombine' : 'AOKCombine'
        If GameVersion[Key1].Has(Ctrl.Text) {
            For Each, Version in GameVersion[Key1][Ctrl.Text] {
                If DirExist('DB\002\' Version) {
                    DirCopy('DB\002\' Version, GameDirectory, 1)
                }
            }
        }
        If DirExist('DB\002\' Ctrl.Text) {
            DirCopy('DB\002\' Ctrl.Text, GameDirectory, 1)
        }
        AnalyzeVersion()
        EnableControls(GameVersion[Key2])
        SoundPlay('DB\000\30 Wololo.mp3')
    } Catch Error As Err {
        MsgBox("Patch failed!`n`n" Err.Message '`n' Err.Line '`n' Err.File, 'Version', 0x10)
    }
}
; Analyzes game versions
AnalyzeVersion() {
    If FileExist(GameDirectory '\empires2.exe') {
        MatchVersion := ''
        Loop Files, 'DB\002\2.*', 'D' {
            Version := A_LoopFileName
            Match := True
            Loop Files, 'DB\002\' Version '\*.*', 'R' {
                PathFile := StrReplace(A_LoopFileDir '\' A_LoopFileName, 'DB\002\' Version '\')
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
            CreateImageButton(GameVersion['AOKHandle'][MatchVersion], 0, [[0xFF3434,, 0xFFFFFF, 4, 0xFF3434, 2],,, [0xFFFFFF,, 0xCCCCCC]]*)
            GameVersion['AOKHandle'][MatchVersion].Redraw()
        }
    } Else {
        For Each, Version in GameVersion['AOKHandle'] {
            Version.Enabled := False
        }
    }
    If FileExist(GameDirectory '\age2_x1\age2_x1.exe') {
        MatchVersion := ''
        Loop Files, 'DB\002\1.*', 'D' {
            Version := A_LoopFileName
            Match := True
            Loop Files, 'DB\002\' Version '\*.*', 'R' {
                PathFile := StrReplace(A_LoopFileDir '\' A_LoopFileName, 'DB\002\' Version '\')
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
            CreateImageButton(GameVersion['AOCHandle'][MatchVersion], 0, [[0x0080FF,, 0xFFFFFF, 4, 0x0080FF, 2],,, [0xFFFFFF,, 0xCCCCCC]]*)
            GameVersion['AOCHandle'][MatchVersion].Redraw()
        }
    } Else {
        For Each, Version in GameVersion['AOCHandle'] {
            Version.Enabled := False
        }
    }
    If FileExist(GameDirectory '\age2_x1\age2_x2.exe') {
        If 'fe3ac4feabf17a91134959b12866b7f6' = HashFile(GameDirectory '\age2_x1\age2_x2.exe') {
            CreateImageButton(GameVersion['FEHandle']['2.2 CD'], 0, [[0x00AC00,, 0xFFFFFF, 4, 0x00AC00, 2],,, [0xFFFFFF,, 0xCCCCCC]]*)
            GameVersion['FEHandle']['2.2 CD'].Redraw()
        }
    } Else {
        For Each, Version in GameVersion['FEHandle'] {
            Version.Enabled := False
        }
    }
}