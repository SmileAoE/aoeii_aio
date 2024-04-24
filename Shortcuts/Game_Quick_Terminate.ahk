;------------------------------------------------------;
; Exits the game using Win + q combination             ;
; Visit https://www.autohotkey.com/docs/v2/Hotkeys.htm ;
; For more information                                 ;
;------------------------------------------------------;
#Requires AutoHotkey v2.0
#SingleInstance Force
GroupAdd('AOEII', 'ahk_exe empires2.exe')
GroupAdd('AOEII', 'ahk_exe age2_x1.exe')
GroupAdd('AOEII', 'ahk_exe age2_x2.exe')
#HotIf WinActive("ahk_group AOEII")
#q:: {
    For Each, App in ['empires2.exe'
                    , 'age2_x1.exe'
                    , 'age2_x2.exe'] {
        If ProcessExist(App) {
            ProcessClose(App)
        }
    }
}
#HotIf