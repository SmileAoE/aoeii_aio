;------------------------------------------------------;
; Unselects one unit on a group selection using        ;
; Alt + Right Mouse Button combination                 ;
; Visit https://www.autohotkey.com/docs/v2/Hotkeys.htm ;
; For more information                                 ;
;------------------------------------------------------;
#Requires AutoHotkey v2.0
#SingleInstance Force
GroupAdd('AOEII', 'ahk_exe empires2.exe')
GroupAdd('AOEII', 'ahk_exe age2_x1.exe')
GroupAdd('AOEII', 'ahk_exe age2_x2.exe')
#HotIf WinActive("ahk_group AOEII")
!RButton:: {
    WinGetPos(,, &W, &H, 'ahk_group AOEII')
    If W != A_ScreenWidth || H != A_ScreenHeight {
        Return
    }
    MouseClick('Right', , , , 0)
    MouseGetPos(&X, &Y)
    SendInput('{LCtrl Down}')
    MouseClick('Left', 315, A_ScreenHeight - 130, , 0)
    SendInput('{Ctrl Up}')
    MouseMove(X, Y, 0)
}
#HotIf