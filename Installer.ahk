#Requires AutoHotkey v2
#SingleInstance Force
; Checks if the script run as admin
If !A_IsAdmin {
    MsgBox('Installer must run as administrator!', 'Warning', 0x30)
    ExitApp
}
Installer := Gui('-MinimizeBox', 'Setup')
Installer.OnEvent('Close', (*) => ExitApp())
Try {
    HIcon := LoadPicture('Shell32.dll', 'Icon123')
    Installer.AddPicture(, 'HBITMAP:*' HIcon)
}
Installer.SetFont('Bold s11', 'Calibri')
Installer.AddText('ym+7 cGreen', 'Age of Empires II Easy Manager Installer')
InstallBtn := Installer.AddButton('xm+100 w100', 'Install')
InstallBtn.OnEvent('Click', Install)
InstallPrg := Installer.AddProgress('xm -Smooth w300 h17')
Installer.Show()
Install(Ctrl, Info) {
    Try {
        InstallBtn.Text := 'Installing...'
        InstallBtn.Enabled := False
        AppDir := A_ProgramFiles '\AoE II AIO'
        If !DirExist(AppDir) {
            DirCreate(AppDir)
        }
        InstallPrg.Value := 10
        Download('https://raw.githubusercontent.com/SmileAoE/aoeii_aio/main/AoE II Manager AIO Ex.ahk', AppDir '\AoE II Manager AIO Ex.ahk')
        InstallPrg.Value := 50
        Download('https://raw.githubusercontent.com/SmileAoE/aoeii_aio/main/SharedLib.ahk', AppDir '\SharedLib.ahk')
        InstallPrg.Value := 90
        Download('https://raw.githubusercontent.com/SmileAoE/aoeii_aio/main/Uninstall.ahk', AppDir '\Uninstall.ahk')
        InstallPrg.Value := 100
        Sleep(1000)
        InstallBtn.Text := 'Installed'
        If 'Yes' = MsgBox('Installation complete!`n`nWant to launch the app now?', 'Setup', 0x40 + 0x4) {
            Run(AppDir '\AoE II Manager AIO Ex.ahk', AppDir)
        }
        ExitApp()
    } Catch Error As Err {
        MsgBox("Installation failed!`n`n" Err.Message '`n' Err.Line '`n' Err.File, 'Version', 0x10)
    }
}