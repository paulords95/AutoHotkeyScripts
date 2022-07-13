FileList := ""
Loop, Files, C:\BPA BACKUP\*.*, FD
    FileList .= A_LoopFileName "`n"

Sort, FileList, R

Loop Parse, FileList, "`n"
{
    if (A_LoopField = "")
        continue

    if (A_Index >= 4)
        FileRemoveDir, C:\BPA BACKUP\%A_LoopField%, 1
        
FileRemoveDir, C:\Download Temp, 1
}

FileCreateDir, C:\BPA BACKUP\%A_YYYY%%A_MM%
FileCopyDir, C:\BPA, C:\BPA BACKUP\%A_YYYY%%A_MM%, 1

Run, bpamag.exe
WinWait, ahk_class TFrmSplash
ControlSend, TEdit2, mestre
ControlSend, TEdit1, a
ControlSend, TEdit1, {enter}

CoordMode, Mouse
Click 345, 38
Click 429, 106
Sleep, 150
hwnd := WinExist("ahk_class TFALT_CMP")

CurrentDate := A_Now
    CurrentDate += -1, D
    FormatTime, TimeString, %CurrentDate%, M/d/yyyy
    StringTrimRight, month, TimeString, 8

Control, Choose,%month%, TComboBox1, ahk_id %hwnd%
Sleep, 150
SendInput {tab}
Sleep, 150
SendInput {enter}

hwnd = false
if WinExist("ahk_class TMessageForm") {
    SendInput {enter}
    hwnd = true
}


if (hwnd != false) {
    SendInput {tab}
    SendInput {enter}
    Sleep, 150
    SendInput {enter}
    SendInput {enter}
    SendInput {tab}
    Sleep, 150
    SendInput {tab}
    SendInput {enter}

}

Sleep, 150
Click 345, 38
Click 378, 57
Sleep, 150
ControlClick, x304  y287, Importação dos arquivos do KIT BDSIA
Sleep, 150
SendInput {enter}

WinWait, ahk_class TMessageForm

Sleep, 200
SendInput {enter}
SendInput {tab}

SendInput {enter}

Sleep, 1000

Process,Close,bpamag.exe
Sleep, 150

Run, bpamag.exe
WinWait, ahk_class TFrmSplash
ControlSend, TEdit2, mestre
ControlSend, TEdit1, a
ControlSend, TEdit1, {enter}

Sleep, 300
SendInput {enter}
SendInput {enter}
Sleep , 300
Click 345, 38
Click 380, 122
Sleep , 300
SendInput {enter}
Sleep , 300
SendInput {enter}
SendInput {tab}
SendInput {enter}

Process,Close,bpamag.exe
Process,Close,bpamag.exe
Process,Close,bpamag.exe
Process,Close,bpamag.exe
Process,Close,bpamag.exe

FileAppend, 1, C:\BPA\atualizado.txt

MsgBox "BPA Atualizado com sucesso"
ExitApp

^b::Process,Close,bpamag.exe
