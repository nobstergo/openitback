#Persistent
#SingleInstance Force
#NoEnv
SetBatchLines -1

global closedHistory := []
global previousWindows := {}

SetTimer, TrackExplorer, 1000

^+t::  ; Ctrl+Shift+T
    if (WinActive("ahk_class CabinetWClass") or WinActive("ahk_class ExploreWClass")) {
        if (closedHistory.Length() > 0) {
            folder := closedHistory.Pop()
            shell := ComObjCreate("Shell.Application")
            shell.Open(folder)
        } else {
            TrayTip, Explorer Reopen, No recently closed folders left.
        }
    }
return



TrackExplorer:
    windows := ComObjCreate("Shell.Application").Windows
    current := {}

    Loop % windows.Count
    {
        try {
            window := windows.Item(A_Index - 1)
            path := window.Document.Folder.Self.Path
            current[path] := true
        }
    }

    ; Compare to previous windows to find closed ones
    for path, _ in previousWindows {
        if !current.HasKey(path) {
            closedHistory.Push(path)
        }
    }

    previousWindows := current
return
