#Requires AutoHotkey >v2.0 

#include D:\SoftX\AHK_Scripts_V2\Lib\WebView2.ahk

; WebView2.ahk -- https://github.com/thqby/ahk2_lib/blob/master/WebView2/WebView2.ahk
;  ComVar.ahk -- https://github.com/thqby/ahk2_lib/blob/master/ComVar.ahk


Persistent
main := Gui() 
main.Title := "WebPagePreviewer"
main.SetFont("S14")
main.OnEvent("Close", _exit_)

ctl:=main.Add("GroupBox", "y+10 w800 h800", "ThisGroup")
main.Show()
ctl.wvc := wvc := WebView2.create(ctl.Hwnd)
wv := wvc.CoreWebView2
ctl.nwr := wv.NewWindowRequested(NewWindowRequestedHandler)
RegisterMsgReceiver()
return 
RegisterMsgReceiver()
{    ;https://wincmd.ru/forum/viewtopic.php?p=110848&sid=0dfde01723b39e508e96d62c00a7a9b5
    OnMessage(0x4a, Receive_WM_COPYDATA)  ; 0x4a is WM_COPYDATA          
}

Receive_WM_COPYDATA(wParam, lParam, msg, hwnd)
{  
    PtrPos:=NumGet(lParam + A_PtrSize * 2,0,"Ptr")
    retVal:=StrGet(PtrPos)
    if(retVal)
        MHTML_Preview(retVal) 
}

MHTML_Preview(strCMD)
{
    
    FilePath:=Trim(strCMD)        
    if(FileExist(FilePath))
    {
        wv.Navigate("file:///" . FilePath)
    }
    else
        MsgBox "File does NOT exist:`n" . FilePath
}
_exit_(*) {
	ExitApp()
}

NewWindowRequestedHandler(handler, wv2, arg) {
	argp := WebView2.NewWindowRequestedEventArgs(arg)
	deferral := argp.GetDeferral()
	main.GetClientPos(,,&w,&h)
	ctl.wvc := WebView2.create(ctl.Hwnd, ControllerCompleted_Invoke, WebView2.Core(wv2).Environment)
	return 0
	ControllerCompleted_Invoke(wvc) {
		argp.NewWindow := wv := wvc.CoreWebView2
		ctl.nwr := wv.NewWindowRequested(NewWindowRequestedHandler)
		deferral.Complete()
	}
}
