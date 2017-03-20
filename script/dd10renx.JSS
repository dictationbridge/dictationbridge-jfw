include "hjglobal.jsh"
include "hjconst.jsh"
include "msaaconst.jsh"
Globals
int giJustEnteredCorrectionBox, ;used to indicate when we have just entered the correction box as aposed to alt-tabbing away.
int giDragonLoading,
int giCorrectionBoxAlreadyActive ;used to indicate if the correction box is already active when autostart/autofixinish fire.

int Function HandleCustomAppWindows (handle hWnd)
if StringCompare (GetWindowName (hwnd), "Loading Dragon NaturallySpeaking", false) ==0
;We are loading dragon.
giDragonLoading =1
endif
if StringCompare ("ClsSpellDialog", GetWindowClass (hWnd), false) ==0 && giJustEnteredCorrectionBox==0 ;only run this code if we have entered the correction box and the correction box has just been opened.
if giJustEnteredCorrectionBox ==0 
SayUsingVoice (vctx_message, "Correction box opened", ot_dialog_name)
giDragonLoading =0
giJustEnteredCorrectionBox  =1
giCorrectionBoxAlreadyActive =1
return true
endif
else 
return HandleCustomAppWindows (hWnd)
endif
EndFunction

void function ProcessSayFocusWindowOnFocusChange(string RealWindowName, handle FocusWindow)
if giJustEnteredCorrectionBox ==1
return
else
return ProcessSayFocusWindowOnFocusChange(RealWindowName, FocusWindow)
Endif
EndFunction

Void Function FocusChangedEvent (handle FocusWindow, handle PrevWindow)
if giJustEnteredCorrectionBox  ==1
let giJustEnteredCorrectionBox  =0
endif
return FocusChangedEvent (FocusWindow, PrevWindow)
EndFunction

Void Function SayHighLightedText (handle hwnd, string buffer)
if (giJustEnteredCorrectionBox ==1)
return
else
return SayHighLightedText (hwnd, buffer)
endif
EndFunction

function AutoFinishEvent()
var
handle correctionBoxWindow =FindTopLevelWindow ("ClsSpellDialog", "")
if !IsWindowVisible (correctionBoxWindow) && giDragonLoading ==0
SayUsingVoice (vctx_message, "Correction box closed.", ot_string)
giCorrectionBoxAlreadyActive =0
giJustEnteredCorrectionBox =0
Endif
endfunction