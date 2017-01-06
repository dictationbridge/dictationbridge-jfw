Globals
	string lastAnnouncedOptions,
	int scheduleID ; Used below to schedule announceCorrectionOptions.

function announceCorrectionOptions()
	; If the spelling dialog is open, then the options are in the only available listbox.
	; This has class ListBox, according to Jaws.
	var
		string newOptions,
		handle optionsListBox
	optionsListBox = FindTopLevelWindow("ClsSpellDialog", "")
optionsListBox = FindWindow(optionsListBox, "ListBox")
	newOptions = "Options are " + GetWindowText(optionsListBox, 0)
	; We need to clean this up.
newOptions = StringReplaceSubstrings(newOptions, "\"Choose ", ", ")
newOptions = StringReplaceSubstrings(newOptions, "\"", " ")
	if newOptions != lastAnnouncedOptions then
		Say(newOptions, 0)
		lastAnnouncedOptions = newOptions
	endIf
	scheduleID = ScheduleFunction("announceCorrectionOptions", 2)
endfunction

function AutoStartEvent()
	lastAnnouncedOptions = ""
	scheduleID = ScheduleFunction("announceCorrectionOptions", 2)
endFunction

function AutoFinishEvent()
	if scheduleID != 0
	UnscheduleFunction(scheduleID)
	endIf
endFunction