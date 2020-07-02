; Jonathan Hopper
; 6/26/2020
; Excel Find Highlighter

; CAUTION: THE FIRST TIME YOU USE THIS ON A SHEET, IT WILL WIPE YOUR UNDO STACK
; Notes:
;	If you want to change the highlight color, I would recommend opening the VBA Developer console,
;	pressing ctr+g to open the immediate window, and type '?RGB(x,y,z)' where x,y,z are RGB values.
;	This will output the VBA color code format needed for this script.

; Known Issues:
;	1. The highlight will fall behind if you quickly go through the search.
;	  1a. If you press enter quick enough, sometimes the cursour can end up ahead of the highlight.
;	2. In a new sheet, using the find dialog will destroy your undo stack. This is an issue in Excel.
;		The same thing will happen if you program a macro within excel to do the same thing.
;		This will also happen if you make a new sheet in a workbook where the formatting already exist.
;	3. If you save and exit while the 'Find and Replace' dialog is open, the highlight will remain when you
;		open the document again.  The highlight can then be removed by opening and closing the 'Find and Replace'
;		dialog box.
;	4. This will only work if you're able to edit the worksheet.  You can be in read-only mode, though.

#Include Acc.ahk
#SingleInstance Force

global sheetArray := Array()
; I wanted to add a global variable for the cell that stores the formatting,
; but I honestly could not get it to work in variable form.

; Excel_Get() There are a lot of versions of this function floating around, but I ended up using
; this one, since some of the others were ending in endless loops when the excel workbook was saved and closed
; when the find dialog box was still open.  This one doesn't do that.
Excel_Get(WinTitle:="ahk_class XLMAIN", Excel7#:=1) {
    static h := DllCall("LoadLibrary", "Str", "oleacc", "Ptr")
    WinGetClass, WinClass, %WinTitle%
    if !(WinClass == "XLMAIN")
        return "Window class mismatch."
    ControlGet, hwnd, hwnd,, Excel7%Excel7#%, %WinTitle%
    if (ErrorLevel)
        return "Error accessing the control hWnd."
    VarSetCapacity(IID_IDispatch, 16)
    NumPut(0x46000000000000C0, NumPut(0x0000000000020400, IID_IDispatch, "Int64"), "Int64")
    if DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", hWnd, "UInt", -16, "Ptr", &IID_IDispatch, "Ptr*", pacc) != 0
        return "Error calling AccessibleObjectFromWindow."
    window := ComObject(9, pacc, 1)
    if ComObjType(window) != 9
        return "Error wrapping the window object."
    Loop
        try return window.Application
        catch e
            if SubStr(e.message, 1, 10) = "0x80010001"
                ControlSend, Excel7%Excel7#%, {Esc}, %WinTitle%
            else
                return "Error accessing the application object."
}

FindFormat_ApplyFormat(xl) {
	; applying new range to the conditional format of the active sheet
	c := xl.ActiveCell.address
	newRange := "$Z$999" . "," . c
    xl.ActiveSheet.Range("$Z$999").FormatConditions(1).ModifyAppliesToRange(xl.Range(newRange))
	
	; crude way of checking if a value is in the sheetArray.
	x := -1
	for index, element in sheetArray
		if (element == xl.ActiveSheet.Index)
			x := 1
		
	if (x == -1)
		sheetArray.Push(xl.ActiveSheet.Index)
	return
}

FindFormat_CreateFormat(xl) {
	; this does take a couple seconds for large sheets (>~100)
	; applying the conditional format to all sheets in the file.
	; THIS PROCESS WILL DELETE YOUR CURRENT UNDO STACK.
	sheetCount := xl.Sheets.Count
	loop % sheetCount {
		c := xl.Sheets(a_index).Range("$Z$999")
		c.FormatConditions.Add(8)
		c.FormatConditions(1).Interior.Color := 6750207
		c.FormatConditions(1).Priority := 1
	}
	return
}

FindFormat_RemoveFormat(xl) {
	; resetting the conditional format on sheets that have been affected.
	for index, element in sheetArray {
		xl.Sheets(element).Range("$Z$999").FormatConditions(1).ModifyAppliesToRange(xl.Range("$Z$999"))
		sheetArray.Delete(index)
	}
}

; Enter keys
~Enter::
~NumpadEnter::
	; ensuring the user is in Microsoft Excel
	WinGet, ProcessTitle, ProcessName, A
	if !(InStr(ProcessTitle, "EXCEL.EXE") > 0)
		return
	
	; ensuring the user has the 'Find and Replace' dialog open
	WinGetTitle, Title, A
	if (InStr(Title, "Find and Replace") > 0) {
		xl := Excel_Get()
		
		; creating formatting if it doesn't exist
		if (xl.ActiveSheet.Range("$Z$999").FormatConditions.Count <= 0)
			FindFormat_CreateFormat(xl)
		
		; the active cell does not switch until the enter key has been released
		KeyWait, Enter
		; need to use 'try' here, in the event the find fails to find anything
		try
			FindFormat_ApplyFormat(xl)
	}
return

; Left mouse button.  This one is a bit different.
~LButton::
	; ensuring the user is in Microsoft Excel
	WinGet, ProcessTitle, ProcessName, A
	if !(InStr(ProcessTitle, "EXCEL.EXE") > 0)
		return
	
	; ensuring the user has the 'Find and Replace' dialog open
	WinGetTitle, Title, A
	if (InStr(Title, "Find and Replace") > 0) {
		xl := Excel_Get()
		; creating formatting if it doesn't exist
		if (xl.ActiveSheet.Range("$Z$999").FormatConditions.Count <= 0)
			FindFormat_CreateFormat(xl)
		
		; getting the active cell before and after the click event
		; this may not be the only way to do this, but it's all I can
		; think of at the moment.
		beforePress := xl.ActiveCell.address
		KeyWait, LButton
		afterPress := xl.ActiveCell.address						

		; else, if the before and after cells are different,
		; then the user must have clicked 'Find Next'.
		if !("" . beforePress == "" . afterPress)
			try
				FindFormat_ApplyFormat(xl)

		; if user closes the dialog box
		if !(WinExist("Find and Replace"))
			FindFormat_RemoveFormat(xl)
	}
	
	; if user closes the dialog box
	; note this section cannot be inside the above if statement,
	;   since you can close the dialog box when you have the sheet/other
	;   windows active.
	else if (WinExist("Find and Replace")) {
		stdout := FileOpen("*", "w")
		stdout.WriteLine("test")
		KeyWait, LButton
		sleep, 50
		; need to use Excel_Get here by itself; I was having issues otherwise.
		if !(WinExist("Find and Replace"))
			FindFormat_RemoveFormat(Excel_Get())
	}
return

; Escape key
~Escape::
~!F4::
	; ensuring the user is in Microsoft Excel
	if !(InStr(ProcessTitle, "EXCEL.EXE") > 0)
		return
	
	; ensuring the user has the 'Find and Replace' dialog open
	WinGetTitle, Title, A
	if (InStr(Title, "Find and Replace") > 0)
		FindFormat_RemoveFormat(xl)
return

;F15::
;	z := WinExist("Find and Replace")
;	MsgBox, %z%
