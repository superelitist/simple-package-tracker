;================================================================================
;   DIRECTIVES, ETC.
;================================================================================
#NoEnv ; Recommended for performance and compatibility 
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Persistent ; Keeps a script running (until closed).
#SingleInstance FORCE ; automatically replaces an old version of the script
SetBatchLines, -1
FileEncoding, UTF-8

;================================================================================
;   GLOBALS
;================================================================================
Global WM_CLOSE := 0x10
Global WM_SYSCOMMAND := 0x112
Global WM_SETICON := 0x80
Global SC_CLOSE := 0xF060
Global tracker := new SimplePackageTracker()
Global add_tracking_number ; this basically ruins the whole point of making an instantiable class in the first place, but I didn't see an alternative...
Global add_description

;================================================================================
;   INIT
;================================================================================
Loop, read, SimplePackageTracker.cfg ; read in the configuration file line-by-line
{
	If (SubStr(A_LoopReadLine, 1, 1) = ";") ; ignore comments (although currently they aren't parsed, and disappear when the file is saved.)
		continue
	desc := ""
	Loop, Parse, A_LoopReadLine, CSV
	{
		If (A_Index = 1) { ; the first element is the tracking number
			tn := A_LoopField
		} Else { ; everything else must be the description. we could later add functionality with more fields...
			desc := desc . A_LoopField ; concatenate each subsequent field
		}
	}
	tracker.AddTrackingNumber(tn, desc)
}

;================================================================================
;   MAIN
;================================================================================
tracker.Start()
return

;================================================================================
;   CLASSES
;================================================================================
class SimplePackageTracker {

	url_base := "https://www.packagetrackr.com/track/ups-packages/"
	re := "O)<h4 class=""media-heading t-info-status status-font print-text-left print-padding-0 status-ibc-\d+"">([A-Za-z\s]+)</h4>"
	packages := []
	hours_to_refresh := [7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17] ; 7-17 checks on the hour from 7am to 5pm
	last_hour_refreshed := -1
	save_counter := 0
	
	__New() {
		OnMessage(WM_SYSCOMMAND, ObjBindMethod(this, "OnSysCommand")) ; register for a message when user clicks 'X'
	}

	DrawGui() {
		Gui pts_gui: Destroy
		Gui pts_gui: +LastFound +E0x20 +AlwaysOnTop +Hwndhwnd_gui
		this.hwnd_gui := hwnd_gui
		Gui pts_gui: Margin, 8, 8
		Gui pts_gui: Font, s14 Bold, Consolas
		Gui pts_gui: Add, Text, ym w192, % "Tracking Number"
		Gui pts_gui: Add, Text, ym w256, % "Description"
		Gui pts_gui: Add, Text, ym w144, % "Last Checked"
		Gui pts_gui: Add, Text, ym w160, % "Status"
		Gui pts_gui: Font, s10 Normal
		for index, element in this.packages {
			Gui pts_gui: Add, Edit, ReadOnly w192 xm section, % element.tracking_number
			Gui pts_gui: Add, Edit, ReadOnly w256 ys, % element.description
			Gui pts_gui: Add, Edit, ReadOnly w144 ys hwndlast, % element.last
			; this is awkward, but if we want to be able to reference this line in the GUI later, we have to add the handle to our package objects.
			this.packages[index].hwnd_last := last
			Gui pts_gui: Add, Edit, ReadOnly w160 ys hwndstatus, % element.status_string
			If (element.status_string == "Delivered") {
				Gui pts_gui: Font, cGreen Bold ; this simply sets the desired style
				GuiControl, Font, % status ; then we have to actually set the content to the previously selected style
				Gui pts_gui: Font, cBlack Normal
			} Else {
				Gui pts_gui: Font, cBlue Bold
				GuiControl, Font, % status
				Gui pts_gui: Font, cBlack Normal
			}
			this.packages[index].hwnd_status := status
			Gui pts_gui: Add, Button, w40 ys hwndremove, -
			temp_func_ref := ObjBindMethod(this, "ItemRemove") ; stupid black magic to call a method from a button
			GuiControl, +g, %remove%, %temp_func_ref%
			this.packages[index].hwnd_remove := remove
		}
		; add an editable line. because we need to be able to access the contents of the two edit boxes, and because we can't use class properties for control variables, we need to define globals. ugh.

		Gui pts_gui: Add, Edit, w192 xm section vadd_tracking_number
		Gui pts_gui: Add, Edit, w256 ys vadd_description
		Gui pts_gui: Add, Edit, ReadOnly w144 ys hwndlast, % "..."
		Gui pts_gui: Add, Edit, ReadOnly w160 ys hwndstatus, % "..."
		Gui pts_gui: Add, Button, w40 ys hwndadd, +
		temp_func_ref := ObjBindMethod(this, "ItemAdd") ; stupid black magic to call a method from a button
		GuiControl, +g, %add%, %temp_func_ref%
		this.hwnd_add := add ; store the handle for our add button so we can disable it later
		; Gui pts_gui: Show, % "Center" . " NoActivate" ; display the GUI
		Gui pts_gui: Show, % "NoActivate" ; display the GUI
	}

	AddTrackingNumber(tn, desc) {
		item := {}
		item.tracking_number := tn
		item.description := desc
		item.delivered := False
		this.packages.Push(item)
	}

	Start() {
		this.DrawGui()
		this.OnTick()
		SetTimer(ObjBindMethod(this, "OnTick"), (Random(180, 900) * 1000)) ; we remember to check anywhere between every three and fifteen minutes. some very basic pattern obfuscation.
	}

	RefreshDeliveryStatus(tn) {
		Try {
			browser := ComObjCreate("WinHttp.WinHttpRequest.5.1")
			browser.Open("GET", this.url_base . StrReplace(tn, A_Space), true)
			browser.Send()
			browser.WaitForResponse()
			re_err := RegExMatch(browser.ResponseText, this.re, response_string)
			status_string := response_string.Value(1)
		} Catch e {
			status_string := "Message: " . e.Message . "`n" . "What: " . e.What . "`n" . "Extra: " . e.Extra . "`n" . "File:Line: " . e.File . " : " . e.Line
		}
		
		Return status_string
	}
	
	OnTick() {
		; we shouldn't spam the website, so we only check on certain hours
		for index, element in this.hours_to_refresh {
			If (element == A_Hour && this.last_hour_refreshed != A_Hour) { ; if it's currently an hour to check and we haven't checked yet
				this.last_hour_refreshed := A_Hour
				this.SuspendButtons(true)
				for index, element in this.packages { ; for each package we're tracking
					If (element.status_string != "Delivered") { ; if it hasn't already been delivered
						this.Update(index)
						; some more basic pattern evasion--it's unlikely that anyone at packagetrackr cares if we scrape the site, but let's not give them an obvious reason so stop us
						Sleep, % Random(1111, 3333) ; it takes between 3 and 8 seconds to enter a new url	
					}
				}
				this.SuspendButtons(false)
			}
		}
	}

	SuspendButtons(bool) {
		for index, element in this.packages {
			if (bool) {
				GuiControl, Disable, % element.hwnd_remove
			} else {
				GuiControl, Enable, % element.hwnd_remove
			}
		}
		if (bool) {
			GuiControl, Disable, % this.hwnd_add
		} else {
			GuiControl, Enable, % this.hwnd_add
		}
	}

	Update(index) {
		package := this.packages[index]
		package.last := GetClockTime() ; set the last checked time to now
		GuiControl pts_gui: , % package.hwnd_last, % package.last ; update the gui
		package.status_string := this.RefreshDeliveryStatus(package.tracking_number) ; actually go check the current status
		If (package.status_string == "Delivered") { ; else if ladder to set colors
			Gui pts_gui: Font, cGreen Bold ; this simply sets the desired style
			GuiControl, Font, % package.hwnd_status ; then we have to actually set the content to the previously selected style
			Gui pts_gui: Font, cBlack Normal
		} Else If (package.status_string == "") {
			package.status_string := "Unknown"
			Gui pts_gui: Font, cOlive Bold
			GuiControl, Font, % package.hwnd_status
			Gui pts_gui: Font, cBlack Normal
		} Else {
			Gui pts_gui: Font, cBlue Bold
			GuiControl, Font, % package.hwnd_status
			Gui pts_gui: Font, cBlack Normal
		}
		GuiControl pts_gui: , % package.hwnd_status, % package.status_string ; update the gui
	}

	ItemRemove(ctrl_hwnd) {
		Gui pts_gui: -AlwaysOnTop
		for index, element in this.packages {
			if element.hwnd_remove = ctrl_hwnd {
				tn := element.tracking_number
				desc := element.description
				MsgBox, 8193, Remove Package, Do you want to remove this package?`n%tn%`n%desc%
				rem := false
				IfMsgBox, OK
					rem := true
				if rem {
					this.packages.Remove(index)
					this.DrawGui()
				}
			}
		}
		return
	}

	ItemAdd() {
		Gui pts_gui: Submit
		global add_tracking_number
		global add_description
		this.AddTrackingNumber(add_tracking_number, add_description)
		this.DrawGui()
		this.Update(this.packages.MaxIndex())
		return
	}

	OnSysCommand(wp, lp, msg, hwnd) {
		if (hwnd = this.hwnd_gui && wp = SC_CLOSE) {
			OnMessage(this.WM_SYSCOMMAND, ObjBindMethod(this, "SysCommand", hGui), 0)
			this.SaveAndExit()
		}
	}

	SaveAndExit() {
		FileDelete("SimplePackageTracker.cfg")
		for index, element in this.packages {
			out_string := """" . element.tracking_number . """,""" . element.description . """`n"
			FileAppend("SimplePackageTracker.cfg", out_string)
		}
		ExitApp
	}
}

;================================================================================
;   FUNCTIONS
;================================================================================
GetClockTime(timestamp := "", format := "hh:mm tt") {
	FormatTime, output_var, % timestamp, % format
	Return output_var
}

FileAppend(filename, text) {
	FileAppend, %text%, %filename%
}

FileDelete(file_pattern) {
	FileDelete, %file_pattern%
}

IniRead(filename, section, key) {
	IniRead, result, %filename%, %section%, %key%
	Return result
}

IniWrite(filename, section, key, value) {
	IniWrite, %value%, %filename%, %section%, %key%
}

Random(min, max) {
	Random, result, % min, % max
	Return result
}

SendMessage(hwnd, msg, w_param, l_param, timeout) {
    ; possible values for Msg: https://www.autohotkey.com/docs/misc/SendMessageList.htm
    SendMessage, %msg%, %w_param%, %l_param%, , ahk_id %hwnd%, , , , %timeout%
}

SetTimer(timer, interval) {
    SetTimer, %timer%, %interval%
}