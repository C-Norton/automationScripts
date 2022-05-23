#Numpad1::

;
; Add a new folder
;

; Connect to outlook
try {
outlookInstance := ComObjActive("Outlook.Application").GetNameSpace("MAPI")
inbox := outlookInstance.GetDefaultFolder(6)
}
catch e{
MsgBox, % "An error occurred attaching to Outlook (Is it open?): " e.message
return
}

CaseNumber := ""
CustomerName := ""
;Get case info
InputBox, CaseNumber, Case Number, Please enter the case number.
InputBox, CustomerName, Customer, Please enter the customer name.

; Decide on our folderName
folderName := CustomerName " - " CaseNumber

;make it
try{
newFolder := inbox.Folders.Add(folderName)
}
catch e{
MsgBox, % "An error occurred creating the outlook folder (Does it already exist?): " e.message
}

;
; So, there's a bug in outlooks COM interface. Creating rules is fine, creating subject based rules actually breaks due to type errors. This isn't an autohotkey issue, as the powershell interface has 
; the exact same issue. VBA works, because that's not going over com, so we have to do our rule creation in an outlook macro. There's no way to trigger outlook VBA directly via that interface, AND 
; directly triggerable macros CAN'T take parameters. As such, we need to do some hackey crap that I don't like here. We need to write our parameters to a text file that the macro will read to get info in
; We then need to switch to outlook, and push a button in the interface to run the macro. Fun.
;

file := FileOpen("OutlookCreateRule.txt","w") ; Carriage returns needed because outlook thinks its 1980, and decides that \n isn't a new line unless \r is also present

if (NOT file){
MsgBox, %A_LastError%
return
}
file.Write(CaseNumber)
file.Write("`r`n")
file.Write("zzz_"folderName)
file.Write("`r`n")
file.Write(folderName)
file.Write("`r`n")
file.Read()
file.Close()
Sleep 85 ; Not officially necessary, but should ensure time to write to disk. There are other, more appropriate methods, such as threading to await on that write, or reading to force a buffer flush, but I don't want to use them here; they may interfere with the
		  ; outlook script's ability to view the file as "closed"
SetTitleMatchMode 2
WinActivate, Outlook ahk_class rctrl_renwnd32 
Sleep 50
SendInput, {LAlt down}
Sleep 300
SendInput, y
Sleep 25
SendInput, 1
sleep 100
SendInput, y
sleep 25
SendInput, {LAlt up}


; Rule Created. Let's Create the folder on Goliath
FileCreateDir, \\Goliath.Rocsoft.com\home\cnorton\Tickets\%folderName%