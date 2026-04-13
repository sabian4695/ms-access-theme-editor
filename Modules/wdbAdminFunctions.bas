option compare database
option explicit

public sub handleerror(modname as string, activecon as string, errdesc as string, errnum as long, optional datatag as string = "")
on error resume next

select case errnum
    case 70
        msgbox "Permissions Error - Check if the file is already in use.", vbinformation, "Error Code: " & errnum
    case 53
        msgbox "File Not Found", vbinformation, "Error Code: " & errnum
        exit sub
    case 3011
        msgbox "Looks like I'm having issues connecting to SharePoint. Please reopen when you can", vbinformation, "Error Code: " & errnum
    case 490, 52, 75
        msgbox "I cannot open this file or location - check if it has been moved or deleted. Or - you do not have proper access to this location", vbinformation, "Error Code: " & errnum
        exit sub
    case 3022
        msgbox "A record with this key already exists. I cannot create another!", vbinformation, "Error Code: " & errnum
    case 3167
        msgbox "Looks like you already deleted that record", vbinformation, "Error Code: " & errnum
        exit sub
    case 94
        msgbox "Hmm. Looks like something is missing. Check for an empty field", vbinformation, "Error Code: " & errnum
    case 3151
        msgbox "You're not connected to Oracle. Just FYI, Oracle connection does not work outside of VMWare.", vbinformation, "Error Code: " & errnum
        exit sub
    case 429
        if modname = "frmCatiaMacros" then
            msgbox "Looks like Catia isn't open", vbinformation, "Error Code: " & errnum
            exit sub
        else
            msgbox errdesc, vbinformation, "Error Code: " & errnum
        end if
    case 3343
        msgbox "Error. Please re-open WorkingDB to reset.", vbcritical, "Error Code: " & errnum
    case else
        msgbox errdesc, vbinformation, "Error Code: " & errnum
end select

end sub

function ap_disableshift()

on error goto errdisableshift
dim db as dao.database
dim prop as dao.property
const conpropnotfound = 3270

set db = currentdb()

db.properties("AllowByPassKey") = false
set db = nothing
exit function

errdisableshift:
if err = conpropnotfound then
    set prop = db.createproperty("AllowByPassKey", dbboolean, false)
    db.properties.append prop
    resume next
    else
    msgbox "Function 'ap_DisableShift' did not complete successfully."
    exit function
end if

end function

function ap_enableshift()

on error goto errenableshift
dim db as dao.database
dim prop as dao.property
const conpropnotfound = 3270

set db = currentdb()

db.properties("AllowByPassKey") = true
set db = nothing
exit function

errenableshift:
if err = conpropnotfound then
    set prop = db.createproperty("AllowByPassKey", dbboolean, true)
    db.properties.append prop
    resume next
    else
    msgbox "Function 'ap_DisableShift' did not complete successfully."
    exit function
end if

end function
