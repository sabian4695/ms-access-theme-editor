attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

function filterit(controlname as string)
on error goto err_handler

me(controlname).setfocus
docmd.runcommand accmdfiltermenu

exit function
err_handler:
    call handleerror(me.name, "filterIt", err.description, err.number)
end function

private sub form_load()
on error goto err_handler

call settheme(me)

me.filter = "completed_date is null"
me.filteron = true

me.orderby = "Due_date"
me.orderbyon = true

exit sub
err_handler:
    call handleerror(me.name, "Form_Load", err.description, err.number)
end sub

private sub newtask_click()
on error goto err_handler

msgbox "No sample form here, just a sample button to show an 'action button'"

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub opendetails_click()
on error goto err_handler

docmd.openform "frmTaskDetails_example", , , "recordId = " & me.recordid

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

public sub refresh_click()
on error goto err_handler

me.requery

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
