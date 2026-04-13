attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub btncancel_click()
on error goto err_handler

me.status = 99

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub closetask_click()
on error goto err_handler

me.status = 6
me.completed_date = date

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub form_load()
on error goto err_handler

call settheme(me)

exit sub
err_handler:
    call handleerror(me.name, "Form_Load", err.description, err.number)
end sub

private sub refresh_click()
on error goto err_handler

me.requery

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub submittask_click()
on error goto err_handler

me.status = 3

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
