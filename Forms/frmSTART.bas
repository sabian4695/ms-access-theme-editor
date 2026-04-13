attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub form_load()
on error goto err_handler

'here is the standard vba call for applying a theme to a form
call settheme(me)

exit sub
err_handler:
    call handleerror(me.name, "Form_Load", err.description, err.numbe)
end sub

private sub linkarticle_click()
on error goto err_handler

followhyperlink "https://www.vbadecoded.com/ms-access-vba/user-themes"

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub linkgithub_click()
on error goto err_handler

followhyperlink "https://github.com/vbadecoded/ms-access-theme-editor"

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub sampletracker_click()
on error goto err_handler

docmd.openform "frmTaskTracker_example"

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub themeeditor_click()
on error goto err_handler

docmd.openform "frmThemeEditor"

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
