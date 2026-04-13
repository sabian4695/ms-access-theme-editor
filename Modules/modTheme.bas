option compare database
option explicit

'---this is an api for the color picker---
'i use this on the frmthemeeditor to select colors in window
declare ptrsafe sub choosecolor lib "msaccess.exe" alias "#53" (byval hwnd as longptr, rgb as long)

'--- theme constants ---
private const dark_forebase as long = 16777215       'white
private const light_forebase as long = 657930
private const dark_btnx_back as long = 4342397
private const light_btnx_back as long = 8947896
private const dark_scalar_back as double = 1.3
private const light_scalar_back as double = 1.1
private const dark_scalar_front as double = 0.9
private const light_scalar_front as double = 0.3
private const default_primary as long = 3355443
private const default_accent as long = 5787704
private const black_substitute as long = &h1a1a1a     'dark gray for black inputs


public function settheme(setform as form)
on error goto err_handler

dim colorlevarr() as string
dim colorlevels(0 to 4) as long

dim scalarback as double, scalarfront as double, darkmode as boolean
dim backbase as long, forebase as long, backaccent as long, backsecondary as long, btnxback as long

initthemedefaults

darkmode = tempvars!thememode = "Dark"

if darkmode then
    forebase = dark_forebase
    btnxback = dark_btnx_back
    scalarback = dark_scalar_back
    scalarfront = dark_scalar_front
else
    forebase = light_forebase
    btnxback = light_btnx_back
    scalarback = light_scalar_back
    scalarfront = light_scalar_front
end if

backbase = clng(tempvars!themeprimary)
backsecondary = clng(tempvars!themesecondary)
backaccent = clng(tempvars!themeaccent)

colorlevarr = split(tempvars!themecolorlevels, ",")
buildcolorlevels colorlevels, backbase, backsecondary, colorlevarr

applyformsections setform, colorlevels
applycontrolthemes setform, colorlevels, colorlevarr, forebase, btnxback, backaccent, scalarback, darkmode

exit function
err_handler:
    call handleerror("modTheme", "setTheme", err.description, err.number)
end function


private sub initthemedefaults()
'if no theme set, apply default theme (for dev mode mostly)
    if nz(tempvars!themeprimary, "") = "" then
        tempvars.add "themePrimary", default_primary
        tempvars.add "themeSecondary", 0
        tempvars.add "themeAccent", default_accent
        tempvars.add "themeMode", "Dark"
        tempvars.add "themeColorLevels", "1.3,1.6,1.9,2.2"
    end if
end sub


private sub buildcolorlevels(colorlevels() as long, backbase as long, backsecondary as long, colorlevarr() as string)
'to achieve the 5 levels of controls, this array is the primary method.
    colorlevels(0) = backbase
    if backsecondary <> 0 then 'primary and secondary color
        colorlevels(1) = shadecolor(backsecondary, cdbl(colorlevarr(0)))
        colorlevels(2) = shadecolor(backbase, cdbl(colorlevarr(1)))
        colorlevels(3) = shadecolor(backsecondary, cdbl(colorlevarr(2)))
        colorlevels(4) = shadecolor(backbase, cdbl(colorlevarr(3)))
    else 'primary color only
        colorlevels(1) = shadecolor(backbase, cdbl(colorlevarr(0)))
        colorlevels(2) = shadecolor(backbase, cdbl(colorlevarr(1)))
        colorlevels(3) = shadecolor(backbase, cdbl(colorlevarr(2)))
        colorlevels(4) = shadecolor(backbase, cdbl(colorlevarr(3)))
    end if
end sub


private sub applyformsections(setform as form, colorlevels() as long)
'set the form parts themes - handles missing header/footer gracefully
'note - this does assume form parts don't use tags for other purposes
    on error resume next

    setform.formheader.backcolor = colorlevels(findcolorlevel(setform.formheader.tag))
    
    on error goto 0
    setform.detail.backcolor = colorlevels(findcolorlevel(setform.detail.tag))
    if len(setform.detail.tag) = 4 then
        setform.detail.alternatebackcolor = colorlevels(findcolorlevel(setform.detail.tag) + 1)
    else
        setform.detail.alternatebackcolor = setform.detail.backcolor
    end if

    on error resume next
    setform.formfooter.backcolor = colorlevels(findcolorlevel(setform.formfooter.tag))
    on error goto 0
end sub


private sub applycontrolthemes(setform as form, colorlevels() as long, colorlevarr() as string, _
    forebase as long, btnxback as long, backaccent as long, scalarback as double, darkmode as boolean)

dim ctl as control
dim fadeback as long, fadefore as long
dim level as long
dim backcol as long, levfore as double
dim disfore as double
dim forelevint as long, maxlev as long
dim btnxbackshade as long

for each ctl in setform.controls
    if ctl.tag like "*.L#*" then
        on error resume next  'tolerate missing properties per-control

        '---for all controls---
        level = findcolorlevel(ctl.tag)
        backcol = colorlevels(level)
        forelevint = level
        if forelevint > 3 then forelevint = 3

        if darkmode then
            levfore = (1 / colorlevarr(forelevint)) + 0.2
            disfore = 1.4 - levfore
        else
            levfore = (colorlevarr(forelevint))
            disfore = 15 - levfore
        end if

        maxlev = level + 1
        if maxlev > 4 then maxlev = 4
        if ctl.tag like "*ContrastBorder*" then
            ctl.bordercolor = colorlevels(maxlev)
        else
            ctl.bordercolor = backcol
        end if

        '--now, find the control type and apply the applicable
        select case ctl.controltype
            '---command button
            case accommandbutton, actogglebutton
                ctl.backcolor = backcol

                '---this is for swapping out button icons for light / dark theme icons - turned off by default---
'                        if (ctl.picture = "") then goto skipahead0
'                        if darkmode then
'                            if instr(ctl.picture, "\Core_theme_light\") then ctl.picture = replace(ctl.picture, "\Core_theme_light\", "\Core\")
'                        else
'                            if instr(ctl.picture, "\Core\") then ctl.picture = replace(ctl.picture, "\Core\", "\Core_theme_light\")
'                        end if
'skipahead0:

                '---test for individual attributes---
                if ctl.tag like "*dis*" then
                    fadefore = shadecolor(forebase, disfore)
                    ctl.forecolor = fadefore
                    ctl.hoverforecolor = fadefore
                    ctl.pressedforecolor = fadefore
                else
                    fadefore = shadecolor(forebase, levfore - 0.2)
                    ctl.forecolor = forebase
                    ctl.hoverforecolor = forebase
                    ctl.pressedforecolor = forebase
                end if

                if ctl.tag like "*btnX*" then
                    fadeback = shadecolor(btnxback, scalarback)
                    btnxbackshade = shadecolor(btnxback, (0.1 * level) + scalarback)
                    ctl.backcolor = btnxbackshade
                    ctl.bordercolor = btnxback
                else
                    fadeback = shadecolor(backcol, scalarback)
                end if

                if ctl.tag like "*accentBtn*" then
                    fadeback = shadecolor(backaccent, (0.2 * level) + scalarback)
                    ctl.backcolor = shadecolor(backaccent, scalarback)
                    ctl.gradient = 17
                end if

                ctl.hovercolor = fadeback
                ctl.pressedcolor = fadeback

                if ctl.tag like "*cardBtn*" then
                    ctl.hovercolor = backcol
                    ctl.pressedcolor = backcol
                end if

            '---label
            case aclabel
                ctl.forecolor = shadecolor(forebase, levfore)
                if ctl.tag like "*lbl_wBack.L#*" then ctl.backcolor = backcol

            '---text box
            case actextbox, accombobox
                ctl.backcolor = backcol
                if ctl.tag like "*txtTransFore*" then
                    ctl.forecolor = backcol
                elseif ctl.tag like "*txtErr*" then
                    ctl.bordercolor = btnxback
                    ctl.borderstyle = 1
                    ctl.forecolor = forebase
                else
                    ctl.forecolor = forebase
                end if

                if ctl.formatconditions.count = 1 then
                    if ctl.formatconditions.item(0).expression1 like "*IsNull*" then
                        ctl.formatconditions.item(0).backcolor = backcol
                        ctl.formatconditions.item(0).forecolor = forebase
                    end if
                end if

            '---box / subform
            case acrectangle, acsubform
                if not ctl.name like "sfrm*" then ctl.backcolor = backcol

            '---tab control
            case actabctl
                ctl.pressedcolor = backcol
                if level > 0 then
                    fadeback = shadecolor(colorlevels(level - 1), scalarback)
                else
                    fadeback = shadecolor(colorlevels(0), scalarback)
                end if
                ctl.hovercolor = fadeback
                ctl.hoverforecolor = forebase
                ctl.pressedforecolor = forebase
                if level = 0 then
                    ctl.backcolor = colorlevels(0)
                    fadefore = shadecolor(forebase, levfore - 0.6)
                    ctl.forecolor = fadefore
                else
                    ctl.backcolor = colorlevels(level - 1)
                    fadefore = shadecolor(forebase, levfore)
                    ctl.forecolor = fadefore
                end if

            '---picture
            case acimage
                ctl.backcolor = backcol
        end select

        on error goto 0
    end if
next

end sub


function findcolorlevel(tagtext as string) as long
on error goto err_handler

    dim pos as long
    findcolorlevel = 0
    if tagtext = "" then exit function

    pos = instr(tagtext, ".L")
    if pos = 0 then exit function

    dim levelchar as string
    levelchar = mid$(tagtext, pos + 2, 1)
    if isnumeric(levelchar) then
        findcolorlevel = clng(levelchar)
    end if

exit function
err_handler:
    call handleerror("modTheme", "findColorLevel", err.description, err.number)
end function


function shadecolor(inputcolor as long, scalar as double) as long
on error goto err_handler

dim temphex as string
dim ior as long, iog as long, iob as long

'black input gets bumped to dark gray so it shades visibly
if inputcolor = 0 then inputcolor = black_substitute

temphex = right$("000000" & hex(inputcolor), 6)

ior = val("&H" & mid$(temphex, 5, 2)) * scalar
iog = val("&H" & mid$(temphex, 3, 2)) * scalar
iob = val("&H" & mid$(temphex, 1, 2)) * scalar

if ior > 255 then ior = 255
if iog > 255 then iog = 255
if iob > 255 then iob = 255

if ior < 0 then ior = 0
if iog < 0 then iog = 0
if iob < 0 then iob = 0

shadecolor = rgb(ior, iog, iob)

exit function
err_handler:
    call handleerror("modTheme", "shadeColor", err.description, err.number)
end function


public function colorpicker(optional lngcolor as long) as long
on error goto err_handler
    choosecolor application.hwndaccessapp, lngcolor
    colorpicker = lngcolor
exit function
err_handler:
    call handleerror("modTheme", "colorPicker", err.description, err.number)
end function
