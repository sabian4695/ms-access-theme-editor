option compare database
option explicit

function dueday(duedate, completeddate) as string
on error resume next

if isnull(duedate) then
    dueday = "N/A"
    exit function
end if

if isnull(completeddate) then
    select case duedate
        case date
            dueday = "Today"
        case date + 1
            dueday = "Tomorrow"
        case is < date
            dueday = "Overdue"
        case is < date + 7
            dueday = weekdayname(weekday(duedate))
        case date + 7
            dueday = "1 Week"
        case is < date + 14
            dueday = "<2 Weeks"
        case date + 14
            dueday = "2 Weeks"
        case is < date + 21
            dueday = "<3 Weeks"
        case date + 21
            dueday = "3 Weeks"
        case is < date + 28
            dueday = "<4 Weeks"
        case date + 28
            dueday = "4 Weeks"
        case is > date + 28
            dueday = ">4 Weeks"
        case else
            dueday = duedate
    end select
else
    dueday = "Complete"
end if

end function

function randomnumber(low as long, high as long) as long

randomize
randomnumber = int((high - low + 1) * rnd() + low)

end function

function generatevalues()

dim db as database
dim rs as recordset

set db = currentdb()
set rs = db.openrecordset("tblTaskTracker_example")

do while not rs.eof
    rs.edit
    
    rs!request_type = randomnumber(1, 7)
    rs!complexity = randomnumber(1, 3)
    rs!assignee = randomnumber(1, 10)
    rs!checker_1 = randomnumber(1, 10)
    rs!checker_2 = randomnumber(1, 10)
    rs!customer = randomnumber(1, 3)
    rs!delay_reason = randomnumber(1, 3)
    rs!status = randomnumber(1, 6)
    
    rs.update
    
    rs.movenext
loop

end function
