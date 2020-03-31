option explicit

sub findColumnsWithNullValuesOnly() ' {

    dim t as dao.tableDef

    for each t in currentDb.tableDefs
      ' debug.print(t.attributes & vbtab & t.name)
        if t.attributes = 0 then
           findColumnsWithNullValuesOnlyInTable t  
        end if
    next t

end sub ' }

sub findColumnsWithNullValuesOnlyInTable(t as dao.tableDef) ' {


    dim stmt as string
    stmt = "select count(*) as cnt"

    dim c as dao.field
    for each c in t.fields
        stmt = stmt & ",count(" & c.name & ") as cnt_" & c.name
'       debug.print("  " & c.name)
    next c

    stmt = stmt & " from " & t.name

    dim rs as dao.recordSet
    set rs = currentDb.openRecordSet(stmt)

    debug.print(t.name & ": " & rs("cnt"))
    for each c in t.fields
    '   debug.print("  " & c.name & ": " & rs("cnt_" & c.name))
        if rs("cnt_" & c.name).value = 0 then
           debug.print("  " & c.name)
        end if
    next c

    set rs = nothing

end sub ' }
