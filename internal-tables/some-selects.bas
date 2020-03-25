' vim: ft=basic
option explicit

sub main()  ' {

    dim rootIds as dao.recordSet

    set rootIds = currentDb.openRecordset( _
       "select distinct parentId from msysObjects where parentId not in " & _
          "(select distinct id from msysObjects)" )


'   children 251658241
'   exit sub

    do until rootIds.eof ' }

       debug.print("Root id: " & rootIds(0))

       children rootIds(0)

       rootIds.moveNext

    loop ' }

end sub ' }

sub children(parentId as long) ' {

    dim sel as dao.queryDef


    set sel = currentDb.createQueryDef("",                         _ 
       "parameters pid long; "                              &      _
       "select * from msysObjects where parentId = pid")     

    sel.parameters("pid") = parentId

    dim rs as dao.recordSet
    set rs = sel.openRecordset

    do until rs.eof ' {
       debug.print(format(rs!id, "               #") & " | " & format(rs!parentId, "              #") & " |  " & rs!name)
       rs.moveNext
    loop ' }

end sub ' }


sub dao_query() ' {

    dim rs as dao.recordSet

    set rs = currentDb.openRecordset("SELECT * FROM MSysWSDPCacheComplexColumnMapping;")

    do until rs.eof ' }
       debug.print(rs("name"))
       rs.moveNext
    loop ' }

end sub ' }
