select *
from
   msysobjects
where
  name in (
  -- 'MSysAccessObjects' ,
  -- 'MSysAccessXML'     ,
     'MSysACEs'          ,
     'MSysComplexColumns',
     'MSysObjects'       ,
     'MSysQueries'       ,
     'MSysRelationships' 
  )
order by
  name
