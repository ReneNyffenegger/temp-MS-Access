select *
from
   msysobjects
where
  name like '*Sys*'
order by
  name
