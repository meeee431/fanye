--函数：从售票表中统计车次已售票数
CREATE FUNCTION [dbo].[GetSaledSeatCount]( 
@BusDate DATETIME,
@BusId VARCHAR(20) 
)RETURNS INT
AS 
BEGIN 
DECLARE @r INT
SET @r=0

SELECT @r = COUNT(*) FROM ticket_sell_lst 
WHERE bus_date = @BusDate AND bus_id = @BusId 
AND status IN (1 ,2 ,33 ,34)

RETURN @r
END 

GO

--函数：统计预订数
CREATE function [dbo].[BookNums]( 
@busdate datetime,
@busid Varchar(20) 
)returns int
as 
begin 
declare @r int
set @r=0
select @r=SUM(book_count) from work_env_bus_station_lst 
where bus_date=@busdate and bus_id=@busid
return @r
end 

GO






