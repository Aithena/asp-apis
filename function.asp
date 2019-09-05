<%
' -------------------------------
' 手机号码验证
' checkPhone(15800002222)
' -------------------------------
function checkPhone(str)
	set reg = new RegExp
	reg.Pattern = "^1[3456789]\d{9}$"
	
	if reg.test(str) then
		result = true
	else
		result = false
	end if
	
	checkPhone = result
end function

' -------------------------------
' 手机号码黑名单
' blacklist(15800002222)
' -------------------------------
function blacklist(str)
	dim list
	list = array()

	for i=0 to ubound(list)
		if list(i) = cstr(str) then
			blacklist = true
			exit function
		else
			blacklist = false
		end if
	next
end function

' ----------------------------------
' 标准时间转化为时间戳(东八区)
' toUnixTime(Now)
' toUnixTime("2018-05-05 15:05:00")
' ----------------------------------
function toUnixTime(strTime)
	if IsEmpty(strTime) or not IsDate(strTime) then
		strTime = Now
	end if

	toUnixTime = DateAdd("h", -8, strTime)
	toUnixTime = DateDiff("s", "1970-01-01 00:00:00", toUnixTime)
end function

' ------------------------------------
' 时间戳转化成标准时间(东八区)
' fromUnixTime(1567415514)
' ------------------------------------
function fromUnixTime(intTime)
	if isEmpty(intTime) or not isNumeric(intTime) then
		fromUnixTime = Now()
		exit function
	end if

	fromUnixTime = DateAdd("s", intTime, "1970-01-01 00:00:00")
	fromUnixTime = DateAdd("h", +8, fromUnixTime)
end function

%>