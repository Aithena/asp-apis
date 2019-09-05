<!--#include file="./json.asp" -->
<!--#include file="./function.asp" -->

<%
' 申明文档
Response.ContentType = "text/json"

' 全局变量
dim data
dim json : set json = jsObject()
dim list : set list = jsArray()

' 连接数据库
set conn = Server.CreateObject("adodb.connection")
set rs = Server.CreateObject("adodb.recordset")
DBPath = Server.MapPath("./#data.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)}; dbq=" & DBPath

' ------------------------------------------------------
' 读取数据
' ------------------------------------------------------
if Request("action") = "list" then
	formid = Request("formid")
	id = Request("id")

	' 合法性
	if formid <> "" or id <> "" then
		if formid <> "" then
			sql = "select * from feedback where formid = " & formid
		end if
		if id <> "" then
			sql = "select * from feedback where id = " & id
		end if

		rs.open sql,conn,1,3

		' 列表数据
		if (rs.eof or rs.bof) then
			json("code") = 201
			json("count") = 0
			json("msg") = "没有找到相关数据"
		else
			json("code") = 200
			set json("data") = list

			Do while not (rs.eof or rs.bof)
				set list(null) = jsObject()
				for each item in rs.fields
					list(null)(item.Name) = item.Value
				next
				rs.movenext
			Loop

			json("count") = rs.recordcount
			json("msg") = ""
		end if

		' 关闭数据库
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
	else
		json("code") = 301
		json("msg") = "关键词缺失"
	end if

	' 输出数据
	Response.Write("callback("& json.jsString &")")

' ------------------------------------------------------
' 添加数据
' ------------------------------------------------------
elseif Request("action") = "add" then
	formid = Request("formid")
	name = Request("name")
	phone = Request("phone")
	company = Request("company")
	dateTime = Request("dateTime")
	postUrl = Request("postUrl")
	ip = Request("ip")
	extend = Request("extend")

	' 默认值
	if IsEmpty(dateTime) then
		dateTime = toUnixTime(Now)
	end if
	
	' 合法性
	if blacklist(phone) then
		json("code") = 202
		json("msg") = "手机黑名单"
		json("url") = ""
	elseif not checkPhone(phone) then
		json("code") = 201
		json("msg") = "请输入正确的手机号码"
		json("url") = ""
	else
		' 写入数据
		sql = "select * from feedback"
		rs.open sql,conn,1,3

		rs.addNew
		rs("formid") = formid
		rs("name") = name
		rs("phone") = phone
		rs("company") = company
		rs("dateTime") = dateTime
		rs("postUrl") = postUrl
		rs("ip") = ip
		rs("extend") = extend
		rs.update
		
		json("code") = 200
		json("msg") = "添加成功"
		json("url") = ""

		' 关闭数据库
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
	end if
	
	' 输出数据
	Response.Write("callback("& json.jsString &")")
	
' ------------------------------------------------------
' 修改数据
' ------------------------------------------------------
elseif Request("action") = "modify" then
	id = Request("id")
	formid = Request("formid")
	name = Request("name")
	phone = Request("phone")
	company = Request("company")
	dateTime = Request("dateTime")
	ip = Request("ip")
	postUrl = Request("postUrl")
	extend = Request("extend")

	' 合法性
	if id <> "" then
		' 更新数据
		sql = "select * from feedback where id = " & id
		rs.open sql,conn,1,3

		if (rs.eof or rs.bof) then
			json("code") = 201
			json("msg") = "未找到数据"
		else
			if formid <> "" then
				rs("formid") = formid
			end if	
			if name <> "" then
				rs("name") = name
			end if
			if phone <> "" then
				rs("phone") = phone
			end if
			if company <> "" then
				rs("company") = company
			end if
			if dateTime <> "" then
				rs("dateTime") = dateTime
			end if
			if postUrl <> "" then
				rs("postUrl") = postUrl
			end if
			if ip <> "" then
				rs("ip") = ip
			end if
			if extend <> "" then
				rs("extend") = extend
			end if

			rs.update

			json("code") = 200
			json("msg") = "修改成功"
		end if

		' 关闭数据库
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
	else
		json("code") = 300
		json("msg") = "关键词缺失"
	end if

	' 输出数据
	Response.Write("callback("& json.jsString &")")
	
' ------------------------------------------------------
' 删除数据
' ------------------------------------------------------
elseif Request("action") = "delete" then
	id = Request("id")

	' 合法性
	if id <> "" then	
		' 删除数据
		sql = "delete from feedback where id = " & id
		' rs.open sql,conn,1,3
		conn.execute sql

		json("code") = 200
		json("msg") = "删除成功"

		' 关闭数据库
		conn.close
		set rs = nothing
		set conn = nothing
	else
		json("code") = 300
		json("msg") = "关键词缺失"
	end if

	' 输出数据
	Response.Write("callback("& json.jsString &")")
	
end if
%>