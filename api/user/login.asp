<!--#include file="../../data/json.asp" -->
<!--#include file="../../data/function.asp" -->

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
DBPath = Server.MapPath("../../data/#data.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)}; dbq=" & DBPath

' ------------------------------------------------------
' 读取数据
' ------------------------------------------------------
userName = Request("username")
password = Request("password")

' 合法性
if userName = "" then
	json("code") = 202
	json("error") = "请输入登录名"
elseif password = "" then
	json("code") = 203
	json("error") = "请输入密码"
else
	sql = "select top 1 * from admin where userName = '" & userName & "' and password = '" & password & "'"
	rs.open sql, conn, 1, 3

	' 列表数据
	if (rs.eof or rs.bof) then
		json("code") = 201
		json("error") = "您还没有注册"
	else
		json("code") = 200
		' set json("data") = list

		Do while not (rs.eof or rs.bof)
			set json("data") = jsObject()
			for each item in rs.fields
				json("data")(item.Name) = item.Value
			next
			json("data")("token") = "TPDg9IJ2NOhnwL2IlOcfFoMm236Wz2hGtzFNop9ftJJb8jEo1"
			rs.movenext
		Loop
	end if

	' 关闭数据库
	rs.close
	conn.close
	set rs = nothing
	set conn = nothing
end if

' 输出数据
Response.Write(json.jsString)
%>