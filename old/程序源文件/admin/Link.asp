<!--#include file="../conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
Head()
Dim admin_flag
admin_flag=",15,"
CheckAdmin(admin_flag)
Dim body,readme,Tlink,rs
Call main()
Footer()

Sub main()
	Select Case request("action")
		Case "saveall"
			Call saveall()
		Case "add" 
			Call addlink()
		Case "edit"
			Call editlink()
		Case "savenew"
			Call savenew()
		Case "savedit"
			Call savedit()
		Case "del"
			Call del()
		Case "orders"
			Call orders()
		Case "updatorders"
			Call updateorders()
		Case Else
			call linkinfo()
	End Select
	If Body<>"" Then
		Dv_suc(body)
	End If
End Sub

Sub addlink()
	Response.Write "<table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=""center"">"
	Response.Write "<form action=""link.asp?action=savenew"" method = post> <tr>"
	Response.Write "<th width=""100%"" colspan=2 style=""text-align:center;"">添加联盟论坛 </th>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td width=""40%"" height=25 class=td1>论坛名称</td>"
	Response.Write "<td width=""60%"" height=25 class=td1>"
	Response.Write "<input type=""text"" name=""name"" size=40>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td width=""40%"" height=25 class=td2>连接URL</td>"
	Response.Write "<td width=""60%"" class=td2>"
	Response.Write "<input type=""text"" name=""url"" size=40>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td width=""40%"" height=25 class=td1>连接LOGO地址</td>"
	Response.Write "<td width=""60%"" class=td1>"
	Response.Write "<input type=""text"" name=""logo"" size=40>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td width=""40%"" height=25 class=td2>论坛简介</td>"
	Response.Write "<td width=""60%"" class=td2>"
	Response.Write "<input type=""text"" name=""readme"" size=40>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td width=""40%"" height=25 class=td1>在首页是文字连接还是LOGO连接</td>"
	Response.Write "<td width=""60%"" class=td1>"
	Response.Write "文字连接<input type=""radio"" class=""radio"" name=""islogo"" value=0 checked>&nbsp;&nbsp;LOGO连接<input type=""radio"" class=""radio"" name=""islogo"" value=1>"
	Response.Write "&nbsp;&nbsp;<Input type=""submit"" class=""button"" name=""Submit"" value=""添 加"">"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</form>"
	Response.Write "</table>"
End Sub

sub editlink()
	Dim rs,sql
	set rs= Dvbbs.iCreateObject ("adodb.recordset")
	sql = "select * from dv_bbslink where id="&Request("id")
	rs.open sql,conn,1,1
	Response.Write "<table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=""center"">"
	Response.Write "<form action=""link.asp?action=savedit"" method=post>"
	Response.Write "<input type=hidden name=id value="
	Response.Write Request("id")
	Response.Write "><tr> <th width=""100%"" colspan=2 style=""text-align:center;"">编辑联盟论坛</th>"
	Response.Write "</tr><tr> "
	Response.Write "<td width=""40%"" class=td1>"
	Response.Write "论坛名称： </td>"
	Response.Write "<td width=""60%"" class=td1>"
	Response.Write "<input type=""text"" name=""name"" size=40 value="
	Response.Write server.htmlencode(rs("boardname"))
	Response.Write ">"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td width=""40%"" class=td1>"
	Response.Write "连接URL： </td>"
	Response.Write "<td width=""60%"" class=td1>"
	Response.Write "<input type=""text"" name=""url"" size=40 value="
	Response.Write server.htmlencode(rs("url"))
	Response.Write ">"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td width=""40%"" class=td1>"
	Response.Write "连接LOGO地址： </td>"
	Response.Write "<td width=""60%"" class=td1>"
	Response.Write "<input type=""text"" name=""logo"" size=40 value="""
	If Rs("logo")<>"" or Not IsNull(Rs("logo")) Then Response.Write server.htmlencode(rs("logo"))
	Response.Write """>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td width=""40%"" class=td1>"
	Response.Write "论坛简介：</td>"
	Response.Write "<td width=""60%"" class=td1>"
	Response.Write "<input type=""text"" name=""readme"" size=40 value="
	If Rs("readme")<>"" or Not IsNull(Rs("readme")) Then Response.Write server.htmlencode(rs("readme"))
	Response.Write ">"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td width=""40%"" height=25 class=td1>在首页是文字连接还是LOGO连接</td>"
	Response.Write "<td width=""60%"" class=td1>"
	Response.Write "文字连接<input type=""radio"" class=""radio"" name=""islogo"" value=0 "
	If rs("islogo")=0 Then
	 	Response.Write " checked"
	End If
	Response.Write ">&nbsp;&nbsp;LOGO连接<input type=""radio"" class=""radio"" name=""islogo"" value=1 "
	If rs("islogo")=1 Then
		Response.Write " checked"
	End If 
	Response.Write ">"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td height=""15"" colspan=""2"" class=td1>"
	Response.Write "<div align=""center"">"
	Response.Write "<input type=""submit"" class=""button"" name=""Submit"" value=""修 改"">"
	Response.Write "</div>"
	Response.Write "</td>"
	Response.Write "</tr></form>"
	Response.Write "</table>"
	Rs.Close
	Set Rs=Nothing
End Sub

Sub linkinfo()
	Dim rs,sql,i
	i=0
	addlink()
	Set rs= Dvbbs.iCreateObject ("adodb.recordset")
	sql = " select * from dv_bbslink order by id"
	rs.open sql,conn,1,1       
	Response.Write "<br><table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=""center"">"
	Response.Write "<form action=""link.asp?action=saveall"" method = post>"
	Response.Write "<tr>"
	Response.Write "<th style=""text-align:center;"" colspan=4>联盟论坛列表批量修改 </th>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td height=""22"" colspan=4 class=td2><b>注意事项：</b><li>你可以编辑所有友情链接信息然后一次性提交。<li>其中序号部分可以按照你的需要修改，不能有重复。<li>其他内容尽量避免使用单引号，以免破坏页面代码执行。</td>"
	Response.Write "</tr>"
	If rs.eof and rs.bof Then
		Response.Write "<tr><td height=""25"" colspan=4 align=""center"" class=td2>尚未添加友情论坛</td></tr>"
	Else
		
	Do While Not Rs.EOF
	Response.Write "<tr align=left>"
    Response.Write "<td height=25 class=td1>"
    Response.Write "<B>序号：</B> <input type=""text"" name=""id"" size=4 value="
	Response.Write Rs("id")
	Response.Write "></td>"
	Response.Write "<td class=td1>"
	Response.Write "<B>名称：</B><input type=""text"" name=""boardname"&i&""" size=30 value="
	Response.Write server.htmlencode(Rs("boardname")&"")
	Response.Write "></td>"
	Response.Write "<td class=td1>"
	Response.Write "<B> URL：</B><input type=""text"" name=""url"&i&""" size=35 value="
	Response.Write server.htmlencode(Rs("url")&"")
	Response.Write "></td>"
	Response.Write "<td class=td1><a href=""link.asp?action=orders&id="
	Response.Write Rs("id")
	Response.Write """>排序</a>  <a href=""link.asp?action=edit&id="
	Response.Write Rs("id")
	Response.Write """>编辑</a>  <a href=""link.asp?action=del&id="
	Response.Write Rs("id")
	Response.Write """ onclick=""{if(confirm('删除后不可恢复，确定删除吗?')){return true;}return false;}"">删除</a></td>"
	Response.Write "</tr><tr>"
	Response.Write "<td class=td2><b>是否图片</b><br>"
	Response.Write "是<input type=""radio"" class=""radio"" name=""islogo"&i&""" value=""1"" "
	If Rs("islogo")=1 Then	
		Response.Write " checked"
	End If
	Response.Write ">"
	
	Response.Write "否<input type=""radio"" class=""radio"" name=""islogo"&i&""" value=""0"" "
	If Rs("islogo")=0 Then	
		Response.Write " checked"
	End If
	Response.Write ">"
	Response.Write "</td><td class=td2>"
	Response.Write "<b>logo：</b>"
	Response.Write "<input type=""text"" name=""logo"&i&""" size=30 value="
	If Rs("logo")<>"" or Not IsNull(Rs("logo")) Then Response.Write server.htmlencode(Rs("logo"))
	Response.Write "></td>"
	Response.Write "<td class=td2 colspan=4><B>简介：</B>"
	Response.Write "<input type=""text"" name=""readme"&i&""" size=50 value="
	If Rs("readme")<>"" or Not IsNull(Rs("readme")) Then Response.Write server.htmlencode(Rs("readme"))
	Response.Write "></td></tr>"
	Response.Write "<tr><th style=""height:1px;"" colspan=4>&nbsp;</td></tr>"
	i=i+1
	rs.movenext
	loop
	
	Response.Write "<tr><td height=""25"" colspan=4 align=""center""><input type=""submit"" class=""button"" name=""Submit"" value=""批量更新""></td></tr>"
	Response.Write "</from>"
	End If
	Response.Write "</table>"
        rs.Close
	set rs=Nothing
End Sub 

sub savenew()
if Request("url")<>"" and Request("readme")<>"" and request("name")<>"" then
	dim linknum,rs,sql
	set rs= Dvbbs.iCreateObject ("adodb.recordset")
	sql = "select * from dv_bbslink order by id desc"
	rs.Open sql,conn,1,3
	if rs.eof and rs.bof then
	linknum=1
	else
	linknum=rs("id")+1
	end if
	sql="insert into dv_bbslink(id,boardname,readme,logo,url,islogo) values("&linknum&",'"&fixjs(Trim(Request.Form ("name")))&"','"&fixjs(Trim(Request.Form ("readme")))&"','"&fixjs(trim(request.Form("logo")))&"','"&fixjs(Request.Form ("url"))&"',"&CInt(request.Form("islogo"))&")"
	Dvbbs.Execute(sql) 
	rs.Close
	set rs=Nothing 
	Call loadlinkList()
	body=body+"<br>"+"更新成功，请继续其他操作。"
else
	body=body+"<br>"+"请输入完整联盟论坛信息。"
end if
end sub

sub savedit()
if Request("url")<>"" and Request("readme")<>"" and request("name")<>"" then
	Dim rs,sql
	set rs= Dvbbs.iCreateObject ("adodb.recordset")
	sql = "select * from dv_bbslink where id="&Dvbbs.CheckNumeric(request("id"))
	rs.Open sql,conn,1,3
	if rs.eof and rs.bof then
	body=body+"<br>"+"错误，没有找到联盟论坛。"
	else
	rs("boardname") = fixjs(Trim(Request.Form ("name")))
	rs("readme") =  fixjs(Trim(Request.Form ("readme")))
	rs("logo")=fixjs(Trim(request.Form("logo")))
	rs("url") = fixjs(Request.Form ("url"))
	rs("islogo")=request.Form("islogo")
	rs.Update
	end if 
	rs.Close
	set rs=nothing
	Call loadlinkList()
	body=body+"<br>"+"更新成功，请继续其他操作。"
Else
	body=body+"<br>"+"请输入完整联盟论坛信息。"
End If
end sub

sub del
	dim id,sql
	id = Dvbbs.CheckNumeric(request("id"))
	sql="delete from dv_bbslink where id="&id
	Dvbbs.Execute(sql)
	body=body+"<br>"+"删除成功，请继续其他操作。"
	Call loadlinkList()
end sub

sub orders()
	Response.Write "<br>"
	Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=center>"
	Response.Write "			<tr><th style=""text-align:center;"">联盟论坛重新排序</th></tr>"
	Response.Write "<tr>"
	Response.Write "<td height=""23"" class=td2>"
	Response.Write "注意：请在相应论坛的排序表单内输入相应的排列序号，<font color=red>注意不能和别的联盟论坛有相同的排列序号</font>。</font>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<td class=td1>"

	Dim rs,sql
	set rs= Dvbbs.iCreateObject ("adodb.recordset")
	sql="select * from dv_bbslink where id="&Dvbbs.CheckNumeric(request("id"))
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "没有找到相应的联盟论坛。"
	else
		response.write "<form action=link.asp?action=updatorders method=post>"
		response.write ""&rs("boardname")&"  <input type=text name=newid size=2 value="&rs("id")&">"
		response.write "<input type=hidden name=id value="&request("id")&">"
		response.write "<input type=submit class=button name=Submit value=修改></form>"
	end if
	rs.close
	set rs=Nothing
	Response.Write"</td>"
	Response.write"</tr>"
	Response.write"<tr>"
	Response.write"<td>"
	Response.write"<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
	Response.write"<tr><th style=""text-align:center;"" colspan=4>当前排序情况</th></tr>"
	Response.Write "<tr>"
	Dim a
	a=0		
	Set rs=Dvbbs.Execute("select id,boardname from dv_bbslink order by id")
	Do While Not Rs.EOF
	Response.Write "<td height=22 width=""25%"">"
	Response.Write Rs(0)
	Response.Write "、"
	Response.Write Rs(1)
	Response.Write "</td>"
	
	Rs.MoveNext
	a=a+1
	If a=4 Then
		a=0
		Response.Write "</tr><tr><th style=""height:1px;"" colspan=4>&nbsp;</td></tr><tr>"		
	End If
	Loop
	Set rs=Nothing
	Response.Write "</tr>"
	Response.Write "<tr><th style=""height:1px;"" colspan=4>&nbsp;</td></tr>"
	Response.Write " </table>"
	Response.Write"</td></tr>"
	Response.write"</table>"			
end sub

sub updateorders()
	Dim rs,sql
	if isnumeric(request("id")) and isnumeric(request("newid")) and request("newid")<>request("id") then
		set rs=Dvbbs.Execute("select id from dv_bbslink where id="&request("newid"))
		if rs.eof and rs.bof then
		sql="update dv_bbslink set id="&request("newid")&" where id="&cstr(request("id"))
		Dvbbs.Execute(sql)
		response.write "更新成功！"
		else
		response.write "更新失败，您指定了和其他联盟论坛相同的序号！"
		end if
	else
		response.write "更新失败！您输入的字符不合法，或者输入了和原来相同的序号！"
	end if
end sub

Sub saveall()
	Dim IDlist,id,i,tmpstr
	ID=Request.form("id")
	id=Replace(id," ","")
	IDlist=","&ID&","
	ID=split(id,",")
	For i=0 to UBound(id)
		tmpstr=","&ID(i)&","
		If InStr(IDlist,tmpstr)>0 Then
			If InStr(Len(tmpstr)-1+InStr(IDlist,tmpstr),IDlist,tmpstr)>0 Then
			Errmsg=ErrMsg + "发现相同的序号："&ID(i)&",请返回仔细检查。"
			Exit For
			End If 
		End If	
	Next 	
	If Errmsg<>"" Then
		dvbbs_error()
	End If
	'清除原来数据表中的数据,打篮球，五进五出了。：）
	Dvbbs.Execute("Delete from dv_bbslink")
	'开始利用循环插入数据
	Dim sql,boardname,readme,url,logo,islogo
	For i= 0 to UBound(id)
		boardname=fixjs(Request.form("boardname"&i))
		readme=fixjs(Request.form("readme"&i))
		url=fixjs(Request.form("url"&i))
		logo=fixjs(Request.form("logo"&i))
		islogo=Request.form("islogo"&i)
		sql="insert into dv_bbslink (id,boardname,readme,url,logo,islogo) values ("&CInt(id(i))&",'"&boardname&"','"&readme&"','"&url&"','"&logo&"',"&CInt(islogo)&")"
		Dvbbs.Execute(sql)	
	Next
	loadlinkList()
	Dv_suc("论坛批量更新成功！")
End Sub 
Sub LoadlinkList()
	Dim rs
	Set Rs=Dvbbs.Execute("select * From Dv_bbslink Order by islogo desc,id ")
	Set Application(Dvbbs.CacheName & "_link")=Dvbbs.RecordsetToxml(rs,"link","bbslink")
	Set Rs=Nothing
End Sub
%>
