<!--#include file =../conn.asp-->
<!-- #include file="inc/const.asp" -->
<%
Head()
Dim admin_flag
admin_flag=",4,"
CheckAdmin(admin_flag)
Select case request("action")
	case "save1"
		save1()
	case "save2"
		save2()
	case "save3"
		save3()
	case "search"
		search()
	case "view"
		view()
	case "edit"
		edit()
	case "del"
		del()
	case else
		consted()
end select
	If founderr then call dvbbs_error()
	footer()

Sub consted()
dim sel
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th colspan="3" style="text-align:center;">论坛帮助管理 | <a href="../boardhelp.asp" target=_blank>预览帮助</a></th>
</tr>
<FORM METHOD=POST ACTION="help.asp?action=search">
<tr>
<td height="25" colspan="3" class="td1">
<B>输入帮助关键字</B>：<input type="text" name="keyword" size=50> <input type=submit class=button name=submit value="搜 索"><BR><BR>输入帮助关键字进行搜索，不输入查询条件为显示所有帮助
</td>
</tr>
</FORM>
</table><BR>
<%
search()
%><BR>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr><th>论坛帮助添加</th></tr>
<tr> 
<td width="48%" class="td2" valign=top>
<li>所添加内容将自动显示于前台的帮助页面；
<li>添加一级分类方法：选取下拉选项“<font color=blue>作为一级分类</font>”，并填写旁边分类标题，标题与内容可以不填写；
<li><font color=blue>所有填写内容均可使用HTML语法填写！</font><BR>
</td>
</tr>
<tr> 
<td width="48%" class="td1" valign=top>
<FORM METHOD=POST ACTION="help.asp?action=save1">
<table width=100% >
<tr><td width=40>标题：</td><td width=*><input type="text" name="title" size=50></td></tr>
<tr><td width=40>分类：</td><td width=*>
<select name="classid">
<%
Dim rs
set rs=Dvbbs.Execute("select * from dv_help where h_type=0 and h_parentid=0")
do while not rs.eof
	Response.Write "<option value="&rs("h_id")&">"&server.htmlencode(rs("h_title"))&"</option>"
rs.movenext
loop
rs.close
set rs=nothing
%>
<option value="0">作为一级分类</option>
</select>&nbsp;<input type="text" name="classtitle" size=30>
</td></tr>
<tr><td width=40>类型：</td><td width=*><input type=checkbox class=checkbox name="stype" checked value="1">&nbsp;选择此项所添加内容将不显示在帮助文件首页中</td></tr>
<tr><td width=40>内容：</td><td width=*>
<textarea name="content" cols="80" rows="8" ID="TDcontent"></textarea><a href="javascript:admin_Size(-8,'TDcontent')"><img src="skins/images/minus.gif" unselectable="on" border='0'></a> <a href="javascript:admin_Size(8,'TDcontent')"><img src="skins/images/plus.gif" unselectable="on" border='0'></a>
</tr>
<tr><td width=40>&nbsp;</td><td width=*><input type=submit class=button name=submit value="保 存"></td></tr>
</table>
</FORM>
</td>
</tr>
</table>
<%
end sub

sub save1()
	dim title,content,parentid,stype,rs,sql
	if Request.form("classid")="0" then
		if request("classtitle")="" then
			Errmsg="如果您选择添加一级分类，请填写选择下拉框的左边输入框"
			founderr=true
			exit sub
		else
			title=request("classtitle")
		end if
		ParentID=0
	else
		if request("title")="" then
			Errmsg="请填写帮助标题"
			founderr=true
			exit sub
		else
			title=request("title")
		end if
		ParentID=Request.form("classid")
	end if
	if request("stype")="1" then
		stype=1
	else
		stype=0
	end if
	set rs=Dvbbs.iCreateObject("adodb.recordset")
	sql="select * from dv_help"
	rs.open sql,conn,1,3
	rs.addnew
	rs("h_parentid")=parentid
	rs("h_title")=FilterJS(title)
	rs("h_content")=replace(FilterJS(request.form("content")),chr(10),"<br>")
	rs("h_bgimg")=FilterJS(request.form("targeturl"))
	rs("h_type")=0
	rs("h_stype")=stype
	rs.update
	rs.close
	set rs=nothing
	dv_suc("保存前台帮助成功！")
end sub

sub save3()
	dim title,content,parentid,stype,rs,SQL
	if request("classid")="0" then
		if request("classtitle")="" then
			Errmsg="如果您选择添加一级分类，请填写选择下拉框的左边输入框"
			founderr=true
			exit sub
		else
			title=request("classtitle")
		end if
		ParentID=0
	else
		if request("title")="" then
			Errmsg="请填写帮助标题"
			founderr=true
			exit sub
		else
			title=request("title")
		end if
		ParentID=request("classid")
	end if
	if request("stype")="1" then
		stype=1
	else
		stype=0
	end if
	set rs=Dvbbs.iCreateObject("adodb.recordset")
	sql="select * from dv_help where h_id="&Dvbbs.CheckNumeric(request("id"))
	rs.open sql,conn,1,3
	if not rs.eof then
	rs("h_parentid")=parentid
	rs("h_title")=FilterJS(title)
	rs("h_content")=replace(FilterJS(request("content")),chr(10),"<br>")
	rs("h_bgimg")=FilterJS(request("targeturl"))
	rs("h_type")=request("ctype")
	rs("h_stype")=stype
	end if
	rs.update
	rs.close
	set rs=nothing
	dv_suc("保存后台帮助成功！")
end sub
Function FilterJS(v)
If not isnull(v) then
dim t
dim re
dim reContent
Set re=new RegExp
re.IgnoreCase =true
re.Global=True
re.Pattern="(javascript)"
t=re.Replace(v,"&#106avascript")
re.Pattern="(jscript:)"
t=re.Replace(t,"&#106script:")
re.Pattern="(js:)"
t=re.Replace(t,"&#106s:")
re.Pattern="(value)"
t=re.Replace(t,"&#118alue")
re.Pattern="(about:)"
t=re.Replace(t,"about&#58")
re.Pattern="(file:)"
t=re.Replace(t,"file&#58")
re.Pattern="(document.cookie)"
t=re.Replace(t,"documents&#46cookie")
re.Pattern="(vbscript:)"
t=re.Replace(t,"&#118bscript:")
re.Pattern="(vbs:)"
t=re.Replace(t,"&#118bs:")
re.Pattern="(on(mouse|exit|error|click|key))"
t=re.Replace(t,"&#111n$2")
re.Pattern="(&#)"
t=re.Replace(t,"＆#")
FilterJS=t
set re=nothing
End if
End Function

function search()
%>
<table width="100%" border="0" cellspacing="1" cellpadding="5" align="center">
<tr> 
<th colspan=2 style="text-align:center;">论坛帮助和后台菜单管理列表</th>
</tr>
<%
dim keyword,currentpage,page_count,totalrec,rs,sql
set rs=Dvbbs.iCreateObject("adodb.recordset")
if request("stype")<>"" then
	sql=" h_type="&request("stype")
end if
if request("keyword")<>"" then
	if sql<>"" then
	sql=sql & " and h_title like '%"&replace(request("keyword"),"'","")&"%' or h_content like '%"&replace(request("keyword"),"'","")&"%'"
	else
	sql=" h_title like '%"&replace(request("keyword"),"'","")&"%' or h_content like '%"&replace(request("keyword"),"'","")&"%'"
	end if
end if
if sql="" then
	sql="select * from dv_help where not h_id=1 order by h_id desc"
else
	sql="select * from dv_help where "&sql&" and not h_id=1 order by h_id desc"
end if
rs.open sql,conn,1,1

currentPage=request("page")
if currentpage="" or not IsNumeric(currentpage) then
	currentpage=1
else
	currentpage=clng(currentpage)
end if
if not rs.eof then
	rs.PageSize = 10
	rs.AbsolutePage=currentpage
	page_count=0
    totalrec=rs.recordcount
	while (not rs.eof) and (not page_count = rs.PageSize)
%>
<tr><td class=td2 width=80% nowrap>
<B><%
Response.Write rs("h_title")
%></B>
</td>
<td class=td2 width=20% align=right nowrap>
[<a href="help.asp?action=view&id=<%=rs(0)%>">查看内容</a>] [<a href="help.asp?action=edit&id=<%=rs(0)%>">编辑</a>] [<a href="help.asp?action=del&id=<%=rs(0)%>">删除</a>]
</td>
</tr>
<tr><td class=td1 colspan=2>
<%
if not isnull(rs("h_content")) and rs("h_content")<>"" then
	Response.Write left(replace(replace(rs("h_content"),"<br>"," "),"<BR>"," "),100)
else
	if rs("H_ParentID")=0 then
	Response.Write "<font color=red>此项为一级分类标题!</font>"
	Else
	Response.Write "本条帮助没有录入内容!"
	End If
end if
%>
</td></tr>
<%
	page_count = page_count + 1
	rs.movenext
	wend
else
%>
<tr> 
<td height="23" class="td2" colspan=2>没有找到任何帮助</td>
</tr>
<%
end if
rs.close
set rs=nothing
dim pcount,endpage,i
  	if totalrec mod 10=0 then
     		Pcount= totalrec \ 10
  	else
     		Pcount= totalrec \ 10+1
  	end if
	response.write "<tr><td valign=middle nowrap class=td2 height=23 colspan=2>"
	response.write "<table width=""100%""><tr><td valign=middle nowrap class=td2>页次：<b>"&currentpage&"</b>/<b>"&Pcount&"</b>页"
	response.write "&nbsp;每页<b>10</b> 总数<b>"&totalrec&"</b></td>"
	response.write "<td valign=middle nowrap align=right class=td2>分页："
	if currentpage > 4 then
	response.write "<a href=""?page=1&action="&request("action")&"&keyword="&request("keyword")&"&stype="&request("stype")&""">[1]</a> ..."
	end if
	if Pcount>currentpage+3 then
	endpage=currentpage+3
	else
	endpage=Pcount
	end if
	for i=currentpage-3 to endpage
	if not i<1 then
		if i = clng(currentpage) then
        response.write " <font color=red>["&i&"]</font>"
		else
		response.write " <a href=""?page="&i&"&action="&request("action")&"&keyword="&request("keyword")&"&stype="&request("stype")&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&"&action="&request("action")&"&keyword="&request("keyword")&"&stype="&request("stype")&""">["&Pcount&"]</a>"
	end if
	response.write "</td></tr></table></td></tr>"
%>

</table>
<%
end function

function view()
%>
<table width="100%" border="0" cellspacing="0" cellpadding="3" align="center">
<tr> 
<th style="text-align:center;">查看论坛帮助</th>
</tr><tr> 
<td height="23" class="td2">
<%
Dim rs
set rs=Dvbbs.Execute("select * from dv_help where h_id="&request("id"))
if rs.eof and rs.bof then
	Response.Write "没有找到帮助"
else
	Response.Write "<BR><center><b>"&rs("h_title")&"</b></center><BR><BR><BR>"
	Response.Write "<blockquote>"&rs("h_content")&"</blockquote>"
	Response.Write "<div align=right>[<a href=help.asp?action=edit&id="&rs(0)&">编辑</a>] [<a href=help.asp?action=del&id="&rs(0)&">删除</a>]</div>"
end if
%>
</td>
</tr>
</table>
<%
end function

function edit()
%>
<table width="100%" border="0" cellspacing="0" cellpadding="3" align="center">
<tr> 
<th style="text-align:center;">论坛帮助编辑</th>
</tr><tr> 
<td height="23" class="td2">
<%
dim trs,rs
set rs=Dvbbs.Execute("select * from dv_help where h_id="&request("id"))
if rs.eof and rs.bof then
	Response.Write "没有找到帮助"
else
%>
<FORM METHOD=POST ACTION="help.asp?action=save3">
<input type=hidden value="<%=request("id")%>" name="id">
<table width=100%>
<tr><td width=40>标题：</td><td width=*><input type="text" name="title" size=35 value="<%if not rs("h_parentid")=0 then Response.Write server.htmlencode(rs("h_title"))%>">&nbsp;&nbsp;
<select name="ctype" size=1>
<option value=0 <%if rs("h_type")=0 then Response.Write "selected"%>>前台</option>
<option value=1 <%if rs("h_type")=1 then Response.Write "selected"%>>后台</option>
</select>
</td></tr>
<tr><td width=40>分类：</td><td width=*>
<select name="classid">
<%
set trs=Dvbbs.Execute("select * from dv_help where h_parentid=0 order by H_type ")
do while not trs.eof
	Response.Write "<option value="&trs("h_id")
	if rs("h_parentid")=trs(0) then Response.Write " selected"
	Response.Write ">"&server.htmlencode(trs("h_title"))
	if Cint(trs("H_type")) = 1 Then Response.Write " [后台]" Else Response.Write " [前台]" 
	Response.Write "</option>"
trs.movenext
loop
trs.close
set trs=nothing
%>
<option value="0" <%if rs("h_parentid")=0 then Response.Write "selected"%>>作为一级分类</option>
</select>&nbsp;<input type="text" name="classtitle" size=30 value="<%if rs("h_parentid")=0 then Response.Write server.htmlencode(rs("h_title"))%>">
</td></tr>
<tr><td width=40>类型：</td><td width=*><input type=checkbox class=checkbox name="stype" value="1" <%if rs("h_stype")=0 then Response.Write "checked"%>>&nbsp;选择此项所添加内容将不显示在左边菜单中(后台有效)</td></tr>
<tr><td width=40>背景：</td><td width=*><input type="text" name="targeturl" size=35 value="<%=server.htmlencode(rs("h_bgimg"))%>"> 填写路径</td></tr>
<tr><td width=40>内容：</td><td width=*>
<textarea name="content" cols="80" rows="8"><%if not isnull(rs("h_content")) and rs("h_content")<>"" then%><%=server.htmlencode(replace(replace(rs("h_content"),"<br>",chr(10)),"<BR>",chr(10)))%><%end if%></textarea>
</tr>
<tr><td width=40>&nbsp;</td><td width=*><input type=submit class=button name=submit value="保存修改"></td></tr>
</table>
</FORM>
<%
end if
%>
</td>
</tr>
</table>
<%
end function

function del()
	Dvbbs.Execute("delete from dv_help where (not h_id=1) and h_id="&Request("id"))
	dv_suc("删除论坛帮助成功！")
end function
%>