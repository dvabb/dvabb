<!--#include file =../conn.asp-->
<!-- #include file="inc/const.asp" -->
<%
Head()
Dim admin_flag
Dim action
Dim sqlstr,l_type
l_type=request("l_type")
admin_flag=",3,"
CheckAdmin(admin_flag)
If Request("action")="dellog" Then
	batch()
Else 
	Select Case l_type
		Case "3"
			sqlstr=" where l_type=3 "
			l_type=3
			main
		Case "4"
			sqlstr=" where l_type=4 "
			l_type=4
			main
		Case "5"
			sqlstr=" where l_type=5 "
			l_type=5
			main
		Case "6"
			sqlstr=" where l_type=6 "
			l_type=6
			main
		Case "0"
			sqlstr=" where l_type=0 "
			l_type=0
			main
		Case "1"
			sqlstr=" where l_type=1 "
			l_type=1
			main
		Case "2"
			sqlstr=" where l_type=2 "
			l_type=2
			main
		Case Else
			sqlstr=""
			l_type=""
			main
	End Select
End If
If founderr then call dvbbs_error()
footer()

Sub main()
Dim l_boardID
l_boardID=Request("l_boardID")
If l_boardID="" Then l_boardID="0"
If l_boardID<> 0 Then
	If sqlstr <> "" Then
		sqlstr=sqlstr &" and l_boardID="&l_boardID
	Else
		sqlstr=" where l_boardID="&l_boardID
	End If
End If
Dim keyword,checkvalue
checkvalue=Dvbbs.Checkstr(Request("checkvalue"))
keyword=Dvbbs.checkstr(Request("keyword"))
If keyword <> "" Then
	If checkvalue="" Then
		If sqlstr <> "" Then
			sqlstr=sqlstr &" and (l_touser like '%"&keyword&"%' Or l_content like '%"&keyword&"%' Or l_ip like '%"&keyword&"%' Or l_username like '%"&keyword&"%')"
		Else
			sqlstr=" where l_touser like '%"&keyword&"%' Or l_content like '%"&keyword&"%' Or l_ip like '%"&keyword&"%' Or l_username like '%"&keyword&"%'"
		End If
	ElseIf checkvalue = "l_touser" Or checkvalue = "l_content" Or checkvalue = "l_ip" Or checkvalue = "l_username" Then
		If sqlstr <> "" Then
			sqlstr=sqlstr &" and "& checkvalue &" like '%"&keyword&"%'"
		Else
			sqlstr=" where "& checkvalue &" like '%"&keyword&"%'"
		End If
	End If
End If
%>
<form action=log.asp method=post">
<table cellPadding=1 cellSpacing=1 align=center width=100%>
<tr><th style="text-align:center;">论坛日志查看</th></tr>
<tr><td width="*">
&nbsp;搜 索 范 围：<Select Name="l_boardid">
</Select>
&nbsp;日志类型：<Select name="l_type">
<Option value="" 
<%If Request("l_type")="" Then Response.Write " selected"%>>全部日志</Option>
<Option value="3"<%If Request("l_type")="3" Then Response.Write " selected"%>>贴子管理</Option>
<Option value="4"<%If Request("l_type")="4" Then Response.Write " selected"%>>固顶操作</Option>
<Option value="5"<%If Request("l_type")="5" Then Response.Write " selected"%>>奖惩操作</Option>
<Option value="6"<%If Request("l_type")="6" Then Response.Write " selected"%>>用户处理</Option>
<Option value="0"<%If Request("l_type")="0" Then Response.Write " selected"%>>后台日志A</Option>
<Option value="1"<%If Request("l_type")="1" Then Response.Write " selected"%>>后台日志B</Option>
<Option value="2"<%If Request("l_type")="2" Then Response.Write " selected"%>>后台日志C</Option>
</Select>
</td>

</tr>
<tr><td class=td1 width="*">&nbsp;&nbsp;&nbsp;&nbsp;关 键 字：<Input Type="text" Name="keyword" value="<%=Request("keyword")%>" Size="28">
&nbsp;关键字匹配：<Select name="checkvalue">
<Option value="" <%If Request("checkvalue")="" Then Response.Write " selected"%>>全部列</Option>
<Option value="l_touser"<%If Request("checkvalue")="l_touser" Then Response.Write " selected"%>>操作对象</Option>
<Option value="l_content"<%If Request("checkvalue")="l_content" Then Response.Write " selected"%>>事件内容</Option>
<Option value="l_ip"<%If Request("checkvalue")="l_ip" Then Response.Write " selected"%>>IP地址</Option>
<Option value="l_username"<%If Request("checkvalue")="l_username" Then Response.Write " selected"%>>操作人</Option>
</Select>
</td>

</tr>
<tr><td width="100%" align=center><input type=submit value="搜 索" class="button">
</td>
</tr>
</table>
</form>
<SCRIPT LANGUAGE="JavaScript">
BoardJumpListSelect('<%=l_boardID%>',"l_boardid","所有论坛","",0)
</SCRIPT>
<br>
<%
Dim pagestr
Dim currentpage,page_count,Pcount,endpage
Dim sql,Rs,totalrec,i
currentPage=request("page")
If currentpage="" or not IsNumeric(currentpage) Then
	currentpage=1
Else
	currentpage=clng(currentpage)
End If
pagestr="?keyword="&Request("keyword")&"&l_type="& Request("l_type") &"&checkvalue="&Request("checkvalue") &"&l_boardID=" &Request("l_boardID")&"&"
Dvbbs.Forum_Setting(11)=50
sql="select * from [dv_log] "&sqlstr&" order by l_addtime desc"
'Response.Write SQL

set rs=Dvbbs.iCreateObject("adodb.recordset")
rs.open sql,conn,1,1

Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"" style=""word-break:break-all"" >"
Response.Write "<form action=log.asp?action=dellog&l_type="&l_type&" method=post name=even>"
Response.Write "<tr align=center>"
Response.Write "<th width=""10%"">"
Response.Write "对象"
Response.Write "</td>"
Response.Write "<th width=""45%"">"
Response.Write "事件内容"
Response.Write "</td>"
Response.Write "<th width=""15%"">"
Response.Write "操作时间"
Response.Write "</td>"
Response.Write "<th width=""15%"">"
Response.Write "IP"
Response.Write "</td>"
Response.Write "<th width=""10%"">"
Response.Write "操作人"
Response.Write "</td>"
Response.Write "<th width=""5%"">"
Response.Write "操作"
Response.Write "</th>"
Response.Write "</tr>"
If Not(Rs.eof or Rs.bof) Then
	rs.PageSize = Dvbbs.Forum_Setting(11)
	rs.AbsolutePage=currentpage
	page_count=0
    	totalrec=rs.recordcount
	While (Not Rs.EOF) And (Not page_count = Rs.PageSize)
	Response.Write "<tr align=left>"
	Response.Write "<td class=""td1"" width=""10%"">"
	Response.Write "<a href=../dispuser.asp?name="
	Response.Write Dvbbs.HTMLEncode(rs("l_touser"))
	Response.Write " target=_blank>"
	Response.Write Dvbbs.HTMLEncode(rs("l_touser"))
	Response.Write "</a>"
	Response.Write "</td>"
	Response.Write "<td class=""td1"" width=""45%"">"
	Response.Write HighLigth(Dvbbs.HTMLEncode(URLDecode(Rs("l_content"))),keyword)
	Response.Write "</td>"
	Response.Write "<td class=""td1"" width=""15%"">"
	Response.Write rs("l_addtime")
	Response.Write "</td>"
	Response.Write "<td class=""td1"" width=""15%"">"
	Response.Write Rs("l_ip")
	Response.Write "</td>"
	Response.Write "<td class=""td1"" width=""10%"">"
	Response.Write "<a href=../dispuser.asp?name="&Dvbbs.HTMLEncode(rs("l_username"))&" target=_blank>"&Dvbbs.HTMLEncode(rs("l_username"))&"</a>"
	Response.Write "&nbsp;</td>"
	Response.Write "<td class=""td1"" width=""5%"">"
	If Rs("l_type")<>2 Then
		Response.Write  "<input type=checkbox class=checkbox name=lid value="&rs("l_id")&">"
	End If
	Response.Write "</td>"
	Response.Write "</tr>"
'	Response.Write "<tr>"
'	Response.Write "<td height=2></td></tr>"
	
	page_count = page_count + 1
	Rs.MoveNext
	Wend
	Response.Write "<tr><td class=td2 colspan=6>请选择要删除的事件，<input type=checkbox class=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">全选 <input type=submit name=act value=删除 class=button onclick=""{if(confirm('您确定执行的操作吗?')){this.document.even.submit();return true;}return false;}"">"
	Response.Write "　<input type=submit class=button name=act onclick=""{if(confirm('确定清除所有的记录吗?')){this.document.even.submit();return true;}return false;}"" value=清空日志></td></tr>"
	If totalrec mod Dvbbs.Forum_Setting(11)=0 Then
		Pcount= totalrec \ Dvbbs.Forum_Setting(11)
  	Else
  		Pcount= totalrec \ Dvbbs.Forum_Setting(11)+1
  	End If
  	Response.Write "<table border=0 cellpadding=0 cellspacing=3 width=""100%"" align=center>"
  	Response.Write "<tr><td valign=middle nowrap>"
	Response.Write "页次：<b>"&currentpage&"</b>/<b>"&Pcount&"</b>页"
	Response.Write "&nbsp;每页<b>"&Dvbbs.Forum_Setting(11)&"</b> 总数<b>"&totalrec&"</b></td>"
	Response.Write "<td valign=middle nowrap align=right>分页："
	If currentpage > 4 Then
		Response.Write "<a href="""&pagestr&"page=1"">[1]</a> ..."
	End If
	If Pcount>currentpage+3 Then
		endpage=currentpage+3
	Else
		endpage=Pcount
	End If
	For i=currentpage-3 to endpage
	If Not i<1 Then
		If i = clng(currentpage) Then
			response.write " <font color="&Dvbbs.mainsetting(1)&">["&i&"]</font>"
		Else
			Response.Write " <a href="""&pagestr&"page="&i&""">["&i&"]</a>"
		End If
	End If
	Next
	If currentpage+3 < Pcount Then   
		Response.Write "... <a href="""&pagestr&"page="&Pcount&""">["&Pcount&"]</a>"
	End If
	Response.Write "</td></tr></table>"
Else
	Response.Write "<tr align=center>"
	Response.Write "<td class=""td1"" width=""100%"" colspan=""6"" >"
	Response.Write "无相关记录。"
	Response.Write "</td>"
	Response.Write "</tr>"
End If
Response.Write "</form>"
Response.Write "</table>"
Rs.close
Set rs=Nothing
End Sub

Sub batch()
	Dim lid
	If request.form("lid")="" Then
		Errmsg=ErrMsg + "请指定相关事件。"
		Founderr = True
		Exit Sub
	End If
	lid=replace(request.Form("lid"),"'","")
	lid=replace(lid,";","")
	lid=replace(lid,"--","")
	lid=replace(lid,")","")
	If request("act")="删除" Then
		Dvbbs.Execute("delete from dv_log where Datediff(""D"",l_addtime, "&SqlNowString&") > 2 and l_id in ("&lid&")")
	ElseIf request("act")="清空日志" Then
		If request("l_type")="" or IsNull(request("l_type")) Then 
			If IsSqlDataBase = 1 Then
			Dvbbs.Execute("delete from dv_log Where Datediff(D,l_addtime, "&SqlNowString&") > 2")
			else
			Dvbbs.Execute("delete from dv_log Where Datediff('D',l_addtime, "&SqlNowString&") > 2")
			end if
		Else
			If IsSqlDataBase = 1 Then
			Dvbbs.Execute("delete from dv_log where  Datediff(D,l_addtime, "&SqlNowString&") > 2 and l_type="&CInt(request("l_type"))&"")
			else
			Dvbbs.Execute("delete from dv_log where  Datediff('D',l_addtime, "&SqlNowString&") > 2 and l_type="&CInt(request("l_type"))&"")
			end if
		End If
	End If
	Dv_suc("成功删除日志。注意：两天内的日志会被系统保留。")
End Sub

'关键字突出显示 by 轻飘飘
Function HighLigth(Str,keyword)
	If keyword="" Then
		HighLigth=Str
		Exit Function
	End IF
	Dim re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	re.Pattern="("&keyword&")"
	HighLigth=re.Replace(Str,"<font color=""red"">$1</font>")
End Function

'URL解码函数 by 轻飘飘
Function URLDecode(enStr)
	On Error Resume Next
	Dim deStr,c,i,v:deStr=""
	For i=1 to len(enStr)
		c=Mid(enStr,i,1)
		If c="%" Then
			v=eval("&h"+Mid(enStr,i+1,2))
			If v<128 Then
				deStr=deStr&Chr(v)
				i=i+2
			Else
				If isvalidhex(Mid(enstr,i,3)) Then
					If isvalidhex(Mid(enstr,i+3,3)) Then	'这个判断检测是否双字节--不是
						v=eval("&h"+Mid(enStr,i+1,2)+Mid(enStr,i+4,2))
						deStr=deStr&Chr(v)
						i=i+5
					Else
						v=eval("&h"+Mid(enStr,i+1,2)+Cstr(Hex(Asc(Mid(enStr,i+3,1)))))	'--是
						deStr=deStr&Chr(v)
						i=i+3 
					End If
				Else 
					destr=destr&c
				End If
			End If
		Else
			If c="+" Then
				deStr=deStr&" "
			Else
				deStr=deStr&c
			End If
		End If
	Next
	URLDecode=deStr
End Function

Function IsValidHex(str)
	Dim c
	IsValidHex=True
	str=UCase(str)
	If Len(str)<>3 Then
		IsValidHex=False
		Exit Function
	End If
	If Left(str,1)<>"%" Then
		IsValidHex=False
		Exit Function
	End If
	c=Mid(str,2,1)
	If Not (((c>="0") And (c<="9")) Or ((c>="A") And (c<="Z"))) Then
		IsValidHex=False
		Exit Function
	End If
	c=Mid(str,3,1)
	If Not (((c>="0") And (c<="9")) Or ((c>="A") And (c<="Z"))) Then
		IsValidHex=False
		Exit Function
	End If
End Function
%>