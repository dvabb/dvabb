<!--#include file =../conn.asp-->
<!--#include file="inc/const.asp"-->
<!--#include file="../inc/dv_clsother.asp"-->
<!--#include file="../Dv_plus/Tools/plus_Tools_const.asp"-->
<%
Head()
Dim admin_flag
admin_flag=",41,"
CheckAdmin(admin_flag)
Main_head()

Select Case Request("action")
	Case "paylist" : UserList(2)
	Case "paylist1" : UserList(3)
	Case "EditInfo" : EditInfo()
	Case "SaveEdit" : SaveEdit()
	Case "Not_Apply_Sale" : Not_Apply_Sale()
	Case "Del_UserTools" : Del_UserTools()
	Case "SendTools" : SendTools()
	Case "SaveSend" : SaveSend()
	Case Else
		UserList(1)
End Select
If founderr then call dvbbs_error()
footer()

'顶部标题
Sub Main_head()
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr><th style="text-align:center;">道具中心用户信息管理</th></tr>
<tr><td class="td2"><B>说明</B>：<BR>
1、列表中的金币和点券指用户的总拥有量，而不是道具价格，查看相关用户道具信息请点击用户名进入<BR>
2、关键字搜索内容：用户名模糊搜索
</td></tr>
<tr><td class="td1"><B>操作</B>：<a href="plus_Tools_User.asp">用户道具管理首页</a> | <a href="?action=paylist">出售中的用户道具</a> <!-- | <a href="?action=paylist1">低于系统售价的出售道具</a> --> | <a href="?action=SendTools">给用户添加道具或金币</a>
</td></tr>
</table><br>
<%
End Sub

Sub UserList(ListType)
	Dim ToolsType,KeyWord,DataDesc,DataDescStr,UpDesc
	Dim ToolsValues_1,SearchStr
	Dim Rs,Sql,i
	Dim Page,MaxRows,Endpage,CountNum,PageSearch,SqlString
	Endpage = 0
	MaxRows = 20
	Page = Request("Page")
	ToolsType = Request("ToolsType")
	KeyWord = Request("keyword")
	DataDesc = Request("DataDesc")
	UpDesc = Request("UpDesc")
	If IsNumeric(Page) = 0 or Page="" Then Page=1
	Page = Clng(Page)
	Response.Write "<script language=""JavaScript"" src=""../inc/Pagination.js""></script>"
	If Not IsNumeric(ToolsType) Or ToolsType = "" Then ToolsType = 0
	ToolsType = Clng(ToolsType)
	If KeyWord <> "" Then KeyWord = Replace(KeyWord,"'","''")
	If ToolsType > 0 Then SearchStr = " ToolsID="&ToolsType&" "
	If Not IsNumeric(DataDesc) Or DataDesc = "" Then DataDesc = 0
	DataDesc = Clng(DataDesc)

	PageSearch = "action="&Request("action")&"&keyword="&KeyWord&"&ToolsType="&ToolsType&"&DataDesc="&DataDesc&"&UpDesc="&UpDesc
	
	Select Case DataDesc
	Case 1
		DataDescStr = " SaleMoney"
		ListType = 2
	Case 2
		DataDescStr = " SaleTicket"
		ListType = 2
	Case 3
		DataDescStr = " ToolsCount"
	Case 4
		DataDescStr = " SaleCount"
		ListType = 2
	Case Else
		DataDescStr = " ID"
	End Select

	If UpDesc="1" Then
		DataDescStr = DataDescStr & " Desc"
	End If

	If ListType > 1 Then ToolsValues_1 = "价格"

	If ListType = 3 Then
		If SearchStr <> "" Then
			SearchStr = SearchStr & " And (Not SaleCount=0)"
		Else
			SearchStr = " (Not SaleCount=0)"
		End If
	ElseIf ListType = 2 Then
		If SearchStr <> "" Then
			SearchStr = SearchStr & " And (Not SaleCount=0)"
		Else
			SearchStr = " (Not SaleCount=0)"
		End If
	End If
	If SearchStr<>"" Then
		If KeyWord<>"" Then SearchStr = SearchStr & " And UserName Like '%"&KeyWord&"%'"
	Else
		If KeyWord<>"" Then SearchStr = SearchStr & " UserName Like '%"&KeyWord&"%'"
	End If
	If SearchStr <> "" Then SearchStr = "Where " & SearchStr
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr>
<FORM METHOD=POST ACTION="plus_Tools_User.asp">
<td class="td2" colspan=5>
用户名 : <input type=text size=15 name="keyword">
<input type=hidden value="<%=ToolsType%>" name="ToolsType">
<input type=hidden value="<%=DataDesc%>" name="DataDesc">
<input type=hidden value="<%=Request("action")%>" name="action">
<input type=submit class="button" name=submit value="搜索">
</td>
</FORM>
<FORM METHOD=POST ACTION="plus_Tools_User.asp">
<input type=hidden value="<%=ToolsType%>" name="ToolsType">
<input type=hidden value="<%=Request("action")%>" name="action">
<td class="td2" colspan=2 align=right>
<Select Name="DataDesc" Size=1 onchange='javascript:submit()'>
<Option value="0" Selected>列表排序</Option>
<Option value="1" <%If DataDesc = 1 Then Response.Write "Selected"%>>出售金币数</Option>
<Option value="2" <%If DataDesc = 2 Then Response.Write "Selected"%>>出售点券数</Option>
<Option value="3" <%If DataDesc = 3 Then Response.Write "Selected"%>>拥有数</Option>
<Option value="4" <%If DataDesc = 4 Then Response.Write "Selected"%>>出售数</Option>
</Select>
<Select Name="UpDesc" Size=1 onchange='javascript:submit()'>
<Option value="0" <%If UpDesc = "0" Then Response.Write "Selected"%>>升序</Option>
<Option value="1" <%If UpDesc = "1" Then Response.Write "Selected"%>>降序</Option>
</Select>
</td>
</FORM>
<FORM METHOD=POST ACTION="plus_Tools_User.asp">
<input type=hidden value="<%=Request("action")%>" name="action">
<input type=hidden value="<%=DataDesc%>" name="DataDesc">
<td class="td2" align=right>
<Select Name="ToolsType" size=1 onchange='javascript:submit()'>
<Option value="0" Selected>拥有某道具用户列表</option>
<%
	SQL = "Select ID,ToolsName From Dv_Plus_Tools_Info Order By ID"
	Set Rs = Dvbbs.iCreateObject ("adodb.recordset")
	If Cint(Dvbbs.Forum_Setting(92))=1 Then
		If Not IsObject(Plus_Conn) Then Plus_ConnectionDatabase
		Rs.Open Sql,Plus_Conn,1,1
	Else
		If Not IsObject(Conn) Then ConnectionDatabase
		Rs.Open Sql,conn,1,1
	End If
	Do While Not Rs.Eof
		Response.Write "<Option value="""&Rs(0)&""" "
		If ToolsType=Rs(0) Then Response.Write "Selected"
		Response.Write ">"&Rs(1)&"</Option>"
	Rs.MoveNext
	Loop
	Rs.Close
	Set Rs=Nothing
%>
</Select>
</td>
</FORM>
</tr>
<tr>
<th width="20%">用户名</th>
<th width="8%">金币<%=ToolsValues_1%></th>
<th width="8%">点券<%=ToolsValues_1%></th>
<th width="15%">道具名</th>
<th width="7%">拥有数</th>
<th width="7%">出售数</th>
<th width="15%">操作时间</th>
<th width="21%">最后登录</th>
</tr>
<%
	'T.ToolsName=0,T.ToolsCount=1,T.SaleCount=2,U.UserName=3,U.UserMoney=4,U.UserTicket=5,U.UserLastIP=6,U.LastLogin=7,T.ID=8
	SQL = "Select ToolsName,ToolsCount,SaleCount,UserName,SaleMoney,SaleTicket,UpdateTime,UpdateTime,ID From Dv_Plus_Tools_Buss "&SearchStr&" Order By " & DataDescStr

	'Response.Write SQL
	Set Rs = Dvbbs.iCreateObject ("adodb.recordset")
	If Cint(Dvbbs.Forum_Setting(92))=1 Then
		If Not IsObject(Plus_Conn) Then Plus_ConnectionDatabase
		Rs.Open Sql,Plus_Conn,1,1
	Else
		If Not IsObject(Conn) Then ConnectionDatabase
		Rs.Open Sql,conn,1,1
	End If

	If Not (Rs.Eof And Rs.Bof) Then
		CountNum = Rs.RecordCount
		If CountNum Mod MaxRows=0 Then
			Endpage = CountNum \ MaxRows
		Else
			Endpage = CountNum \ MaxRows+1
		End If
		Rs.MoveFirst
		If Page > Endpage Then Page = Endpage
		If Page < 1 Then Page = 1
		If Page >1 Then 				
			Rs.Move (Page-1) * MaxRows
		End if
		SQL=Rs.GetRows(MaxRows)
	Else
		Response.Write "<tr><td class=""td1"" colspan=""8"" align=center>还未有信息！</td></tr></table>"
		Exit Sub
	End If
	Rs.close:Set Rs = Nothing
	For i=0 To Ubound(SQL,2)
%>
<tr>
<td class="td1" align=center><a href="?action=EditInfo&ID=<%=SQL(8,i)%>"><%=SQL(3,i)%></a></td>
<td class="td1" align=center><%=SQL(4,i)%></td>
<td class="td1" align=center><%=SQL(5,i)%></td>
<td class="td1" align=center><%=SQL(0,i)%></td>
<td class="td1" align=center><%=SQL(1,i)+SQL(2,i)%></td>
<td class="td1" align=center><%=SQL(2,i)%></td>
<td class="td1" align=center><%=SQL(6,i)%></td>
<td class="td1" align=center><%=SQL(7,i)%></td>
</tr>
<%
	Next

	PageSearch=Replace(Replace(PageSearch,"\","\\"),"""","\""")
	Response.Write "<SCRIPT>PageList("&Page&",3,"&MaxRows&","&CountNum&","""&PageSearch&""",1);</SCRIPT>"
End Sub

Sub EditInfo()
	Dim ID,rs
	ID=Request("ID")
	If Not IsNumeric(ID) Or ID="" Then
		Errmsg=ErrMsg + "<BR><li>非法的参数。"
		founderr=True
	Else
		ID=Clng(ID)
	End If
	If founderr Then Exit Sub
	Set Rs=Dvbbs.Plus_Execute("Select * From Dv_Plus_Tools_Buss Where ID=" & ID)

	If Rs.Eof And Rs.Bof Then
		Errmsg=ErrMsg + "<BR><li>没有找到相关的数据。"
		founderr=True
	End If
	If founderr Then Exit Sub
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<FORM METHOD=POST ACTION="?action=SaveEdit">
<input type=hidden value="<%=ID%>" name="ID">
<tr><th style="text-align:center;" colspan=2><%=Rs("UserName")%> 的论坛道具 <%=Rs("ToolsName")%> 管理</th></tr>
<tr><td class="td1" colspan=2><B>操作</B>：
<a href="plus_Tools_Info.asp?action=Editinfo&EditID=<%=Rs("ToolsID")%>">查看该道具资料</a> | <a href="?action=Not_Apply_Sale&ID=<%=Rs("ID")%>">取消该道具销售状态</a> | <a href="?action=Del_UserTools&ID=<%=Rs("ID")%>" onclick="{if(confirm('删除此用户该道具信息，操作将不可恢复，确定吗?')){return true;}return false;}">删除此用户该道具信息</a>
</td></tr>
<tr>
<td class="td1" width="20%" align=right><B>道具名</B>：</td>
<td class="td1" width="80%"><%=Rs("ToolsName")%></td>
</tr>
<tr>
<td class="td1" width="20%" align=right><B>所属用户</B>：</td>
<td class="td1" width="80%"><%=Rs("UserName")%></td>
</tr>
<tr>
<td class="td1" width="20%" align=right><B>最后操作时间</B>：</td>
<td class="td1" width="80%"><%=Rs("UpdateTime")%></td>
</tr>
<tr>
<td class="td1" width="20%" align=right><B>可使用数量</B>：</td>
<td class="td1" width="80%"><input type=text size=10 value="<%=Rs("ToolsCount")%>" name="ToolsCount"> 个</td>
</tr>
<tr>
<td class="td1" width="20%" align=right><B>出售中数量</B>：</td>
<td class="td1" width="80%"><input type=text size=10 value="<%=Rs("SaleCount")%>" name="SaleCount"> 个</td>
</tr>
<tr>
<td class="td1" width="20%" align=right><B>出售价格（金币）</B>：</td>
<td class="td1" width="80%"><input type=text size=10 value="<%=Rs("SaleMoney")%>" name="SaleMoney"> 个</td>
</tr>
<tr>
<td class="td1" width="20%" align=right><B>出售价格（点券）</B>：</td>
<td class="td1" width="80%"><input type=text size=10 value="<%=Rs("SaleTicket")%>" name="SaleTicket"> 张</td>
</tr>
<tr><td class="td2" colspan=2 align=center>
<input type=submit class="button" value="保存修改" name=submit>
此部分关系用户个人财产，修改请慎重，所做修改将进道具日志
</td></tr>
</FORM>
</table>
<%
	Rs.Close
	Set Rs=Nothing
End Sub

Sub SaveEdit()
	Dim ID
	Dim ToolsCount,SaleCount,SaleMoney,SaleTicket
	ID=Request("ID")
	If Not IsNumeric(ID) Or ID="" Then
		Errmsg=ErrMsg + "<BR><li>非法的参数。"
		founderr=True
	Else
		ID=Clng(ID)
	End If
	ToolsCount=Request("ToolsCount")
	If Not IsNumeric(ToolsCount) Or ToolsCount="" Then
		Errmsg=ErrMsg + "<BR><li>非法的参数。"
		founderr=True
	Else
		ToolsCount=Clng(ToolsCount)
	End If
	SaleCount=Request("SaleCount")
	If Not IsNumeric(SaleCount) Or SaleCount="" Then
		Errmsg=ErrMsg + "<BR><li>非法的参数。"
		founderr=True
	Else
		SaleCount=Clng(SaleCount)
	End If
	SaleMoney=Request("SaleMoney")
	If Not IsNumeric(SaleMoney) Or SaleMoney="" Then
		Errmsg=ErrMsg + "<BR><li>非法的参数。"
		founderr=True
	Else
		SaleMoney=Clng(SaleMoney)
	End If
	SaleTicket=Request("SaleTicket")
	If Not IsNumeric(SaleTicket) Or SaleTicket="" Then
		Errmsg=ErrMsg + "<BR><li>非法的参数。"
		founderr=True
	Else
		SaleTicket=Clng(SaleTicket)
	End If
	If founderr Then Exit Sub
	Dim Rs,SQL,Trs
	Set Rs = Dvbbs.iCreateObject ("adodb.recordset")
	SQL = "Select * From Dv_Plus_Tools_Buss Where ID=" & ID
	If Cint(Dvbbs.Forum_Setting(92))=1 Then
		If Not IsObject(Plus_Conn) Then Plus_ConnectionDatabase
		Rs.Open Sql,Plus_Conn,1,3
	Else
		If Not IsObject(Conn) Then ConnectionDatabase
		Rs.Open Sql,Conn,1,3
	End If
	If Rs.Eof And Rs.Bof Then
		Errmsg=ErrMsg + "<BR><li>没有找到相关的数据。"
		founderr=True
	End If
	If founderr Then Exit Sub
	'ToolsCount,SaleCount,SaleMoney,SaleTicket
	Set Trs=Dvbbs.Execute("Select UserMoney,UserTicket From Dv_User Where UserID=" & Rs("UserID"))
	Sql = "Insert into [Dv_MoneyLog] (ToolsID,AddUserName,AddUserID,Log_IP,Log_Type,BoardID,Conect,HMoney) values ("&Rs("ToolsID")&",'"&Rs("UserName")&"','"&Rs("UserID")&"','"&Dvbbs.UserTrueIP&"',0,-1,'"&Dvbbs.Membername&"管理员编辑用户道具资料，可使用数量变动<B>"&ToolsCount-Rs("ToolsCount")&"</B>、出售数量变动<B>"&SaleCount-Rs("SaleCount")&"</B>、出售价格变动（金币）<B>"&SaleMoney-Rs("SaleMoney")&"</B>、出售价格变动（点券）<B>"&SaleTicket-Rs("SaleTicket")&"</B>','"&Trs("UserMoney")&"|"&Trs("UserTicket")&"')"
	Dvbbs.Plus_Execute(SQL)
	Rs("ToolsCount")=ToolsCount
	Rs("SaleCount")=SaleCount
	Rs("SaleMoney")=SaleMoney
	Rs("SaleTicket")=SaleTicket
	Rs.UpDate
	Rs.Close
	Set Rs=Nothing
	Trs.Close
	Set Trs=Nothing
	Dv_Suc("修改用户道具资料成功！")
	Footer()
	Response.End
End Sub

Sub Not_Apply_Sale()
	Dim ID,SQL,Trs,Rs
	ID=Request("ID")
	If Not IsNumeric(ID) Or ID="" Then
		Errmsg=ErrMsg + "<BR><li>非法的参数。"
		founderr=True
	Else
		ID=Clng(ID)
	End If
	If founderr Then Exit Sub
	SQL = "UpDate Dv_Plus_Tools_Buss Set ToolsCount=ToolsCount + SaleCount,SaleMoney=0,SaleTicket=0,SaleCount=0 Where ID=" & ID
	Dvbbs.Plus_Execute(SQL)
	Set Rs=Dvbbs.Plus_Execute("Select * From Dv_Plus_Tools_Buss Where ID=" & ID)
	Set Trs=Dvbbs.Execute("Select UserMoney,UserTicket From Dv_User Where UserID=" & Rs("UserID"))
	Sql = "Insert into [Dv_MoneyLog] (ToolsID,AddUserName,AddUserID,Log_IP,Log_Type,BoardID,Conect,HMoney) values ("&Rs("ToolsID")&",'"&Rs("UserName")&"','"&Rs("UserID")&"','"&Dvbbs.UserTrueIP&"',0,-1,'"&Dvbbs.Membername&"管理员编辑用户道具资料，取消此用户该道具销售资格。','"&Trs("UserMoney")&"|"&Trs("UserTicket")&"')"
	Dvbbs.Plus_Execute(SQL)
	Rs.Close
	Set Rs=Nothing
	Trs.Close
	Set Trs=Nothing
	Dv_Suc("修改用户道具资料成功！")
	Footer()
	Response.End
End Sub

Sub Del_UserTools()
	Dim ID,SQL,Trs,Rs
	ID=Request("ID")
	If Not IsNumeric(ID) Or ID="" Then
		Errmsg=ErrMsg + "<BR><li>非法的参数。"
		founderr=True
	Else
		ID=Clng(ID)
	End If
	If founderr Then Exit Sub
	Set Rs=Dvbbs.Plus_Execute("Select * From Dv_Plus_Tools_Buss Where ID=" & ID)
	Set Trs=Dvbbs.Execute("Select UserMoney,UserTicket From Dv_User Where UserID=" & Rs("UserID"))
	Sql = "Insert into [Dv_MoneyLog] (ToolsID,AddUserName,AddUserID,Log_IP,Log_Type,BoardID,Conect,HMoney) values ("&Rs("ToolsID")&",'"&Rs("UserName")&"','"&Rs("UserID")&"','"&Dvbbs.UserTrueIP&"',0,-1,'"&Dvbbs.Membername&"管理员编辑用户道具资料，删除此用户该道具信息。','"&Trs("UserMoney")&"|"&Trs("UserTicket")&"')"
	Dvbbs.Plus_Execute(SQL)
	Dvbbs.Plus_Execute("Delete From Dv_Plus_Tools_Buss Where ID=" & ID)
	Rs.Close
	Set Rs=Nothing
	Trs.Close
	Set Trs=Nothing
	Dv_Suc("删除用户道具资料成功！")
	Footer()
	Response.End
End Sub

Sub SendTools()
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<FORM METHOD=POST ACTION="?action=SaveSend">
<tr><th style="text-align:center;" colspan=2>给用户赠送道具</th></tr>
<tr><td height=24 class="td1" colspan=2>说明：部分系统不出售的内部道具请慎重赠送</td></tr>
<tr>
<td class="td1" width="30%" align=right>
<B>道具名称</B>：
</td>
<td class="td1" width="70%">
<Select Size=1 Name="ToolsID">
<Option value="0" Selected>请选择您要赠送的道具</Option>
<%
	Dim rs
	Set Rs=Dvbbs.Plus_Execute("Select * From Dv_Plus_Tools_Info Order By ID")
	Do While Not Rs.Eof
		Response.Write "<Option value="""&Rs("ID")&""">"&Rs("ToolsName")&"</Option>"
	Rs.MoveNext
	Loop
	Rs.Close
	Set Rs=Nothing
%>
</Select>
</td>
</tr>
<tr>
<td class="td1" width="30%" align=right>
<B>目标用户</B>：
</td>
<td class="td1" width="70%">
<input type=text size=20 name="SendName">
</td>
</tr>
<tr>
<td class="td1" width="30%" align=right>
<B>赠送数量</B>：
</td>
<td class="td1" width="70%">
<input type=text size=10 value=0 name="ToolsNum">
</td>
</tr>
<tr>
<td class="td1" width="30%" align=right>
<B>同时送出金币</B>：
</td>
<td class="td1" width="70%">
<input type=text size=10 value=0 name="ToolsMoney">
</td>
</tr>
<tr><td height=24 class="td1" colspan=2 align=center><input type=submit class="button" name=submit value="送出道具或金币"></td></tr>
</form>
</table>
<%
End Sub

Sub SaveSend()
	Dim ToolsID,ToolsNum,SendName,ToolsName,ToolsMoney
	Dim Trs,rs,sql
	ToolsID = Dv_Tools.CheckNumeric(Request("ToolsID"))
	ToolsNum = Dv_Tools.CheckNumeric(Request("ToolsNum"))
	ToolsMoney = Dv_Tools.CheckNumeric(Request("ToolsMoney"))
	SendName = Request("SendName")
	If (ToolsID = 0 Or ToolsNum = 0 Or SendName = "") And ToolsMoney = 0 Then
		Errmsg=ErrMsg + "<BR><li>非法的道具或数量或用户名参数。"
		founderr=True
	End If
	If founderr Then Exit Sub
	SendName = Replace(SendName,"'","''")
	If SendName<>"" And ToolsID>0 And ToolsNum>0 Then
	Set Rs=Dvbbs.Plus_Execute("Select ToolsName From Dv_Plus_Tools_Info Where ID="&ToolsID)
	If Rs.Eof And Rs.Bof Then
		Errmsg=ErrMsg + "<BR><li>您要赠送的系统道具并不存在。"
		Exit Sub
	Else
		ToolsName = Rs(0)
	End If
	Set Rs=Dvbbs.Execute("Select UserID,UserName,UserMoney,UserTicket From Dv_User Where UserName='"&SendName&"'")
	If Rs.Eof And Rs.Bof Then
		Errmsg=ErrMsg + "<BR><li>您要赠送的目标用户名并不存在。"
		founderr=True
		Exit Sub
	Else
		'更新用户道具记录
		Set Trs=Dvbbs.Plus_Execute("Select ID From [Dv_Plus_Tools_Buss] Where UserID="& Rs(0) &" and ToolsID="& ToolsID)
		If Trs.Eof And Trs.Bof Then
			Sql = "Insert Into [Dv_Plus_Tools_Buss] (UserID,UserName,ToolsID,ToolsName,ToolsCount) Values ("&Rs(0)&",'"&SendName&"',"&ToolsID&",'"&ToolsName&"',"&ToolsNum&")"
			Dvbbs.Plus_Execute(Sql)
		Else
			Sql = "Update [Dv_Plus_Tools_Buss] Set ToolsCount = ToolsCount+"&ToolsNum&" Where UserID="& Rs(0) &" and ToolsID="& ToolsID
			Dvbbs.Plus_Execute(Sql)
		End If
		'插入论坛日志
		Dvbbs.Plus_Execute("Insert into [Dv_MoneyLog] (ToolsID,AddUserName,AddUserID,Log_IP,Log_Type,BoardID,Conect,HMoney) values ("&ToolsID&",'"&Rs("UserName")&"','"&Rs("UserID")&"','"&Dvbbs.UserTrueIP&"',0,-1,'"&Dvbbs.Membername&"管理员赠送<B>"&ToolsNum&"</B>个<B>"&ToolsName&"</B>道具给<B>"&Rs("UserName")&"</B>用户。','"&Rs("UserMoney")&"|"&Rs("UserTicket")&"')")
	End If
	Rs.Close
	Set Rs=Nothing
	Sql = "Update [Dv_Plus_Tools_Info] Set UserStock =  UserStock+"&ToolsNum&" Where ID="& ToolsID
	Dvbbs.Plus_Execute(Sql)
	Trs.Close
	Set Trs=Nothing
	End If
	If ToolsMoney<>0 And SendName<>"" Then
		Set Rs=Dvbbs.Execute("Select UserID,UserName,UserMoney,UserTicket From Dv_User Where UserName='"&SendName&"'")
		If Rs.Eof And Rs.Bof Then
			Errmsg=ErrMsg + "<BR><li>您要赠送的目标用户名并不存在。"
			founderr=True
			Exit Sub
		Else
			Dvbbs.Execute("Update Dv_User Set UserMoney = UserMoney + "&ToolsMoney&" Where UserID="&Rs(0))
			'插入论坛日志
			Dvbbs.Plus_Execute("Insert into [Dv_MoneyLog] (ToolsID,AddUserName,AddUserID,Log_IP,Log_Type,BoardID,Conect,HMoney) values ("&ToolsID&",'"&Rs("UserName")&"','"&Rs("UserID")&"','"&Dvbbs.UserTrueIP&"',0,-1,'"&Dvbbs.Membername&"管理员赠送<B>"&ToolsMoney&"</B>个金币给<B>"&Rs("UserName")&"</B>用户。','"&Rs("UserMoney")+ToolsMoney&"|"&Rs("UserTicket")&"')")
		End If
	End If
	Dv_Suc("赠送用户道具成功！")
	Footer()
	Response.End
End Sub
%>