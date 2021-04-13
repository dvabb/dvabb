<!--#include file =../conn.asp-->
<!--#include file="inc/const.asp"-->
<!--#include file="../inc/dv_clsother.asp"-->
<!--#include file="../Dv_plus/Tools/plus_Tools_const.asp"-->
<%
Head()
Dim admin_flag
If Trim(Request("action"))="Setting" Then Tools_Setting()
admin_flag=",40,"
CheckAdmin(admin_flag)
Main_head()
Select Case Trim(Request("action"))
	'Case "Setting" : Tools_Setting
	Case "AllStarSetting" : AllStarSetting
	Case "Editinfo" : AddTools
	Case "SaveTools" : SaveTools
	Case "List" : Tools_List()
	Case "UpdateUserStock" : UpdateUserStock
	Case Else
		Tools_List()
End Select
If founderr then call dvbbs_error()
footer()

'顶部标题
Sub Main_head()
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr><th>道具中心管理</th></tr>
<tr><td class="td2"><B>道具资料设置说明</B>：<BR>
	1、<font color=red>建议管理员定期添加库存，请按论坛市场实际需求进行调节，不建议频繁添加。</font>（通常在道具总拥有量非常少的情况下进行添加库存，防止会员趁机抬高价格。）
	<BR>
	2、<font color=blue>系统库存设置为-1则表示该道具为系统道具</font>，系统不出售，一般为论坛使用过程产生，如果用户得到可自己转让或出售</td></tr>
</table>
<br>
<%
End Sub

Sub Tools_Setting()
	Response.Redirect "setting.asp#settingxu"
End Sub

'道具列表
Sub Tools_List()
	Dim Orders
	Select case Trim(Request.QueryString("orders"))
		Case "0" : Orders = "SysStock"
		Case "1" : Orders = "UserMoney"
		Case "2" : Orders = "UserStock"
		Case "3" : Orders = "IsStar"
		Case Else : Orders = "SysStock"
	End Select
	%>
	<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
	<tr><th colspan="6" style="text-align:center;">道具资料列表</th></tr>
	<tr><td class=td1 colspan="6">
	[<a href="?action=UpdateUserStock"><font color=blue>更新用户库存</font></a>]
	</td></tr>
	<tr><td class="td1" colspan="6"><ol type="A">
	<li>点击道具名称进行修改其详细资料。</li>
	<li>点击带链接的标题栏目查看相应的排序。</li>
	</ol></td></tr>
	<tr>
		<th width="17%">名称</th>
		<th width="45%">说明</th>
		<th width="15%" id=TableTitleLink><A HREF="?orders=1" Title="按金币由少到多排序">价格</A></th>
		<th width="5%" id=TableTitleLink><A HREF="?orders=0" Title="按系统库存由少到多排序">库存</A></th>
		<th width="10%" id=TableTitleLink><A HREF="?orders=2" Title="按用户拥有库存由少到多排序">用户库存</A></th>
		<th width="8%" id=TableTitleLink><A HREF="?orders=3" Title="按关闭到开启排序">启用</A></th>
	</tr>
	<form action="?action=AllStarSetting" method=post> 
	<%
	Dim Rs,Sql,i
	Sql = "Select ID,ToolsName,ToolsInfo,IsStar,SysStock,UserStock,UserMoney,UserTicket,BuyType From [Dv_Plus_Tools_Info] ORDER BY "& Orders
	Set Rs = Dvbbs.Plus_Execute(Sql)
	If Not Rs.eof Then
		SQL=Rs.GetRows(-1)
	Else
		Response.Write "<tr><td class=""td2"" colspan=""6"" align=center>道具还未添加！</td></tr></form></table>"
		Exit Sub
	End If
	Rs.close:Set Rs = Nothing
	For i=0 To Ubound(SQL,2)
	%>
	<tr>
		<td class="td2"><a href="?action=Editinfo&EditID=<%=SQL(0,i)%>"><%=Server.Htmlencode(SQL(1,i))%></a></td>
		<td class="td2"><%=Server.Htmlencode(SQL(2,i)&"")%></td>
		<td class="td2">
		<%
		Select Case SQL(8,i)
		Case 0
			Response.Write SQL(6,i)
		Case 1
			Response.Write SQL(7,i)
		Case 2
			Response.Write SQL(6,i) & " And " & SQL(7,i)
		Case 3
			Response.Write SQL(6,i) & " Or " & SQL(7,i)
		End Select
		%>
		</td>
		<td class="td2" align=center>
		<%
		If SQL(4,i)="-1" Then
		Response.Write "系统"
		Else
		Response.Write SQL(4,i)
		End If
		%>
		</td>
		<td class="td2" align=center><%=SQL(5,i)%></td>
		<td class="td2" align=center><INPUT TYPE="checkbox" class="checkbox" NAME="Star" <%If SQL(3,i)="1" Then Response.Write "checked"%> value="<%=SQL(0,i)%>">
		</td>
	</tr>
	<%
	Next
	Response.Write "<tr><td class=""td1"" colspan=""6"" align=right><INPUT TYPE=""submit"" class=""button"" value=""保存修改设置""></td></tr></form></table>"
End Sub

'批量修改道具开启或关闭设置
Sub AllStarSetting()
	Dim EditID,Sql
	EditID = Trim(Request.Form("Star"))
	If CheckID(EditID) = False Then
		Response.Write "修改中止，请确认提交的参数是否正确后重新提交！<a href="&Request.ServerVariables("HTTP_REFERER")&" ><<返回上一页</a>"
		Exit Sub
	End If
	'先将启动设置还原为关闭状态
	Sql = "Update [Dv_Plus_Tools_Info] Set IsStar=0"
	Dvbbs.Plus_Execute(Sql)

	'开启所选项
	Sql = "Update [Dv_Plus_Tools_Info] Set IsStar=1 Where ID in (" & EditID & ")"
	Dvbbs.Plus_Execute(Sql)

	Dv_suc("批量修改道具开关设置成功！")
End Sub

'添加，修改道具信息
Sub AddTools()
Dim EditID,Rs,Sql
EditID = Trim(Request.QueryString("EditID"))
If EditID<>"" and IsNumeric(EditID) Then
	EditID = Cint(EditID)
Else
	Response.Write "修改中止，请确认提交的参数是否正确后重新提交！<a href="&Request.ServerVariables("HTTP_REFERER")&" ><<返回上一页</a>"
	Exit Sub
End If

'ID=0 ,ToolsName=1 ,ToolsInfo=2 ,IsStar=3 ,SysStock=4 ,UserStock=5 ,UserMoney=6 ,UserPost=7 ,UserWealth=8 ,UserEp=9 ,UserCp=10 ,UserGroupID=11 ,BoardID=12,UserTicket=13,BuyType=14,ToolsImg=15,ToolsSetting=16
Dim ToolsSetting
Sql = "Select ID,ToolsName,ToolsInfo,IsStar,SysStock,UserStock,UserMoney,UserPost,UserWealth,UserEp,UserCp,UserGroupID,BoardID,UserTicket,BuyType,ToolsImg,ToolsSetting From [Dv_Plus_Tools_Info] Where ID="& EditID

Set Rs = Dvbbs.Plus_Execute(Sql)
If Rs.Eof Then
	Response.Write "查找的道具数据不存在！<a href="&Request.ServerVariables("HTTP_REFERER")&" ><<返回上一页</a>"
	Exit Sub
Else
	Sql = Rs.GetString(,1, "§§§", "", "")
	Sql = Split(Sql,"§§§")
End If
Rs.Close
Set Rs = Nothing
ToolsSetting = Split(Sql(16),",")
If SQL(15)="" Then SQL(15)="Dv_plus/Tools/pic/None.jpg"
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<form action="?action=SaveTools" method=post name=PlusTools onsubmit="GetData();">
<input type=hidden name="EditID" value="<%=SQL(0)%>">
<tr><th colspan="2" style="text-align:center;">道具资料设置</th></tr>
<tr><td colspan="2" height=23 class="td1" align=center>
<img src="../<%=Server.Htmlencode(SQL(15))%>" border=0></td></tr>
<tr>
<td class="td2" width="20%" align=right><b>道具名称</b></td>
<td class="td2" width="80%"><INPUT TYPE="text" NAME="ToolsName" value="<%=Server.Htmlencode(SQL(1))%>" size=35> 不能超过50个字符！</td>
</tr>
<tr>
<td class="td2" align=right><b>道具简介</b></td>
<td class="td2">
<INPUT  NAME="ToolsInfo" TYPE="text" Style="width:80%" value="<%=Server.Htmlencode(SQL(2))%>"> 不能超过250个字符！</td>
</tr>
<tr>
<td class="td2" align=right><b>道具图标</b></td>
<td class="td2">
<INPUT  NAME="ToolsImg" TYPE="text" Style="width:80%" value="<%=Server.Htmlencode(SQL(15))%>"> 请正确填写图片路径</td>
</tr>
<tr>
<td class="td2" align=right><b>当前道具状态</b></td>
<td class="td2">
<input type=radio class="radio" name="IsStar" value=0 <%If Sql(3)="0" Then%>checked<%End If%>>关闭&nbsp;
<input type=radio class="radio" name="IsStar" value=1 <%If Sql(3)="1" Then%>checked<%End If%>>开启&nbsp;
</td>
</tr>
<tr><th colspan="2" style="text-align:center;">道具交易设置</th></tr>
<tr>
<td class="td2" align=right><b>金币价格</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsMoney" value="<%=SQL(6)%>" size=10></td>
</tr>
<tr>
<td class="td2" align=right><b>点券价格</b></td>
<td class="td2"><INPUT TYPE="text" NAME="UserTicket" value="<%=SQL(13)%>" size=10></td>
</tr>
<tr>
<td class="td2" align=right><b>当前系统库存</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsStock" value="<%=SQL(4)%>" size=10>&nbsp;当值为：-1，为系统道具，不允许交易。</td>
</tr>
<tr>
<td class="td2" align=right><b>购买方式</b></td>
<td class="td2">
<SELECT NAME="ToolsBuyType">
	<option value="0"<%If Sql(14)=0 Then%> Selected<%End If%>>只需金币
	<option value="1"<%If Sql(14)=1 Then%> Selected<%End If%>>只需点券
	<option value="2"<%If Sql(14)=2 Then%> Selected<%End If%>>金币+点券
	<option value="3"<%If Sql(14)=3 Then%> Selected<%End If%>>金币或点券
</option>
</SELECT>
</td>
</tr>

<tr><th colspan="2" style="text-align:center;">道具使用权限设置</th></tr>
<tr>
<td class="td2" align=right><b>用户帖子数限制</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsPost" value="<%=SQL(7)%>" size=10></td>
</tr>
<tr>
<td class="td2" align=right><b>用户金钱数限制</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsWealth" value="<%=SQL(8)%>" size=10></td>
</tr>
<tr>
<td class="td2" align=right><b>用户积分值限制</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsEP" value="<%=SQL(9)%>" size=10></td>
</tr>
<tr>
<td class="td2" align=right><b>用户魅力值限制</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsCP" value="<%=SQL(10)%>" size=10></td>
</tr>
<tr>
<td class="td2" align=right><b>使用目标用户帖子数限制</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsSetting(0)" value="<%=ToolsSetting(0)%>" size=10> 部分道具只有目标用户满足条件才能使用，下同</td>
</tr>
<tr>
<td class="td2" align=right><b>使用目标用户金钱数限制</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsSetting(1)" value="<%=ToolsSetting(1)%>" size=10></td>
</tr>
<tr>
<td class="td2" align=right><b>使用目标用户积分值限制</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsSetting(2)" value="<%=ToolsSetting(2)%>" size=10></td>
</tr>
<tr>
<td class="td2" align=right><b>使用目标用户魅力值限制</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsSetting(3)" value="<%=ToolsSetting(3)%>" size=10></td>
</tr>
<tr>
<td class="td2" align=right><b>道具使用反点奖励</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsSetting(4)" value="<%=ToolsSetting(4)%>" size=10><BR> 部分道具可设置，设置后使用该道具可获得一定金币的奖励，对多用户有效的道具每产生一笔交易则给使用用户一定金币的奖励，<font color=blue>在金币购买帖子相应道具中，此设置为百分比（如0.3），如购买此贴需要10个金币，则返回10*0.3的金币给发贴者</font>，<font color=red>在查税卡中为向目标用户征收的总金币数百分比</font></td>
</tr>
<tr>
<td class="td2" align=right><b>允许使用的用户组ID</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsGroupID" value="<%=Trim(SQL(11))%>" size=40 Disabled>
<INPUT TYPE="hidden" NAME="TrueToolsGroupID" value="<%=Trim(SQL(11))%>">
<input type="button" class="button" value="设置" onclick="PlusOpen('../plus_Tools_InfoSetting.asp?orders=0&id=<%=SQL(0)%>',650,500)"></td>
</tr>
<tr>
<td class="td2" align=right><b>允许使用的版块ID</b></td>
<td class="td2">
<INPUT TYPE="text" NAME="ToolsBoardID" value="<%=Trim(SQL(12))%>" size=40 Disabled>
<INPUT TYPE="hidden" NAME="TrueToolsBoardID" value="<%=Trim(SQL(12))%>">
<input type="button" class="button" value="设置" onclick="PlusOpen('../plus_Tools_InfoSetting.asp?orders=1&id=<%=SQL(0)%>',650,500)"></td>
</tr>
<tr><td class="td1" colspan="2" align=center><INPUT TYPE="submit" class="button" value="保存修改设置"></td></tr>
</form>
</table>
<SCRIPT LANGUAGE="JavaScript">
<!--
function GetData(){
PlusTools.TrueToolsGroupID.value=PlusTools.ToolsGroupID.value;
PlusTools.TrueToolsBoardID.value=PlusTools.ToolsBoardID.value;
}
//-->
</SCRIPT>
<%
End Sub

'保存添加，修改道具信息
Sub SaveTools()
	Dim EditID,ToolsName,ToolsInfo,ToolsImg
	Dim ToolsGroupID,ToolsBoardID
	Dim Rs,Sql,i
	Dim ToolsSetting,ChanSetting
	EditID = CheckNumeric(Request.Form("EditID"))
	ToolsName = Left(Trim(Request.Form("ToolsName")),50)
	ToolsInfo = Left(Trim(Request.Form("ToolsInfo")),250)
	ToolsImg = Left(Trim(Request.Form("ToolsImg")),150)
	ToolsGroupID = Trim(Replace(Request.Form("TrueToolsGroupID")," ",""))
	ToolsBoardID = Trim(Replace(Request.Form("TrueToolsBoardID")," ",""))
	If Right(ToolsGroupID,1)="," Then ToolsGroupID = Left(ToolsGroupID,Len(ToolsGroupID)-1)
	If Right(ToolsBoardID,1)="," Then ToolsBoardID = Left(ToolsBoardID,Len(ToolsBoardID)-1)
	If EditID = 0 Then
		Response.Write "更新的道具数据不存在！<a href="&Request.ServerVariables("HTTP_REFERER")&" ><<返回上一页</a>"
		Exit Sub
	End If
	If CheckID(ToolsGroupID)=False Then ToolsGroupID = ""
	If CheckID(ToolsBoardID)=False Then ToolsBoardID = ""

	If ToolsName="" or ToolsInfo="" Then
		Response.Write "修改中止，道具名称或简介不能为空！<a href="&Request.ServerVariables("HTTP_REFERER")&" ><<返回上一页</a>"
		Exit Sub
	Else
		ToolsName = Replace(ToolsName,"§§§","")
		ToolsInfo = Replace(ToolsInfo,"§§§","")
	End If
	For i=0 To 60
		If Request.Form("ToolsSetting("&i&")")="" Then
			ChanSetting = 0
		Else
			ChanSetting = Replace(Request.Form("ToolsSetting("&i&")"),",","")
		End If
		If i = 0 Then
			ToolsSetting = ChanSetting
		Else
			ToolsSetting = ToolsSetting & "," & ChanSetting
		End If
	Next

	Set Rs = Dvbbs.iCreateObject("adodb.recordset")
	Sql = "Select * From [Dv_Plus_Tools_Info] where ID="& EditID
	If Cint(Dvbbs.Forum_Setting(92))=1 Then
		If Not IsObject(Plus_Conn) Then Plus_ConnectionDatabase
		Rs.Open Sql,Plus_Conn,1,3
	Else
		If Not IsObject(Conn) Then ConnectionDatabase
		Rs.Open Sql,conn,1,3
	End IF
	If Rs.eof and Rs.bof then
		Response.Write "查找的道具数据不存在！<a href="&Request.ServerVariables("HTTP_REFERER")&" ><<返回上一页</a>"
		Exit Sub
	Else
		Rs("ToolsName") = ToolsName
		Rs("ToolsInfo") = ToolsInfo
		Rs("ToolsImg") = ToolsImg
		Rs("IsStar") = CheckNumeric(Request.Form("IsStar"))
		Rs("SysStock") = CheckNumeric(Request.Form("ToolsStock"))
		Rs("UserTicket") = CheckNumeric(Request.Form("UserTicket"))
		Rs("UserMoney") = CheckNumeric(Request.Form("ToolsMoney"))
		Rs("UserPost") = CheckNumeric(Request.Form("ToolsPost"))
		Rs("UserWealth") = CheckNumeric(Request.Form("ToolsWealth"))
		Rs("UserEp") = CheckNumeric(Request.Form("ToolsEP"))
		Rs("UserCp") = CheckNumeric(Request.Form("ToolsCP"))
		Rs("UserGroupID") = ToolsGroupID
		Rs("BoardID") = ToolsBoardID
		Rs("BuyType") = CheckNumeric(Request.Form("ToolsBuyType"))
		Rs("ToolsSetting") = ToolsSetting
		Rs.Update
	End If 
	Rs.Close
	Set Rs = Nothing
	Dvbbs.Plus_Execute("UPDATE [Dv_Plus_Tools_Buss] Set ToolsName ='"& Dvbbs.Checkstr(ToolsName) &"' where ToolsID="&EditID)
	Dv_suc(ToolsName&"道具开关设置成功！")
End Sub

'删除道具信息
Sub DllTools()

End Sub

'更新道具的用户拥有库存
Sub UpdateUserStock()
	Dim Rs,Sql,Totals
	If IsSqlDataBase = 1 Then
		Sql = "Update [Dv_Plus_Tools_Info] Set UserStock = (Select Count(*) From [Dv_Plus_Tools_Buss] where ToolsID=Dv_Plus_Tools_Info.ID)"
		Dvbbs.Plus_Execute(Sql)
	Else
		Sql = "Select ID From [Dv_Plus_Tools_Info]"
		Set Rs = Dvbbs.Plus_Execute(Sql)
		Do while Not Rs.eof
			Totals = Dvbbs.Plus_Execute("Select Count(*) From [Dv_Plus_Tools_Buss] where ToolsID="&Rs(0))(0)
			Dvbbs.Plus_Execute("Update [Dv_Plus_Tools_Info] Set UserStock = "& Totals &" where id="&Rs(0))
		Rs.movenext
		loop
		Rs.Close
		Set Rs = Nothing
	End If
	Dv_suc("用户拥有库存更新成功！")
End Sub

Function CheckID(CHECK_ID)
	Dim Fixid
	CheckID = False
	Fixid = Replace(CHECK_ID,",","")
	Fixid = Trim(Replace(Fixid," ",""))
	If IsNumeric(Fixid) and Fixid<>"" Then CheckID = True
End Function

Function CheckNumeric(CHECK_ID)
	If CHECK_ID<>"" and IsNumeric(CHECK_ID) Then
		CHECK_ID = cCur(CHECK_ID)
	Else
		CHECK_ID = 0
	End If
	CheckNumeric = CHECK_ID
End Function
%>