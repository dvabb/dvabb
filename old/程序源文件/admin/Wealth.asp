<!--#include file=../conn.asp-->
<!-- #include file="inc/const.asp" -->
<!--#include file="../inc/dv_clsother.asp"-->
<%
Head()
Dim admin_flag
admin_flag=",5,"
CheckAdmin(admin_flag)

If request("action")="save" Then
	Call savegrade()
Else
	Call grade()
End If
If  founderr Then dvbbs_error()
Footer()

sub grade()
dim sel
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<th colspan="2" style="text-align:center;">用户积分设置</th>
</tr>
<tr> 
<td width="100%" class=td2 colspan=2>
<B>说明</B>：<br />1、复选框中选择的为当前的使用设置模板，点击可查看该模板设置，点击别的模板直接查看该模板并修改设置。您可以将您下面的设置保存在多个论坛版面中<br />2、您也可以将下面设定的信息保存并应用到具体的分论坛版面设置中，可多选<br />3、如果您想在一个版面引用别的版面的配置，只要点击该版面名称，保存的时候选择要保存到的版面名称名称即可。<br />
4、默认模板中的积分设置为论坛所有页面（<font color=blue>不包括具体的论坛版面</font>）使用，如登录和注册的相关分值；具体的论坛版面可以有不同的积分设置，如发贴、删贴等，当然您也可以根据上面的设定方法设定所有版面的积分设置都是一样的。
</td>
</tr>
<FORM METHOD=POST ACTION="">
<tr> 
<td width="100%" class="td2" colspan=2>
查看分版面积分设置，请选择左边下拉框相应版面&nbsp;&nbsp;
<select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
<option value="">查看分版面广告请选择</option>
<%
Dim ii,rs
set rs=Dvbbs.Execute("select boardid,boardtype,depth from dv_board order by rootid,orders")
do while not rs.eof
Response.Write "<option "
if rs(0)=dvbbs.boardid then
Response.Write " selected"
end if
Response.Write " value=""wealth.asp?boardid="&rs(0)&""">"
Select Case rs(2)
	Case 0
		Response.Write "╋"
	Case 1
		Response.Write "&nbsp;&nbsp;├"
End Select
If rs(2)>1 Then
	For ii=2 To rs(2)
		Response.Write "&nbsp;&nbsp;│"
	Next
	Response.Write "&nbsp;&nbsp;├"
End If
Response.Write rs(1)
Response.Write "</option>"
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
</td>
</tr>
</FORM>
</table><br />

<form method="POST" action=wealth.asp?action=save>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">

<tr> 
<td width="100%" class=td2 colspan=2>
<input type=checkbox class=checkbox name="getskinid" value="1" <%if request("getskinid")="1" or request("boardid")="" then Response.Write "checked"%>><a href="wealth.asp?getskinid=1">论坛默认积分</a><br /> 点击此处返回论坛默认积分设置，默认积分设置包含所有<FONT COLOR="blue">除</FONT>包含具体版面内容（如发贴、回帖、精华等）<FONT COLOR="blue">以外</FONT>的页面。<hr size=1 width="90%" color=blue>
</td>
</tr>
<tr>
<td width="200" class="td1">
版面积分设置保存选项<br />
请按 CTRL 键多选<br />
<select name="getboard" size="40" style="width:100%" multiple>
<%
set rs=Dvbbs.Execute("select boardid,boardtype,depth from dv_board order by rootid,orders")
do while not rs.eof
Response.Write "<option "
if rs(0)=dvbbs.boardid then
Response.Write " selected"
end if
Response.Write " value="&rs(0)&">"
Select Case rs(2)
	Case 0
		Response.Write "╋"
	Case 1
		Response.Write "&nbsp;&nbsp;├"
End Select
If rs(2)>1 Then
	For ii=2 To rs(2)
		Response.Write "&nbsp;&nbsp;│"
	Next
	Response.Write "&nbsp;&nbsp;├"
End If
Response.Write rs(1)
Response.Write "</option>"
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
</td>
<td class="td1" valign=top>
<table width=100%>
<tr> 
<th colspan="2" style="text-align:center;">用户金钱设定</th>
</tr>
<tr> 
<td width="40%" class=td1>注册金钱数</td>
<td width="60%" class=td1> 
<input type="text" name="wealthReg" size="35" value="<%=Dvbbs.Forum_user(0)%>">
</td>
</tr>
<tr> 
<td width="40%" class=td1>登录增加金钱</td>
<td width="60%" class=td1> 
<input type="text" name="wealthLogin" size="35" value="<%=Dvbbs.Forum_user(4)%>">
</td>
</tr>
<tr> 
<td width="40%" class=td1>发帖增加金钱</td>
<td width="60%" class=td1> 
<input type="text" name="wealthAnnounce" size="35" value="<%=Dvbbs.Forum_user(1)%>">
</td>
</tr>
<tr> 
<td width="40%" class=td1>跟帖增加金钱</td>
<td width="60%" class=td1> 
<input type="text" name="wealthReannounce" size="35" value="<%=Dvbbs.Forum_user(2)%>">
</td>
</tr>
<tr> 
<td width="40%" class=td1>精华增加金钱</td>
<td width="60%" class=td1> 
<input type="text" name="BestWealth" size="35" value="<%=Dvbbs.Forum_user(15)%>">
</td>
</tr>
<tr> 
<td width="40%" class=td1>删帖减少金钱</td>
<td width="60%" class=td1> 
<input type="text" name="wealthDel" size="35" value="<%=Dvbbs.Forum_user(3)%>">
</td>
</tr>
<tr> 
<th colspan="2" style="text-align:center;">用户积分设定</th>
</tr>
<tr> 
<td width="40%" class=td1>注册积分值</td>
<td width="60%" class=td1> 
<input type="text" name="epReg" size="35" value="<%=Dvbbs.Forum_user(5)%>">
</td>
</tr>
<tr> 
<td width="40%" class=td1>登录增加积分值</td>
<td width="60%" class=td1> 
<input type="text" name="epLogin" size="35" value="<%=Dvbbs.Forum_user(9)%>">
</td>
</tr>
<tr> 
<td width="40%" class=td1>发帖增加积分值</td>
<td width="60%" class=td1> 
<input type="text" name="epAnnounce" size="35" value="<%=Dvbbs.Forum_user(6)%>">
</td>
</tr>
<tr> 
<td width="40%" class=td1>跟帖增加积分值</td>
<td width="60%" class=td1> 
<input type="text" name="epReannounce" size="35" value="<%=Dvbbs.Forum_user(7)%>">
</td>
</tr>
<tr> 
<td width="40%" class=td1>精华增加积分值</td>
<td width="60%" class=td1> 
<input type="text" name="bestuserep" size="35" value="<%=Dvbbs.Forum_user(17)%>">
</td>
</tr>
<tr> 
<td width="40%" class=td1>删帖减少积分值</td>
<td width="60%" class=td1> 
<input type="text" name="epDel" size="35" value="<%=Dvbbs.Forum_user(8)%>">
</td>
</tr>
<tr> 
<th colspan="2" style="text-align:center;">用户魅力设定</th>
</tr>
<tr> 
<td width="40%" class=td1>注册魅力值</td>
<td width="60%" class=td1> 
<input type="text" name="cpReg" size="35" value="<%=Dvbbs.Forum_user(10)%>">
</td>
</tr>
<tr> 
<td width="40%" class=td1>登录增加魅力值</td>
<td width="60%" class=td1> 
<input type="text" name="cpLogin" size="35" value="<%=Dvbbs.Forum_user(14)%>">
</td>
</tr>
<tr> 
<td width="40%" class=td1>发帖增加魅力值</td>
<td width="60%" class=td1> 
<input type="text" name="cpAnnounce" size="35" value="<%=Dvbbs.Forum_user(11)%>">
</td>
</tr>
<tr> 
<td width="40%" class=td1>跟帖增加魅力值</td>
<td width="60%" class=td1> 
<input type="text" name="cpReannounce" size="35" value="<%=Dvbbs.Forum_user(12)%>">
</td>
</tr>
<tr> 
<td width="40%" class=td1>精华增加魅力值</td>
<td width="60%" class=td1> 
<input type="text" name="bestusercp" size="35" value="<%=Dvbbs.Forum_user(16)%>">
</td>
</tr>
<tr> 
<td width="40%" class=td1>删帖减少魅力值</td>
<td width="60%" class=td1> 
<input type="text" name="cpDel" size="35" value="<%=Dvbbs.Forum_user(13)%>">
</td>
</tr>
<tr> 
<td width="40%" class=td1>&nbsp;</td>
<td width="60%" class=td1> 
<div align="center"> 
<input type="submit" class="button" name="Submit" value="提 交">
</div>
</td>
</tr>
</table>
</td>
</tr>
</table>
</form>
<%
end sub

Sub savegrade()
	Dim Forum_user,iforum_setting,forum_setting,rs,sql,BoardIdStr,i
	Forum_user=Dvbbs.CheckNumeric(request.form("wealthReg")) & "," & Dvbbs.CheckNumeric(request.form("wealthAnnounce")) & "," & Dvbbs.CheckNumeric(request.form("wealthReannounce")) & "," & Dvbbs.CheckNumeric(request.form("wealthDel")) & "," & Dvbbs.CheckNumeric(request.form("wealthLogin")) & "," & Dvbbs.CheckNumeric(request.form("epReg")) & "," & Dvbbs.CheckNumeric(request.form("epAnnounce")) & "," & Dvbbs.CheckNumeric(request.form("epReannounce")) & "," & Dvbbs.CheckNumeric(request.form("epDel")) & "," & Dvbbs.CheckNumeric(request.form("epLogin")) & "," & Dvbbs.CheckNumeric(request.form("cpReg")) & "," & Dvbbs.CheckNumeric(request.form("cpAnnounce")) & "," & Dvbbs.CheckNumeric(request.form("cpReannounce")) & "," & Dvbbs.CheckNumeric(request.form("cpDel")) & "," & Dvbbs.CheckNumeric(request.form("cpLogin")) & "," & Dvbbs.CheckNumeric(request.form("BestWealth")) & "," & Dvbbs.CheckNumeric(request.form("BestuserCP")) & "," & Dvbbs.CheckNumeric(request.form("BestuserEP"))
	'response.write Forum_user
	
	'forum_info|||forum_setting|||forum_user|||copyright|||splitword|||stopreadme
	Set rs=Dvbbs.execute("select forum_setting from dv_setup")
	iforum_setting=split(rs(0),"|||")
	forum_setting=iforum_setting(0) & "|||" & iforum_setting(1) & "|||" & forum_user & "|||" & iforum_setting(3) & "|||" & iforum_setting(4) & "|||" & iforum_setting(5)
	forum_setting=dvbbs.checkstr(forum_setting)
	
	If request("getskinid")="1" Then
		sql = "update dv_setup set Forum_setting='"&forum_setting&"'"
		Dvbbs.Execute(sql)
		Dvbbs.Name="setup"
		Dvbbs.loadSetup()
	End If

	For i = 1 TO request("getboard").Count
		If isNumeric(request("getboard")(i)) Then
			If BoardIdStr = "" Then
				BoardIdStr = request("getboard")(i)
			Else
				BoardIdStr = BoardIdStr & "," & request("getboard")(i)
			End If
		End If
	Next

	If request("getboard")<>"" Then
		sql = "update dv_board set board_user='"&Forum_user&"' where boardid in ("&BoardIdStr&")"
		Dvbbs.Execute(sql)
		Dvbbs.ReloadBoardCache request("getboard")
	End If
	Dv_suc("论坛积分设置成功！")
End  Sub

%>