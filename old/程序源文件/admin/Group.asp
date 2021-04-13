<!--#include file=../conn.asp-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/ArrayList.asp"-->
<!--#include file="../inc/GroupPermission.asp"-->
<%
Head()
Dim admin_flag
admin_flag=",17,"
CheckAdmin(admin_flag)
Main()
Footer()

Sub Main()
	Select Case Request("action")
	Case "editgroup"
		EditGroup()
	Case "saveusergroup"
		SaveUserGroup()
	Case "savesysgroup"
		SaveSysGroup()
	Case "delusergroup"
		DelUserGroup()
	Case "online"
		GroupOnline()
	Case "saveonline"
		SaveGroupOnline()
	Case Else
		UserGroup()
	End Select
End Sub

Sub UserGroup()
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th style="text-align:center;" align=left>&nbsp;操作提示</th>
</tr>
<tr align=left>
<td height="23" class="td1" style="LINE-HEIGHT: 140%">
<li>动网论坛用户组分为系统用户组、特殊用户组、注册用户组、多属性用户组四种类型
<li>系统用户组为内置固定用户组，不能添加，供论坛管理之用，不能随意更改，如删除则会引起论坛运行异常
<li>特殊用户组不随用户等级升降而变更，通常建立来分配给一些对论坛有特殊贡献或操作的人员
<li>多属性用户组不随用户等级升降而变更，该组用户<U>可设置享有多个不同用户组的权限</U>，通常建立来分配给一些对论坛有特殊贡献或操作的人员
<li>注册用户组即为传统的用户等级，每个组(等级)可设定不同的权限
<li>默认权限为添加新的用户组时使用其中一些定义好的权限设置，通常新添加用户组后都要再次定义其权限
</td>
</tr>
<tr align=left>
<td height="23" class="td2" style="LINE-HEIGHT: 140%">
<B>快捷操作</B>：<a href="#1">系统组</a> | <a href="#2">特殊组</a> | <a href="#3">多属性组</a>
 | <a href="#4">Vip用户组管理</a>
| <a href="?action=editgroup&groupid=4">编辑默认组资料</a> | <a href="?action=online">在线图例定制</a>
</td>
</tr>
</table>
<BR>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th style="text-align:center;" colspan="6">注册用户组(等级)管理</th>
</tr>
<tr><td colspan=6 height=25 class="td1">
小提示：点击权限您可以分别设定每个注册用户组(等级)分别拥有不同的论坛权限
</td></tr>
<tr>
<td height="23" width="5%" class=bodytitle><B>组ID</B></td>
<td height="23" width="20%" class=bodytitle><B>用户组(等级)名称</B></td>
<td width="15%" class=bodytitle><B>最少发贴</B></td>
<td width="30%" class=bodytitle><B>组(等级)图片</B></td>
<td height="23" width="10%" class=bodytitle><B>用户数</B></td>
<td width="20%" class=bodytitle><B>操作</B></td>
</tr>
<FORM METHOD=POST ACTION="?action=saveusergroup">
<%
Dim Trs,rs
Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where ParentGID=3 Order By MinArticle")
Do While Not Rs.Eof
%>
<input type=hidden value="<%=rs("UserGroupID")%>" name="usertitleid">
<tr>
<td class=td1 align=center><%=Rs("UserGroupID")%></td>
<td height="23" class=td1><input size=15 value="<%=rs("usertitle")%>" name="usertitle" type=text></td>
<td class=td1><input size=5 value="<%=rs("MinArticle")%>" name="minarticle" type=text></td>
<td class=td1><input size=15 value="<%=rs("grouppic")%>" name="titlepic" type=text><img src="../<%=Dvbbs.Forum_PicUrl%>star/<%=rs("grouppic")%>" border="0"></td>
<td class=td1>
<B><%
Set Trs=Dvbbs.Execute("Select Count(*) From [Dv_User] Where UserGroupID="&Rs("UserGroupID"))
Response.Write Trs(0)
%></B>
</td>
<td class=td1><a href="?action=editgroup&groupid=<%=rs("UserGroupID")%>">编辑</a> | <a href="user.asp?action=userSearch&userSearch=10&usergroupid=<%=rs("usergroupid")%>">列出用户</a> | <a href="?action=delusergroup&id=<%=rs("UserGroupID")%>" onclick="{if(confirm('删除操作将会自动更新一部分用户的等级，并且不可恢复，确定吗?')){return true;}return false;}">删除</a></td>
</tr>
<%
Rs.MoveNext
Loop
Set Rs=Nothing
Set Trs=Nothing
%>
<input type=hidden value="0" name="usertitleid">
<tr>
<td class=td1 align=center><font color=blue>新</font></td>
<td height="23" class=td1><input size=15 value="" name="usertitle" type=text></td>
<td class=td1><input size=5 value="0" name="minarticle" type=text></td>
<td class=td1><input size=15 value="level0.gif" name="titlepic" type=text></td>
<td class=td1>
<B>0</B>
</td>
<td width="20%" class=td1>&nbsp;</td>
</tr>
<tr align=center>
<td colspan=6 height=25 class="td2">
<input type=submit class="button" name=submit value="提交更改">
</td></tr>
</FORM>
</table>
<BR>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th style="text-align:center;" colspan="6">系统用户组管理<a name="1"></a></th>
</tr>
<tr><td colspan=6 height=25 class="td1">
小提示：点击权限您可以分别设定每个系统用户组分别拥有不同的论坛权限，系统组头衔和图标显示在前台用户信息中
</td></tr>
<tr>
<td height="23" width="5%" class=bodytitle><B>组ID</B></td>
<td height="23" width="20%" class=bodytitle><B>系统组头衔</B></td>
<td width="15%" class=bodytitle><B>系统中名称</B></td>
<td height="23" width="30%" class=bodytitle><B>系统组图标</B></td>
<td height="23" width="10%" class=bodytitle><B>用户数</B></td>
<td width="20%" class=bodytitle><B>操作</B></td>
</tr>
<FORM METHOD=POST ACTION="?action=savesysgroup">
<input type=hidden value="1" name="ParentID">
<%
Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where ParentGID=1 Order By UserGroupID")
Do While Not Rs.Eof
%>
<input type=hidden value="<%=rs("UserGroupID")%>" name="usertitleid">
<input value="<%=rs("title")%>" name="title" type=hidden>
<input type=hidden value="<%=rs("IsSetting")%>" name="issetting">
<tr>
<td class=td1 align=center><%=Rs("UserGroupID")%></td>
<td height="23" class=td1><input size=15 value="<%=rs("usertitle")%>" name="usertitle" type=text></td>
<td class=td1><%=Rs("Title")%></td>
<td class=td1><input size=15 value="<%=rs("grouppic")%>" name="titlepic" type=text>
<img src="../<%=Dvbbs.Forum_PicUrl%>star/<%=rs("grouppic")%>" border="0">
</td>
<td class=td1>
<B><%
Set Trs=Dvbbs.Execute("Select Count(*) From [Dv_User] Where UserGroupID="&Rs("UserGroupID"))
Response.Write Trs(0)
%></B>
</td>
<td class=td1><a href="?action=editgroup&groupid=<%=rs("UserGroupID")%>">编辑</a> | <a href="user.asp?action=userSearch&userSearch=10&usergroupid=<%=rs("usergroupid")%>">列出用户</a></td>
</tr>
<%
Rs.MoveNext
Loop
Set Rs=Nothing
Set Trs=Nothing
%>
<tr align=center>
<td colspan=6 height=25 class="td2">
<input type=submit class="button" name=submit value="提交更改">
</td></tr>
</FORM>
</table>

<BR>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th style="text-align:center;" colspan="6">特殊用户组管理<a name="2"></a></th>
</tr>
<tr><td colspan=6 height=25 class="td1">
小提示：点击权限您可以分别设定每个特殊用户组分别拥有不同的论坛权限，通常建立来分配给论坛上比较特殊的用户群体，特殊组头衔和图标显示在前台用户信息中
</td></tr>
<tr>
<td width="5%" class=bodytitle><B>组ID</B></td>
<td height="23" width="15%" class=bodytitle><B>特殊组头衔</B></td>
<td width="15%" class=bodytitle><B>系统中名称</B></td>
<td width="30%" class=bodytitle><B>特殊组图片</B></td>
<td height="23" width="10%" class=bodytitle><B>用户数</B></td>
<td width="20%" class=bodytitle><B>操作</B></td>
</tr>
<FORM METHOD=POST ACTION="?action=savesysgroup">
<input type=hidden value="2" name="ParentID">
<%
Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where ParentGID=2 Order By UserGroupID")
Do While Not Rs.Eof
%>
<input type=hidden value="<%=rs("UserGroupID")%>" name="usertitleid">
<input type=hidden value="<%=rs("IsSetting")%>" name="issetting">
<tr>
<td class=td1 align=center><%=Rs("UserGroupID")%></td>
<td height="23" class=td1><input size=15 value="<%=rs("usertitle")%>" name="usertitle" type=text></td>
<td class=td1><input size=15 value="<%=rs("title")%>" name="title" type=text></td>
<td class=td1><input size=15 value="<%=rs("grouppic")%>" name="titlepic" type=text>
<img src="../<%=Dvbbs.Forum_PicUrl%>star/<%=rs("grouppic")%>" border="0">
</td>
<td class=td1>
<B><%
Set Trs=Dvbbs.Execute("Select Count(*) From [Dv_User] Where UserGroupID="&Rs("UserGroupID"))
Response.Write Trs(0)
%></B>
</td>
<td class=td1><a href="?action=editgroup&groupid=<%=rs("UserGroupID")%>">编辑</a> | <a href="user.asp?action=userSearch&userSearch=10&usergroupid=<%=rs("usergroupid")%>">列出用户</a> | <a href="?action=delusergroup&id=<%=rs("UserGroupID")%>" onclick="{if(confirm('删除操作将会自动更新一部分用户的等级，并且不可恢复，确定吗?')){return true;}return false;}">删除</a></td>
</tr>
<%
Rs.MoveNext
Loop
Set Rs=Nothing
Set Trs=Nothing
%>
<input type=hidden value="0" name="usertitleid">
<input type=hidden value="" name="issetting">
<tr>
<td class=td1 align=center><font color=blue>新</font></td>
<td height="23" class=td1><input size=15 value="" name="usertitle" type=text></td>
<td class=td1><input size=15 value="" name="title" type=text ></td>
<td class=td1><input size=15 value="level0.gif" name="titlepic" type=text></td>
<td class=td1>
<B>0</B>
</td>
<td class=td1>&nbsp;</td>
</tr>
<tr align=center>
<td colspan=6 height=25 class="td2">
<input type=submit class="button" name=submit value="提交更改">
</td></tr>
</FORM>
</table>
<!--
<BR>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th style="text-align:center;" colspan="7">多属性用户组管理<a name="3"></a></th>
</tr>
<tr><td colspan=7 height=25 class="td1">
小提示：点击权限您可以分别设定每个多属性用户组的默认论坛权限，通常建立来分配给论坛上比较特殊的用户群体，多属性组头衔和图标显示在前台用户信息中，多属性用户组的用户可同时拥有多个用户组的权限。<BR><font color=blue>包含组ID请慎重填写，组ID的获取可参考上面的各个组列表，内容用竖线分隔(如：2|8)，如果发现不能更新，请仔细检查所填写组ID是否正确</font>
</td></tr>
<tr>
<td width="5%" class=bodytitle><B>组ID</B></td>
<td height="23" width="13%" class=bodytitle><B>多属性组头衔</B></td>
<td width="10%" class=bodytitle><B>系统中名称</B></td>
<td width="17%" class=bodytitle><B>包含组ID</B></td>
<td width="28%" class=bodytitle><B>多属性组图片</B></td>
<td height="23" width="8%" class=bodytitle><B>用户数</B></td>
<td width="25%" class=bodytitle><B>操作</B></td>
</tr>
<FORM METHOD=POST ACTION="?action=savesysgroup">
<input type=hidden value="4" name="ParentID">
<%
Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where ParentGID=4 Order By UserGroupID")
Do While Not Rs.Eof
%>
<input type=hidden value="<%=rs("UserGroupID")%>" name="usertitleid">
<tr>
<td class=td1 align=center><%=Rs("UserGroupID")%></td>
<td height="23" class=td1><input size=15 value="<%=rs("usertitle")%>" name="usertitle" type=text></td>
<td class=td1><input size=15 value="<%=rs("title")%>" name="title" type=text></td>
<td class=td1><input size=15 value="<%=rs("IsSetting")%>" name="issetting" type=text> *</td>
<td class=td1><input size=15 value="<%=rs("grouppic")%>" name="titlepic" type=text>
<img src="../<%=Dvbbs.Forum_PicUrl%>star/<%=rs("grouppic")%>" border="0">
</td>
<td class=td1>
<B><%
Set Trs=Dvbbs.Execute("Select Count(*) From [Dv_User] Where UserGroupID="&Rs("UserGroupID"))
Response.Write Trs(0)
%></B>
</td>
<td class=td1><a href="?action=editgroup&groupid=<%=rs("UserGroupID")%>">编辑</a> | <a href="user.asp?action=userSearch&userSearch=10&usergroupid=<%=rs("usergroupid")%>">列出用户</a> | <a href="?action=delusergroup&id=<%=rs("UserGroupID")%>" onclick="{if(confirm('删除操作将会自动更新一部分用户的等级，并且不可恢复，确定吗?')){return true;}return false;}">删除</a></td>
</tr>
<%
Rs.MoveNext
Loop
Set Rs=Nothing
Set Trs=Nothing
%>
<input type=hidden value="0" name="usertitleid">
<tr>
<td class=td1 align=center><font color=blue>新</font></td>
<td height="23" class=td1><input size=15 value="" name="usertitle" type=text></td>
<td class=td1><input size=15 value="" name="title" type=text ></td>
<td class=td1><input size=15 value="" name="issetting" type=text> *</td>
<td class=td1><input size=15 value="level0.gif" name="titlepic" type=text></td>
<td class=td1>
<B>0</B>
</td>
<td class=td1>&nbsp;</td>
</tr>
<tr align=center>
<td colspan=7 height=25 class="td2">
<input type=submit class="button" name=submit value="提交更改">
</td></tr>
</FORM>
</table>
<BR>
//-->
<%
Dim FoundVipGroup
FoundVipGroup = False
%>
<BR>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th style="text-align:center;" colspan="6">Vip用户组管理<a name="4"></a></th>
</tr>
<tr><td colspan=6 height=25 class="td1">
小提示：VIP用户将有权限限期控制，当该用户的使用权限过期，系统将会自动将会员转到默认注册组。
</td></tr>
<tr>
<td width="5%" class=bodytitle><B>组ID</B></td>
<td height="23" width="15%" class=bodytitle><B>特殊组头衔</B></td>
<td width="15%" class=bodytitle><B>系统中名称</B></td>
<td width="30%" class=bodytitle><B>特殊组图片</B></td>
<td height="23" width="10%" class=bodytitle><B>用户数</B></td>
<td width="20%" class=bodytitle><B>操作</B></td>
</tr>
<FORM METHOD=POST ACTION="?action=savesysgroup">
<input type=hidden value="5" name="ParentID">
<%
Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where ParentGID=5 Order By UserGroupID")
Do While Not Rs.Eof
FoundVipGroup = True
%>
<input type=hidden value="<%=rs("UserGroupID")%>" name="usertitleid">
<input type=hidden value="<%=rs("IsSetting")%>" name="issetting">
<tr>
<td class=td1 align=center><%=Rs("UserGroupID")%></td>
<td height="23" class=td1><input size=15 value="<%=rs("usertitle")%>" name="usertitle" type=text></td>
<td class=td1><input size=15 value="<%=rs("title")%>" name="title" type=text></td>
<td class=td1><input size=15 value="<%=rs("grouppic")%>" name="titlepic" type=text>
<img src="../<%=Dvbbs.Forum_PicUrl%>star/<%=rs("grouppic")%>" border="0">
</td>
<td class=td1>
<B><%
Set Trs=Dvbbs.Execute("Select Count(*) From [Dv_User] Where UserGroupID="&Rs("UserGroupID"))
Response.Write Trs(0)
%></B>
</td>
<td class=td1><a href="?action=editgroup&groupid=<%=rs("UserGroupID")%>">编辑</a> | <a href="user.asp?action=userSearch&userSearch=10&usergroupid=<%=rs("usergroupid")%>">列出用户</a> | <a href="?action=delusergroup&id=<%=rs("UserGroupID")%>" onclick="{if(confirm('删除后VIP用户将失去相关的VIP权限，并且不可恢复，确定吗?')){return true;}return false;}">删除</a></td>
</tr>
<%
Rs.MoveNext
Loop
Set Rs=Nothing
Set Trs=Nothing
%>
<input type=hidden value="0" name="usertitleid">
<input type=hidden value="" name="issetting">
<%
If Not FoundVipGroup Then
%>
<tr>
<td class=td1 align=center><font color=blue>新</font></td>
<td height="23" class=td1><input size=15 value="" name="usertitle" type=text></td>
<td class=td1><input size=15 value="Vip用户组" name="title" type=text ></td>
<td class=td1><input size=15 value="level0.gif" name="titlepic" type=text></td>
<td class=td1>
<B>0</B>
</td>
<td class=td1>&nbsp;</td>
</tr>
<%
Else
%>
<input size=15 value="" name="usertitle" type=hidden>
<input size=15 value="Vip用户组" name="title" type=hidden>
<input size=15 value="level0.gif" name="titlepic" type=hidden>
<%
End If
%>
<tr align=center>
<td colspan=6 height=25 class="td2">
<input type=submit class="button" name=submit value="提交更改">
</td></tr>
</FORM>
</table>
<br/>
<%
End Sub

'保存注册用户组(等级)批量更改信息
Sub SaveUserGroup()
	Server.ScriptTimeout=99999999
	Dim UserTitleID,UserTitle,MinArticle,TitlePic,i,rs
	For i=1 To Request.Form("usertitleid").Count
		UserTitleID=Replace(Request.Form("usertitleid")(i),"'","")
		UserTitle=Replace(Request.Form("usertitle")(i),"'","")
		MinArticle=Replace(Request.Form("minarticle")(i),"'","")
		TitlePic=Replace(Request.Form("titlepic")(i),"'","")
		If IsNumeric(UserTitleID) And UserTitle<>"" And IsNumeric(MinArticle) And TitlePic<>"" Then
			Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where ParentGID=3 And UserGroupID="&UserTitleID)
			If Not (Rs.Eof And Rs.Bof) Then
				If Rs("UserTitle")<>Trim(UserTitle) Or Rs("GroupPic")<>Trim(TitlePic) Then
					Dvbbs.Execute("Update [Dv_User] Set UserClass='"&UserTitle&"',TitlePic='"&TitlePic&"' Where UserGroupID="&UserTitleID)
				End If
				Dvbbs.Execute("Update Dv_UserGroups Set UserTitle='"&UserTitle&"',MinArticle="&MinArticle&",GroupPic='"&TitlePic&"' Where UserGroupID="&UserTitleID)
			End If
			'新加入用户组(等级)
			If Clng(UserTitleID) = 0 Then
				Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where UserGroupID=4")
				Dvbbs.Execute("Insert Into Dv_UserGroups (Title,UserTitle,GroupSetting,Orders,MinArticle,TitlePic,GroupPic,ParentGID) Values ('"&Rs("Title")&"','"&UserTitle&"','"&Rs("GroupSetting")&"',0,"&MinArticle&",'"&Rs("TitlePic")&"','"&TitlePic&"',3)")
			End If
		End If
	Next
	Dv_suc("批量更新用户组（等级）资料成功！")
	Set Rs=Nothing
	Dvbbs.LoadGroupSetting
End Sub

'保存系统、特殊、多属性用户组批量更改信息
Sub SaveSysGroup()
	Server.ScriptTimeout=99999999
	Dim UserTitleID,UserTitle,TitlePic,ParentID,Title,IsSetting,FoundIsSetting,mIsSetting,GroupIDList,k,rs,sql,i
	SQL = "Select UserGroupID From Dv_UserGroups"
	Set Rs = Dvbbs.Execute(SQL)
		GroupIDList = Rs.GetString(,, "", ",", "")
	Rs.close
	Set Rs = Nothing
	GroupIDList = "," & GroupIDList
	GroupIDList = Replace(GroupIDList,",","|")
	ParentID = Request.Form("ParentID")
	If Not IsNumeric(ParentID) Or ParentID="" Then
		Errmsg = ErrMsg + "<BR><li>非法的用户组参数。"
		Dvbbs_Error()
		Exit Sub
	End If
	ParentID = Cint(ParentID)
	FoundIsSetting = True
	For i=1 To Request.Form("usertitleid").Count
		UserTitleID=Replace(Request.Form("usertitleid")(i),"'","")
		UserTitle=Replace(Request.Form("usertitle")(i),"'","")
		Title=Replace(Request.Form("title")(i),"'","")
		TitlePic=Replace(Request.Form("titlepic")(i),"'","")
		IsSetting=Replace(Request.Form("issetting")(i),"'","")
		If IsNumeric(UserTitleID) And UserTitle<>"" And TitlePic<>"" Then
			Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where ParentGID="&ParentID&" And UserGroupID="&UserTitleID)
			If Not (Rs.Eof And Rs.Bof) Then
				If Rs("UserTitle")<>Trim(UserTitle) Or Rs("GroupPic")<>Trim(TitlePic) Then
					Dvbbs.Execute("Update [Dv_User] Set UserClass='"&UserTitle&"',TitlePic='"&TitlePic&"' Where UserGroupID="&UserTitleID)
				End If
			End If
			If ParentID = 4 And Trim(IsSetting)<>"" Then
				mIsSetting = Split(IsSetting,"|")
				For k = 0 To Ubound(mIsSetting)
					'多属性用户组，填写的UserGroupID不存在则不更新
					If InStr(GroupIDList,"|" & mIsSetting(k) & "|") = 0 Then
						FoundIsSetting = False
						Exit For
					End If
				Next
				If FoundIsSetting Then
					Dvbbs.Execute("Update Dv_UserGroups Set Title='"&Title&"',UserTitle='"&UserTitle&"',GroupPic='"&TitlePic&"',IsSetting='"&IsSetting&"' Where UserGroupID="&UserTitleID)
					'新加入用户组
					If Clng(UserTitleID) = 0 Then
						Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where UserGroupID=4")
						Dvbbs.Execute("Insert Into Dv_UserGroups (Title,UserTitle,GroupSetting,Orders,MinArticle,TitlePic,GroupPic,ParentGID,IsSetting) Values ('"&Title&"','"&UserTitle&"','"&Rs("GroupSetting")&"',0,0,'"&Rs("TitlePic")&"','"&TitlePic&"',"&ParentID&",'"&IsSetting&"')")
					End If
				Else
					Dvbbs.Execute("Update Dv_UserGroups Set Title='"&Title&"',UserTitle='"&UserTitle&"',GroupPic='"&TitlePic&"' Where UserGroupID="&UserTitleID)
				End If
				FoundIsSetting = True
			Else
				Dvbbs.Execute("Update Dv_UserGroups Set Title='"&Title&"',UserTitle='"&UserTitle&"',GroupPic='"&TitlePic&"' Where UserGroupID="&UserTitleID)
				'新加入用户组
				If Clng(UserTitleID) = 0 Then
					Dim tGroupSetting	'修正下标越界，轻飘飘
					Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where UserGroupID=4")
					tGroupSetting=Rs("GroupSetting")
					tGroupSetting=Split(tGroupSetting,",")
					tGroupSetting(71)="0§0§0§0"
					tGroupSetting=Join(tGroupSetting,",")
					Dvbbs.Execute("Insert Into Dv_UserGroups (Title,UserTitle,GroupSetting,Orders,MinArticle,TitlePic,GroupPic,ParentGID) Values ('"&Title&"','"&UserTitle&"','"&Rs("GroupSetting")&"',0,0,'"&Rs("TitlePic")&"','"&TitlePic&"',"&ParentID&")")
				End If
			End If
		End If
	Next
	Dvbbs.LoadGroupSetting():iGroupSetting_UserName()
	Dv_suc("批量用户组资料成功！")
	Set Rs=Nothing
End Sub

'删除注册用户组(等级)信息
Sub DelUserGroup()
	Dim UserTitleID,tRs,rs
	UserTitleID = Request("id")
	If Not IsNumeric(UserTitleID) Or UserTitleID = "" Then
		Errmsg = ErrMsg + "<BR><li>请指定要删除的用户组(等级)。"
		Dvbbs_Error()
		Exit Sub
	End If
	UserTitleID = Clng(UserTitleID)
	'检测用户组是否存在以及取得临近用户组的信息
	'如果用户组为特殊、多属性组，则更新其用户信息为最低等级，用户登陆后会自动重新更新
	Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where (Not ParentGID=1) And UserGroupID = " & UserTitleID)
	If Rs.Eof And Rs.Bof Then
		Errmsg = ErrMsg + "<BR><li>指定要删除的用户组(等级)不存在。"
		Dvbbs_Error()
		Exit Sub
	ElseIf Not Rs("UserGroupID") = 8 And Rs("ParentGID") = 2 Then
		'删除特殊用户组（等级）之判断 2005-4-9 Dv.Yz
		Set tRs = Dvbbs.Execute("SELECT TOP 1 * FROM Dv_UserGroups WHERE ParentGID = 3 ORDER BY MinArticle Desc")
		If tRs.Eof And tRs.Bof Then
			Errmsg = ErrMsg + "<BR><li>注册用户组(等级)为空，不能删除，请先添加至少一个注册用户组(等级)。"
			Dvbbs_Error()
			Exit Sub
		Else
			Dvbbs.Execute("UPDATE Dv_User SET UserClass = '" & tRs("UserTitle") & "', TitlePic = '" & tRs("GroupPic") & "', UserGroupID = " & tRs("UserGroupID") & " WHERE UserGroupID = " & UserTitleID)
			Dvbbs.Execute("DELETE FROM Dv_UserGroups WHERE UserGroupID = " & UserTitleID)
		End If
	Else
		Set tRs=Dvbbs.Execute("Select Top 1 * From Dv_UserGroups Where ParentGID=3 And (Not UserGroupID="&UserTitleID&") And MinArticle<="&Rs("MinArticle")&" Order By MinArticle Desc")
		If tRs.Eof And tRs.Bof Then
			Errmsg = ErrMsg + "<BR><li>该用户组(等级)为最后一个注册用户组，不能删除。"
			Dvbbs_Error()
			Exit Sub
		Else
			Dvbbs.Execute("Update Dv_User Set UserClass='"&tRs("UserTitle")&"',TitlePic='"&tRs("GroupPic")&"',UserGroupID="&tRs("UserGroupID")&" Where UserGroupID="&UserTitleID)
			Dvbbs.Execute("Delete From Dv_UserGroups Where UserGroupID = " & UserTitleID)
		End If
		Set tRs=Nothing
	End If
Dvbbs.LoadGroupSetting():iGroupSetting_UserName()
	Dv_suc("用户组（等级）资料删除成功！")
	Set Rs=Nothing
End Sub
Function GetUserGroup() 'Add By reoaiq at 090926
Dim Rs_G
Set Rs_G=Dvbbs.Execute("Select UserGroupID,Title,UserTitle From Dv_UserGroups where ParentGid>0  Order by ParentGid,UserGroupID")
If Not Rs_G.eof Then
	GetUserGroup=Rs_G.GetRows(-1)
	Rs_G.close:Set Rs_G = Nothing
Else
	Exit Function
End If
End Function 

Sub SaveGroup(groupid) 'Add By reoaiq at 090926 Start
Dim i,CheckBoxList,CheckBoxList2,CheckBoxList3
Dim Rs
Dim GroupSetting,GroupSetting58,GroupSetting71
Dim GroupSetting_Arr
Dim GroupSetting_ArrayList
Set Rs=Dvbbs.Execute("select groupsetting from dv_usergroups where Usergroupid="&groupid&"")
If Rs.eof Then 
	Errmsg=ErrMsg + "<BR><li>查询数据不存在。"
	dvbbs_error()
	exit Sub
Else 
	GroupSetting=Rs(0)
	GroupSetting_Arr=Split(GroupSetting,",")
End If 
Rs.Close:Set Rs=Nothing 
Set GroupSetting_ArrayList=new ArrayList
For i=0 To 74
	If Request("CheckGroupSetting("&i&")")="on" Then 
	CheckBoxList=CheckBoxList&i&","
	End If 
Next 
i=0
CheckBoxList2=Split(CheckBoxList,",")
For i=0 To UBound(CheckBoxList2)-1
	If i=UBound(CheckBoxList2)-1 Then 
	CheckBoxList3=CheckBoxList3&CheckBoxList2(i)
	Else 
	CheckBoxList3=CheckBoxList3&CheckBoxList2(i)&","
	End If 
Next
i=0
If Request("CheckGroupPic")="on" Then 
		Dim title,grouppic
		title=Dvbbs.CheckStr(Request("title"))
		grouppic=Dvbbs.CheckStr(Request("grouppic"))	
End If 
Rem Fish 2010-2-1，定时设置用户组读写
Dim TyClsGroupM,TyClsGroup
If Request("ChkTyClsGroup")="on" Then 
	TyClsGroup=Dvbbs.CheckStr(Request("TyClsGroup"))
	TyClsGroupM = Request("TyClsGroupM0")
	For i = 1 to 23
		TyClsGroupM = TyClsGroupM & "|" & Request("TyClsGroupM"&i)
	Next
	i=0
	TyClsGroupM=Dvbbs.CheckStr(TyClsGroupM)
	'Response.write TyClsGroupM:exit sub
	if TyClsGroup<>"0" and TyClsGroupM="" then
		Errmsg = ErrMsg + "<BR><li>您没有选择设置该用户组访问权限的具体时间。" 
		Dvbbs_Error()
		Exit Sub
	end if
End If 
If CheckBoxList3="" And grouppic="" And TyClsGroup="" Then 
		Errmsg = ErrMsg + "<BR><li>未选中任何设置。"
		Dvbbs_Error()
		Exit Sub
Else 
		If grouppic<>"" Then 
		Dvbbs.Execute("Update Dv_Usergroups set UserTitle='"&title&"',grouppic='"&grouppic&"' where usergroupid="&Dvbbs.CheckStr(Request("groupid"))&"")
		End If
		Rem Fish
		If TyClsGroup<>"" Then 
		Dvbbs.Execute("Update Dv_Usergroups set TyClsGroup="&TyClsGroup&",TyClsGroupM='"&TyClsGroupM&"' where usergroupid="&Dvbbs.CheckStr(groupid)&"")
		End If
		If CheckBoxList3<>"" Then 
			GroupSetting_ArrayList.AddArray(GroupSetting_Arr)
			CheckBoxList3=Split(CheckBoxList3,",")
			For Each i In 	CheckBoxList3
				If i=58 Then 
					If Request("CheckGroupSetting(58)")="on" Then 
					GroupSetting58=Replace(Dvbbs.CheckStr(Request("GroupSetting(58)A")&"§"&Request("GroupSetting(58)B")),",","")
					End If 
				End If 
				If i=71 Then 
					If Request("CheckGroupSetting(71)")="on" Then 
					GroupSetting71=Replace(Dvbbs.CheckStr(Request("GroupSetting(71)A")&"§"&Request("GroupSetting(71)B")&"§"&Request("GroupSetting(71)C")&"§"&Request("GroupSetting(71)D")),",","")
					End If 
				End If 
				GroupSetting_ArrayList.Update i,Replace(Dvbbs.CheckStr(CLng(Request("GroupSetting("&i&")"))),",","")
			Next 		
			If GroupSetting58<>"" Then 
			GroupSetting_ArrayList.Update 58,GroupSetting58
			End If 
			If GroupSetting71<>"" Then 
			GroupSetting_ArrayList.Update 71,GroupSetting71
			End If 
			GroupSetting=GroupSetting_ArrayList.Implode(",")
			Dvbbs.Execute("update dv_usergroups set groupsetting='"&GroupSetting&"' where usergroupid="&groupid&"")
		End If 
End If 
Set GroupSetting_ArrayList=Nothing 
End Sub  'Add By reoaiq at 090926 End



Sub EditGroup()
Dim i
	If Not IsNumeric(Replace(Request("groupid"),",","")) Then
		Errmsg = ErrMsg + "<BR><li>请选择对应的用户组。"
		Dvbbs_Error()
		Exit Sub
	End If
	If Request("groupaction")="yes" Then
	    'add by reoaiq at 090927 Start
		Dim usergroup_arr,usergroupid 
	    usergroup_arr=Split(Dvbbs.CheckStr(request("Select_groupid")),",")
		For Each usergroupid In usergroup_arr
		SaveGroup usergroupid
		Next 
		Dv_suc("用户组（等级）资料修改成功！")
		Dvbbs.LoadGroupSetting():iGroupSetting_UserName()
		'add by reoaiq at 090927 End
	Else
		Dim reGroupSetting,Rs
		Set Rs=Dvbbs.Execute("Select * From Dv_Usergroups Where UserGroupID="&Cint(Request("groupid")))
		If Rs.Eof And Rs.Bof Then
			Errmsg = ErrMsg + "<BR><li>未找到该用户组！"
			Dvbbs_Error()
			Exit Sub
		End If
		reGroupSetting=Split(Rs("GroupSetting"),",")
%>
<FORM METHOD=POST ACTION="?action=editgroup" name="TheForm">
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr>
<th style="text-align:center;" colspan="5">
论坛用户组权限设置
</th>
</tr>
<tr><td colspan=5 height=25 class="td1"><B>说明</B>：
<BR>①在这里您可以设置各个用户组（等级）在论坛中的默认权限；
<BR>②可以删除和编辑新添加的用户组；
<BR>③<font color="red">在更新多个用户组设置时，请选取最左边的复选表单，只有选取的设置项目才会更新；</font>
<BR>④不执行多用户组更新时，不需要选取左边的用户组列表。
</td></tr>
<tr><td rowspan="800" valign="top">
应用用户组保存选项<BR>
请按 CTRL 键多选<BR>
<%
Dim GroupArr
GroupArr=GetUserGroup()
%>
<SELECT NAME="select_groupid" style="width:200" multiple size="<%=Ubound(GroupArr,2)+1%>">
<%
For i=0 To Ubound(GroupArr,2)
%>
<option value="<%=GroupArr(0,i)%>" <%If CLng(Dvbbs.CheckStr(Request("groupid")))=GroupArr(0,i) Then Response.Write "Selected"%>> <%=GroupArr(1,i)%>--<%=GroupArr(2,i)%></option>
<%
Next
%>
</SELECT>
</td></tr>



<tr><td colspan=5 height=25 class="td1">
<b>设置功能</b>：
[<a href="#setting1">编辑用户组(等级)资料信息</a>] 
[<a href="#setting2">浏览相关选项</a>] 
[<a href="#setting3">发帖权限</a>] 
[<a href="#setting4">帖子/主题编辑权限</a>] 
[<a href="#setting5">上传权限设置</a>] 
[<a href="#setting6">管理权限</a>] 
[<a href="#setting7">短信权限</a>] 
[<a href="#setting8">其他权限</a>] 
[<a href="#setting9">重要权限设置</a>] 
</td></tr>
<tr>
<th style="text-align:center;" colspan="5" align=left>
<a name="setting1"></a>
<INPUT TYPE="checkbox" class="checkbox" NAME="chkall" onclick="CheckAll(this.form);">[全选]
编辑用户组(等级)资料信息
</th>
</tr>
<tr><td colspan=5 height=25 class="bodytitle">
<B><a href="?">用户组(等级)管理</a> >> <%=SysGroupName(Rs("ParentGID"))%></B>
<%=Rs("UserTitle")%>
<input name="groupid" type="hidden" value="<%=Request("groupid")%>">
</td></tr>
<tr>
<td width="1%" class=td1>&nbsp;</td>
<td height="23" width="40%" class=td1>用户组(等级)名称</td>
<td height="23" class=td1 colspan=2><input size=35 name="title" type=text value="<%=Rs("UserTitle")%>"></td>
</tr>
<tr>
<td width="1%" class=td1><INPUT TYPE="checkbox" class="checkbox" NAME="CheckGroupPic"></td>
<td height="23" class=td1>用户组(等级)图片</td>
<td height="23"class=td1 colspan=2><input size=35 name="grouppic" type=text value="<%=rs("grouppic")%>"></td>
</tr>
<!--定时开关,Dv.Fish 2010-2-1 -->
<tr> 
<td width="1%" class=td1><INPUT TYPE="checkbox" class="checkbox" NAME="ChkTyClsGroup"></td>
<td class=td1>论坛定时设置</td>
<td class=td1>
<input type=radio class="radio" name="TyClsGroup" value="0" <%If rs("TyClsGroup")="0" or isnull(rs("TyClsGroup")) Then %>checked <%End If%>>关 闭</option>
<input type=radio class="radio" name="TyClsGroup" value="1" <%If rs("TyClsGroup")="1" Then %>checked <%End If%>>定时关闭</option>
<input type=radio class="radio" name="TyClsGroup" value="2" <%If rs("TyClsGroup")="2" Then %>checked <%End If%>>定时只读</option>
</td>
<input type="hidden" id="b10" value="<b>定时设置选择:</b><br><li>在这里您可以设置是否起用定时的各种功能，如果开启了本功能，请设置好下面选项中的论坛设置时间，论坛该版面将在您规定的时间内有指定的设置">
<td class=td2><a href=# onclick="helpscript(b10);return false;" class="helplink"><img src="../images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td width="1%" class=td1><INPUT TYPE="checkbox" class="checkbox" NAME="ChkTyClsGroupM"></td>
<td colspan=1 class=td1>
定时设置<BR>请根据需要选择开或关</td></td>
<td colspan=1 class=td1>
<%
Rem 新增两字段 TyClsGroup 及 TyClsGroupM  by 动网.小易
Dim TyClsGroupM
TyClsGroupM=rs("TyClsGroupM")
If instr(TyClsGroupM,"|")=0 or IsNull(TyClsGroupM) then TyClsGroupM="|"
TyClsGroupM=split(TyClsGroupM,"|")

If UBound(TyClsGroupM)<2 Then 
	TyClsGroupM="1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1"
	TyClsGroupM=split(TyClsGroupM,"|")
End If
For i= 0 to UBound(TyClsGroupM)
If i<10 Then Response.Write "&nbsp;"
%>
 <%=i%>点：<input type="checkbox" class="checkbox" name="TyClsGroupM<%=i%>" value="1" <%If TyClsGroupM(i)="1" Then %>checked<%End If%>>开
   
 <%
 If (i+1) mod 4 = 0 Then Response.Write "<br>"
 Next
 %>
</td>
<input type="hidden" id="b11" value="<b>论坛开放时间</b><br><li>设置了本选项必须同时打开是否起用定时开关论坛设置才有效，设置了此选项，论坛该版面将在您规定的时间内给用户开放">
<td class=td2><a href=# onclick="helpscript(b11);return false;" class="helplink"><img src="../images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<!--定时开关,Dv.Fish 2010-2-1 end-->
<%
GroupPermission(rs("GroupSetting"))
Dim test
test=Split(rs("GroupSetting"),",")

%>
<input type=hidden value="yes" name="groupaction">
</FORM>
</table>
<%
		Set Rs=Nothing
	End If
End Sub

Sub GroupOnline()
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th style="text-align:center;" align=left>操作提示</th>
</tr>
<tr align=left>
e;" class="helplink"><img src="../images/help.gif" border=0 title="鐐瑰嚮鏌ラ槄绠＄悊甯