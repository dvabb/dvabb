<!--#include file=../conn.asp-->
<!--#include file="inc/const.asp"-->
<!--#include file="../inc/GroupPermission.asp"-->
<%
	Head()
dim admin_flag
admin_flag=",39,"
CheckAdmin(admin_flag)
Call main()
Footer()

Sub main()
	If  request("action")="save" Then
		call savenew()
	ElseIf request("action")="savedit" Then
		call savedit()
	ElseIf request("action")="del" Then
		call del()
	ElseIf request("action")="AddNew" OR request("action")="edit" Then
		AddNew()
	Else
		call gradeinfo()
	End If
End Sub

Sub AddNew()
dim trs,rs
Dim PSetting
Dim Groupids
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr><th width="100%" style="text-align:center;" colspan=2>论坛插件菜单管理
</th>
</tr>
<tr>
<td height="23" colspan="2" class=td1>注意：这里所添加内容将自动显示于论坛前台的顶部菜单</td>
</tr>

<tr>
<td height="23" colspan="2">
<a href="plus.asp">菜单管理首页</a>
<%if request("action")="edit" then
Set tRs=Dvbbs.Execute("Select * From Dv_Plus Where id="&request("id")&"")
PSetting=Split(Server.HTMLEncode(tRs("Plus_Setting")),"|||")
PSetting(0)=split(PSetting(0),"|")
%>
 | 编辑菜单 | <a href="plus.asp?action=AddNew">新建菜单</a></td>
</tr>
<FORM METHOD=POST ACTION="?action=savedit">
<input type=hidden value="<%=trs("id")%>" name="id">
<tr>
<td height="23" colspan="2" class=td1>
标题： <input type=text size=50 name="title" value="<%=Server.HtmlEncode(tRs("Plus_name"))%>"> 可用HTML语法
</td></tr>
<tr>
<td height="23" colspan="2" class=td1>
是否在导航中显示： 是  <Input type="radio" class="radio" name="Isuse" value="1"
<%
If trs("Isuse")=1 Then 
%>
 checked
<%
End If 
%>
>  否  <input type="radio" class="radio" name="Isuse" value="0" 
<%
If trs("Isuse")=0 Then 
%>
 checked
<%
End If 
%>
>
</td></tr>
<tr>
<td height="23" colspan="2" class=td1>
分类：
<Select Name="stype" size=1>
<%
Set Rs=Dvbbs.Execute("Select * From Dv_Plus Where Plus_type='0'")
If Rs.Eof And Rs.bof Then
	Response.Write "<option value=0>作为一级菜单</option>"
Else
	If Clng(tRs("Plus_type"))=0 Then
		Response.Write "<option value=0 selected>作为一级菜单</option>"
	Else
		Response.Write "<option value=0>作为一级菜单</option>"
	End If
	Do While Not Rs.Eof
		If CStr(request("id")) <> CStr(Rs("id")) Then 
			Response.Write "<option value="&rs("id")
			If Clng(tRs("Plus_type"))=Rs("ID") Then Response.Write " selected "
			Response.Write ">"&Server.htmlencode(rs("plus_name"))&"</option>"
		End If
	Rs.MoveNext
	Loop
End If
%>
</select>
不选择则将做为一级菜单<BR>
</td>
</tr>
<tr>
<td height="23" colspan="2" class=td1>
注释：
<input type=text size=50 name="readme" value="<%=server.htmlencode(trs("plus_copyright"))%>"> 显示于链接上的title注释，也是插件的版权信息<BR>
</td>
</tr>
<tr>
<td height="23" colspan="2" class=td1>
模式：
<select name="windowtype" size="1">
<option value="0" <%If PSetting(0)(0)=0 Then Response.Write "selected"%>>原窗口</option>
<option value="1" <%If PSetting(0)(0)=1 Then Response.Write "selected"%>>新窗口</option>
<option value="2" <%If PSetting(0)(0)=2 Then Response.Write "selected"%>>固定大小窗口</option>
<option value="3" <%If PSetting(0)(0)=3 Then Response.Write "selected"%>>全屏</option>
</select>&nbsp;&nbsp;
窗口宽：<input type=text name="windowwidth" value="<%=PSetting(0)(1)%>" size=5>
&nbsp;&nbsp;
窗口高：<input type=text name="windowheight" value="<%=PSetting(0)(2)%>" size=5>
</td>
</tr>
<tr>
<td height="23" colspan="2" class=td1>
链接：
<input type=text size=50 name="url" value="<%=server.htmlencode(trs("mainpage"))%>"> <BR>
</td>
</tr>
<tr>
<td height="23" colspan="2" class=td1>
后台管理链接：
<input type=text size=50 name="plus_adminpage" value="<%=server.htmlencode(trs("plus_adminpage")&"")%>"> <BR>
</td>
</tr>
<tr><th style="text-align:center;" colspan="2" >插件常规设置</th></tr>
<tr><td height="23" colspan="2" class=td1> 插件ID&nbsp;&nbsp;<input type=text name="plusID" value="<%=Trs("plus_ID")%>" size="20">这是你插件的唯一的标识，注意不能有重复的。
<a href="http://bbs.dvbbs.net/Dv_plusInfo.asp" target="_blank" title="到动网官方站查询关于插件的信息" >获得插件ID和信息</a> </td></tr>
<%
If UBound(PSetting)>2 Then
	PSetting(3)=Split(PSetting(3),",")
%>
<tr><td height="23" colspan="2" class=td1>是否定时开放&nbsp;&nbsp; 是  <Input type="radio" class="radio" name="useTime" value="1"
<%
If PSetting(3)(0)="1" Then Response.Write "checked"
%>
>  否  <input type="radio" class="radio" name="useTime" value="0" 
<%
If PSetting(3)(0)="0" Then Response.Write "checked"

Groupids = PSetting(3)(2)
%>
></td></tr>
<tr><td height="23" colspan="2" class=td1>定时开放起止时间&nbsp;&nbsp; <input type=text name="timesetting" value="<%=PSetting(3)(1)%>" size=10></td></tr>
<tr><td height="23" colspan="2" class=td1>能使用插件的用户组&nbsp;&nbsp; 
<input type=text name="groupid" value="<%=Replace(Groupids&"","@",",")%>" size=30>这里设置可以使用插件的用户组
<input type="button" class="button" value="选择用户组" onclick="getGroup('Select_Group');">
</td></tr>
<tr><td height="23" colspan="2" class=td1>管理人员&nbsp;&nbsp; <textarea  name="plusmaster" cols="50" rows="5" ><%=Replace(PSetting(3)(3),"|",vbCrLf)%></textarea><br>这里设置插件的管理员，每个用户名用回车分隔开.系统默认论坛管理员可以管理插件，如果您不需要另外设置管理员，此项可以不填。</td></tr>
<tr><th style="text-align:center;" colspan="2" >限制项</th></tr>
<tr><td height="23" colspan="2" class=td1> 能使用插件的最少文章&nbsp;&nbsp;<input type=text name="Plus_UserPost" value=<%=PSetting(3)(4)%> size=5></td></tr>
<tr><td height="23" colspan="2" class=td1> 能使用插件的最低金钱&nbsp;&nbsp;<input type=text name="Plus_userWealth" value=<%=PSetting(3)(5)%> size=5></td></tr>
<tr><td height="23" colspan="2" class=td1> 能使用插件的最低积分&nbsp;&nbsp;<input type=text name="Plus_UserEP" value=<%=PSetting(3)(6)%> size=5></td></tr>
<tr><td height="23" colspan="2" class=td1> 能使用插件的最低魅力&nbsp;&nbsp;<input type=text name="Plus_UserCP" value=<%=PSetting(3)(7)%> size=5></td></tr>
<tr><td height="23" colspan="2" class=td1> 能使用插件的最低威望&nbsp;&nbsp;<input type=text name="Plus_UserPower" value=<%=PSetting(3)(8)%> size=5></td></tr>
<tr><th style="text-align:center;" colspan="2" >更新项</th></tr>
<tr><td height="23" colspan="2" class=td1> 每次使用插件金钱变化&nbsp;&nbsp;<input type=text name="Plus_ADDuserWealth" value=<%=PSetting(3)(9)%> size=5></td></tr>
<tr><td height="23" colspan="2" class=td1> 每次使用插件积分变化&nbsp;&nbsp;<input type=text name="Plus_ADDUserEP" value=<%=PSetting(3)(10)%> size=5></td></tr>
<tr><td height="23" colspan="2" class=td1> 每次使用插件魅力变化&nbsp;&nbsp;<input type=text name="Plus_ADDUserCP" value=<%=PSetting(3)(11)%> size=5></td></tr>
<tr><td height="23" colspan="2" class=td1> 每次使用插件威望变化&nbsp;&nbsp;<input type=text name="Plus_ADDUserPower" value=<%=PSetting(3)(12)%> size=5></td></tr>
<%
Else 
%>
<tr><td height="23" colspan="2" class=td1>是否定时开放&nbsp;&nbsp; 是  <Input type="radio" class="radio" name="useTime" value="1">  否  <input type="radio" class="radio" name="useTime" value="0" checked></td></tr>
<tr><td height="23" colspan="2" class=td1>定时开放起止时间&nbsp;&nbsp; <input type=text name="timesetting" value="0|24" size=10></td></tr>
<tr><td height="23" colspan="2" class=td1>能使用插件的用户组&nbsp;&nbsp; 
<input type=text name="groupid" value="" size=30>
这里设置可以使用插件的用户组
<input type="button" class="button" value="选择用户组" onclick="getGroup('Select_Group');">
</td></tr>
<tr><td height="23" colspan="2" class=td1>管理人员&nbsp;&nbsp; <textarea  name="plusmaster" cols="50" rows="5" ></textarea><br>这里设置插件的管理员，每个用户名用回车分隔开.系统默认论坛管理员可以管理插件，如果您不需要另外设置管理员，此项可以不填。</td></tr>
<tr><th style="text-align:center;" colspan="2" >限制项</th></tr>
<tr><td height="23" colspan="2" class=td1> 能使用插件的最少文章&nbsp;&nbsp;<input type=text name="Plus_UserPost" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=td1> 能使用插件的最低金钱&nbsp;&nbsp;<input type=text name="Plus_userWealth" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=td1> 能使用插件的最低积分&nbsp;&nbsp;<input type=text name="Plus_UserEP" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=td1> 能使用插件的最低魅力&nbsp;&nbsp;<input type=text name="Plus_UserCP" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=td1> 能使用插件的最低威望&nbsp;&nbsp;<input type=text name="Plus_UserPower" value=0 size=5></td></tr>
<tr><th style="text-align:center;" colspan="2" >更新项</th></tr>
<tr><td height="23" colspan="2" class=td1> 每次使用插件金钱变化&nbsp;&nbsp;<input type=text name="Plus_ADDuserWealth" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=td1> 每次使用插件积分变化&nbsp;&nbsp;<input type=text name="Plus_ADDUserEP" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=td1> 每次使用插件魅力变化&nbsp;&nbsp;<input type=text name="Plus_ADDUserCP" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=td1> 每次使用插件威望变化&nbsp;&nbsp;<input type=text name="Plus_ADDUserPower" value=0 size=5></td></tr>
<%
End If
%>
<tr><td height="23" colspan="2" class=td1>注意，如果用户组中的客人组(ID为7)设置为可进入，那么所有的限制项无效。<Br>更新项中如果是客人可使用，对客人无效</td></tr>
<tr><th style="text-align:center;" colspan="2" >插件自定权限设置</th></tr>
<tr><td height="23" colspan="2" class=td1>
<textarea name="Plus_Setting" cols="80" rows="20">
<%	
	If UBound(PSetting)>1 Then 
		Dim i
		PSetting(1)=Split(PSetting(1),",")
		PSetting(2)=Split(PSetting(2),",")
		For i=0 to UBound (PSetting(1))
			Response.Write PSetting(2)(i)&"="&PSetting(1)(i)
			Response.Write vbCrLf
		Next
	Else
%>
设置字段1=0
设置字段2=0
设置字段3=0
设置字段4=0
设置字段5=0
设置字段6=0
设置字段7=0
设置字段8=0
设置字段9=0
设置字段10=0
设置字段11=0
设置字段12=0
设置字段13=0
设置字段14=0
设置字段15=0
设置字段16=0
设置字段17=0
设置字段18=0
设置字段19=0

<%
End If
%>
</textarea>
</td></tr>
<tr><td height="25" colspan="2" >说明：由于每个插件的设置不可能完全一样，设置字段的定义也不一样，这些都交给插件作者自行修改了。</td></tr>
<tr>
<td height="23" colspan="2" class=td2>
<input type=submit class="button" name=submit value="提交">
</td></tr>
</FORM>
<%else%>
<tr>
<th style="text-align:center;" colspan="2">添加菜单</th>
</tr>
<FORM METHOD=POST ACTION="?action=save">
<tr>
<td height="23" colspan="2" class=td1>
标题： <input type=text size=50 name="title"> 可用HTML语法
</td></tr>
<tr>
<td height="23" colspan="2" class=td1>
是否在导航中显示： 是  <Input type="radio" class="radio" name="Isuse" value="1" checked>  否  <input type="radio" class="radio" name="Isuse" value="0" >
</td></tr>
<tr>
<td height="23" colspan="2" class=td1>
分类：
<Select Name="stype" size=1>
<%
Set Rs=Dvbbs.Execute("Select * From Dv_Plus Where Plus_type='0'")
If Rs.Eof And Rs.bof Then
	Response.Write "<option value=0>作为一级菜单</option>"
Else
	Response.Write "<option value=0>作为一级菜单</option>"
	Do While Not Rs.Eof
		Response.Write "<option value="&rs("id")&">"&Server.htmlencode(rs("plus_name"))&"</option>"
	Rs.MoveNext
	Loop
End If
%>
</select>
不选择则将做为一级菜单<BR>
</td>
</tr>
<tr>
<td height="23" colspan="2" class=td1>
注释：
<input type=text size=50 name="readme"> 显示于链接上的title注释,也是插件的版权信息<BR>
</td>
</tr>
<tr>
<td height="23" colspan="2" class=td1>
模式：
<select name="windowtype" size="1">
<option value="0">原窗口</option>
<option value="1">新窗口</option>
<option value="2">固定大小窗口</option>
<option value="3">全屏</option>
</select>&nbsp;&nbsp;
窗口宽：<input type=text name="windowwidth" value=0 size=5>
&nbsp;&nbsp;
窗口高：<input type=text name="windowheight" value=0 size=5>
</td>
</tr>
<tr>
<td height="23" colspan="2" class=td1>
链接：
<input type=text size=50 name="url"> <BR>
</td>
</tr>
<tr>
<td height="23" colspan="2" class=td1>
后台管理链接：
<input type=text size=50 name="plus_adminpage"> <BR>
</td>
</tr>
<tr><th style="text-align:center;" colspan="2" >插件常规设置</th></tr>
<tr><td height="23" colspan="2" class=td1> 插件ID&nbsp;&nbsp;<input type=text name="plusID" value="newplus1" size="20">这是你插件的唯一的标识，注意不能有重复的。
<a href="http://bbs.dvbbs.net/Dv_plusInfo.asp" target="_blank" title="到动网官方站查询关于插件的信息" >获得插件ID和信息</a> </td></tr>
<tr><td height="23" colspan="2" class=td1>是否定时开放&nbsp;&nbsp; 是  <Input type="radio" class="radio" name="useTime" value="1">  否  <input type="radio" class="radio" name="useTime" value="0" checked></td></tr>
<tr><td height="23" colspan="2" class=td1>定时开放起止时间&nbsp;&nbsp; <input type=text name="timesetting" value="0|24" size=10></td></tr>
<tr><td height="23" colspan="2" class=td1>
能使用插件的用户组&nbsp;&nbsp; 
<input type=text name="groupid" value="" size=30>这里设置可以使用插件的用户组
<input type="button" class="button" value="选择用户组" onclick="getGroup('Select_Group');">
</td>
</tr>
<tr><td height="23" colspan="2" class=td1>管理人员&nbsp;&nbsp; <textarea  name="plusmaster" cols="50" rows="5" ></textarea><br>这里设置插件的管理员，每个用户名用回车分隔开.系统默认论坛管理员可以管理插件，如果您不需要另外设置管理员，此项可以不填。</td></tr>
<tr><th style="text-align:center;" colspan="2" >限制项</th></tr>
<tr><td height="23" colspan="2" class=td1> 能使用插件的最少文章&nbsp;&nbsp;<input type=text name="Plus_UserPost" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=td1> 能使用插件的最低金钱&nbsp;&nbsp;<input type=text name="Plus_userWealth" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=td1> 能使用插件的最低积分&nbsp;&nbsp;<input type=text name="Plus_UserEP" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=td1> 能使用插件的最低魅力&nbsp;&nbsp;<input type=text name="Plus_UserCP" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=td1> 能使用插件的最低威望&nbsp;&nbsp;<input type=text name="Plus_UserPower" value=0 size=5></td></tr>
<tr><th style="text-align:center;" colspan="2" >更新项</th></tr>
<tr><td height="23" colspan="2" class=td1> 每次使用插件金钱变化&nbsp;&nbsp;<input type=text name="Plus_ADDuserWealth" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=td1> 每次使用插件积分变化&nbsp;&nbsp;<input type=text name="Plus_ADDUserEP" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=td1> 每次使用插件魅力变化&nbsp;&nbsp;<input type=text name="Plus_ADDUserCP" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=td1> 每次使用插件威望变化&nbsp;&nbsp;<input type=text name="Plus_ADDUserPower" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=td1>注意，如果用户组中的客人组(ID为7)设置为可进入，那么所有的限制项无效。<Br>更新项中如果是客人可使用，对客人无效</td></tr>
<tr><th style="text-align:center;" colspan="2" >插件自定义扩展设置</th></tr>
<tr><td height="23" colspan="2" class=td1>
<textarea name="Plus_Setting" cols="80" rows="20">
设置字段1=0
设置字段2=0
设置字段3=0
设置字段4=0
设置字段5=0
设置字段6=0
设置字段7=0
设置字段8=0
设置字段9=0
设置字段10=0
设置字段11=0
设置字段12=0
设置字段13=0
设置字段14=0
设置字段15=0
设置字段16=0
设置字段17=0
设置字段18=0
设置字段19=0

</textarea>
</td></tr>
<tr><td height="25" colspan="2" >说明：由于每个插件的设置不可能完全一样，设置字段的定义也不一样，这些都交给插件作者自行修改了。</td></tr>
<tr>
<td height="23" colspan="2" class=td2>
<input type=submit class="button" name=submit value="提交">
</td></tr>
</FORM>
<%
end if
%>
</table><BR>
<%
Call Select_Group(Replace(Groupids&"","@",","))
End Sub 

sub gradeinfo()
dim trs,rs
Dim PSetting
%>

<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr><th width="100%" style="text-align:center;" colspan=5>论坛插件菜单管理
</th>
</tr>
<tr>
<td height="23" colspan="5" class=td1>注意：这里所添加内容将自动显示于论坛前台的顶部菜单</td>
</tr>
<tr>
<td height="23" colspan="5" class=td1><a href="plus.asp?action=AddNew">新建菜单</a> | <a href="plus.asp?action=posttodb">导出插件模板数据</a> | <a href="plus.asp?action=getfromdb">导入插件模板数据</a></td>
</tr>
<tr>
<th height="23" style="text-align:center;">标题</th>
<th>分类</th>
<th>窗口属性</th>
<th>是否显示</th>
<th>操作</th>
</tr>
<%
Set Rs=Dvbbs.Execute("Select * From Dv_Plus Order by ID Desc")
Do While Not Rs.Eof
PSetting=Split(Rs("Plus_Setting"),"|")
%>
<tr>
<td height="23" class=td1><%=Rs("Plus_Name")%></td>
<td class=td1>
<%If Rs("Plus_type")=0 Then%>
一级菜单
<%Else%>
<%
Set tRs=Dvbbs.Execute("Select * From Dv_Plus Where id="&Rs("Plus_Type")&"")
If tRs.Eof And tRs.Bof Then
	Response.Write "该菜单分类有误，请编辑修正"
Else
	Response.Write tRs("Plus_name")
End If
%>
<%End If%>
</td>
<td class=td1>
<%
Select Case PSetting(0)
Case 0
	Response.Write "当前窗口"
Case 1
	Response.Write "新窗口"
Case 2
	Response.Write "固定大小窗口，宽"&PSetting(1)&"，高 "&PSetting(2)&""
Case 3
	Response.Write "全屏"
End Select
%>
</td>
<td class=td1 align=center><%If Rs("isuse")=1 Then%>Yes<%Else%>No<%End If%></td>

<td class=td1 align=center><%
If Rs("plus_adminpage") <> "" Then  
%>
<a href="<%=Rs("plus_adminpage")%>">管理</a> | 
<%
End If
%>
<a href="?action=edit&id=<%=rs("id")%>">编辑</a> | <a href="?action=del&id=<%=rs("id")%>">删除</a></td>
</tr>
<%
Rs.MoveNext
Loop
Set Rs=Nothing
end sub

sub savenew()
	Dim plusID,plus_adminpage,Isuse,rs,sql
	plusID=Trim(Request("plusID"))
	If InStr(plusID,"'") >0 Then 
		Response.Write "插件ID中不允许有单引号"
		exit sub
	End If
	If request("title")="" then
		Errmsg = Errmsg + "请输入菜单的标题！"
		Dvbbs_Error()
		exit sub
	End If
	If plusID="" Then
		Errmsg = Errmsg + "请设定插件ID"
		Dvbbs_Error()
		exit sub
	End If
	SQL="Select count(*) From Dv_plus where plus_ID='"&plusID&"'"
	Set rs=Dvbbs.execute(SQL)
	If Rs(0) >0 Then
		Errmsg = Errmsg + "你设置的插件ID已经存在，请另行设置。"
		Dvbbs_Error()
		exit sub
	End If
	Isuse=Request("Isuse")
	If Isuse<>"1" And Isuse<>"0" Then Isuse="1"
	Isuse=CInt(Isuse)
	plus_adminpage=Dvbbs.checkStr(Request("plus_adminpage"))
	Dim Plus_SettingData,Plus_Setting,i,tmpstr
	Plus_Setting=Request("Plus_Setting")
	Plus_Setting=Split(Plus_Setting,vbCrLf)
	Plus_SettingData=""
	For i=0 to UBound(Plus_Setting)
		Plus_Setting(i)=Split(Plus_Setting(i),"=")
		If UBound(Plus_Setting(i))=1 Then
			If Plus_SettingData="" Then 
				Plus_SettingData=Trim(Plus_Setting(i)(1))
				tmpstr=Trim(Plus_Setting(i)(0))
			Else
				Plus_SettingData=Plus_SettingData&","&Trim(Plus_Setting(i)(1))
				tmpstr=tmpstr&","&Trim(Plus_Setting(i)(0))
			End If
		End If
	Next
	Plus_SettingData=Plus_SettingData&"|||"&tmpstr&"|||"
	Dim plusmaster,masterlist
	plusmaster=Request("plusmaster")
	plusmaster=split(plusmaster,vbCrLf)
	masterlist=""
	For i=0 to UBound(plusmaster)
		If Trim(plusmaster(i)) <>"" Then
			If masterlist="" Then
				masterlist=plusmaster(i)
			Else
				masterlist=masterlist&"|"&plusmaster(i)
			End If
		End If
	Next
	Dim useTime,timesetting,Groupsetting,Plus_UserPost,Plus_userWealth,Plus_UserEP,Plus_UserCP
	Dim Plus_UserPower,Plus_ADDuserWealth,Plus_ADDUserEP,Plus_ADDUserCP,Plus_ADDUserPower,guestuse
	useTime=Request("useTime")
	If useTime="" Then useTime=0
	timesetting=Trim(Request("timesetting"))
	If timesetting="" Then timesetting="0|24"
	Groupsetting=Replace(Trim(Request("groupid"))&"",",","@")

	Plus_UserPost=Trim(Request("Plus_UserPost"))
	If Plus_UserPost="" Then Plus_UserPost=0
	Plus_userWealth=Trim(Request("Plus_userWealth"))
	If Plus_userWealth="" Then Plus_userWealth=0
	Plus_UserEP=Trim(Request("Plus_UserEP"))
	If Plus_UserEP="" Then Plus_UserEP=0
	Plus_UserCP=Trim(Request("Plus_UserCP"))
	If Plus_UserCP="" Then Plus_UserCP=0
	Plus_UserPower=Trim(Request("Plus_UserPower"))
	If Plus_UserPower="" Then Plus_UserPower=0
	Plus_ADDuserWealth=Trim(Request("Plus_ADDuserWealth"))
	If Plus_ADDuserWealth="" Then Plus_ADDuserWealth=0
	Plus_ADDUserEP=Trim(Request("Plus_ADDUserEP"))
	If Plus_ADDUserEP="" Then Plus_ADDUserEP=0
	Plus_ADDUserCP=Trim(Request("Plus_ADDUserCP"))
	If Plus_ADDUserCP="" Then Plus_ADDUserCP=0
	Plus_ADDUserPower=Trim(Request("Plus_ADDUserPower"))
	If Plus_ADDUserPower="" Then Plus_ADDUserPower=0
	guestuse=Request("guestuse")
	tmpstr=useTime&","&timesetting&","&Groupsetting&","&masterlist&","&plus_UserPost&","
	tmpstr=tmpstr&Plus_userWealth&","&Plus_UserEP&","&Plus_UserCP&","&Plus_UserPower&","
	tmpstr=tmpstr&Plus_ADDuserWealth&","&Plus_ADDUserEP&","&Plus_ADDUserCP&","&Plus_ADDUserPower&","&guestuse
	Plus_SettingData=Plus_SettingData&tmpstr
	set rs=Dvbbs.iCreateObject("adodb.recordset")
	sql="select * from dv_plus"
	rs.open sql,conn,1,3
	rs.addnew
	Rs("plus_ID")=plusID
	rs("plus_type")=request("stype")
	rs("plus_name")=replace(request("title"),CHR(34),"")
	rs("isuse")=Isuse
	rs("IsShowMenu")=1
	rs("Mainpage")=replace(request("url"),CHR(34),"")
	rs("plus_Copyright")=replace(request("readme"),CHR(34),"")
	rs("Plus_Setting")=request("windowtype") & "|" & request("windowwidth") & "|" & request("windowheight")&"|||"&Plus_SettingData
	Rs("plus_adminpage")=plus_adminpage
	rs.update
	rs.close
	set rs=nothing
	dv_suc("新建论坛菜单成功")
	LoadForumPlusMenuCache
End sub
sub savedit()
	Dim plusID,plus_adminpage,Isuse,rs,sql
	plusID=Trim(Request("plusID"))
	If InStr(plusID,"'")>0 Then 
		Errmsg = Errmsg + "插件ID中不允许有单引号"
		Dvbbs_Error()
		exit sub
	End If
	If request("title")="" then
		Errmsg = Errmsg + "请输入菜单的标题！"
		Dvbbs_Error()
		exit sub
	End If
	If plusID="" Then
		Errmsg = Errmsg + "请设定插件ID"
		Dvbbs_Error()
		exit sub
	End If
	SQL="Select count(*) From Dv_plus where plus_ID='"&plusID&"' and id<>"&Dvbbs.CheckNumeric(request("id"))
	Set rs=Dvbbs.execute(SQL)
	If Rs(0) >0 Then
		Errmsg = Errmsg + "你设置的插件ID已经存在，请另行设置。"
		Dvbbs_Error()
		exit sub
	End If
	Isuse=Request("Isuse")
	If Isuse<>"1" And Isuse<>"0" Then Isuse="1"
	Isuse=CInt(Isuse)
	plus_adminpage=Dvbbs.Checkstr(Request("plus_adminpage"))
	Dim Plus_SettingData,Plus_Setting,i,tmpstr
	Plus_Setting=Request("Plus_Setting")
	Plus_Setting=Split(Plus_Setting,vbCrLf)
	Plus_SettingData=""
	For i=0 to UBound(Plus_Setting)
		Plus_Setting(i)=Split(Plus_Setting(i),"=")
		If UBound(Plus_Setting(i))=1 Then
			If Plus_SettingData="" Then 
				Plus_SettingData=Trim(Plus_Setting(i)(1))
				tmpstr=Trim(Plus_Setting(i)(0))
			Else
				Plus_SettingData=Plus_SettingData&","&Trim(Plus_Setting(i)(1))
				tmpstr=tmpstr&","&Trim(Plus_Setting(i)(0))
			End If
		End If
	Next
	Plus_SettingData=Plus_SettingData&"|||"&tmpstr&"|||"
	Dim plusmaster,masterlist
	plusmaster=Request("plusmaster")
	plusmaster=split(plusmaster,vbCrLf)
	masterlist=""
	For i=0 to UBound(plusmaster)
		If Trim(plusmaster(i)) <>"" Then
			If masterlist="" Then
				masterlist=plusmaster(i)
			Else
				masterlist=masterlist&"|"&plusmaster(i)
			End If
		End If
	Next
	Dim useTime,timesetting,Groupsetting,Plus_UserPost,Plus_userWealth,Plus_UserEP,Plus_UserCP
	Dim Plus_UserPower,Plus_ADDuserWealth,Plus_ADDUserEP,Plus_ADDUserCP,Plus_ADDUserPower,guestuse
	useTime=Request("useTime")
	If useTime="" Then useTime=0
	timesetting=Trim(Request("timesetting"))
	If timesetting="" Then timesetting="0|24"
	Groupsetting=Replace(Trim(Request("groupid"))&"",",","@")
	Plus_UserPost=Trim(Request("Plus_UserPost"))
	If Plus_UserPost="" Then Plus_UserPost=0
	Plus_userWealth=Trim(Request("Plus_userWealth"))
	If Plus_userWealth="" Then Plus_userWealth=0
	Plus_UserEP=Trim(Request("Plus_UserEP"))
	If Plus_UserEP="" Then Plus_UserEP=0
	Plus_UserCP=Trim(Request("Plus_UserCP"))
	If Plus_UserCP="" Then Plus_UserCP=0
	Plus_UserPower=Trim(Request("Plus_UserPower"))
	If Plus_UserPower="" Then Plus_UserPower=0
	Plus_ADDuserWealth=Trim(Request("Plus_ADDuserWealth"))
	If Plus_ADDuserWealth="" Then Plus_ADDuserWealth=0
	Plus_ADDUserEP=Trim(Request("Plus_ADDUserEP"))
	If Plus_ADDUserEP="" Then Plus_ADDUserEP=0
	Plus_ADDUserCP=Trim(Request("Plus_ADDUserCP"))
	If Plus_ADDUserCP="" Then Plus_ADDUserCP=0
	Plus_ADDUserPower=Trim(Request("Plus_ADDUserPower"))
	If Plus_ADDUserPower="" Then Plus_ADDUserPower=0
	guestuse=Request("guestuse")
	tmpstr=useTime&","&timesetting&","&Groupsetting&","&masterlist&","&plus_UserPost&","
	tmpstr=tmpstr&Plus_userWealth&","&Plus_UserEP&","&Plus_UserCP&","&Plus_UserPower&","
	tmpstr=tmpstr&Plus_ADDuserWealth&","&Plus_ADDUserEP&","&Plus_ADDUserCP&","&Plus_ADDUserPower&","&guestuse
	Plus_SettingData=Plus_SettingData&tmpstr
	set rs=Dvbbs.iCreateObject("adodb.recordset")
	sql="select * from dv_plus where id="&Dvbbs.CheckNumeric(request("id"))
	rs.open sql,conn,1,3
	Rs("plus_ID")=plusID
	rs("plus_type")=request("stype")
	rs("plus_name")=replace(request("title"),CHR(34),"")
	rs("isuse")=Isuse
	rs("IsShowMenu")=1
	rs("Mainpage")=replace(request("url"),CHR(34),"")
	rs("plus_Copyright")=replace(request("readme"),CHR(34),"")
	rs("Plus_Setting")=request("windowtype") & "|" & request("windowwidth") & "|" & request("windowheight")&"|||"&Plus_SettingData
	Rs("plus_adminpage")=plus_adminpage
	rs.update
	rs.close
	set rs=nothing
	dv_suc("修改论坛菜单成功")
	LoadForumPlusMenuCache
end sub


sub del()
	Dvbbs.Execute("Delete From Dv_Plus Where ID="&Dvbbs.CheckNumeric(Request("id")))
	dv_suc("删除论坛菜单成功")
	LoadForumPlusMenuCache
end Sub
Sub LoadForumPlusMenuCache()
	Dvbbs.Name="Plus_Settingts"
	Dim Rs,SQL
	SQL = "select plus_ID,Plus_Setting,Plus_Name,plus_Copyright from [Dv_plus] Order By ID"
	Set Rs = Dvbbs.Execute(SQL)	
	If Not Rs.Eof Then
		Dvbbs.Name="Plus_Settingts"
		Dvbbs.value = Rs.GetRows(-1)
	End If
	Set Rs = Nothing
	Dvbbs.LoadPlusMenu()
End Sub
Sub FixPlusTable()
	Dim Rs,SQL
	SQL="select * From Dv_plus"
	Set Rs=Dvbbs.Execute(SQL)
	If Rs.Fields.Count < 10 Then
		Set Rs=Nothing
		Dvbbs.Execute("alter table [Dv_plus] add plus_adminpage varchar(100)")
		Dvbbs.Execute("alter table [Dv_plus] add plus_id varchar(100)")
		Set Rs=Dvbbs.Execute(SQL)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
				Dvbbs.execute("update [Dv_plus] set plus_id='newplus"&Rs(0)&"' Where ID="&Rs(0)&"")
				Rs.MoveNext
			Loop
		End If
	End If
	Set Rs=Nothing
End Sub
%>