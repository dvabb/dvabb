<!--#include file="../conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="../inc/dv_clsother.asp"-->
<!--#include file="../inc/md5.asp"-->
<!--#include file="../inc/GroupPermission.asp"-->
<!--#include file="../dv_dpo/cls_dvapi.asp"-->
<%
Head()
Dim admin_flag,sqlstr,myrootid
FoundErr=False 
admin_flag=",16,"
CheckAdmin(admin_flag)

Dim tRs,UserInfo,UserTitle
UserMain(1)
Select Case Request("action")
Case "fix"
	Fixuser()
Case "userSearch"
	UserSearch()
Case "touser"
	ToUser()
Case "modify"
	UserModify()
Case "saveuserinfo"
	SaveUserInfo()
Case "UserPermission"
	UserPermission()
Case "UserBoardPermission"
	UserBoardPermission()
Case "saveuserpermission"
	SaveUserPermission()
Case "uniteuser"
	UniteUser()
Case "audituser" '审核用户
	audituser()
Case "saveaudit" '保存审核
	saveaudit()
Case Else
	UserIndex()
End Select

UserMain(0)
Footer()

'用户管理通用头部
Sub UserMain(Str)
	If Str = 1 Then
%>

<table cellpadding="2" cellspacing="1" border="0" width="100%" align=center>
<tr>
<th colspan=8 style="text-align:center;">用户管理</th>
</tr>
<tr>
<td width="20%" class=td2 align="center"><button Style="width:80;height:50;border: 1px outset;" class="button">注意事项</button></td>
<td width="80%" class=td2 colspan=7><li>①点删除按钮将删除所选定的用户，此操作是不可逆的；<li>②您可以批量移动用户到相应的组；<li>③点用户名进行相应的资料操作；<li>④点用户最后登陆IP可进行锁定IP操作；<li>⑤点用户Email将给该用户发送Email；<li>⑥点修复贴子将会修复该用户所发的贴子数据并更新其文章数，用于误删ID用户贴的修复。</td>
</tr>
<tr>
<td width=100% class=td2 colspan=8>
快速查看：<a href="user.asp">用户管理首页</a> | <a href="?action=userSearch&userSearch=1"><%If Request("userSearch")="1" Then%><font color=red><%End If%>所有用户<%If Request("userSearch")="1" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=2"><%If Request("userSearch")="2" Then%><font color=red><%End If%>发贴TOP100<%If Request("userSearch")="2" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=3"><%If Request("userSearch")="3" Then%><font color=red><%End If%>发贴END100<%If Request("userSearch")="3" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=4"><%If Request("userSearch")="4" Then%><font color=red><%End If%>24H内登录<%If Request("userSearch")="4" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=5"><%If Request("userSearch")="5" Then%><font color=red><%End If%>24H内注册<%If Request("userSearch")="5" Then%></font><%End If%></a><BR>
　　　　　<a href="?action=userSearch&userSearch=6"><%If Request("userSearch")="6" Then%><font color=red><%End If%>等待验证会员<%If Request("userSearch")="6" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=7"><%If Request("userSearch")="7" Then%><font color=red><%End If%>邮件验证<%If Request("userSearch")="7" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=8"><%If Request("userSearch")="8" Then%><font color=red><%End If%>管理　团队<%If Request("userSearch")="8" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=11"><%If Request("userSearch")="11" Then%><font color=red><%End If%>屏蔽　用户<%If Request("userSearch")="11" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=12"><%If Request("userSearch")="12" Then%><font color=red><%End If%>锁定 用户<%If Request("userSearch")="12" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=14"><%If Request("userSearch")="13" Then%><font color=red><%End If%>自定义权限用户<%If Request("userSearch")="13" Then%></font><%End If%></a>
 | <a href="?action=userSearch&userSearch=15"><%If Request("userSearch")="15" Then%><font color=red><%End If%>VIP用户<%If Request("userSearch")="15" Then%></font><%End If%></a>
 | <a href="?action=userSearch&userSearch=16"><%If Request("userSearch")="16" Then%><font color=red><%End If%>审核会员(包含用户组)<%If Request("userSearch")="16" Then%></font><%End If%></a>
 | <a href="?action=userSearch&userSearch=17"><%If Request("userSearch")="17" Then%><font color=red><%End If%>自定义审核会员<%If Request("userSearch")="17" Then%></font><%End If%></a>
</td>
</tr>
<tr>
<td width=100% class=td2 colspan=8>
功能选项：<a href="?action=uniteuser">合并用户</a> | <a href="update_user.asp">奖惩用户管理</a> <!--| <a href="boardmastergrade.asp">版主工作情况</a>-->
</td>
</tr>
<%
	Else
%>
</table>
<p></p>
<%
	End If
End Sub

'用户管理首页，搜索项
Sub UserIndex()
%>
<form action="?action=userSearch" method=post>
<tr>
<th colspan=7 style="text-align:center;">高级查询</th>
</tr>
<tr>
<td width="20%" class=td1>注意事项</td>
<td width="80%" class=td1 colspan=5>在记录很多的情况下搜索条件越多查询越慢，请尽量减少查询条件；最多显示记录数也不宜选择过大</td>
</tr>
<tr>
<td width="20%" class=td1>最多显示记录数</td>
<td width="80%" class=td1 colspan=5><input size=45 name="searchMax" type=text value=100></td>
</tr>
<tr>
<td width="20%" class=td1>用户名</td>
<td width="80%" class=td1 colspan=5><input size=45 name="username" type=text>&nbsp;<input type=checkbox class=checkbox name="usernamechk" value="yes" checked>用户名完整匹配</td>
</tr>
<tr>
<td width="20%" class=td1>用户组</td>
<td width="80%" class=td1 colspan=5>
<select size=1 name="usergroups">
<option value=0>任意</option>
<%
Dim rs
set rs=Dvbbs.Execute("select usergroupid,UserTitle,ParentGID from dv_usergroups Where Not ParentGID=0 order by ParentGID,usergroupid")
do while not rs.eof
response.write "<option value="&rs(0)&">"&SysGroupName(Rs(2)) & rs(1)&"</option>"
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
</td>
</tr>
<tr>
<td width="20%" class=td1>Email包含</td>
<td width="80%" class=td1 colspan=5><input size=45 name="userEmail" type=text></td>
</tr>
<tr>
<td width="20%" class=td1>用户IM包含</td>
<td width="80%" class=td1 colspan=5><input size=45 name="userim" type=text> 包括主页、OICQ、UC、ICQ、YAHOO、AIM、MSN</td>
</tr>
<tr>
<td width="20%" class=td1>登录IP包含</td>
<td width="80%" class=td1 colspan=5><input size=45 name="lastip" type=text></td>
</tr>
<tr>
<td width="20%" class=td1>头衔包含</td>
<td width="80%" class=td1 colspan=5><input size=45 name="usertitle" type=text></td>
</tr>
<tr>
<td width="20%" class=td1>签名包含</td>
<td width="80%" class=td1 colspan=5><input size=45 name="sign" type=text></td>
</tr>
<tr>
<td width="20%" class=td1>详细资料包含</td>
<td width="80%" class=td1 colspan=5><input size=45 name="userinfo" type=text></td>
</tr>
<!--shinzeal加入特殊搜索-->
<tr>
<th colspan=7 style="text-align:center;">特殊查询&nbsp;（注意： <多于> 或 <少于> 已默认包含 <等于>；条件留空则不使用此条件 ）</th>
</tr>
<tr>
<td class=td1 colspan=7>
<table ID="Table1" width="100%">
<tr>
<td width="50%" class=td2>登录次数:<input type=radio class=radio value=more name="loginR" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="loginR">&nbsp;少于&nbsp;&nbsp;<input size=5 name="loginT" type=text> 次</td>
<td width="50%" class=td2>消失天数:<input type=radio class=radio value=more name="vanishR" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="vanishR">&nbsp;少于&nbsp;&nbsp;<input size=5 name="vanishT" type=text> 天</td>
</tr>
<tr>
<td>注册天数:<input type=radio class=radio value=more name="regR" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="regR">&nbsp;少于&nbsp;&nbsp;<input size=5 name="regT" type=text> 天</td>
<td>发表帖数:<input type=radio class=radio value=more name="artcleR" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="artcleR">&nbsp;少于&nbsp;&nbsp;<input size=5 name="artcleT" type=text> 篇</td>
</tr>

<tr>
<td class=td2>用户金钱:<input type=radio class=radio value=more name="UWealth" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="UWealth">&nbsp;少于&nbsp;&nbsp;<input size=5 name="UWealth_value" type=text></td>
<td class=td2>用户积分:<input type=radio class=radio value=more name="UEP" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="UEP">&nbsp;少于&nbsp;&nbsp;<input size=5 name="UEP_value" type=text></td>
</tr>
<tr>
<td>用户魅力:<input type=radio class=radio value=more name="UCP" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="UCP">&nbsp;少于&nbsp;&nbsp;<input size=5 name="UCP_value" type=text></td>
<td>用户威望:<input type=radio class=radio value=more name="UPower" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="UPower">&nbsp;少于&nbsp;&nbsp;<input size=5 name="UPower_value" type=text></td>
</tr>
<tr>
<td class=td2>用户金币:<input type=radio class=radio value=more name="UMoney" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="UMoney">&nbsp;少于&nbsp;&nbsp;<input size=5 name="UMoney_value" type=text></td>
<td class=td2>用户点券:<input type=radio class=radio value=more name="UTicket" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="UTicket">&nbsp;少于&nbsp;&nbsp;<input size=5 name="UTicket_value" type=text></td>
</tr>
<tr>
<td class=td1><LI>以下条件请选取相应的VIP用户组进行查询</LI></td>
</tr>
<tr>
<td class=td2>Vip登记时间:<input type=radio class=radio value=more name="UVipStarTime" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="UVipStarTime">&nbsp;少于&nbsp;&nbsp;<input size=5 name="UVipStarTime_value" type=text></td>
<td class=td2>Vip截止时间:<input type=radio class=radio value=more name="UVipEndTime" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="UVipEndTime">&nbsp;少于&nbsp;&nbsp;<input size=5 name="UVipEndTime_value" type=text></td>
</tr>
</table>
</td></tr>

<!--特殊搜索结束-->
<tr>
<td width="100%" class=td1 align=center colspan=7><input name="submit" type=submit class=button value="   搜  索   "></td>
</tr>
<input type=hidden value="9" name="userSearch">
</form>
<%
End Sub
'用户批量审核
Sub audituser()
Dim Groupids
%>
<FORM METHOD=POST ACTION="?action=saveaudit" name="formaudituser">
<tr><th colspan=8 style="text-align:center;">用户批量设置审核</th></tr><tr><td width=20% class=td1>注意事项</td><td width=80% class=td1 colspan=5>多每个用户或每个用户组必须以<font color=red>","</font>符号分隔</td></tr>
<tr><td width=20% class=td1>例如：用户</td><td width=80% class=td1 colspan=5>用户1,用户2,...,用户n</td></tr>
<tr><td width=20% class=td1>例如：用户组</td><td width=80% class=td1 colspan=5>用户组ID1,用户组ID2,...,用户组IDn</td></tr>
<tr><td width=20% class=td1>例如：自定义用户</td><td width=80% class=td1 colspan=5>自定义用户,自定义用户,...,自定义用户n</td></tr>
<tr><td width=20% class=td1>审核权限区分:</td><td width=80% class=td1 colspan=5>自定义用户<font color=red>大于</font>用户组<font color=red>大于</font>用户</td></tr>
<tr><td width=20% class=td1>&nbsp;</td><td width=80% class=td1 colspan=5>
(1)如果自定义用户应用审核，则审核
(2)如果自定义用户不应用审核，则根据用户组是否应用审核
(3)如果用户组不应用审核，则根据用户是否应用审核</td></tr>
<tr><td width=20% class=td1>&nbsp;</td><td width=80% class=td1 colspan=5><input type="radio" id="audittype" name="audittype" value="1" checked>应用审核&nbsp;&nbsp;&nbsp;<input type="radio" id="audittype" name="audittype" value="0">取消审核</td></td></tr>
<tr><td width=20% class=td1>用户</td><td width=80% class=td1 colspan=5><input type="text" id="usernames" name="usernames" size="100"></td></tr>
<tr><td width=20% class=td1>用户组</td><td width=80% class=td1 colspan=5><input type=text name="groupid" value="" size="100"><input type="button" class="button" value="选择用户组" onclick="getGroup('Select_Group');"></td></tr>
<tr><td width=20% class=td1>自定义用户</td><td width=80% class=td1 colspan=5><input type="text" id="customusernames" name="customusernames" size="100"></td></tr>
<tr><td width=20% class=td1>所有用户</td><td width=80% class=td1 colspan=5><input type="checkbox" id="userall" name="userall" >&nbsp针对所有用户</td></tr>
<tr><td width=20% class=td1>&nbsp;</td><td width=80% class=td1 colspan=5><INPUT TYPE="submit" value="提交">&nbsp;&nbsp;&nbsp;<INPUT TYPE="reset" value="重新设置"></td></tr>
<%
Call Select_Audit_Group(Replace(Groupids&"","@",","))
End Sub 
Sub saveaudit()
Dim i
Dim audittype,usernames,groupid,customusernames,userall
audittype=Dvbbs.CheckStr(Request("audittype"))
usernames=Dvbbs.CheckStr(request("usernames"))
groupid=Dvbbs.CheckStr(request("groupid"))
customusernames=Dvbbs.CheckStr(request("customusernames"))
If Dvbbs.CheckStr(Request("userall"))="on" Then userall=True 
If Not userall And usernames="" And groupid="" And customusernames="" Then 
			ErrMsg =ErrMsg& "请设置用户或用户组或自定义用户。"
			founderr=true
			Dvbbs_Error()        
Else 
		If userall Then 
			Dvbbs.Execute("Update Dv_User set UserIsAudit_Custom="&audittype&",userisaudit="&audittype&"")
            Dvbbs.Execute("Update Dv_UserGroups set UserGroupIsAudit="&audittype&"")
			Dv_suc("所有用户设置审核成功！")
		Else 
			If usernames<>"" Then 
			updateaudit usernames,0,audittype
			Dv_suc("用户设置审核成功！")
			End If 
			If groupid<>"" Then 
			updateaudit groupid,1,audittype
			Dv_suc("用户组设置审核成功！")
			End If 
			If customusernames<>"" Then 
			updateaudit customusernames,2,audittype
			Dv_suc("自定义用户设置审核成功！")
			End If 
		End If 
End If 
End Sub 
Function updateaudit(value,type1,audittype) '参1=表单值 参2=用户或用户组或自定义用户 参3=操作类型
Dim usernames,groupid,SQL,Rs,i,ToUserID,ToUserGroupId
Select Case type1
Case "0"
				usernames=Split(value,",")
				If Ubound(usernames)>100 Then
				ErrMsg =ErrMsg& "限制一次不能超过100位目标用户。"
				founderr=true
				Dvbbs_Error()
				Exit Function 
				End If 
				For i=0 To Ubound(usernames)
					SQL = "Select UserID From [Dv_user] Where UserName = '"&usernames(i)&"' order by userid"
					SET Rs = Dvbbs.Execute(SQL)
					If Not Rs.eof Then
						If i=0 or ToUserID="" Then
							ToUserID = ToUserID & Rs(0)
						Else
							ToUserID = ToUserID &","& Rs(0)
						End If
					Else 
						ErrMsg =ErrMsg&"<font color=red>"&usernames(i)& "</font>用户不存在。"
						founderr=true
						Dvbbs_Error()
						Exit Function
					End If
				Next
				Rs.Close : Set Rs = Nothing
				Dvbbs.Execute("Update dv_user set userisaudit="&audittype&" where userid in ("&ToUserID&")")
Case "1"
			groupid=Split(value,",")
			For i=0 To UBound(groupid)
				SQL = "select usergroupid,title from dv_usergroups where usergroupid="&groupid(i)&""
				SET Rs = Dvbbs.Execute(SQL)
				If Rs.eof Then 
							ErrMsg =ErrMsg&"ID为<font color=red>"&groupid(i)& "</font>的用户组不存在。"
							founderr=true
							Dvbbs_Error()
							Exit Function
				End If 
			Next 
			Rs.Close : Set Rs = Nothing
			Dvbbs.Execute("update Dv_UserGroups set UserGroupIsAudit="&audittype&" where UserGroupID in ("&value&")")
Case "2"
				usernames=Split(value,",")
				If Ubound(usernames)>100 Then
				ErrMsg =ErrMsg& "限制一次不能超过100位自定义用户。"
				founderr=true
				Dvbbs_Error()
				Exit Function 
				End If 
				For i=0 To Ubound(usernames)
					SQL = "Select UserID From [Dv_user] Where UserName = '"&usernames(i)&"' order by userid"
					SET Rs = Dvbbs.Execute(SQL)
					If Not Rs.eof Then
						If i=0 or ToUserID="" Then
							ToUserID = ToUserID & Rs(0)
						Else
							ToUserID = ToUserID &","& Rs(0)
						End If
					Else 
						ErrMsg =ErrMsg&"<font color=red>"&usernames(i)& "</font>用户不存在。"
						founderr=true
						Dvbbs_Error()
						Exit Function
					End If
				Next
				Rs.Close : Set Rs = Nothing
				Dvbbs.Execute("Update dv_user set UserIsAudit_Custom="&audittype&" where userid in ("&ToUserID&")")
End Select 
End Function

%>
</FORM>
<%
Sub UserSearch()
%>
<tr>
<th colspan=8 style="text-align:center;">搜索结果</th>
</tr>
<%
	dim currentpage,page_count,Pcount
	dim totalrec,endpage
	Dim rs,sql
	currentPage=request("page")
	if currentpage="" or not IsNumeric(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
		if err then
			currentpage=1
			err.clear
		end if
	end if
	Sql = " Userid, Username, Useremail, LastLogin, UserLastIP, UserPost, UserGroupID,Vip_StarTime,Vip_EndTime"
	Set Rs = Dvbbs.iCreateObject("ADODB.Recordset")
	Select Case Request("UserSearch")
	Case 1
		Sql = "SELECT " & Sql & " FROM [Dv_User] ORDER BY UserID DESC"
	Case 2
		Sql = "SELECT TOP 100 " & Sql & " FROM [Dv_User] ORDER BY UserPost DESC"
	case 3
		sql="select top 100 " & Sql & " from [dv_user]  order by UserPost"
	case 4
		If IsSqlDataBase=1 Then
		sql="select " & Sql & " from [dv_user]  where datediff(hour,LastLogin,"&SqlNowString&")<25 order by lastlogin desc"
		else
		sql="select " & Sql & " from [dv_user]  where datediff('h',LastLogin,"&SqlNowString&")<25 order by lastlogin desc"
		end if
	case 5
		If IsSqlDataBase=1 Then
		sql="select " & Sql & " from [dv_user]  where datediff(hour,JoinDate,"&SqlNowString&")<25 order by UserID desc"
		else
		sql="select " & Sql & " from [dv_user]  where datediff('h',JoinDate,"&SqlNowString&")<25 order by UserID desc"
		end if
	case 6
		sql="select " & Sql & " from [dv_user]  where usergroupid=5 order by UserID desc"
	case 7
		sql="select " & Sql & " from [dv_user]  where usergroupid=6 order by UserID desc"
	case 8
		sql="select " & Sql & " from [dv_user]  where usergroupid<4 order by usergroupid"
	case 10
		Sql = "select " & Sql & " from [dv_user]  where usergroupid="&request("usergroupid")&" order by UserID desc"
	case 11
		sql="select " & Sql & " from [dv_user]  where lockuser=2 order by userid desc"
	case 12
		sql="select " & Sql & " from [dv_user]  where lockuser=1 order by userid desc"
	case 13
		sql="select " & Sql & " from [dv_user]  where IsChallenge=1 order by userid desc"
	case 14
		Sql = "SELECT " & Sql & " FROM [Dv_User] WHERE UserID IN (SELECT Uc_UserID FROM Dv_UserAccess) ORDER BY Userid DESC"
	case 15
		Sql = "SELECT " & Sql & " FROM [dv_user]  WHERE UserGroupid IN (SELECT UserGroupID FROM Dv_UserGroups WHERE ParentGID=5) ORDER BY Vip_EndTime desc,UserID desc"
	case 16
        Sql="select u.Userid, u.Username, u.Useremail, u.LastLogin, u.UserLastIP, u.UserPost, u.UserGroupID,u.Vip_StarTime,u.Vip_EndTime from dv_user u inner join Dv_UserGroups g on u.usergroupid=g.Usergroupid where g.UserGroupIsAudit=1 or u.userisaudit=1 order by u.userid desc"
	case 17
        Sql="select u.Userid, u.Username, u.Useremail, u.LastLogin, u.UserLastIP, u.UserPost, u.UserGroupID,u.Vip_StarTime,u.Vip_EndTime from dv_user u where u.UserIsAudit_Custom=1 order by u.userid desc"
	case 9
		sqlstr=""
		if request("username")<>"" then
			if request("usernamechk")="yes" then
			sqlstr=" username='"&request("username")&"'"
			else
			sqlstr=" username like '%"&request("username")&"%'"
			end if
		end if
		if cint(request("usergroups"))>0 then
			if sqlstr="" then
			sqlstr=" usergroupid="&request("usergroups")&""
			else
			sqlstr=sqlstr & " and usergroupid="&CheckNumeric(request("usergroups"))
			end if
		end if
		'if request("userclass")<>"0" then
		'	if sqlstr="" then
		'	sqlstr=" userclass='"&request("userclass")&"'"
		'	else
		'	sqlstr=sqlstr & " and userclass='"&request("userclass")&"'"
		'	end if
		'end if

		'======shinzeal加入特殊搜索=======
		dim Tsqlstr
		if request("loginT")<>"" then
		   	if request("loginR")="more" then
			 Tsqlstr=" userlogins >= "&CheckNumeric(request("loginT"))
			else
			 Tsqlstr=" userlogins <= "&CheckNumeric(request("loginT"))
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if 
		end if

		if request("vanishT")<>"" then
		   	if request("vanishR")="more" then
				If IsSqlDataBase=1 Then
					Tsqlstr=" datediff(d,lastlogin,"&SqlNowString&") >= "&CheckNumeric(request("vanishT"))
				Else
					Tsqlstr=" datediff('d',lastlogin,"&SqlNowString&") >= "&CheckNumeric(request("vanishT"))
				End If
			else
				If IsSqlDataBase=1 Then
					Tsqlstr=" datediff(d,lastlogin,"&SqlNowString&") <= "&CheckNumeric(request("vanishT"))
				Else
					Tsqlstr=" datediff('d',lastlogin,"&SqlNowString&") <= "&CheckNumeric(request("vanishT"))
				End If
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if 
		end if

		if request("regT")<>"" then
		   	if request("regR")="more" then
				If IsSqlDataBase=1 Then
					Tsqlstr=" datediff(d,JoinDate,"&SqlNowString&") >= "&CheckNumeric(request("regT"))
				Else
					Tsqlstr=" datediff('d',JoinDate,"&SqlNowString&") >= "&CheckNumeric(request("regT"))
				End If
			else
				If IsSqlDataBase=1 Then
					Tsqlstr=" datediff(d,JoinDate,"&SqlNowString&") <= "&CheckNumeric(request("regT"))
				Else
					Tsqlstr=" datediff('d',JoinDate,"&SqlNowString&") <= "&CheckNumeric(request("regT"))
				End If
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if 
		end if

		if request("artcleT")<>"" then
		   	if request("artcleR")="more" then
			 Tsqlstr=" UserPost >= "&CheckNumeric(request("artcleT"))
			else
			 Tsqlstr=" UserPost <= "&CheckNumeric(request("artcleT"))
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if 
		end if

		if request("UWealth_value")<>"" then
			if request("UWealth")="more" then
				Tsqlstr=" userWealth >= "&CheckNumeric(Request("UWealth_value"))
			else
				Tsqlstr=" userWealth <= "&CheckNumeric(Request("UWealth_value"))
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if
		end if

		if request("UEP_value")<>"" then
			if request("UEP")="more" then
				Tsqlstr=" userEP >= "&CheckNumeric(Request("UEP_value"))
			else
				Tsqlstr=" userEP <= "&CheckNumeric(Request("UEP_value"))
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if
		end if

		if request("UCP_value")<>"" then
			if request("UCP")="more" then
				Tsqlstr=" userCP >= "&CheckNumeric(Request("UCP_value"))
			else
				Tsqlstr=" userCP <= "&CheckNumeric(Request("UCP_value"))
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if
		end if

		if request("UPower_value")<>"" then
			if request("UPower")="more" then
				Tsqlstr=" UserPower >= "&CheckNumeric(Request("UPower_value"))
			else
				Tsqlstr=" UserPower <= "&CheckNumeric(Request("UPower_value"))
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if
		end if

		if request("UMoney_value")<>"" then
			if request("UMoney")="more" then
				Tsqlstr=" UserMoney >= "&CheckNumeric(Request("UMoney_value"))
			else
				Tsqlstr=" UserMoney <= "&CheckNumeric(Request("UMoney_value"))
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if
		end if

		if request("UTicket_value")<>"" then
			if request("UTicket")="more" then
				Tsqlstr=" UserTicket >= "&CheckNumeric(Request("UTicket_value"))
			else
				Tsqlstr=" UserTicket <= "&CheckNumeric(Request("UTicket_value"))
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if
		end if

		if request("UVipStarTime_value")<>"" then
		   	if request("UVipStarTime")="more" then
				If IsSqlDataBase=1 Then
					Tsqlstr=" datediff(d,Vip_StarTime,"&SqlNowString&") >= "&CheckNumeric(request("UVipStarTime_value"))
				Else
					Tsqlstr=" datediff('d',Vip_StarTime,"&SqlNowString&") >= "&CheckNumeric(request("UVipStarTime_value"))
				End If
			else
				If IsSqlDataBase=1 Then
					Tsqlstr=" datediff(d,Vip_StarTime,"&SqlNowString&") <= "&CheckNumeric(request("UVipStarTime_value"))
				Else
					Tsqlstr=" datediff('d',Vip_StarTime,"&SqlNowString&") <= "&CheckNumeric(request("UVipStarTime_value"))
				End If
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if 
		end if
		if request("UVipEndTime_value")<>"" then
		   	if request("UVipEndTime")="more" then
				If IsSqlDataBase=1 Then
					Tsqlstr=" datediff(d,Vip_EndTime,"&SqlNowString&") >= "&CheckNumeric(request("UVipEndTime_value"))
				Else
					Tsqlstr=" datediff('d',Vip_EndTime,"&SqlNowString&") >= "&CheckNumeric(request("UVipEndTime_value"))
				End If
			else
				If IsSqlDataBase=1 Then
					Tsqlstr=" datediff(d,Vip_EndTime,"&SqlNowString&") <= "&CheckNumeric(request("UVipEndTime_value"))
				Else
					Tsqlstr=" datediff('d',Vip_EndTime,"&SqlNowString&") <= "&CheckNumeric(request("UVipEndTime_value"))
				End If
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if 
		end if

		'======特殊搜索结束======
		if request("useremail")<>"" then
			if sqlstr="" then
			sqlstr=" useremail like '%"&request("useremail")&"%'"
			else
			sqlstr=sqlstr & " and useremail like '%"&request("useremail")&"%'"
			end if
		end if
		if request("userim")<>"" then
			if sqlstr="" then
			sqlstr=" UserIM like '%"&request("userim")&"%'"
			else
			sqlstr=sqlstr & " and UserIM like '%"&request("userim")&"%'"
			end if
		end if
		if request("lastip")<>"" then
			if sqlstr="" then
			sqlstr=" UserLastIP like '%"&request("lastip")&"%'"
			else
			sqlstr=sqlstr & " and UserLastIP like '%"&request("lastip")&"%'"
			end if
		end if
		if request("userinfo")<>"" then
			if sqlstr="" then
			sqlstr=" UserInfo like '%"&request("userinfo")&"%'"
			else
			sqlstr=sqlstr & " and UserInfo like '%"&request("userinfo")&"%'"
			end if
		end if
		'修正不能用头衔搜索 2005-4-9 Dv.Yz
		If Request("usertitle") <> "" Then
			If Sqlstr = "" Then
				Sqlstr = " UserTitle LIKE '%" & Request("usertitle") & "%'"
			Else
				Sqlstr = Sqlstr & " AND UserTitle LIKE '%" & Request("usertitle") & "%'"
			End If
		End If
		if request("sign")<>"" then
			if sqlstr="" then
			sqlstr=" usersign like '%"&request("sign")&"%'"
			else
			sqlstr=sqlstr & " and usersign like '%"&request("sign")&"%'"
			end if
		end if

		If Sqlstr = "" Then
			Response.Write "<tr><td colspan=8 class=td1>请指定搜索参数！</td></tr>"
			Response.End
		End If
		If Request("Searchmax") = "" Or Not Isnumeric(Request("Searchmax")) Then
			Sql = "SELECT TOP 1 "& Sql &" FROM [Dv_User] WHERE " & Sqlstr & " ORDER BY UserID DESC"
		Else
			Sql = "SELECT TOP " & Request("Searchmax") & Sql &" FROM [Dv_User] WHERE " & Sqlstr & " ORDER BY UserID DESC"
		End If
	case else
		Response.Write "<tr><td colspan=8 class=td1>错误的参数。</td></tr>"
		Response.End
	End Select
	'Response.Write sql
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "<tr><td colspan=8 class=td1>没有找到相关记录。"
		If Request("userSearch")="15" Then
			Response.Write "（若未添加VIP用户组，请<a href=""group.asp""><font color=red>点击进入论坛用户组管理</font></a>进行添加。）"
		End If
		Response.Write "</td></tr>"
	else
%>
<FORM METHOD=POST ACTION="?action=touser">
<tr align=center height=23>
<td class=td2 width="10%"><B>用户名</B></td>
<td class=td2 width="15%"><B>Email</B></td>
<td class=td2 width="8%"><B>权限</B></td>
<td class=td2 width="8%"><B>数据修复</B></td>
<td class=td2 width="15%"><B>最后IP</B></td>
<td class=td2 width="15%"><B>最后登录</B></td>
<td class=td2 width="20%"><B>登记/终止日期</B></td>
<td class=td2><B>操作</B></td>
</tr>
<%
		rs.PageSize = Cint(Dvbbs.Forum_Setting(11))
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		while (not rs.eof) and (not page_count = Cint(Dvbbs.Forum_Setting(11)))
%>
<tr>
<td class=td1><a href="?action=modify&userid=<%=rs("userid")%>"><%=rs("username")%></a></td>
<td class=td1><a href="mailto:<%=rs("useremail")%>"><%=rs("useremail")%></a></td>
<td class=td1 align=center><a href="?action=UserPermission&userid=<%=rs("userid")%>&username=<%=rs("username")%>">编辑</a></td>
<td class=td1 align=center><a href="?action=fix&userid=<%=rs("userid")%>&username=<%=rs("username")%>">修复</a></td>
<td class=td1><a href="lockIP.asp?userip=<%=rs("UserLastIP")%>" title="点击锁定该用户IP"><%=rs("userlastip")%></a>&nbsp;</td>
<td class=td1><%if rs("lastlogin")<>"" and isdate(rs("lastlogin")) then%><%=rs("lastlogin")%><%end if%></td>
<td class=td1 align=center>
<%=rs("Vip_StarTime")%>/
<%=rs("Vip_EndTime")%>
</td>
<td class=td1 align=center><input type="checkbox" class=checkbox name="userid" value="<%=rs("userid")%>" <%if rs("userGroupid")=1 then response.write "disabled"%>></td>
</tr>
<%
		page_count = page_count + 1
		rs.movenext
		wend
		Pcount=rs.PageCount
%>
<tr><td colspan=8 class=td1 align=center>分页：
<%
Dim Searchstr,i
'修正头衔搜索用户的分页错误。
'修正最后登陆IP搜索用户的分页错误 2005.10.12 By Winder
Searchstr = "?userSearch=" & Request("userSearch") & "&username=" & Request("username") & "&useremail=" & Request("useremail") & "&userim=" & Request("userim") & "&lastip=" & Request("lastip") & "&usertitle=" & Request("usertitle") & "&sign=" & Request("sign") & "&userinfo=" & Request("userinfo") & "&action=" & Request("action") & "&loginR=" & Request("loginR") & "&loginT=" & Request("loginT") & "&vanishR=" & Request("vanishR") & "&vanishT=" & Request("vanishT") & "&regR=" & Request("regR") & "&regT=" & Request("regT") & "&artcleR=" & Request("artcleR") & "&artcleT=" & Request("artcleT") & "&UWealth=" & Request("UWealth") & "&UWealth_value=" & Request("UWealth_value") & "&UEP=" & Request("UEP") & "&UEP_value=" & Request("UEP_value") & "&UCP=" & Request("UCP") & "&UCP_value=" & Request("UCP_value") & "&UPower=" & Request("UPower") & "&UPower_value=" & Request("UPower_value") & "&UMoney=" & Request("UMoney") & "&UMoney_value=" & Request("UMoney_value") & "&UTicket=" & Request("UTicket") & "&UTicket_value=" & Request("UTicket_value") & "&searchmax=" & Request("searchmax") & "&UVipStarTime=" & Request("UVipStarTime") & "&UVipStarTime_value=" & Request("UVipStarTime_value") & "&UVipEndTime=" & Request("UVipEndTime") & "&UVipEndTime_value=" & Request("UVipEndTime_value")&"&usergroups="&Request("usergroups")&"&usergroupid="&Request("usergroupid")

	if currentpage > 4 then
		response.write "<a href="""&Searchstr&"&page=1"">[1]</a> ..."
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
        response.write " <a href="""&Searchstr&"&page="&i&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
		response.write "... <a href="""&Searchstr&"&page="&Pcount&""">["&Pcount&"]</a>"
	end if
%>
</td></tr>
<tr><td colspan=5 class=td1 align=center><B>请选择您需要进行的操作</B>：<input type="radio" class=radio name="useraction" value=1> 删除&nbsp;&nbsp;<input type="radio" class=radio name="useraction" value=3> 删除用户所有帖子&nbsp;&nbsp;<input type="radio" class=radio name="useraction" value=2 checked> 移动到用户组
<select size=1 name="selusergroup">
<%
set trs=Dvbbs.Execute("select usergroupid,UserTitle,ParentGID from dv_usergroups where not (usergroupid=1 or usergroupid=7) and (Not ParentGID=0) order by ParentGID,usergroupid")
do while not trs.eof
response.write "<option value="&trs(0)&">"&SysGroupName(tRs(2))&trs(1)&"</option>"
trs.movenext
loop
trs.close
set trs=nothing
%>
</select>
</td>
<td class=td1 colspan=8 align=center>全部选定<input type=checkbox class=checkbox value="on" name="chkall" onclick="CheckAll(this.form)">
</td>
</tr>
<tr><td colspan=8 class=td1 align=center>
<input type=submit class=button name=submit value="执行选定的操作"  onclick="{if(confirm('确定执行选择的操作吗?')){return true;}return false;}">
</td></tr>
</FORM>
<%
	end if
	rs.close
	set rs=nothing
End Sub

'操作用户，删除用户信息相关操作
Sub ToUser()
	Dim SQL,rs
	response.write "<tr><th colspan=8 style=""text-align:center;"">执行结果</th></tr>"
	if request("useraction")="" then
		response.write "<tr><td colspan=8 class=td1>请指定相关参数。</td></t粍
<select size=1 name="selusergroup">
<%
set trs=Dvbbs.Execute("select usergroupid,UserTitle,ParentGID from dv_usergroups where not (usergroupid=1 or usergroupid=7) and (Not ParentGID=0) order by ParentGID,usergroupid")
do while not trs.eof
response.write "<option value="&trs(0)&">"&SysGroupName(tRs(2))&trs(1)&"</option>"
trs.movenext
loop
trs.close
set trs=nothing
%>
</select>
</td>
<td class=td1 colspan=8 align=center>鍏ㄩ儴閫夊畾<input type=checkbox class=checkbox value="on" name="chkall" onclick="CheckAll(this.form)">
</td>
</tr>
<tr><td colspan=8 class=td1 align=center>
<input type=submit class=button name=submit value="鎵ц