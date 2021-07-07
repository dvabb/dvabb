<!--#include file="../conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
Head()
Server.ScriptTimeout=9999999
dim iboardid(2000),idepth(2000),iboardname(2000)
dim k,n,i
dim admin_flag
admin_flag="24"
CheckAdmin(admin_flag)
Dim body
Call main()
Footer()

Erase iboardid
Erase idepth
Erase iboardname

sub main()
Dim rs
i=0
set rs=Dvbbs.Execute("select boardid,depth,boardtype from dv_board order by rootid,orders")
if rs.eof and rs.bof then
	iboardid(0)=0
	idepth(0)=0
	iboardname(0)="没有论坛"
else
	do while not rs.eof
		iboardid(i)=rs(0)
		idepth(i)=rs(1)
		iboardname(i)=rs(2)
		i=i+1
		rs.movenext
	loop
end if
set rs=nothing
select case request("action")
case "alldel"
	call alldel()
case "userdel"
	call del()
case "alldelTopic"
	call alldelTopic()
case "delUser"
	call deluser()
case "moveinfo"
	call moveinfo()
case "MoveUserTopic"
	call moveusertopic()
case "MoveDateTopic"
	call movedatetopic()
case else
%>
<table cellpadding=3 cellspacing=1 border=0 width=100% align="center">
	<tr>
    <td width="100%" valign=top class=td1>
<B>注意</B>：下面操作将大批量删除论坛帖子，<font color=red>并且所有操作不可恢复！</font>如果您确定这样做，请仔细检查您输入的信息。
</td>
</tr>
</table><br />
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<form action="alldel.asp?action=alldel" method="post">
	<th valign=middle colspan=2>删除指定日期内帖子</b>(本功能不扣除用户帖子数和积分)</th>
	<tr>
	<td valign=middle width=40% class=td1>删除多少天前的帖子(填写数字)</td><td class=td1><input name="TimeLimited" value=100 size=30>&nbsp;<input type=submit class="button" name="submit" value="提 交"></td></tr>
	<tr>
	<td valign=middle width=40%  class=td1>论坛版面</td>
	<td class=td1>
		<select name="delboardid" size=1>
		<%
		for k=0 to i-1
			if iboardid(k)=0 then
				response.write "<option value=0>没有论坛</option>"
			elseif k=0 then
				response.write "<option value=all>全部论坛</option>"
			end if
			response.write "<option value="&iboardid(k)&">"
			if idepth(k)>0 then
			for n=1 to idepth(k)
			response.write "－"
			next
			end if
			response.write iboardname(k)&"</option>"
		next
		%>
		</select>
	</td></tr>
</form>
<form action="alldel.asp?action=alldelTopic" method="post">
	<th valign=middle colspan=2>删除指定日期内没有回复的主题(本功能不扣除用户帖子数和积分)</th>
	<tr>
	<td valign=middle width=40% class=td1>删除多少天前的帖子(填写数字)</td><td class=td1><input name="TimeLimited" value=100 size=30>&nbsp;<input type=submit class="button" name="submit" value="提 交"></td></tr>
	<tr>
	<td valign=middle width=40% class=td1>论坛版面</td>
	<td class=td1>
		<select name="delboardid" size=1>
		<%
		for k=0 to i-1
			if iboardid(k)=0 then
				response.write "<option value=0>没有论坛</option>"
			elseif k=0 then
				response.write "<option value=all>全部论坛</option>"
			end if
			response.write "<option value="&iboardid(k)&">"
			if idepth(k)>0 then
			for n=1 to idepth(k)
			response.write "－"
			next
			end if
			response.write iboardname(k)&"</option>"
		next
		%>
		</select>
	</td></tr>
</form>
<form action="alldel.asp?action=userdel" method="post">
	<th valign=middle colspan=2>删除某用户的所有帖子</th>
	<tr>
	<td valign=middle width=40% class=td1>请输入用户名</td><td class=td1><input type=text name="username" size=30>&nbsp;<input type=submit class="button" name="submit" value="提 交"></td></tr>
	<tr>
	<td valign=middle width=40%  class=td1>论坛版面</td><td class=td1>
		<select name="delboardid" size=1>
		<%
		for k=0 to i-1
			if iboardid(k)=0 then
				response.write "<option value=0>没有论坛</option>"
			elseif k=0 then
				response.write "<option value=all>全部论坛</option>"
			end if
			response.write "<option value="&iboardid(k)&">"
			if idepth(k)>0 then
			for n=1 to idepth(k)
			response.write "－"
			next
			end if
			response.write iboardname(k)&"</option>"
		next
		%>
		</select>
	</td></tr>
</form>

<form action="alldel.asp?action=delUser" method="post">
	<th valign=middle colspan=2>删除指定日期内没有登录的用户</th>
	<tr>
	<td class=td1 valign=middle>指定日期</td>
	<td class=td1 valign=middle>
		<select name=TimeLimited size=1> 
		<option value=1>删除一天前的
		<option value=2>删除两天前的
		<option value=7>删除一星期前的
		<option value=15>删除半个月前的
		<option value=30>删除一个月前的
		<option value=60>删除两个月前的
		<option value=180>删除半年前的
		</select>
		<input type=submit class="button" name="submit" value="提 交">
	</td></tr>
</form>

</table>
<%end select%>
<%if founderr then Call dvbbs_error()%>
<%
end sub

Sub Moveinfo()
%>
<table cellpadding=3 cellspacing=1 border=0 width=100% align=center>
	<tr>
    <td width="100%" valign=top>
<B>注意</B>：这里只是移动帖子，而不是拷贝或者删除！
            <br>下面操作将删除原论坛帖子，并移动到您指定的论坛中。如果您确定这样做，请仔细检查您输入的信息。<br />您可以将一个论坛下属论坛的帖子移动到上级论坛，也可以将上级论坛的帖子移动到下级论坛，但作为分类的论坛由于论坛设置很可能不能发布帖子（只能浏览）
</td>
</tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<form action="alldel.asp?action=MoveDateTopic" method="post">
	<th valign=middle colspan=2>按日期移动</th>
	<tr>
	<td valign=middle width=40% class=td1>移动多少天前的帖子(填写数字)<li>填写为0，即分版所有帖子。</td><td class=td1><input name="TimeLimited" value=0 size=30>天&nbsp;<input type=submit class="button" name="submit" value="提 交"></td></tr>
	<tr>
	<td valign=middle width=40% class=td1>原论坛</td><td class=td1>
		<select name="outboardid" size=1>
		<%
		for k=0 to i-1
			if iboardid(k)=0 then
				response.write "<option value=0>没有论坛</option>"
			end if
			response.write "<option value="&iboardid(k)&">"
			if idepth(k)>0 then
			for n=1 to idepth(k)
			response.write "－"
			next
			end if
			response.write iboardname(k)&"</option>"
		next
		%>
		</select>
	</td></tr>
	<tr>
	<td valign=middle width=40% class=td1>目标论坛</td><td class=td1>
		<select name="inboardid" size=1>
		<%
		for k=0 to i-1
			if iboardid(k)=0 then
				response.write "<option value=0>没有论坛</option>"
			end if
			response.write "<option value="&iboardid(k)&">"
			if idepth(k)>0 then
			for n=1 to idepth(k)
			response.write "－"
			next
			end if
			response.write iboardname(k)&"</option>"
		next
		%>
		</select>
	</td></tr>
</form>
<form action="alldel.asp?action=MoveUserTopic" method="post">
	<th valign=middle colspan=2>按用户移动</th>
	<tr>
	<td valign=middle width=40% class=td1>请填写用户名</td><td class=td1><input name="username" size=30>&nbsp;<input type=submit class="button" name="submit" value="提 交"></td></tr>
	<tr>
	<td valign=middle width=40% class=td1>原论坛</td><td class=td1>
		<select name="outboardid" size=1>
		<%
		for k=0 to i-1
			if iboardid(k)=0 then
				response.write "<option value=0>没有论坛</option>"
			end if
			response.write "<option value="&iboardid(k)&">"
			if idepth(k)>0 then
			for n=1 to idepth(k)
			response.write "－"
			next
			end if
			response.write iboardname(k)&"</option>"
		next
		%>
		</select>
	</td></tr>
	<tr>
	<td valign=middle width=40% class=td1>目标论坛</td><td class=td1>
		<select name="inboardid" size=1>
		<%
		for k=0 to i-1
			if iboardid(k)=0 then
				response.write "<option value=0>没有论坛</option>"
			end if
			response.write "<option value="&iboardid(k)&">"
			if idepth(k)>0 then
			for n=1 to idepth(k)
			response.write "－"
			next
			end if
			response.write iboardname(k)&"</option>"
		next
		%>
		</select>
	</td></tr>
</form>
</table>
<%
	end sub

'删除某用户的所有帖子
Sub Del()
	dim titlenum,delboardid,PostUserID,delboardida,rs,sql,i
	Dim Dnum '删除帖子数 2005-10-30 Dv.Yz
	If request("delboardid")="0" then
		founderr=true
		Errmsg=ErrMsg + "<br /><li>非法的版面参数。"
		exit sub
	Elseif request("delboardid")="all" then
		delboardid=""
		delboardida=""
	Else
		delboardid=" boardid="&Dvbbs.CheckNumeric(request("delboardid"))&" and "
		delboardida=" F_boardid="&Dvbbs.CheckNumeric(request("delboardid"))&" and "
	End if
	If Request("username")="" then
		founderr=true
		Errmsg=ErrMsg + "<br /><li>请输入被帖子删除用户名。"
		exit sub
	End If
	Set Rs=Dvbbs.Execute("Select UserID,UserGroupID From Dv_User Where UserName='"&replace(request("username"),"'","")&"'")
	If Rs.Eof And Rs.Bof Then
		founderr=true
		Errmsg=ErrMsg + "<br /><li>目标用户不存在，请重新输入。"
		exit sub
	End If
	If Rs(1)=1 Or Rs(1)=2 Or Rs(1)=3 Then
		founderr=true
		Errmsg=ErrMsg + "<br /><li>对管理员、超级版主、版主的贴子不能进行批量删除操作。"
		exit sub
	End If
		PostUserID=Rs(0)
		Set Rs=Nothing
		titlenum=0
		for i=0 to ubound(allposttable)
		set rs=Dvbbs.Execute("Select Count(*) from "&allposttable(i)&" where "&delboardid&" PostUserID="&PostUserID) 
   		titlenum=titlenum+rs(0)

		sql="Delete From "&allposttable(i)&" where "&delboardid&" PostUserID="&PostUserID
		Dvbbs.Execute(sql)
		next
		Set Rs=Nothing
		'精华
		Dvbbs.Execute("delete from dv_besttopic where "&delboardid&" PostUserID="&PostUserID)
		'上传
		Dvbbs.Execute("delete from Dv_UpFile where "&delboardida&" F_UserID="&PostUserID)
		'该用户发表的主题、连带跟贴一起删除
		set rs=Dvbbs.Execute("select topicid,posttable from dv_topic where "&delboardid&" PostUserID="&PostUserID)
		do while not rs.eof
			Dvbbs.Execute("Delete From "&rs(1)&" where rootid="&rs(0))
		rs.movenext
		loop
		Set Rs = Dvbbs.Execute("SELECT COUNT(TopicID) FROM Dv_Topic WHERE " & Delboardid & " PostUserID = " & PostUserID)
		Dnum = Rs(0)
		Set Rs=Nothing
		Dvbbs.Execute("Delete From dv_topic where "&delboardid&" PostUserID="&PostUserID)
		if isnull(titlenum) then titlenum=0
		sql="update [dv_user] set userpost=userpost-"&titlenum&",userWealth=userWealth-"&titlenum*Dvbbs.Forum_user(3)&",userEP=userEP-"&titlenum*Dvbbs.Forum_user(8)&",userCP=userCP-"&titlenum*Dvbbs.Forum_user(13)&" where UserID="&PostUserID
		Dvbbs.Execute(sql)
		Response.write "该用户的" & Dnum & "个主题及跟帖删除成功！<br />建议您到更新论坛数据中更新一下论坛数据，或者<a href=alldel.asp>返回</a>"
	Response.Flush
End Sub

'删除指定日期内帖子
Sub Alldel()
	Dim TimeLimited,Delboardid,DelSql,rs,i
	Dim Dnum '删除帖子数 2005-10-30 Dv.Yz
	If Request("delboardid")="0" Then
		'founderr=true
		Errmsg=ErrMsg + "<br /><li>非法的版面参数。"
		Exit Sub
	Elseif Request("delboardid")="all" Then
		Delboardid=""
	Else
		'Delboardid="And boardid="&Clng(Request("delboardid"))
		Delboardid=" boardid="&Clng(Request("delboardid"))&" and "
	End If
	TimeLimited=Request.Form("TimeLimited")
	If Not Isnumeric(TimeLimited) Then
		'founderr=true
		Errmsg=ErrMsg + "<br /><li>非法的参数。"
		Exit Sub
	Else
		For i=0 to Ubound(allposttable)
			If IsSqlDataBase=1 Then
				Set Rs = Dvbbs.Execute("SELECT COUNT(AnnounceID) FROM " & Allposttable(i) & " WHERE " & Delboardid & " Datediff(d, DateAndTime, " & SqlNowString & ") > " & TimeLimited)
				Dnum = Rs(0)
				Dvbbs.Execute("DELETE FROM "&Allposttable(i)&" WHERE "&Delboardid&" Datediff(d,DateAndTime,"&SqlNowString&")>"&TimeLimited)
			Else
				Set Rs = Dvbbs.Execute("SELECT COUNT(AnnounceID) FROM " & Allposttable(i) & " WHERE " & Delboardid & " Datediff('d', DateAndTime, " & SqlNowString & ") > " & TimeLimited)
				Dnum = Rs(0)
				Dvbbs.Execute("DELETE FROM "&Allposttable(i)&" WHERE "&Delboardid&" Datediff('d',DateAndTime,"&SqlNowString&")>"&TimeLimited)
			End if
			Response.Write Allposttable(i)&"表" & Dnum & "张帖子删除完成！<br />"
			Response.Flush
		Next
		If IsSqlDataBase=1 Then
			Set Rs = Dvbbs.Execute("SELECT COUNT(TopicID) FROM Dv_topic WHERE " & Delboardid & " Datediff(d, DateAndTime, " & SqlNowString & ") > " & TimeLimited)
			Dnum = Rs(0)
			Dvbbs.Execute("DELETE FROM Dv_topic WHERE "&Delboardid&" Datediff(d,DateAndTime,"&SqlNowString&")>"&TimeLimited)
			Dvbbs.Execute("delete from dv_besttopic where "&Delboardid&" datediff(d,DateAndTime,"&SqlNowString&")>"&TimeLimited)
		Else
			Set Rs = Dvbbs.Execute("SELECT COUNT(TopicID) FROM Dv_topic WHERE " & Delboardid & " Datediff('d', DateAndTime, " & SqlNowString & ") > " & TimeLimited)
			Dnum = Rs(0)
			Dvbbs.Execute("DELETE FROM Dv_topic WHERE "&Delboardid&" Datediff('d',DateAndTime,"&SqlNowString&")>"&TimeLimited)
			Dvbbs.Execute("DELETE FROM Dv_besttopic WHERE "&Delboardid&" Datediff('d',DateAndTime,"&SqlNowString&") > "&TimeLimited)
		End If
			Response.Write "Dv_Topic表" & Dnum & "个主题删除完成！<br />"
			Response.Flush
	End if
	Response.Write "删除成功！<br />建议您到更新论坛数据中更新一下论坛数据，或者<a href=alldel.asp>返回</a>"
	Response.Flush
End sub

sub alldelTopic()
	Dim TimeLimited,delboardid,rs
	Dim Dnum '删除帖子数 2005-10-30 Dv.Yz
	if request("delboardid")="0" then
		'founderr=true
		Errmsg=ErrMsg + "<br /><li>非法的版面参数。"
		exit sub
	elseif request("delboardid")="all" then
		delboardid=""
	else
		delboardid=" boardid="&Dvbbs.CheckNumeric(request("delboardid"))&" and "
	end if
	TimeLimited=request.form("TimeLimited")
	if not isnumeric(TimeLimited) then
		'founderr=true
		Errmsg=ErrMsg + "<br /><li>非法的参数。"
		exit sub
	else
	if IsSqlDataBase=1 then
		set rs=Dvbbs.Execute("select Topicid,PostTable from dv_topic where "&delboardid&"   datediff(d,DateAndTime,"&SqlNowString&")>"&TimeLimited&" and Child=0")
	else
		set rs=Dvbbs.Execute("select Topicid,PostTable from dv_topic where "&delboardid&"   datediff('d',DateAndTime,"&SqlNowString&")>"&TimeLimited&" and Child=0")
	end if
	do while not rs.eof
		Dvbbs.Execute("Delete From "&rs(1)&" where rootid="&rs(0))
		Dvbbs.Execute("delete from dv_besttopic where rootid="&rs(0))
	rs.movenext
	loop
	if IsSqlDataBase=1 then
		Set Rs = Dvbbs.Execute("SELECT COUNT(TopicID) FROM Dv_Topic WHERE " & Delboardid & " DATEDIFF(d, DateAndTime, " & SqlNowString & ") > " & TimeLimited & " AND Child = 0")
		Dnum = Rs(0)
		Dvbbs.Execute("Delete From dv_topic where "&delboardid&"   datediff(d,DateAndTime,"&SqlNowString&")>"&TimeLimited&" and Child=0")
	else
		Set Rs = Dvbbs.Execute("SELECT COUNT(TopicID) FROM Dv_Topic WHERE " & Delboardid & " DATEDIFF('d', DateAndTime, " & SqlNowString & ") > " & TimeLimited & " AND Child = 0")
		Dnum = Rs(0)
		Dvbbs.Execute("Delete From dv_topic where "&delboardid&"   datediff('d',DateAndTime,"&SqlNowString&")>"&TimeLimited&" and Child=0")
	end if
	set rs=nothing
	end if
	response.write "Dv_Topic表" & Dnum & "个主题删除成功！<br />建议您到更新论坛数据中更新一下论坛数据，或者<a href=alldel.asp>返回</a>"
	end sub

Sub DelUser()
	Dim TimeLimited,rs,sql,i
	Dim Dnum '删除用户数 2005-10-30 Dv.Yz
	TimeLimited=Replace(request.form("TimeLimited"),"'","")
	if TimeLimited="all" then
	response.Write "算了吧，想开点吧，这样做会连管理员都删掉的！"
	else
	if IsSqlDataBase=1 then
	set rs=Dvbbs.Execute("select userid,username,usergroupid from [dv_user] where datediff(d,LastLogin,"&SqlNowString&")>"&Dvbbs.CheckNumeric(TimeLimited)&"")
	else
	set rs=Dvbbs.Execute("select userid,username,usergroupid from [dv_user] where datediff('d',LastLogin,"&SqlNowString&")>"&Dvbbs.CheckNumeric(TimeLimited)&"")
	end if
	'shinzeal加入删除用户的同时自动删除其帖子（包括精华贴）的功能
	do while not rs.eof
		If rs(2)>3 then
		for i=0 to ubound(allposttable)
		sql="Delete From "&allposttable(i)&" where postuserid="&rs(0)
		Dvbbs.Execute(sql)
		next
		Dvbbs.Execute("delete from dv_besttopic where postuserid="&rs(0))
		Dvbbs.Execute("Delete From Dv_UpFile Where F_UserID="&rs(0))
		Dvbbs.Execute("Delete From Dv_Message Where Sender='"&Replace(Rs(1),"'","''")&"'")
		Dvbbs.Execute("Delete From Dv_Friend Where F_UserID="&rs(0))
		Dvbbs.Execute("Delete From Dv_BookMark Where UserName='"&Replace(Rs(1),"'","''")&"'")
		dim rrs
		set rrs=Dvbbs.Execute("select topicid,posttable from dv_topic where postuserid="&rs(0))
		do while not rrs.eof
		Dvbbs.Execute("Delete From "&rrs(1)&" where rootid="&rrs(0))
		rrs.movenext
		loop
		set rrs=nothing
		Dvbbs.Execute("Delete From dv_topic where postuserid="&rs(0))
		end if
	rs.movenext
	loop
	set rs=nothing
	if IsSqlDataBase=1 then
		Set Rs = Dvbbs.Execute("SELECT COUNT(UserID) FROM Dv_User WHERE Datediff(d, LastLogin, " & SqlNowString & ") > " & TimeLimited & "")
		Dnum = Rs(0)
		Dvbbs.Execute("delete from [dv_user] where datediff(d,LastLogin,"&SqlNowString&")>"&TimeLimited&"")
	else
		Set Rs = Dvbbs.Execute("SELECT COUNT(UserID) FROM Dv_User WHERE Datediff('d', LastLogin, " & SqlNowString & ") > " & TimeLimited & "")
		Dnum = Rs(0)
		Dvbbs.Execute("delete from [dv_user] where datediff('d',LastLogin,"&SqlNowString&")>"&TimeLimited&"")
	end if
	end if
	response.write "删除" & Dnum & "个用户成功！<br />建议您到更新论坛数据中更新一下论坛数据，或者<a href=alldel.asp>返回</a>"
	Response.Flush
End Sub

Sub MoveUserTopic()
	Dim PostUserID,Sql,rs,i
	If Not Isnumeric(Request("Inboardid")) Then
		Response.Write "错误的版面参数。"
		Exit Sub
	End If
	If Not Isnumeric(Request("Outboardid")) Then
		Response.Write "错误的版面参数。"
		Exit Sub
	End If
	If Request("Username") = "" Then
		Response.Write "请填写用户名。"
		Exit Sub
	End If
	If Cint(Request("Outboardid")) = Cint(Request("Inboardid")) Then
		Response.Write "不能在相同版面进行移动操作！"
		Exit Sub
	End If
	Set Rs = Dvbbs.Execute("Select UserID From Dv_User Where UserName = '" & Replace(Request("Username"), "'", "''") & "'")
	If Rs.Eof And Rs.Bof Then
		Response.Write "目标用户名并不存在，请重新输入！"
		Exit Sub
	End If
	PostUserID = Rs(0)
	For i = 0 To Ubound(Allposttable)
		Dvbbs.Execute("UPDATE " & Allposttable(i) & " SET Boardid = " & Request("Inboardid") & " WHERE Boardid = " & Request("Outboardid") & " AND PostUserID = " & PostUserID)
	Next
	Rs.Close:Set Rs = Nothing
	REM 修改批量移动方式 2004-4-25 Dvbbs.YangZheng
	SET Rs = Dvbbs.Execute("SELECT Topicid, Posttable, Istop FROM Dv_Topic WHERE Boardid = " & Request("Outboardid") & " AND PostUserID = " & PostUserID)
	Rem Topicid:0, Posttable:1, Istop:2
	If Not(Rs.Eof And Rs.Bof) Then
		Sql = Rs.GetRows(-1)
		Rs.Close:Set Rs = Nothing
		Dim Yrs, TopstrinfoN, TopstrinfoO
		For i = 0 To Ubound(Sql,2)
			Dvbbs.Execute("UPDATE " & Sql(1,i) & " SET Boardid = " & Request("Inboardid") & " WHERE Rootid = " & Sql(0,i))
			Dvbbs.Execute("UPDATE Dv_Topic SET Boardid = " & Request("Inboardid") & " WHERE Boardid = " & Request("Outboardid") & " AND Topicid = " & Sql(0,i))
			If Sql(2,0) > 0 Then
				'读取新旧版面的固顶信息
				Set Yrs = Dvbbs.Execute("SELECT BoardTopStr From Dv_Board Where Boardid = " & Request("Outboardid"))
				TopstrinfoO = Yrs(0)
				Set Yrs = Dvbbs.Execute("SELECT BoardTopStr From Dv_Board Where Boardid = " & Request("Inboardid"))
				TopstrinfoN = Yrs(0)
				Yrs.Close:Set Yrs = Nothing
				'删除原固顶主题ID
				TopstrinfoO = Replace(TopstrinfoO, Cstr(Sql(0,i))&",", "")
				TopstrinfoO = Replace(TopstrinfoO, ","&Cstr(Sql(0,i)), "")
				TopstrinfoO = Replace(TopstrinfoO, Cstr(Sql(0,i)), "")
				If TopstrinfoN = "" Or Isnull(TopstrinfoN) Then
					TopstrinfoN = Cstr(Sql(0,i))
				ElseIf TopstrinfoN = Cstr(Sql(0,i)) Then
					TopstrinfoN = TopstrinfoN
				ElseIf Instr(TopstrinfoN, ","&Cstr(Sql(0,i))) > 0 Then
					TopstrinfoN = TopstrinfoN
				Else
					TopstrinfoN = TopstrinfoN & "," & Cstr(Sql(0,i))
				End If
				'更新当前版面固顶信息及缓存
				Dvbbs.Execute("UPDATE Dv_Board SET BoardTopStr = '" & TopstrinfoO & "' WHERE BoardID = " & Request("Outboardid"))
				'更新新版面固顶信息及缓存
				Dvbbs.Execute("UPDATE Dv_Board SET BoardTopStr = '" & TopstrinfoN & "' WHERE Boardid = " & Request("Inboardid"))
			End If
		Next
		Dvbbs.ReloadBoardInfo(Request("Outboardid")&","&Request("Inboardid"))
	End If
	Dvbbs.Execute("UPDATE Dv_Besttopic SET Boardid = " & Request("Inboardid") & " WHERE Boardid = " & Request("Outboardid") & " AND PostUserID = " & PostUserID)
	'shinzeal加入移动上传文件数据
	Dvbbs.Execute("UPDATE Dv_Upfile SET F_Boardid = " & Request("Inboardid") & " WHERE F_Boardid = " & Request("Outboardid") & " AND F_UserID = " & PostUserID)
	Response.Write "移动成功！<br>在“重计论坛数据和修复”中“更新论坛数据”。"
End Sub

Sub MoveDateTopic()
	Dim TimeLimited,rs,sql,i
	TimeLimited = Request.Form("TimeLimited")
	If Not Isnumeric(TimeLimited) Then
		Response.Write "错误的日期参数。"
		Exit Sub
	Else
		TimeLimited = Clng(TimeLimited)
	End If
	If Not Isnumeric(Request("Inboardid")) Then
		Response.Write "错误的版面参数。"
		Exit Sub
	End If
	If Not Isnumeric(Request("Outboardid")) Then
		Response.Write "错误的版面参数。"
		Exit Sub
	End If
	If Cint(Request("Outboardid")) = Cint(Request("Inboardid")) Then
		Response.Write "不能在相同版面进行移动操作！"
		Exit Sub
	End If
	Rem 修改移动方式 2004-4-25 Dvbbs.YangZheng
	Sql = "SELECT PostTable,Isbest,IsTop,TopicID FROM Dv_Topic WHERE Boardid = " & Request("Outboardid")
	If TimeLimited > 0 Then
		If IsSqlDataBase = 1 Then
			Sql = Sql & " AND DATEDIFF(d, DateAndTime, " & SqlNowString & ") >= " & TimeLimited
		Else
			Sql = Sql & " AND DATEDIFF('d', DateAndTime, " & SqlNowString & ") >= " & TimeLimited
		End If
	End If
	Rem PostTable:0, Isbest:1, IsTop:2, TopicID:3
	Set Rs = Dvbbs.Execute(Sql)
	Dim Sqlstr
	Dim Yrs, TopstrinfoN, TopstrinfoO
	If Not(Rs.Eof And Rs.Bof) Then
		Sql = Rs.Getrows(-1)
		Rs.Close:Set Rs = Nothing
		For i = 0 To Ubound(Sql,2)
			Sqlstr = "UPDATE " & Sql(0,i) & " SET BoardID = " & Request("Inboardid") & " WHERE BoardID = " & Request("Outboardid") & " AND RootID = " & Clng(Sql(3,i))
			Dvbbs.Execute(Sqlstr)
			Dvbbs.Execute("UPDATE Dv_Topic SET BoardID = " & Request("Inboardid") & " WHERE BoardID = " & Request("Outboardid") & " AND TopicID = " & Sql(3,i))
			If Sql(1,i) = 1 Then
				Dvbbs.Execute("UPDATE Dv_Besttopic Set BoardID = " & Request("Inboardid") & " WHERE BoardID = " & Request("Outboardid") & " AND RootID = " & Sql(3,i))
			End If
			If Sql(2,i) > 0 Then
				
				'读取新旧版面的固顶信息
				Set Yrs = Dvbbs.Execute("SELECT BoardTopStr From Dv_Board Where Boardid = " & Request("Outboardid"))
				TopstrinfoO = Yrs(0)
				Set Yrs = Dvbbs.Execute("SELECT BoardTopStr From Dv_Board Where Boardid = " & Request("Inboardid"))
				TopstrinfoN = Yrs(0)
				Yrs.Close:Set Yrs = Nothing
				'删除原固顶主题ID
				TopstrinfoO = Replace(TopstrinfoO, Cstr(Sql(3,i))&",", "")
				TopstrinfoO = Replace(TopstrinfoO, ","&Cstr(Sql(3,i)), "")
				TopstrinfoO = Replace(TopstrinfoO, Cstr(Sql(3,i)), "")
				If TopstrinfoN = "" Or Isnull(TopstrinfoN) Then
					TopstrinfoN = Cstr(Sql(3,i))
				ElseIf TopstrinfoN = Cstr(Sql(3,i)) Then
					TopstrinfoN = TopstrinfoN
				ElseIf Instr(TopstrinfoN, ","&Cstr(Sql(3,i))) > 0 Then
					TopstrinfoN = TopstrinfoN
				Else
					TopstrinfoN = TopstrinfoN & "," & Cstr(Sql(3,i))
				End If
				'更新原版面固顶信息及缓存
				Sqlstr = "UPDATE Dv_Board SET BoardTopStr = '" & TopstrinfoO & "' WHERE BoardID = " & Request("Outboardid")
				Dvbbs.Execute(Sqlstr)
				
				'更新新版面固顶信息及缓存
				Sqlstr = "UPDATE Dv_Board SET BoardTopStr = '" & TopstrinfoN & "' WHERE Boardid = " & Request("Inboardid")
				Dvbbs.Execute(Sqlstr)
			End If
			Dvbbs.Execute("UPDATE Dv_Upfile SET F_Boardid = " & Request("Inboardid") & " WHERE F_Boardid = " & Request("Outboardid") &" And F_AnnounceID Like '"&Sql(3,i)&"|%'") 
		Next
		Dvbbs.ReloadBoardInfo(Request("Outboardid")&","&Request("Inboardid"))
	End If
	Response.Write "移动成功！<br>在“重计论坛数据和修复”中“更新论坛数据”。"
End Sub
%>