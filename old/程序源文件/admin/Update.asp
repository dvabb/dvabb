<!--#include file="../conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="../inc/dv_clsother.asp"-->
<!--#include file="../inc/ubblist.asp"-->
<%
Head()
Server.ScriptTimeout=9999999
dim admin_flag 
admin_flag=Split("14,20",",")
CheckAdmin(","&admin_flag(0)&",")
CheckAdmin(","&admin_flag(1)&",")

Dim tmprs,body
Call main()
Footer()

sub main()
Dim i
%>
<table cellpadding="0" cellspacing="0" border="0" width="100%" align="center">
<tr>
<th style="text-align:center;" colspan=2>论坛数据处理</th>
</tr>
<tr>
<td width="20%" class="td1" height=25>注意事项</td>
<td width="80%" class="td1">下面有的操作可能将非常消耗服务器资源，而且更新时间很长，请仔细确认每一步操作后执行。</td>
</tr>
<%
	If request("action")="updat" Then
		If request("submit")="更新分版面数据" Or request("submit")="更新论坛数据" Then
			call updateboard()
		ElseIf request("submit")="修 复" Then
			call fixtopic()
		ElseIf request("submit")="清空在线用户" Then
			call Delallonline()
		ElseIf request("submit")="更新收藏夹" Then
			call Updatebm()
		Else
			call updateall()
		End If
		If founderr Then
			response.write errmsg
		Else
			response.write body
		End If
	ElseIf request("action")="fix"  Then
		Call Fixbbs()
		If founderr Then
			response.write errmsg
		Else
			response.write body
		End If
	ElseIf request("action")="delboard" then
		if isnumeric(request("boardid")) then
		Dvbbs.Execute("update dv_topic set boardid=444 where boardid="&request("boardid"))
		for i=0 to ubound(AllPostTable)
		Dvbbs.Execute("update "&AllPostTable(i)&" set boardid=444 where boardid="&request("boardid"))
		next
		end if
		response.write "<tr><td align=left colspan=2 height=23 class=td1>清空论坛数据成功，请返回更新帖子数据！</td></tr>"
	elseif request("action")="updateuser" then
%>
<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr> 
<th style="text-align:center;" colspan=2>更新用户数据</th>
</tr>
<tr>
<td width="20%" class="td1">重新计算用户发贴</td>
<td width="80%" class="td1">执行本操作将按照<font color=red>当前论坛数据库</font>发贴重新计算所有用户发表帖子数量。</td>
</tr>
<tr>
<td width="20%" class="td1">开始用户ID</td>
<td width="80%" class="td1"><input type=text name="beginID" value="1" size=10>&nbsp;用户ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="td1">结束用户ID</td>
<td width="80%" class="td1"><input type=text name="endID" value="100" size=10>&nbsp;将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="td1">&nbsp;</td>
<td width="80%" class="td1"><input type="submit" class="button" name="Submit" value="重新计算用户发贴"></td>
</tr>
</form>

<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr>
<td width="20%" class="td1" valign=top>更新用户等级</td>
<td width="80%" class="td1">执行本操作将按照<font color=red>当前论坛数据库</font>用户发贴数量和论坛的等级设置重新计算用户等级，本操作不影响等级为贵宾、版主、总版主的数据。</td>
</tr>
<tr>
<td width="20%" class="td1">开始用户ID</td>
<td width="80%" class="td1"><input type=text name="beginID" value="1" size=10>&nbsp;用户ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="td1">结束用户ID</td>
<td width="80%" class="td1"><input type=text name="endID" value="100" size=10>&nbsp;将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="td1">&nbsp;</td>
<td width="80%" class="td1"><input type="submit" class="button" name="Submit" value="更新用户等级"></td>
</tr>
</form>

<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr>
<td width="20%" class="td1" valign=top>更新用户金钱/积分/魅力</td>
<td width="80%" class="td1">执行本操作将按照<font color=red>当前论坛数据库</font>用户的发贴数量和论坛的相关设置重新计算用户的金钱/积分/魅力，本操作也将重新计算贵宾、版主、总版主的数据<BR>注意：不推荐用户进行本操作，本操作在数据很多的时候请尽量不要使用，并且本操作对各个版面删除帖子等所扣相应分值不做运算，只是按照发贴和总的论坛分值设置进行运算，请大家慎重操作，<font color=red>而且本项操作将重置用户因为奖励、惩罚等原因管理员对用户分值的修改。</font></td>
</tr>
<tr>
<td width="20%" class="td1">开始用户ID</td>
<td width="80%" class="td1"><input type=text name="beginID" value="1" size=10>&nbsp;用户ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="td1">结束用户ID</td>
<td width="80%" class="td1"><input type=text name="endID" value="100" size=10>&nbsp;将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="td1">&nbsp;</td>
<td width="80%" class="td1"><input type="submit" class="button" name="Submit" value="更新用户金钱/积分/魅力"></td>
</tr>
</FORM>
<%
	elseif request("action")="updateuserinfo" then
		if request("submit")="重新计算用户发贴" then
		call updateTopic()
		elseif request("submit")="更新用户等级" then
		call updategrade()
		else
		call updatemoney()
		end if
		if founderr then
		response.write errmsg
		else
		response.write body
		end if
	else
	'主题数,帖子数,用户数,今日贴,昨日贴,总固顶,最后注册
%>
<tr> 
<th style="text-align:center;" colspan=2>更新论坛数据</th>
</tr>

<form action="update.asp?action=updat" method=post>
<tr>
<td width="20%" class="td2">更新总论坛数据</td>
<td width="80%" class="td2">
<input type="checkbox" class="checkbox" name="u1" value="1">
主题数
<input type="checkbox" class="checkbox" name="u2" value="1">
帖子数
<input type="checkbox" class="checkbox" name="u3" value="1">
用户数
<input type="checkbox" class="checkbox" name="u4" value="1" checked>
今日帖
<input type="checkbox" class="checkbox" name="u5" value="1" checked>
昨日帖
<input type="checkbox" class="checkbox" name="u6" value="1">
总固顶
<input type="checkbox" class="checkbox" name="u7" value="1">
最后注册
<BR><BR><input type="submit" class="button" name="Submit" value="更新论坛总数据"><BR><BR>这里将重新计算整个论坛的帖子主题和回复数，今日帖子，最后加入用户等，建议每隔一段时间运行一次。<hr size=1></td>
</tr>
<tr>
<td width="20%" class="td1">更新分版面数据</td>
<td width="80%" class="td1"><input type="submit" class="button" name="Submit" value="更新分版面数据"><BR><BR>这里将重新计算每个版面的帖子主题和回复数，今日帖子，最后回复信息等，建议每隔一段时间运行一次。<hr size=1>
</td>
</tr>
<tr>
<td width="20%" class="td2">更新论坛收藏夹</td>
<td width="80%" class="td2"><input type="submit" class="button" name="Submit" value="更新收藏夹"><BR><BR>这里将重新整理论坛的收藏夹，删除不存在用户的收藏记录，重新指向被移动的帖子收藏地址，删除已被删除的帖子收藏记录。
</td>
</tr>
<tr> 
<th style="text-align:center;" colspan=2>修复帖子(修复指定范围内帖子的最后回复数据)</th>
</tr>
<tr>
<td width="20%" class="td1">开始的ID号</td>
<td width="80%" class="td1"><input type=text name="beginID" value="1" size=10>&nbsp;帖子主题ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="td2">结束的ID号</td>
<td width="80%" class="td2"><input type=text name="EndID" value="1000" size=10>&nbsp;将更新开始到结束ID之间的帖子数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="td1">&nbsp;</td>
<td width="80%" class="td1"><input type="submit" class="button" name="Submit" value="修 复"></td>
</tr>
</form>
<form name=Fix action="update.asp?action=fix" method=post>
<tr> 
<th style="text-align:center;" colspan=2>修正贴子UBB标签(修复指定范围贴子UBB标签)</th>
</tr>
<tr>
<td width="20%" class="td1">开始的ID号</td>
<td width="80%" class="td1"><input type=text name="beginID" value="1" size=10>&nbsp;帖子主题ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="td2">结束的ID号</td>
<td width="80%" class="td2"><input type=text name="EndID" value="1000" size=10>&nbsp;将更新开始到结束ID之间的帖子数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="td1">新老贴的标识日期</td>
<td width="80%" class="td1"><input type="text" name="updatedate" value="2003-12-1">(格式：YYYY-M-D) 就是论坛升级到v7.0的日期，如果不填写，一律按老贴处理</td>
</tr>
<tr>
<td width="20%" class="td2">去掉贴子中的HTML标记</td>
<td width="80%" class="td2">是 <input type="radio" class="radio" name="killhtml" value="1">
  否 <input type="radio" class="radio" name="killhtml" value="0" checked>&nbsp;<br>选是的话，贴子中的HTML标记将会自动被清除，有利于减少数据库的大小，但是会失去原来的HTML效果。</td>
</tr>
<tr>
<td width="20%" class="td1">&nbsp;</td>
<td width="80%" class="td1"><input type="submit" class="button" name="Submit" value="修 正"></td>
</tr>
</form>

<%
	end if
%>
</table><BR><BR>
<%
	end sub

Sub updateboard()
	'先按照所有版面ID得出帖子数，然后计算各个有下属论坛的帖子总和
	Dim allarticle
	Dim alltopic
	Dim alltoday
	Dim allboard
	Dim trs,Esql,ars,rs,i
	Dim Maxid
	Dim LastTopic,LastRootid,LastPostTime,LastPostUser
	Dim LastPost,uploadpic_n,Lastpostuserid,Lastid
	Dim ParentStr
	Dim C,C1,C2
	Dim reBoard_Setting,BoardTopStr,IsGroupSetting
	Dim UserAccessCount,UpGroupSetting,ii
	Dim Slastpost
	ii=0
	'设置打开数据时间
	conn.CommandTimeout=3600

	'获得要更新的总数
	If IsNumeric(request("boardid")) And request("boardid")<>"" Then
		Set Rs=Dvbbs.Execute("Select Count(*) From [Dv_board] Where BoardID="&request("boardid"))
		C1=rs(0)
		If Isnull(C1) Then C1=0
	Else
		Set Rs=Dvbbs.Execute("Select Count(*) From [Dv_board]")
		C1=rs(0)
		If Isnull(C1) Then C1=0
	End If
	Set Rs=Nothing
%>
</table><BR>
<table cellpadding="0" cellspacing="0" border="0" width="100%" align="center">
<tr><td colspan=2 class=td1>
下面开始更新论坛版面资料，共有<%=C1%>个版面需要更新
<table width="400" border="0" cellspacing="1" cellpadding="1">
<tr> 
<td bgcolor=#000000>
<table width="400" border="0" cellspacing="0" cellpadding="1">
<tr> 
<td bgcolor=#ffffff height=9><img src="../skins/default/bar/bar3.gif" width=0 height=16 id=img2 name=img2 align=absmiddle></td></tr></table>
</td></tr></table> <span id=txt2 name=txt2 style="font-size:9pt">0</span><span style="font-size:9pt">%</span></td></tr>
</table>
<%
	Response.Flush

	'排序按照Child和Orders，以便先更新下级论坛的数据才循环到上级版面，这时上级版面读取的就是下级版面的最新数据
	If IsNumeric(request("boardid")) And request("boardid")<>"" Then
		Set Rs=Dvbbs.Execute("Select BoardID,BoardType,Child,ParentStr,RootID,Board_Setting,BoardTopStr,IsGroupSetting From Dv_Board Where BoardID="&Request("BoardID"))
	Else
		Call Boardchild()	'统计更新下属论坛个数 YZ-2004-2-26注
		Set Rs=Dvbbs.Execute("Select BoardID,BoardType,Child,ParentStr,RootID,Board_Setting,BoardTopStr,IsGroupSetting From Dv_Board Order by Child,RootID,Orders Desc")
	End If
	Dim SQL, LastPostArr 'LastPostArr XG增加2007-04-11
	If Not Rs.EOF Then 
		SQL=Rs.GetRows(-1)
		Set Rs=Nothing
		For i=0 to UBound(SQL,2)
	'Do While Not Rs.Eof
		
		reBoard_Setting=Split(SQL(5,i)&"",",")
		AllBoard = 0
		'所有主题和帖子
		Set Trs=Dvbbs.Execute("Select Count(*),Sum(Child) From Dv_Topic Where BoardID="&SQL(0,i))
		AllTopic=Trs(0)
		AllArticle=Trs(1)
		If IsNull(AllTopic) Then AllTopic = 0
		If IsNull(AllArticle) Then AllArticle = 0
		AllArticle = AllArticle + AllTopic
		Set Trs=Nothing
		'所有今日贴
		If IsSqlDataBase = 1 Then
			Set Trs=Dvbbs.Execute("Select Count(*) From "&Dvbbs.NowUseBBS&" Where BoardID="&SQL(0,i)&" and datediff(d,dateandtime,"&SqlNowString&")=0")
		Else
			Set Trs=Dvbbs.Execute("Select Count(*) From "&Dvbbs.NowUseBBS&" Where BoardID="&SQL(0,i)&" and datediff('d',dateandtime,"&SqlNowString&")=0")
		End If
		AllToday=Trs(0)
		Set Trs=Nothing
		If IsNull(AllToday) Then AllToday=0
		'最后回复信息
		Set Trs=Dvbbs.Execute("Select Top 1 LastPost,TopicID,Title,PostTable From Dv_Topic Where BoardID="&SQL(0,i)&" Order by LastPostTime Desc")
		If Not (Trs.Eof And Trs.Bof) Then
			LastPostArr = Split(Trs(0)&"","$")
			If UBound(LastPostArr)<7 Then
				ReDim LastPostArr(7)
				LastPostArr(1)=Trs(3)
				LastPostArr(3)= Replace(cutStr(Dvbbs.Replacehtml(Trs(2)),20),"$","&#36;")
				LastPostArr(6)=Trs(1)
				Trs.Close:Set Trs=Nothing
				Set Trs = Dvbbs.Execute("Select Top 1 AnnounceID,BoardID,UserName,DateAndTime,PostUserID From "&LastPostArr(1)&" Where RootID="&LastPostArr(6)&" Order by DateAndTime")
				If Not (Trs.Eof And Trs.Bof) Then
					LastPostArr(0) = Trs(2)
					LastPostArr(1) = Trs(0)
					LastPostArr(2) = Trs(3)
					LastPostArr(4) = ""
					LastPostArr(5) = Trs(4)
					LastPostArr(7) = Trs(1)
				Else
					LastPostArr(0) = "无"
					LastPostArr(1) = 0
					LastPostArr(2) = now()
					LastPostArr(3) = "无"
					LastPostArr(4) = ""
					LastPostArr(5) = ""
					LastPostArr(6) = ""
					LastPostArr(7) = ""
				End if
			Else
				LastPostArr(3)= Replace(cutStr(Trs(2),20),"$","&#36;")
			End if
			Trs.Close:Set Trs=Nothing
			LastPost=Join(LastPostArr,"$")
			LastPost=Replace(Dvbbs.Replacehtml(LastPost&""),"'","''")
			'LastPost=Replace(Dvbbs.Replacehtml(Trs(0)&""),"'","''")
		Else
			Trs.Close:Set Trs=Nothing
			LastPost="无$0$"&Now()&"$无$$$$"
		End If
		'更新当前版面数据
		'SLastPost = Split(LastPost,"$")
		'If Ubound(SLastPost) < 7 Then LastPost = LastPost & "$"
		Dvbbs.Execute("Update [Dv_board] Set PostNum="&AllArticle&",TopicNum="&AllTopic&",TodayNum="&AllToday&",LastPost='"&Dvbbs.ChkBadWords(LastPost)&"' Where BoardID="&SQL(0,i))
		'如果当前版面有下属论坛，则更新其数据为下属论坛数据
		If SQL(2,i)>0 Then
			'帖子总数，主题总数，今日贴总数，下属版面数
			If SQL(3,i)=0 Then
				ParentStr=SQL(0,i)
				'Set Trs=Dvbbs.Execute("Select Sum(PostNum),Sum(TopicNum),Sum(TodayNum),Count(*) From Dv_board Where (Not BoardID="&SQL(0,i)&") And RootID="&SQL(0,i))
			Else
				ParentStr=SQL(3,i) & "," & SQL(0,i)
				'Set Trs=Dvbbs.Execute("Select Sum(PostNum),Sum(TopicNum),Sum(TodayNum),Count(*) From Dv_board Where ParentStr Like '%"&ParentStr&"%'")
			End If
			Set Trs=Dvbbs.Execute("Select Sum(PostNum),Sum(TopicNum),Sum(TodayNum),Count(*) From Dv_board Where ParentStr Like '%"&ParentStr&"%'")

			If Not (Trs.Eof And Trs.Bof) Then
				'如果该版面允许发贴，则帖子数应该是该版面贴数+下属版面帖子数
				If reBoard_Setting(43)="0" Then
					If Not IsNull(Trs(0)) Then AllArticle = Trs(0) + AllArticle
					If Not IsNull(Trs(1)) Then AllTopic = Trs(1) + AllTopic
					If Not IsNull(Trs(2)) Then AllToday = Trs(2) + AllToday
					If Not IsNull(Trs(3)) Then AllBoard = Trs(3) + AllBoard
				Else
					AllArticle=Trs(0)
					AllTopic=Trs(1)
					AllToday=Trs(2)
					AllBoard=Trs(3)
					If IsNull(AllArticle) Then AllArticle=0
					If IsNull(AllTopic) Then AllTopic=0
					If IsNull(AllToday) Then AllToday=0
					If IsNull(AllBoard) Then AllBoard=0
				End If
			End If
			Set Trs=Nothing
			'下属版块ID
			ParentStr = Sql(0,i)
			Set Trs = Dvbbs.Execute("SELECT Boardid FROM Dv_Board WHERE ParentID = "&Sql(0,i))
			If Not (Trs.Eof And Trs.Bof) Then
				Do While Not Trs.Eof
					ParentStr = ParentStr & "," & Trs(0)
					Trs.Movenext
				Loop
			End If
			Set Trs=Nothing
			'最后回复信息
			Set Trs=Dvbbs.Execute("Select Top 1 LastPost,TopicID,Title,PostTable From Dv_Topic Where BoardID In ("&ParentStr&") Order by LastPostTime Desc")
			If Not (Trs.Eof And Trs.Bof) Then
				LastPostArr = Split(Trs(0)&"","$")
				If UBound(LastPostArr)<>7 Then
					ReDim LastPostArr(7)
					LastPostArr(1)=Trs(3)
					LastPostArr(3)= Replace(cutStr(Dvbbs.Replacehtml(Trs(2)),20),"$","&#36;")
					LastPostArr(6)=Trs(1)
					Trs.Close:Set Trs=Nothing
					Set Trs = Dvbbs.Execute("Select Top 1 AnnounceID,BoardID,UserName,DateAndTime,PostUserID From "&LastPostArr(1)&" Where RootID="&LastPostArr(6)&" Order by DateAndTime")
					If Not (Trs.Eof And Trs.Bof) Then
						LastPostArr(0) = Trs(2)
						LastPostArr(1) = Trs(0)
						LastPostArr(2) = Trs(3)
						LastPostArr(4) = ""
						LastPostArr(5) = Trs(4)
						LastPostArr(7) = Trs(1)
					Else
						LastPostArr(0) = "无"
						LastPostArr(1) = 0
						LastPostArr(2) = now()
						LastPostArr(3) = "无"
						LastPostArr(4) = ""
						LastPostArr(5) = ""
						LastPostArr(6) = ""
						LastPostArr(7) = ""
					End if
				Else
					LastPostArr(3)= Replace(cutStr(Trs(2),20),"$","&#36;")
				End if
				Trs.Close:Set Trs=Nothing
				LastPost=Join(LastPostArr,"$")
				LastPost=Replace(Dvbbs.Replacehtml(LastPost&""),"'","''")
				'LastPost=Replace(Dvbbs.Replacehtml(Trs(0)&""),"'","''")
			Else
				Trs.Close:Set Trs=Nothing
				LastPost="无$0$"&Now()&"$无$$$$"
			End If
			'更新版面数据
			'SLastPost = Split(LastPost,"$")
			'If Ubound(SLastPost) < 7 Then LastPost = LastPost & "$"
			Dvbbs.Execute("Update [Dv_board] Set PostNum="&AllArticle&",TopicNum="&AllTopic&",TodayNum="&AllToday&",LastPost='"&Dvbbs.ChkBadWords(LastPost)&"' Where BoardID="&SQL(0,i))
		End If
		'更新IsGroupSetting
		'IsGroupSetting=SQL(7,i)
		'Set Trs=Dvbbs.Execute("Select Count(*) From Dv_UserAccess Where uc_BoardID="&SQL(0,i))
		'UserAccessCount = Trs(0)
		'If IsNull(UserAccessCount) Or UserAccessCount="" Then UserAccessCount=0
		'If UserAccessCount>0 Then UpGroupSetting="0"
		'Set Trs=Dvbbs.Execute("Select GroupID From Dv_BoardPermission Where BoardID="&SQL(0,i))
		'If Not Trs.Eof Then
		'	Do While Not Trs.Eof
		'		If UpGroupSetting="" Then
		'			UpGroupSetting = Trs(0)
		'		Else
		'			UpGroupSetting = UpGroupSetting & "," & Trs(0)
		'		End If
		'	Trs.MoveNext
		'	Loop
		'End If
		'更新和清理固顶贴数据(固顶和区域固顶)
		'Set Trs=Dvbbs.Execute("Select TopicID From Dv_Topic Where BoardID="&Rs(0)&" And IsTop In (1,2)")
		If Not IsNull(SQL(6,i)) And SQL(6,i)<>"" Then
		Set Trs=Dvbbs.Execute("Select TopicID,BoardID,IsTop From Dv_Topic Where TopicID In ("&SQL(6,i)&")")
		If tRs.Eof And tRs.Bof Then
			BoardTopStr=""
		Else
			Do While Not Trs.Eof
				If Trs(1)<>444 And Trs(1)<>777 And Trs(2)>0 And Trs(2)<>3 Then
					If BoardTopStr="" Then
						BoardTopStr = Trs(0)
					Else
						BoardTopStr = BoardTopStr & "," & Trs(0)
					End If
				End If
			Trs.MoveNext
			Loop
		End If
		End If
		Dvbbs.Execute("Update Dv_Board Set BoardTopStr='"&BoardTopStr&"' Where BoardID="&SQL(0,i))
		UserAccessCount=""
		IsGroupSetting=""
		UpGroupSetting=""
		BoardTopStr=""
		ii=ii+1
		'If (i mod 100) = 0 Then
			Response.Write "<script>img2.width=" & Fix((ii/C1) * 400) & ";" & VbCrLf
			Response.Write "txt2.innerHTML=""" & FormatNumber(ii/C1*100,4,-1) & """;" & VbCrLf
			Response.Write "img2.title=""" & SQL(0,i) & "(" & ii & ")"";</script>" & VbCrLf
			Response.Flush
		'End If
		body="<table cellpadding=0 cellspacing=0 border=0 width=100% align=center><tr><td colspan=2 class=td1>更新论坛数据成功，"&SQL(1,i)&"共有"&AllArticle&"篇贴子，"&AllTopic&"篇主题，今日有"&AllToday&"篇帖子。</td></tr></table>"
		Response.Write body
		Response.Flush
	'Rs.MoveNext
	'Loop
	Next
	Set Trs=Nothing
	
	End If 
	body=""
	Response.Write "<script>img2.width=400;txt2.innerHTML=""100"";</script>"
	Dvbbs.loadSetup()
	Dim Board
	Dvbbs.LoadBoardList()
	For Each board in Application(Dvbbs.CacheName&"_boardlist").documentElement.selectNodes("board/@boardid")
		Dvbbs.LoadBoardData board.text
		Dvbbs.LoadBoardinformation board.text
	Next
End Sub

Rem 统计下属论坛函数 2004-5-3 Dvbbs.YangZheng
Sub Boardchild()
	Dim cBoardNum, cBoardid
	Dim Trs,rs,Sql,i
	Dim Bn
	Dvbbs.Execute("UPDATE Dv_Board SET Child = 0")
	Set Rs = Dvbbs.Execute("SELECT Boardid, Rootid, ParentID, Depth, Child, ParentStr FROM Dv_Board ORDER BY Boardid DESC")
	If Not (Rs.Eof And Rs.Bof) Then
		Sql = Rs.GetRows(-1)
		Rs.Close:Set Rs = Nothing
		For Bn = 0 To Ubound(Sql,2)
			If Isnull(Sql(4,Bn)) And Cint(Sql(3,Bn)) > 0 Then
				Dvbbs.Execute("UPDATE Dv_Board SET Child = 0 WHERE Boardid = " & Sql(0,Bn))
			End If
			If Cint(Sql(2,Bn)) = 0 And Cint(Sql(3,Bn)) = 0 Then
				Set Trs = Dvbbs.Execute("SELECT COUNT(*) FROM Dv_Board WHERE RootID = " & Sql(1,Bn))
				Cboardnum = Trs(0) - 1
				Trs.Close:Set Trs = Nothing
				If Isnull(Cboardnum) Or Cboardnum < 0 Then Cboardnum = 0
				Dvbbs.Execute("UPDATE Dv_Board SET Child = " & Cboardnum & " WHERE Boardid = " & Sql(0,Bn))
			Elseif Cint(Sql(3,Bn)) > 1 Then
				cBoardid = Split(Sql(5,Bn)&"",",")
				For i = 1 To Ubound(cBoardid)
					Dvbbs.Execute("UPDATE Dv_Board SET Child = Child + 1 WHERE Boardid = " & cBoardid(i))
				Next
			End If
		Next
	End If
End Sub

Sub Updateall()
	'主题数,帖子数,用户数,今日贴,昨日贴,总固顶,最后注册
	Body = "<tr><td colspan=2 class=td1>更新总论坛数据成功。"
	Dim AllTopNum,PostNum,TopicNum,LastUser
	Dim TodayNum,UserNum, YesterdayNum,SqlStr,Sql_a,sql
	If Request.Form("u1") = "1" Or Request("index")="1" Then
		TopicNum	= GetTopicnum()
		If Sql_a = "" Then
			Sql_a = "Forum_TopicNum = " & TopicNum & ""
		Else
			Sql_a = Sql_a & ",Forum_TopicNum = " & TopicNum & ""
		End If
		If SqlStr = "" Then
			SqlStr = "论坛共有 " & TopicNum & " 篇主题"
		Else
			SqlStr = SqlStr & "，" & TopicNum & " 篇主题"
		End If
	End If
	If Request.Form("u2") = "1" Or Request("index")="1" Then
		PostNum		= Announcenum()
		If Sql_a = "" Then
			Sql_a = "Forum_PostNum = " & PostNum & ""
		Else
			Sql_a = Sql_a & ",Forum_PostNum = " & PostNum & ""
		End If
		If SqlStr = "" Then
			SqlStr = "论坛共有 " & PostNum & " 篇帖子"
		Else
			SqlStr = SqlStr & "，" & PostNum & " 篇帖子"
		End If
	End If
	If Request.Form("u3") = "1" Or Request("index")="1" Then
		UserNum		= Allusers()
		If Sql_a = "" Then
			Sql_a = "Forum_UserNum = " & UserNum & ""
		Else
			Sql_a = Sql_a & ",Forum_UserNum = " & UserNum & ""
		End If
		If SqlStr = "" Then
			SqlStr = "论坛共有 " & UserNum & " 个用户"
		Else
			SqlStr = SqlStr & "，" & UserNum & " 个用户"
		End If
	End If
	If Request.Form("u4") = "1" Or Request("index")="1"  Then
		TodayNum	= Alltodays()
		If Sql_a = "" Then
			Sql_a = "Forum_TodayNum = " & TodayNum & ""
		Else
			Sql_a = Sql_a & ",Forum_TodayNum = " & TodayNum & ""
		End If
		If SqlStr = "" Then
			SqlStr = "论坛共有 " & TodayNum & " 篇今日帖"
		Else
			SqlStr = SqlStr & "，" & TodayNum & " 篇今日帖"
		End If
	End If
	If Request.Form("u5") = "1" Or Request("index")="1"  Then
		YesterdayNum	= Allyesterdays()
		If Sql_a = "" Then
			Sql_a = "Forum_YesterdayNum = " & YesterdayNum & ""
		Else
			Sql_a = Sql_a & ",Forum_YesterdayNum = " & YesterdayNum & ""
		End If
		If SqlStr = "" Then
			SqlStr = "论坛共有 " & YesterdayNum & " 篇昨日帖"
		Else
			SqlStr = SqlStr & "，" & YesterdayNum & " 篇昨日帖"
		End If
	End If
	If Request.Form("u6") = "1" Or Request("index")="1" Then
		AllTopNum	= Forum_AllTopNum()
		If Sql_a = "" Then
			Sql_a = "Forum_AllTopNum = '" & AllTopNum & "'"
		Else
			Sql_a = Sql_a & ",Forum_AllTopNum = '" & AllTopNum & "'"
		End If
		If SqlStr = "" Then
			SqlStr = "论坛共有 " & UBound(Split(AllTopNum&"", ",")) + 1 & " 个固顶主题"
		Else
			SqlStr = SqlStr & "，" & UBound(Split(AllTopNum&"", ",")) + 1 & " 个固顶主题"
		End If
	End If
	If Request.Form("u7") = "1" Or Request("index")="1" Then
		LastUser	= Newuser()
		If Sql_a = "" Then
			Sql_a = "Forum_lastUser = '" & Dvbbs.CheckStr(Dvbbs.HtmlEncode(LastUser)) & "'"
		Else
			Sql_a = Sql_a & ",Forum_lastUser = '" & Dvbbs.CheckStr(Dvbbs.HtmlEncode(LastUser)) & "'"
		End If
		If SqlStr = "" Then
			SqlStr = "论坛最新加入用户为 " & LastUser & ""
		Else
			SqlStr = SqlStr & "，最新加入用户为 " & LastUser & ""
		End If
	End If

	Body = Body & SqlStr
	Body = Body & "</td></tr>"

	If Sql_a = "" Then Exit Sub

	Sql = "UPDATE Dv_Setup SET " & Sql_a
	Dvbbs.Execute(Sql)

	Dvbbs.Name="setup"
	Dvbbs.loadSetup()
End sub

Sub fixtopic()
if not isnumeric(request.form("beginid")) then
	body="<tr><td colspan=2 class=td1>错误的开始参数！</td></tr>"
	exit sub
End If
if not isnumeric(request.form("endid")) then
	body="<tr><td colspan=2 class=td1>错误的结束参数！</td></tr>"
	exit sub
end if
if clng(request.form("beginid"))>clng(request.form("endid")) then
	body="<tr><td colspan=2 class=td1>开始ID应该比结束ID小！</td></tr>"
	exit sub
end if
dim TotalUseTable,Ers,sql,rs,i
dim username,dateandtime,rootid,announceid,postuserid,lastpost,topic
'set rs=Dvbbs.iCreateObject("adodb.recordset")
'Dvbbs.Execute("update Dv_topic set PostTable='dv_bbs1'")
Dim C1
C1=clng(request.form("endid"))-clng(request.form("beginid"))
%>
</table>
&nbsp;<BR>
<table cellpadding="0" cellspacing="0" border="0" width="100%" align="center">
<tr><td colspan=2 class=td1>
下面开始更新论坛帖子资料，预计本次共有<%=C1%>个帖子需要更新
<table width="400" border="0" cellspacing="1" cellpadding="1">
<tr> 
<td bgcolor=#000000>
<table width="400" border="0" cellspacing="0" cellpadding="1">
<tr> 
<td bgcolor=#ffffff height=9><img src="../skins/default/bar/bar3.gif" width=0 height=16 id=img2 name=img2 align=absmiddle></td></tr></table>
</td></tr></table> <span id=txt2 name=txt2 style="font-size:9pt">0</span><span style="font-size:9pt">%</span></td>
</tr>
</table>

<table cellpadding="0" cellspacing="0" border="0" width="100%" align="center">
<%
Response.Flush
sql="select topicid,PostTable from Dv_topic where topicid>="&request.form("beginid")&" and topicid<="&request.form("endid")

set rs=Dvbbs.Execute(sql)
if rs.eof and rs.bof then
	body="<tr><td colspan=2 class=td1>已经到记录的最尾端，请结束更新！</td></tr>"
	exit sub
end if
do while not rs.eof
	sql="select top 1 username,dateandtime,topic,Announceid,PostUserID,rootid,body,boardid from "&rs(1)&" where rootid="&rs(0)&" order by Announceid desc"
	set ers=Dvbbs.Execute(sql)
	if not (ers.eof and ers.bof) then
		username=Ers("username")
		dateandtime=Ers("dateandtime")
		rootid=Ers("rootid")
		topic=left(Ers("body"),20)
		Announceid=ers("Announceid")
		postuserid=ers("postuserid")
		LastPost=username & "$" & Announceid & "$" & dateandtime & "$" & replace(topic,"$","") & "$$" & postuserid & "$" & rootid & "$" & ers("BoardID") & "$"
		LastPost=Dvbbs.Checkstr(LastPost)
		Dvbbs.Execute("update [DV_topic] set LastPost='"&replace(LastPost,"'","")&"',LastPostTime='"&dateandtime&"' where topicid="&rs(0))
		i=i+1
		'If (i mod 100) = 0 Then
		Response.Write "<script>img2.width=" & Fix((i/C1) * 400) & ";" & VbCrLf
		Response.Write "txt2.innerHTML=""更新完"&server.htmlencode(ers(2)&"")&"的数据，正在更新下一个帖子数据，" & FormatNumber(i/C1*100,4,-1) & """;" & VbCrLf
		Response.Write "img2.title=""" & server.htmlencode(eRs(2)&"") & "(" & i & ")"";</script>" & VbCrLf
		Response.Flush
		'End If
	end if
	'计算回帖数 2004-8-2
	Sql = "SELECT COUNT(*) FROM " & Rs(1) & " WHERE Rootid = " & Rs(0) & " AND Boardid <> 444 AND Boardid <> 777"
	Set Ers = Dvbbs.Execute(Sql)
	Dvbbs.Execute("UPDATE Dv_Topic SET Child = " & Ers(0)-1 & " WHERE Topicid = " & Rs(0) & "")
	Rs.Movenext
loop
set ers=nothing
set rs=nothing
Response.Write "<script>img2.width=400;txt2.innerHTML=""100"";</script>"
%>
<form action="update.asp?action=updat" method=post>
<tr> 
<th style="text-align:center;" colspan=2>继续修复帖子(修复指定范围内帖子的最后回复数据)</th>
</tr>
<tr>
<td width="20%" class="td1">开始的ID号</td>
<td width="80%" class="td1"><input type=text name="beginID" value="<%=request.form("endid")+1%>" size=5>&nbsp;帖子主题ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="td1">结束的ID号</td>
<td width="80%" class="td1"><input type=text name="EndID" value="<%=request.form("endid")+(request.form("endid")-request.form("beginid"))+1%>" size=5>&nbsp;将更新开始到结束ID之间的帖子数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="td1">&nbsp;</td>
<td width="80%" class="td1"><input type="submit" class="button" name="Submit" value="修 复"></td>
</tr>
</form>
<%
end sub

'分论坛今日帖子
REM 修改查询所有帖子表数据 2004-8-26.Dv.Yz
Function Todays(Boardid)
	Todays = 0
	If IsSqlDataBase = 1 Then
		For i = 0 To Ubound(AllPostTable)
			Set Tmprs = Dvbbs.Execute("SELECT COUNT(Announceid) FROM " & AllPostTable(i) & " WHERE Boardid = " & Boardid & " AND DATEDIFF(day, Dateandtime, " & SqlNowString & ") = 0")
			Todays = Todays + Tmprs(0)
		Next
	Else
		For i = 0 To Ubound(AllPostTable)
			Set Tmprs = Dvbbs.Execute("SELECT COUNT(Announceid) FROM " & AllPostTable(i) & " WHERE Boardid = " & Boardid & " AND DATEDIFF('d', Dateandtime, " & SqlNowString & ") = 0")
			Todays = Todays + Tmprs(0)
		Next
	End If
	Set Tmprs = Nothing
End Function

'全部论坛今日帖子
REM 修改查询所有帖子表数据 2004-8-26.Dv.Yz
Function Alltodays()
	Dim i
	Alltodays = 0
	If IsSqlDataBase = 1 Then
		For i = 0 To Ubound(AllPostTable)
			Set Tmprs = Dvbbs.Execute("SELECT COUNT(Announceid) FROM " & AllPostTable(i) & " WHERE NOT Boardid IN (444,777) AND DATEDIFF(day, Dateandtime, " & SqlNowString & ") = 0")
			Alltodays = Alltodays + Tmprs(0)
		Next
	Else
		For i = 0 To Ubound(AllPostTable)
			Set Tmprs = Dvbbs.Execute("SELECT COUNT(Announceid) FROM " & AllPostTable(i) & " WHERE NOT Boardid IN (444,777) AND DATEDIFF('d', Dateandtime, " & SqlNowString & ") = 0")
			Alltodays = Alltodays + Tmprs(0)
		Next
	End If
	Set Tmprs = Nothing
End Function

'论坛昨天的帖子 2004-8-31.Dv.Yz
Function Allyesterdays()
	Dim i
	Allyesterdays = 0
	If IsSqlDataBase = 1 Then
		For i = 0 To Ubound(AllPostTable)
			Set Tmprs = Dvbbs.Execute("SELECT COUNT(Announceid) FROM " & AllPostTable(i) & " WHERE NOT Boardid IN (444,777) AND DATEDIFF(day, Dateandtime, " & SqlNowString & ") = 1")
			Allyesterdays = Allyesterdays + Tmprs(0)
		Next
	Else
		For i = 0 To Ubound(AllPostTable)
			Set Tmprs = Dvbbs.Execute("SELECT COUNT(Announceid) FROM " & AllPostTable(i) & " WHERE NOT Boardid IN (444,777) AND DATEDIFF('d', Dateandtime, " & SqlNowString & ") = 1")
			Allyesterdays = Allyesterdays + Tmprs(0)
		Next
	End If
	Set Tmprs = Nothing
End Function

'所有注册用户数量
function allusers() 
	allusers=Dvbbs.Execute("Select count(userid) from [Dv_user]")(0) 
	If IsNull(allusers) Then allusers=0 
End function
'最新注册用户
Function newuser()
	Dim sql
	sql="Select top 1 username from [Dv_user] order by userid desc"
	Set tmprs=Dvbbs.Execute(sql)
	If tmprs.eof and tmprs.bof Then
		newuser="没有会员"
	Else
   		newuser=tmprs("username")
	End If
	Set tmprs=Nothing 
End function 

'所有论坛帖子
function AnnounceNum()
	dim AnnNum,i
	AnnNum=0
	AnnounceNum=0
	For i=0 to ubound(AllPostTable)
		AnnNum=Dvbbs.Execute("Select Count(announceID) from "&AllPostTable(i)&" where not boardid in (444,777)")(0) 
		if isnull(AnnNum) then AnnNum=0
		AnnounceNum=AnnounceNum + AnnNum
	next
end function
'分论坛帖子
function BoardAnnounceNum(boardid)
	dim BoardAnnNum
	BoardAnnNum=0
	BoardAnnounceNum=0
	For i=0 to ubound(AllPostTable)
		BoardAnnNum=Dvbbs.Execute("Select Count(announceID) from "&AllPostTable(i)&" where boardid="&boardid)(0) 
		if isnull(BoardAnnNum) then BoardAnnNum=0
		BoardAnnounceNum=BoardAnnounceNum + BoardAnnNum
	next
end function

'所有论坛主题
function GetTopicnum()
	Dim TopicNum
	TopicNum=Dvbbs.Execute("Select Count(topicid) from DV_topic where not boardid in (444,777)")(0)
	if isnull(TopicNum) then TopicNum=0
	GetTopicnum = TopicNum
end function

'分论坛主题
function BoardTopicNum(boardid) 
	BoardTopicNum=Dvbbs.Execute("Select Count(topicid) from [Dv_topic] where boardid="&boardid)(0) 
	if isnull(BoardTopicNum) then BoardTopicNum=0 
end function

'论坛总固顶主题数
function Forum_AllTopNum()
	Set tmprs=Dvbbs.Execute("Select TopicID From Dv_Topic Where Not BoardID In (444,777) And IsTop=3")
	If tmprs.eof and tmprs.bof Then
		Forum_AllTopNum=""
	Else
		Do While Not tmprs.Eof
			If Forum_AllTopNum="" Then
				Forum_AllTopNum = tmprs(0)
			Else
				Forum_AllTopNum = Forum_AllTor]")(0) 
	If IsNull(allusers) Then allusers=0 
End function
'鏈€鏂版敞鍐岀敤鎴