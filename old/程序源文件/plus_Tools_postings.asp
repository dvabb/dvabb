<!-- #include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp"-->
<!-- #include file="Dv_plus/Tools/plus_Tools_const.asp" -->
<%
Dim ToUserID,TopicID,ReplyID,Action,ChkAction,LogMsg
Dvbbs.ErrType = 1 '设置错误提示信息显示模式
ChkAction = True
ToUserID = Dv_Tools.CheckNumeric(Request("ToUserID"))	'目标用户
TopicID = Dv_Tools.CheckNumeric(Request("TopicID"))		'主题ID
ReplyID = Dv_Tools.CheckNumeric(Request("ReplyID"))		'回复ID
Action = Dv_Tools.CheckNumeric(Request("Action"))		'执行分类
If TopicID = 0 or ReplyID = 0 or Dvbbs.BoardID = 0 Then ChkAction = False
Dvbbs.stats = "论坛道具使用"
If Action=0 Then
	Dv_Tools.ChkToolsLogin
	Dvbbs.stats = "论坛道具使用=="&Dv_Tools.ToolsInfo(1)
End If
Dvbbs.LoadTemplates("")
Dvbbs.Head()
ToolsMain
Dvbbs.Showerr()
Dvbbs.mainsetting(0)="98%"
Dvbbs.Footer()
Dvbbs.PageEnd()
'---------------------------------------------------
'Dv_Tools.ToolsInfo 道具系统信息
'ID=0 ,ToolsName=1 ,ToolsInfo=2 ,IsStar=3 ,SysStock=4 ,UserStock=5 ,UserMoney=6 ,UserPost=7 ,UserWealth=8 ,UserEp=9 ,UserCp=10 ,UserGroupID=11 ,BoardID=12,UserTicket=13,BuyType=14,ToolsImg=15
'---------------------------------------------------
'事件记录过程：Call Dvbbs.ToolsLog(道具ID，发生数量，金币发生额，点券发生额，记录事件类型，备注内容，用户最后剩余金币和点券（金币|点券）)
'---------------------------------------------------
Sub ToolsMain()
	Dv_Tools.ChkUseTools '检查道具使用权限
	Select Case Dv_Tools.ToolsID
	Case 1 :  Tools_1
	Case 2 :  Tools_2
	Case 3 :  Tools_3
	Case 4 :  Tools_4
	Case 5 :  Tools_5
	Case 6 :  Tools_6
	Case 7 :  Tools_7
	Case 8 :  Tools_8
	Case 9 :  Tools_9
	Case 10 : Tools_10
	Case 11 : Tools_11
	Case 12 : Tools_12
	Case 13 : Tools_13
	Case 14 : Tools_14
	Case 16 : Tools_16
	Case 17 : Tools_17
	Case 18 : Tools_18
	Case 19 : Tools_19
	Case 20 : Tools_20
	Case 21 : Tools_21
	Case 22 : Tools_22
	Case 23 : Tools_23
	Case 24 : Tools_24
	Case 25 : Tools_25
	Case 26 : Tools_26
	Case 27 : Tools_27
	Case 28 : Tools_28
	Case 29 : Tools_29
	Case Else
		Dv_Tools.ShowErr(3)
	End Select
End Sub

'------------------------------------------------------------------------------------------------------
'道具处理过程
'------------------------------------------------------------------------------------------------------

'---------------------------------------------------
'道具:转让器，可进行道具、金币和点券的转让
'---------------------------------------------------
Sub Tools_1()
	Dim Rs,Sql
	Dim T_Title,T_UseTools,T_PostTable
	Dim iUserInfo
	ChkAction = True
	If ToUserID = 0 Then ChkAction = False
	If ChkAction = False Then Dvbbs.AddErrCode(42) : Exit Sub
	'判断目标用户使用权限并取出目标用户信息
	Dv_Tools.ChkToUseTools(Request("ToUserID"))

	If Request("ToolsAction")="SendTools" Then
		Dim SendToolsID,SendToolsNum,SendMoneyNum,SendTicketNum
		SendToolsID = Dv_Tools.CheckNumeric(Request("SendToolsID"))
		SendToolsNum = Dv_Tools.CheckNumeric(Request("SendToolsNum"))
		SendMoneyNum = CCur(Abs(Dv_Tools.CheckNumeric(Request("SendMoneyNum"))))
		SendTicketNum = CCur(Abs(Dv_Tools.CheckNumeric(Request("SendTicketNum"))))
		If (SendToolsID=0 Or SendToolsNum=0) And SendMoneyNum=0 And SendTicketNum=0 Then
			LogMsg = "由于您没有正确填写相应的转让内容，使用道具不成功！"
		Else
			If Dvbbs.UserID = Clng(Dv_Tools.ToUserInfo(0)) Then
				Dv_Tools.ShowErr(14)
				Exit Sub
			End If
			LogMsg = "使用：<B>"& Dv_Tools.ToolsInfo(1) &"</B>成功"
			'金币转让
			If SendMoneyNum > 0 Then
				If CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text) < SendMoneyNum Then Dv_Tools.ShowErr(17) : Exit Sub
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text = cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text) - cCur(SendMoneyNum)
				LogMsg = LogMsg & "，转给"&Dv_Tools.ToUserInfo(1)&"<B>"&SendMoneyNum&"</B>个金币"
				Dvbbs.Execute("Update Dv_User Set UserMoney = UserMoney - "&SendMoneyNum&" Where UserID=" & Dvbbs.UserID)
				Dvbbs.Execute("Update Dv_User Set UserMoney = UserMoney + "&SendMoneyNum&" Where UserID=" & Dv_Tools.ToUserInfo(0))
			End If
			'点券转让
			If SendTicketNum > 0 Then
				If CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text) < SendTicketNum Then Dv_Tools.ShowErr(17) : Exit Sub
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text = cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text) - cCur(SendTicketNum)
				LogMsg = LogMsg & "，转给"&Dv_Tools.ToUserInfo(1)&"<B>"&SendTicketNum&"</B>张点券"
				Dvbbs.Execute("Update Dv_User Set UserTicket = UserTicket - "&SendTicketNum&" Where UserID=" & Dvbbs.UserID)
				Dvbbs.Execute("Update Dv_User Set UserTicket = UserTicket + "&SendTicketNum&" Where UserID=" & Dv_Tools.ToUserInfo(0))
			End If
			'道具转让
			If SendToolsID > 0 And SendToolsNum > 0 Then
				Dim Trs,UserToolsNum
				UserToolsNum = 0
				Sql = "Select ID,UserID,UserName,ToolsID,ToolsName,ToolsCount,SaleCount,UpdateTime From [Dv_Plus_Tools_Buss] Where ToolsCount>0 and UserID="& Dvbbs.UserID &" and ToolsID="& SendToolsID
				Set Trs = Dvbbs.Plus_Execute(Sql)
				If Trs.Eof Then
					Response.redirect "showerr.asp?ErrCodes=<li>所选取转让的道具不存在，请购买了相应的道具再执行转让！&action=NoHeadErr"
					Exit Sub
				Else
					UserToolsNum = Trs(5)
					If UserToolsNum<SendToolsNum Then
						Response.redirect "showerr.asp?ErrCodes=<li>你目前只能转让("&UserToolsNum&")个道具！&action=NoHeadErr"
						Exit Sub
					End If
				End If
				Trs.Close
				Set Trs = Dvbbs.Plus_Execute("Select ToolsName From Dv_Plus_Tools_Info Where ID=" & SendToolsID)
				If Not (Trs.Eof And Trs.Bof) Then
					LogMsg = LogMsg & "，转给"&Dv_Tools.ToUserInfo(1)&"<B>"&SendToolsNum&"</B>个"&Trs(0)&"道具"
				End If
				Trs.Close
				Set Trs=Nothing
				'更新用户和系统使用数量
				Call UpdateUserTools(Dvbbs.UserID,Dv_Tools.ToolsID,1)
				'更新用户道具数量
				Call UpdateBussTools(Dvbbs.UserID,SendToolsID,SendToolsNum)	
				Call UpdateBussTools(Dv_Tools.ToUserInfo(0),SendToolsID,-SendToolsNum)
			End If
			Call Dvbbs.ToolsLog(Dv_Tools.ToolsID,1,0,0,2,LogMsg,Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text&"|"&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)
		End If
		Dvbbs.Dvbbs_Suc(LogMsg)
	Else
%>
<table border="0" cellpadding=3 cellspacing=1 align=center class=Tableborder1 Style="Width:99%">
	<tr>
	<th height=23 colspan=2>使用道具 <%=Dv_Tools.ToolsInfo(1)%></th></tr>
	<tr><td height=23 class=Tablebody1 colspan=2>
	<B>说明</B>：<br />1、使用本道具可将您自己的金钱、点券或道具转让给目标用户<br />2、目标用户的选择方法：通常在论坛的各种位置只要点击用户名连接即可进入该用户资料页面，浏览帖子过程可点击该贴用户“信息”图标，进入用户资料页面后点击“使用道具”连接即可进入具体的道具操作页面</td></tr>
	<tr>
	<td height=23 class=Tablebody1 width="30%" align=right>目标用户：</td>
	<td height=23 class=Tablebody1 width="70%"><B><%=Dv_Tools.ToUserInfo(1)%></B></td>
	</tr>
	<FORM METHOD=POST ACTION="?ToolsAction=SendTools">
	<input type=hidden value="<%=ToUserID%>" name="ToUserID">
	<input type=hidden value="<%=Dv_Tools.ToolsID%>" name="ToolsID">
	<tr>
	<td height=23 class=Tablebody1 width="30%" align=right>转让道具：</td>
	<td height=23 class=Tablebody1 width="70%">
	<Select Size=1 Name="SendToolsID">
	<Option value=0 selected>请选择要转让的道具</option>
<%
	Set Rs=Dvbbs.Plus_Execute("Select ToolsID,ToolsName,ToolsCount From [Dv_Plus_Tools_Buss] where UserID="& Dvbbs.UserID &" ORDER BY ToolsCount Desc")
	Do While Not Rs.Eof
		Response.Write "<option value="""&Rs(0)&""">拥有"&Rs(1)&Rs(2)&"个</option>"
	Rs.MoveNext
	Loop
	Rs.Close
	Set Rs=Nothing
%>
	</Select>
	转让数量：
	<input type=text size=5 value="0" name="SendToolsNum">
	个
	</td>
	</tr>
	<tr><td height=23 class=Tablebody1 colspan=2 align=center>
	您有 <B><font color=red><%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text%></font></B> 个金币和 <B><font color=red><%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text%></font></B> 张点券可供转让
	</td></tr>
	<tr>
	<td height=23 class=Tablebody1 width="30%" align=right>转让金币：</td>
	<td height=23 class=Tablebody1 width="70%">
	<input type=text size=5 value="0" name="SendMoneyNum">
	个</td>
	</tr>
	<tr>
	<td height=23 class=Tablebody1 width="30%" align=right>转让点券：</td>
	<td height=23 class=Tablebody1 width="70%">
	<input type=text size=5 value="0" name="SendTicketNum">
	个</td>
	</tr>
	<tr><td height=23 class=Tablebody2 colspan=2 align=center>
	<input type=submit value="确认转让" name=submit>
	</td></tr>
	</FORM>
</table>
<%
	End If
End Sub
'---------------------------------------------------
'道具:后悔药，可删除自己发表的帖子，有回复则不能删
'---------------------------------------------------
Sub Tools_2()
	Dim Rs,Sql
	Dim T_Title,T_UseTools,T_PostTable,ToolsParnetID,ToolsIsToday
	ToolsIsToday = 0
'	If ToUserID = 0 Then ChkAction = False
'	If ChkAction = False Then Dvbbs.AddErrCode(42) : Exit Sub
	'判断目标用户使用权限并取出目标用户信息
	Dv_Tools.ChkToUseTools(Dvbbs.UserID)
	If Dvbbs.UserID <> Clng(Dv_Tools.ToUserInfo(0)) Then
		Dv_Tools.ShowErr(15)
		Exit Sub
	End If
	Sql = "Select Title,UseTools,PostTable,Child From [Dv_Topic] Where TopicID="&TopicID&" And PostUserID="&Dvbbs.UserID
	Set Rs = Dvbbs.Execute(Sql)
	If Rs.Eof Then 
		Response.redirect "showerr.asp?ErrCodes=<li>该主题不存在！&action=NoHeadErr"
		Exit Sub
	Else
		If Rs(3)>0 Then
			Response.redirect "showerr.asp?ErrCodes=<li>该贴已有人回复，不能删除，您可自行编辑清除该贴相关内容！&action=NoHeadErr"
			Exit Sub
		End If
		T_PostTable = Rs(2)
	End If
	Rs.Close
	Set Rs=Dvbbs.Execute("Select Topic,UseTools,Body,ParentID,DateAndTime From "&T_PostTable&" Where AnnounceID="&ReplyID&" And PostUserID=" & Dvbbs.UserID)
	If Rs.Eof Then
		Response.redirect "showerr.asp?ErrCodes=<li>该帖子不存在！&action=NoHeadErr"
		Exit Sub
	Else
		If Rs(0)="" Or IsNull(Rs(0)) Then
			T_Title = Left(Rs(2),25)
		Else
			T_Title = Rs(0)
		End If
		ToolsParnetID = Rs(3)
		T_UseTools = LoadUserTools(Rs(1),Dv_Tools.ToolsID)
		If DateDiff("d",Rs(4),Now())=0 Then ToolsIsToday = 1
	End If
	Rs.Close
	If ToolsParnetID = 0 Then
		Sql = "Update Dv_Topic Set BoardID=444,locktopic="&Dvbbs.BoardID&",UseTools='"& T_UseTools &"' Where TopicID=" & TopicID
		Dvbbs.Execute(Sql)
		Sql = "Update "&T_PostTable&" Set BoardID=444,locktopic="&Dvbbs.BoardID&",UseTools='"& T_UseTools &"' Where AnnounceID=" & ReplyID
		Dvbbs.Execute(Sql)
	Else
		Sql = "Update "&T_PostTable&" Set BoardID=444,locktopic="&Dvbbs.BoardID&",UseTools='"& T_UseTools &"' Where AnnounceID=" & ReplyID
		Dvbbs.Execute(Sql)
	End If
	'更新所有版面帖子数
	AllboardNumSub ToolsIsToday,1,1
	'更新相关版面帖子数
	Call BoardNumSub(Dvbbs.BoardID,1,1,ToolsIsToday)
	Call UpdateUserTools(Dvbbs.UserID,Dv_Tools.ToolsID,1)
	LogMsg = "使用：<B>"& Dv_Tools.ToolsInfo(1) &"</B>成功，"&LoadTitle(T_Title)&"已成功删除入论坛回收站！"
	Call Dvbbs.ToolsLog(Dv_Tools.ToolsID,1,0,0,1,LogMsg,Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text&"|"&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)
	Dvbbs.Dvbbs_Suc(LogMsg)

End Sub
'---------------------------------------------------
'道具:一级特赦令，可解除单贴屏蔽
'---------------------------------------------------
Sub Tools_3()
	Dim Rs,Sql
	Dim T_Title,T_UseTools,T_PostTable
'	If ToUserID = 0 Then ChkAction = False
'	If ChkAction = False Then Dvbbs.AddErrCode(42) : Exit Sub
	'判断目标用户使用权限并取出目标用户信息
'	Dv_Tools.ChkToUseTools()
	Sql = "Select Title,UseTools,PostTable From [Dv_Topic] Where TopicID="&TopicID
	Set Rs = Dvbbs.Execute(Sql)
	If Rs.Eof Then 
		Response.redirect "showerr.asp?ErrCodes=<li>该主题不存在！&action=NoHeadErr"
		Exit Sub
	Else
		T_PostTable = Rs(2)
	End If
	Rs.Close
	Set Rs=Dvbbs.Execute("Select topic,UseTools,Body,postuserid From "&T_PostTable&" Where AnnounceID="&ReplyID&" And LockTopic=2")
	If Rs.Eof Then
		Response.redirect "showerr.asp?ErrCodes=<li>该帖子不存在或不是屏蔽状态！&action=NoHeadErr"
		Exit Sub
	Else
		If Rs(0)="" Or IsNull(Rs(0)) Then
			T_Title = Left(Rs(2),25)
		Else
			T_Title = Rs(0)
		End If
		T_UseTools = LoadUserTools(Rs(1),Dv_Tools.ToolsID)
	End If
	'判断目标用户使用权限并取出目标用户信息
	Dv_Tools.ChkToUseTools(Rs(3))
	Rs.Close
	Sql = "Update "&T_PostTable&" Set LockTopic=0,UseTools='"& T_UseTools &"' Where AnnounceID=" & ReplyID
	Dvbbs.Execute(Sql)
	Call UpdateUserTools(Dvbbs.UserID,Dv_Tools.ToolsID,1)
	LogMsg = "使用：<B>"& Dv_Tools.ToolsInfo(1) &"</B>成功，"&LoadTitle(T_Title)&"已成功解除单贴屏蔽状态！"
	Call Dvbbs.ToolsLog(Dv_Tools.ToolsID,1,0,0,1,LogMsg,Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text&"|"&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)
	Dvbbs.Dvbbs_Suc(LogMsg)
End Sub
'---------------------------------------------------
'道具:二级特赦令，可解除主题锁定
'---------------------------------------------------
Sub Tools_4()
	Dim Rs,Sql
	Dim T_Title,T_UseTools,T_PostTable
	If ChkAction = False Then Dvbbs.AddErrCode(42) : Exit Sub
	Sql = "Select Title,UseTools,PostTable From [Dv_Topic] Where TopicID="&TopicID&" And LockTopic=1"
	Set Rs = Dvbbs.Execute(Sql)
	If Rs.Eof Then 
		Response.redirect "showerr.asp?ErrCodes=<li>该主题不存在或不是锁定状态！&action=NoHeadErr"
		Exit Sub
	Else
		T_Title = Rs(0)
		T_UseTools = LoadUserTools(Rs(1),Dv_Tools.ToolsID)
		T_PostTable = Rs(2)
	End If
	Rs.Close
	Sql = "Update [Dv_Topic] Set LockTopic=0,UseTools='"& T_UseTools &"' Where TopicID="&TopicID
	Dvbbs.Execute(Sql)
	Call UpdateUserTools(Dvbbs.UserID,Dv_Tools.ToolsID,1)
	LogMsg = "使用：<B>"& Dv_Tools.ToolsInfo(1) &"</B>成功，"&LoadTitle(T_Title)&"已成功解除锁定！"
	Call Dvbbs.ToolsLog(Dv_Tools.ToolsID,1,0,0,1,LogMsg,Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text&"|"&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)
	Dvbbs.Dvbbs_Suc(LogMsg)
End Sub

'---------------------------------------------------
'道具:三级特赦令，解除自己或他人的屏蔽或锁定状态
'---------------------------------------------------
Sub Tools_5()
	Dim Rs,Sql
	Dim T_Title,T_UseTools,T_PostTable
	ChkAction = True
	If ToUserID = 0 Then ChkAction = False
	If ChkAction = False Then Dvbbs.AddErrCode(42) : Exit Sub
	'判断目标用户使用权限并取出目标用户信息
	Dv_Tools.ChkToUseTools(ToUserID)
	Sql = "Select UserID From Dv_User Where UserID="&ToUserID&" And LockUser>0"
	Set Rs = Dvbbs.Execute(Sql)
	If Rs.Eof Then 
		Response.redirect "showerr.asp?ErrCodes=<li>该用户不存在或不是屏蔽或锁定状态！&action=NoHeadErr"
		Exit Sub
	Else
		Dvbbs.Execute("Update Dv_User Set LockUser=0 Where UserID="& Rs(0))
	End If
	Rs.Close
	Call UpdateUserTools(Dvbbs.UserID,Dv_Tools.ToolsID,1)
	LogMsg = "使用：<B>"& Dv_Tools.ToolsInfo(1) &"</B>成功，用户<B>"&Dv_Tools.ToUserInfo(1)&"</B>已成功解除锁定或屏蔽状态！"
	Call Dvbbs.ToolsLog(Dv_Tools.ToolsID,1,0,0,1,LogMsg,Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text&"|"&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)
	Dvbbs.Dvbbs_Suc(LogMsg)
End Sub
'---------------------------------------------------
'道具:吖噗鸡，可使帖子提升到第一页
'---------------------------------------------------
Sub Tools_6()
	Dim Rs,Sql
	Dim T_Title,T_UseTools,T_PostTable
	If ChkAction = False Then Dvbbs.AddErrCode(42) : Exit Sub
	Sql = "Select Title,UseTools,PostTable From [Dv_Topic] Where TopicID="&TopicID&" And LockTopic=1"
	Set Rs = Dvbbs.Execute(Sql)
	If Rs.Eof Then 
		Response.redirect "showerr.asp?ErrCodes=<li>该主题不存在或不是锁定状态！&action=NoHeadErr"
		Exit Sub
	Else
		T_Title = Rs(0)
		T_UseTools = LoadUserTools(Rs(1),Dv_Tools.ToolsID)
		T_PostTable = Rs(2)
	End If
	Rs.Close
	Sql = "Update [Dv_Topic] Set LastPostTime="&SqlNowString&",UseTools='"& T_UseTools &"' Where TopicID="&TopicID
	Dvbbs.Execute(Sql)
	Call UpdateUserTools(Dvbbs.UserID,Dv_Tools.ToolsID,1)
	LogMsg = "使用：<B>"& Dv_Tools.ToolsInfo(1) &"</B>成功，"&LoadTitle(T_Title)&"已成功提升到第一页！"
	Call Dvbbs.ToolsLog(Dv_Tools.ToolsID,1,0,0,1,LogMsg,Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text&"|"&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)
	Dvbbs.Dvbbs_Suc(LogMsg)
End Sub
'---------------------------------------------------
'道具:醒目灯，可将主题变色
'---------------------------------------------------
Sub Tools_7()
	Dim Rs,Sql,i
	Dim T_Title,T_UseTools,T_PostTable,ToolsColorList
	If ChkAction = False Then Dvbbs.AddErrCode(42) : Exit Sub
	Sql = "Select Title,UseTools,PostTable From [Dv_Topic] Where TopicID="&TopicID
	Set Rs = Dvbbs.Execute(Sql)
	If Rs.Eof Then 
		Response.redirect "showerr.asp?ErrCodes=<li>该主题不存在或不是锁定状态！&action=NoHeadErr"
		Exit Sub
	Else
		T_Title = Rs(0)
		T_UseTools = LoadUserTools(Rs(1),Dv_Tools.ToolsID)
		T_PostTable = Rs(2)
	End If
	Rs.Close
	ToolsColorList = "#000000,#F0F8FF,#FAEBD7,#00FFFF,#7FFFD4,#F0FFFF,#F5F5DC,#FFE4C4,#000000,#FFEBCD,#0000FF,#8A2BE2,#A52A2A,#DEB887,#5F9EA0,#7FFF00,#D2691E,#FF7F50,#6495ED,#FFF8DC,#DC143C,#00FFFF,#00008B,#008B8B,#B8860B,#A9A9A9,#006400,#BDB76B,#8B008B,#556B2F,#FF8C00,#9932CC,#8B0000,#E9967A,#8FBC8F,#483D8B,#2F4F4F,#00CED1,#9400D3,#FF1493,#00BFFF,#696969,#1E90FF,#B22222,#FFFAF0,#228B22,#FF00FF,#DCDCDC,#F8F8FF,#FFD700,#DAA520,#808080,#008000,#ADFF2F,#F0FFF0,#FF69B4,#CD5C5C,#4B0082,#FFFFF0,#F0E68C,#E6E6FA,#FFF0F5,#7CFC00,#FFFACD,#ADD8E6,#F08080,#E0FFFF,#FAFAD2,#90EE90,#D3D3D3,#FFB6C1,#FFA07A,#20B2AA,#87CEFA,#778899,#B0C4DE,#FFFFE0,#00FF00,#32CD32,#FAF0E6,#FF00FF,#800000,#66CDAA,#0000CD,#BA55D3,#9370DB,#3CB371,#7B68EE,#00FA9A,#48D1CC,#C71585,#191970,#F5FFFA,#FFE4E1,#FFE4B5,#FFDEAD,#000080,#FDF5E6,#808000,#6B8E23,#FFA500,#FF4500,#DA70D6,#EEE8AA,#98FB98,#AFEEEE,#DB7093,#FFEFD5,#FFDAB9,#CD853F,#FFC0CB,#DDA0DD,#B0E0E6,#800080,#FF0000,#BC8F8F,#4169E1,#8B4513,#FA8072,#F4A460,#2E8B57,#FFF5EE,#A0522D,#C0C0C0,#87CEEB,#6A5ACD,#708090,#FFFAFA,#00FF7F,#4682B4,#D2B48C,#008080,#D8BFD8,#FF6347,#40E0D0,#EE82EE,#F5DEB3,#FFFFFF,#F5F5F5,#FFFF00,#9ACD32"
	If Request("ToolsAction")="SendColor" Then
		If Instr("," & ToolsColorList & ",","," & Request("color") & ",")=0 Then
			Response.redirect "showerr.asp?ErrCodes=<li>错误的颜色参数！&action=NoHeadErr"
			Exit Sub
		End If
		T_Title = "<font color="&Request("color")&">"&T_Title&"</font>"
		Dvbbs.Execute("Update Dv_Topic Set Title='"&Replace(T_Title,"'","''")&"',TopicMode=1,UseTools='"& T_UseTools &"' Where TopicID=" & TopicID)
		Dvbbs.Execute("Update "&T_PostTable&" Set Topic='"&Replace(T_Title,"'","''")&"',UseTools='"& T_UseTools &"' Where RootID="&TopicID&" And ParentID=0")
		Call UpdateUserTools(Dvbbs.UserID,Dv_Tools.ToolsID,1)
		LogMsg = "使用：<B>"& Dv_Tools.ToolsInfo(1) &"</B>成功，"&Replace(Replace(LoadTitle(T_Title),"&lt;","<"),"&gt;",">")&"已成功操作！"
		Call Dvbbs.ToolsLog(Dv_Tools.ToolsID,1,0,0,1,LogMsg,Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text&"|"&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)
		Dvbbs.Dvbbs_Suc(LogMsg)
	Else
	ToolsColorList = Split(ToolsColorList,",")
%>
<table border="0" cellpadding=3 cellspacing=1 align=center class=Tableborder1 Style="Width:99%">
	<tr>
	<th height=23 colspan=2>使用道具 <%=Dv_Tools.ToolsInfo(1)%></th></tr>
	<tr><td height=23 class=Tablebody1 colspan=2>
	<B>说明</B>：本道具可使目标帖子标题变成您所选择的颜色，请在下面选择您所需要的颜色</td></tr>
	<FORM METHOD=POST ACTION="?ToolsAction=SendColor" name="theForm">
<!--	<input type=hidden value="<%=ToUserID%>" name="ToUserID">  -->
	<input type=hidden value="<%=Dvbbs.BoardID%>" name="BoardID">
	<input type=hidden value="<%=TopicID%>" name="TopicID">
	<input type=hidden value="<%=ReplyID%>" name="ReplyID">
	<input type=hidden value="<%=Dv_Tools.ToolsID%>" name="ToolsID">
	<tr>
	<td height=23 class=Tablebody1 width="30%" align=right>颜色列表：</td>
	<td height=23 class=Tablebody1 width="70%">
	<SELECT onChange="document.getElementById('TopicColor').color=options[selectedIndex].value;" name="color"> 
	<%
	For i=0 To Ubound(ToolsColorList)
		Response.Write "<option style=""background-color:"&ToolsColorList(i)&";color: "&ToolsColorList(i)&""" value="""&ToolsColorList(i)&""">"&ToolsColorList(i)&"</option>"
	Next
	%>
	</SELECT>
	</td>
	</tr>
	<tr>
	<td height=23 class=Tablebody1 width="30%" align=right>使用效果：</td>
	<td height=23 class=Tablebody1 width="70%"><font id=TopicColor><%=Server.HtmlEncode(T_Title)%></font></td>
	</tr>
	<tr><td height=23 class=Tablebody2 colspan=2 align=center>
	<input type=submit value="确认使用" name=submit>
	</td></tr>
	</FORM>
</table>
<%
	End If
End Sub

'---------------------------------------------------
'道具:水晶球，可查看发贴用户IP
'---------------------------------------------------
Sub Tools_8()
	Dim Rs,Sql
	Dim T_Title,T_UseTools,T_PostTable,ToUserToolsIP
'	If ToUserID = 0 Then ChkAction = False
'	If ChkAction = False Then Dvbbs.AddErrCode(42) : Exit Sub
	'判断目标用户使用权限并取出目标用户信息
'	Dv_Tools.ChkToUseTools()
	Sql = "Select Title,UseTools,PostTable From [Dv_Topic] Where TopicID="&TopicID
	Set Rs = Dvbbs.Execute(Sql)
	If Rs.Eof Then 
		Response.redirect "showerr.asp?ErrCodes=<li>该主题不存在！&action=NoHeadErr"
		Response.write "1"
		Exit Sub
	Else
		T_PostTable = Rs(2)
	End If
	Rs.Close
	Set Rs=Dvbbs.Execute("Select Topic,UseTools,Body,IP,postuserid From "&T_PostTable&" Where AnnounceID="&ReplyID)
	If Rs.Eof Then
		Response.redirect "showerr.asp?ErrCodes=<li>该帖子不存在！&action=NoHeadErr"
		Exit Sub
	Else
		If Rs(0)="" Or IsNull(Rs(0)) Then
			T_Title = Left(Rs(2),25)
		Else
			T_Title = Rs(0)
		End If
		T_UseTools = LoadUserTools(Rs(1),Dv_Tools.ToolsID)
		ToUserToolsIP = Rs(3)
	End If
	'判断目标用户使用权限并取出目标用户信息
	Dv_Tools.ChkToUseTools(Rs(4))
	Rs.Close
	Sql = "Update "&T_PostTable&" Set UseTools='"& T_UseTools &"' Where AnnounceID=" & ReplyID
	Dvbbs.Execute(Sql)
	Call UpdateUserTools(Dvbbs.UserID,Dv_Tools.ToolsID,1)
	LogMsg = "使用：<B>"& Dv_Tools.ToolsInfo(1) &"</B>成功，"&LoadTitle(T_Title)&"中帖子编号为"&ReplyID&"的发贴IP是："&ToUserToolsIP&"！"
	Call Dvbbs.ToolsLog(Dv_Tools.ToolsID,1,0,0,1,LogMsg,Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text&"|"&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)
	Dvbbs.Dvbbs_Suc(LogMsg)
End Sub
'---------------------------------------------------
'道具:追踪器，可查看发贴用户的IP和来源
'---------------------------------------------------
Sub Tools_9()
	Dim Rs,Sql
	Dim T_Title,T_UseTools,T_PostTable,ToUserToolsIP,ToUserToolsIP_1,ToUserToolsAddress
'	If ToUserID = 0 Then ChkAction = False
'	If ChkAction = False Then Dvbbs.AddErrCode(42) : Exit Sub
	'判断目标用户使用权限并取出目标用户信息
'	Dv_Tools.ChkToUseTools()
	Sql = "Select Title,UseTools,PostTable From [Dv_Topic] Where TopicID="&TopicID
	Set Rs = Dvbbs.Execute(Sql)
	If Rs.Eof Then 
		Response.redirect "showerr.asp?ErrCodes=<li>该主题不存在！&action=NoHeadErr"
		Exit Sub
	Else
		T_PostTable = Rs(2)
	End If
	Rs.Close
	Set Rs=Dvbbs.Execute("Select Topic,UseTools,Body,IP,postuserid From "&T_PostTable&" Where AnnounceID="&ReplyID)
	If Rs.Eof Then
		Response.redirect "showerr.asp?ErrCodes=<li>该帖子不存在！&action=NoHeadErr"
		Exit Sub
	Else
		If Rs(0)="" Or IsNull(Rs(0)) Then
			T_Title = Left(Rs(2),25)
		Else
			T_Title = Rs(0)
		End If
		T_UseTools = LoadUserTools(Rs(1),Dv_Tools.ToolsID)
		ToUserToolsIP = Rs(3)
	End If
	'判断目标用户使用权限并取出目标用户信息
	Dv_Tools.ChkToUseTools(Rs(4))
	Rs.Close
	Sql = "Update "&T_PostTable&" Set UseTools='"& T_UseTools &"' Where AnnounceID=" & ReplyID
	Dvbbs.Execute(Sql)
	ToUserToolsIP_1 = ToUserToolsIP
	ToUserToolsAddress = lookaddress(ToUserToolsIP_1)
	Call UpdateUserTools(Dvbbs.UserID,Dv_Tools.ToolsID,1)
	LogMsg = "使用：<B>"& Dv_Tools.ToolsInfo(1) &"</B>成功，"&LoadTitle(T_Title)&"中帖子编号为"&ReplyID&"的发贴IP是："&ToUserToolsIP&"，来源是："&ToUserToolsAddress&"！"
	Call Dvbbs.ToolsLog(Dv_Tools.ToolsID,1,0,0,1,LogMsg,Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text&"|"&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)
	Dvbbs.Dvbbs_Suc(LogMsg)

End Sub
'---------------------------------------------------
'道具:一星龙珠，可将用户所有负分转为0
'---------------------------------------------------
Sub Tools_10()
	Dim Rs,Sql
	Dim T_Title,T_UseTools,T_PostTable
	ChkAction = True
	If ToUserID = 0 Then ChkAction = False
	If ChkAction = False Then Dvbbs.AddErrCode(42) : Exit Sub
	'判断目标用户使用权限并取出目标用户信息
	Dv_Tools.ChkToUseTools(ToUserID)
	'更新用户分值信息
	Sql = "Select UserWealth,UserEP,UserCP,UserPower,UserDel From Dv_User Where UserID= " & Dv_Tools.ToUserInfo(0)
	Set Rs = Dvbbs.iCreateObject ("adodb.recordset")
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open Sql,Conn,1,3
	If Rs("UserWealth") < 0 Then Rs("UserWealth") = 0
	If Rs("UserEP") < 0 Then Rs("UserEP") = 0
	If Rs("UserCP") < 0 Then Rs("UserCP") = 0
	If Rs("UserPower") < 0 Then Rs("UserPower") = 0
	If Rs("UserDel") < 0 Then Rs("UserDel") = 0
	Rs.Update
	Rs.Close
	Set Rs=Nothing
	'更新用户和系统使用数量
	Call UpdateUserTools(Dvbbs.UserID,Dv_Tools.ToolsID,1)
	LogMsg = "使用：<B>"& Dv_Tools.ToolsInfo(1) &"</B>成功，成功将用户<b>"&Dv_Tools.ToUserInfo(1)&"</b>的所有负分转正！"
	Call Dvbbs.ToolsLog(Dv_Tools.ToolsID,1,0,0,1,LogMsg,Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text&"|"&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)
	Dvbbs.Dvbbs_Suc(LogMsg)

End Sub
'---------------------------------------------------
'道具:二星龙珠，可将用户积分负分转为0
'---------------------------------------------------
Sub Tools_11()
	Dim Rs,Sql
	Dim T_Title,T_UseTools,T_PostTable
	ChkAction = True
	If ToUserID = 0 Then ChkAction = False
	If ChkAction = False Then Dvbbs.AddErrCode(42) : Exit Sub
	'判断目标用户使用权限并取出目标用户信息
	Dv_Tools.ChkToUseTools(ToUserID )
	'更新用户分值信息
	Sql = "Select UserEP From Dv_User Where UserID= " & Dv_Tools.ToUserInfo(0)
	Set Rs = Dvbbs.iCreateObject ("adodb.recordset")
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open Sql,Conn,1,3
	If Rs("UserEP") < 0 Then Rs("UserEP") = 0
	Rs.Update
	Rs.Close
	Set Rs=Nothing
	'更新用户和系统使用数量
	Call UpdateUserTools(Dvbbs.UserID,Dv_Tools.ToolsID,1)
	LogMsg = "使用：<B>"& Dv_Tools.ToolsInfo(1) &"</B>成功，成功将用户<b>"&Dv_Tools.ToUserInfo(1)&"</b>的积分负分转正！"
	Call Dvbbs.ToolsLog(Dv_Tools.ToolsID,1,0,0,1,LogMsg,Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text&"|"&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)
	Dvbbs.Dvbbs_Suc(LogMsg)
End Sub
'---------------------------------------------------
'道具:狗仔队，可在用户上线第一时间获知
'---------------------------------------------------
Sub Tools_12()
	Dim Rs,Sql
	Dim T_Title,T_UseTools,T_PostTable
	ChkAction = True
	If ToUserID = 0 Then ChkAction = False
	If ChkAction = False Then Dvbbs.AddErrCode(42) : Exit Sub
	'判断目标用户使用权限并取出目标用户信息
	Dv_Tools.ChkToUseTools(ToUserID )
	If Dvbbs.UserID = Clng(Dv_Tools.ToUserInfo(0)) Then
		Dv_Tools.ShowErr(14)
		Exit Sub
	End If
	'更新用户信息
	Sql = "Select FollowMsgID From Dv_User Where UserID= " & Dv_Tools.ToUserInfo(0)
	Set Rs = Dvbbs.iCreateObject ("adodb.recordset")
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open Sql,Conn,1,3
	If Rs(0)="" Or IsNull(Rs(0)) Then
		Rs(0) = Dvbbs.Membername
	Else
		Rs(0) = Rs(0) & "," & Dvbbs.Membername
	End If
	Rs.Update
	Rs.Close
	Set Rs=Nothing
	'更新用户和系统使用数量
	Call UpdateUserTools(Dvbbs.UserID,Dv_Tools.ToolsID,1)
	LogMsg = "使用：<B>"& Dv_Tools.ToolsInfo(1) &"</B>成功，成功跟踪用户<b>"&Dv_Tools.ToUserInfo(1)&"</b>，用户上线后会第一时间通知您！"
	Call Dvbbs.ToolsLog(Dv_Tools.ToolsID,1,0,0,1,LogMsg,Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text&"|"&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)
	Dvbbs.Dvbbs_Suc(LogMsg)
End Sub

'---------------------------------------------------
'道具:救生圈，可将帖子固顶6小时
'---------------------------------------------------
Sub Tools_13()
	Dim Rs,Sql
	Dim T_Title,T_UseTools,T_PostTable,LastPostTime
	If ChkAction = False Then Dvbbs.AddErrCode(42) : Exit Sub
	Sql = "Select Title,UseTools,PostTable From [Dv_Topic] Where TopicID="&TopicID
	Set Rs = Dvbbs.Execute(Sql)
	If Rs.Eof Then 
		Response.redirect "showerr.asp?ErrCodes=<li>该主题不存在！&action=NoHeadErr"
		Exit Sub
	Else
		T_Title = Rs(0)
		T_UseTools = LoadUserTools(Rs(1),Dv_Tools.ToolsID)
		T_PostTable = Rs(2)
	End If
	Rs.Close
	LastPostTime = DateAdd("h",6,now)
	Sql = "Update [Dv_Topic] Set LastPostTime='"&LastPostTime&"',UseTools='"& T_UseTools &"' Where TopicID="&TopicID
	Dvbbs.Execute(Sql)
	Call UpdateUserTools(Dvbbs.UserID,Dv_Tools.ToolsID,1)
	LogMsg = "使用：<B>"& Dv_Tools.ToolsInfo(1) &"</B>成功，"&LoadTitle(T_Title)&"已成功固顶6小时！"
	Call Dvbbs.ToolsLog(Dv_Tools.ToolsID,1,0,0,1,LogMsg,Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text&"|"&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)
	Dvbbs.Dvbbs_Suc(LogMsg)
End Sub
'---------------------------------------------------
'道具:大救生圈，可将帖子固顶12小时
'---------------------------------------------------
Sub Tools_14()
	Dim Rs,Sql
	Dim T_Title,T_UseTools,T_PostTable,LastPostTime
	If ChkAction = False Then Dvbbs.AddErrCode(42) : Exit Sub
	Sql = "Select Title,UseTools,PostTable From [Dv_Topic] Where TopicID="&TopicID
	Set Rs = Dvbbs.Execute(Sql)
	If Rs.Eof Then 
		Response.redirect "showerr.asp?ErrCodes=<li>该主题不存在！&action=NoHeadErr"
		Exit Sub
	Else
		T_Title = Rs(0)
		T_UseTools = LoadUserTools(Rs(1),Dv_Tools.ToolsID)
		T_PostTable = Rs(2)
	End If
	Rs.Close
	LastPostTime = DateAdd("h",12,now)
	Sql = "Update [Dv_Topic] Set LastPostTime='"&LastPostTime&"',UseTools='"& T_UseTools &"' Where TopicID="&TopicID
	Dvbbs.Execute(Sql)
	Call UpdateUserTools(Dvbbs.UserID,Dv_Tools.ToolsID,1)
	LogMsg = "使用：<B>"& Dv_Tools.ToolsInfo(1) &"</B>成功，"&LoadTitle(T_Title)&"已成功固顶12小时！"
	Call Dvbbs.ToolsLog(Dv_Tools.ToolsID,1,0,0,1,LogMsg,Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text&"|"&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)
	Dvbbs.Dvbbs_Suc(LogMsg)
End Sub

'---------------------------------------------------
'道具:时空转移机 可将自已的帖子移动到任意版面（隐含、特殊限定版面除外）。
'---------------------------------------------------
Sub Tools_15()

End Sub

'---------------------------------------------------
'道具:照妖镜 可查看匿名发帖用户名。
'---------------------------------------------------
Sub Tools_16()
	Dim Rs,Sql
	Dim T_Title,T_UseTools,T_PostTable,ToUserToolsName
	Sql = "Select Title,UseTools,PostTable From [Dv_Topic] Where TopicID="&TopicID
	Set Rs = Dvbbs.Execute(Sql)
	If Rs.Eof Then 
		Response.redirect "showerr.asp?ErrCodes=<li>该主题不存在！&action=NoHeadErr"
		Exit Sub
	Else
		T_PostTable = Rs(2)
	End If
	Rs.Close
	Set Rs=Dvbbs.Execute("Select Topic,UseTools,Body,postuserid From "&T_PostTable&" Where AnnounceID="&ReplyID)
	If R
	End If
	Rs.Close
	LastPostTime = DateAdd("h",12,now)
	Sql = "Update [Dv_Topic] Set LastPostTime='"&LastPostTime&"',UseTools='"& T_UseTools &"' Where TopicID="&TopicID
	Dvbbs.Execute(Sql)
	Call UpdateUserTools(Dvbbs.UserID,Dv_Tools.ToolsID,1)
	LogMsg = "浣跨敤锛