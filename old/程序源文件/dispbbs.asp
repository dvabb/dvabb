<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="inc/dv_ubbcode.asp"-->
<!--#include file="inc/dv_template.inc"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/code_encrypt.asp"-->
<!--#include file="inc/dv_pageclass.asp"-->
<%
If Dvbbs.BoardID < 1 Then Response.Write "参数错误":Response.End
If Request("page") <> "" And CStr(Dvbbs.CheckNumeric(Request("page"))) <> Request("page") Then
    Response.Write "参数错误"
    Response.End
End If
If Dvbbs.GroupSetting(2)="0" Then Dvbbs.AddErrcode(31):Dvbbs.ShowErr():response.End
Dim PostUserid, G_TopicTitle, G_IsVote, G_Childs, G_PollID, G_LockTopic, G_Hits, G_Expression,FlashId
Dim G_ItemList, G_ItemsPerPage, G_CurrentPage, G_Pages, G_Moved
Dim G_UserList
Dim G_UserItemQuery
Dim G_Floor
Dim G_CanReply
Dim Dv_ubb
Dim CanRead,TrueMaster,Skin
'以下定义的变量在Dv_ubbcode.asp页面会用到
Dim EmotPath
Dim TotalUsetable
Dim PostBuyUser
Dim UserName
Dim T_GetMoneyType
Dim AnnounceID, ReplyID, Replyid_a, AnnounceID_a, RootID_a
Dim IsThisBoardMaster '确定当前用户是否本版版主，防止下面的操作影响到 Dvbbs.BoardMaster导致出错
IsThisBoardMaster = Dvbbs.BoardMaster
'浏览购买帖权限
CanRead=False
TrueMaster=False
Rem 为兼顾管理菜单显示,对有管理权限的暂时当版主等级处理,为的是显示管理菜单.
If Not Dvbbs.BoardMaster Then
	If Dvbbs.UserID > 0 Then
		If Dvbbs.GroupSetting(18) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(19) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(20) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(21) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(22) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(23) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(24) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(25) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(26) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(27) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(28) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(29) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(30) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(31) = "1" Then
			Dvbbs.BoardMaster=True
		End If
	End If
Else
	TrueMaster=True
End If
If Dvbbs.BoardMaster Then CanRead=True
Dim authorid
authorid = Dvbbs.CheckNumeric(Request("authorid"))

'初始数据
AnnounceID		= 0			'主题ID
G_UserItemQuery = "userid,username,useremail,userpost,usertopic,usersign,usersex,userface,userwidth,userheight,joindate,lastlogin,userlogins,lockuser,userclass,userwealth,userep,usercp,userpower,userdel,userisbest,usertitle,userhidden,usermoney,userticket,titlepic,usergroupid,userim,useremail" '查询用户的字段列表
'0-userid,1-username,2-useremail,3-userpost,4-usertopic,5-usersign,6-usersex,7-userface,8-userwidth,9-userheight,10-joindate,11-lastlogin,12-userlogins,13-lockuser,14-userclass,15-userwealth,16-userep,17-usercp,18-userpower,19-userdel,20-userisbest,21-usertitle,22-userhidden,23-usermoney,24-userticket,25-titlepic,26-UserGroupID,27-userim

Rem 增加勋章插件字段查询，Fish 2010-3-11
if Cint(dvbbs.Forum_Setting(104))=1 then G_UserItemQuery = G_UserItemQuery & ",UserMedal"
if Cint(dvbbs.Forum_Setting(102))=1 then G_UserItemQuery = G_UserItemQuery & ",RLActTimeT" 

'--------------------荣誉勋章------------------------
Dim G_MedalData
G_MedalData = GetMedalData

Function GetMedalData()
	if Cint(dvbbs.Forum_Setting(104))=0 Then Exit Function
	Dim Rs,dTemp
	Dvbbs.Name = "Medal"
	If Dvbbs.ObjIsEmpty() Then
		Set Rs = Dvbbs.Execute("SELECT id,MedalName,MedalPic,MedalDesc FROM Dv_Medal")
		If Not Rs.Eof Then
			Dvbbs.Value = Rs.GetRows(-1)
		End If
		Rs.Close : Set Rs = Nothing
	End If
	GetMedalData = Dvbbs.Value
End Function
'----------------------End----------------------------

LoadTopicInfo
LoadBBSListData
LoadUserListData
Dvbbs.LoadTemplates("dispbbs")
Dvbbs.Nav()
Dvbbs.Head_var 1,"","",""
Response.Write GetForumTextAd(2)
Dvbbs.ActiveOnline()

EmotPath=Split(Dvbbs.Forum_emot,"|||")(0)		'em心情路径
Set Dv_ubb=new Dvbbs_UbbCode
Dv_ubb.PostType=1
TPL_Scan	Template.html(0)'Dvbbs.ReadTextFile("dispbbsnew.tpl")'
TPL_Flush
Set Dv_ubb=Nothing

Sub LoadTopicInfo()
	AnnounceID	= Dvbbs.CheckNumeric(Request("ID"))
	If 0=AnnounceID Then Dvbbs.AddErrCode(30):Dvbbs.Showerr()
	G_CurrentPage = Dvbbs.CheckNumeric(Request("star"))
	If 0=G_CurrentPage Then G_CurrentPage=1
	Skin=Request("Skin")
	If Skin="" Or Not IsNumeric(Skin) Then Skin=Dvbbs.Board_setting(42)
	Dim Rs, SQL, iLockSet, iTopicMode, sMove
	sMove		= request("move")
	iLockSet	= Dvbbs.CheckNumeric(Dvbbs.Board_Setting(71))
	SQL="Select top 1 TopicID,boardid,title,hits,isvote,child,pollid,LockTopic,PostTable,TopicMode,DateAndTime,Expression,GetMoneyType,PostUserid From Dv_topic where "
	If ""=sMove Then
		SQL	= SQL & ("topicID=" & AnnounceID)
	Else
		SQL	= SQL & "BoardID=" & Dvbbs.BoardID & " and topicID"
		If "next"=sMove Then
			SQL	= SQL & ("<" & AnnounceID & " order by topicID desc")
		Else
			SQL	= SQL & (">" & AnnounceID & " order by topicID")
		End If
	End If
	If Not IsObject(Conn) Then ConnectionDatabase
	Set Rs=Dvbbs.iCreateObject("Adodb.RecordSet")
	Rs.Open SQL,Conn,1,3
	Dvbbs.SqlQueryNum=Dvbbs.SqlQueryNum+1
	If Rs.eof Or Rs.bof Then
		If ""<>sMove Then
			Response.Write "<script language=""javascript"">alert(""已经是最后一条主题了！"");history.go(-1);</script>"
			Rs.Close
			Set Rs=Nothing
			Dvbbs.PageEnd
			Response.End
		Else
			Dvbbs.AddErrcode(32)
		End If
	Else
		If CStr(Rs("BoardID"))<>CStr(Dvbbs.BoardID) Then Dvbbs.AddErrCode(29)
		G_Hits			= Dvbbs.CheckNumeric(Rs("hits"))
		Rs("hits")		= G_Hits+1
		G_LockTopic		= Rs("LockTopic")
		If 0=G_LockTopic And iLockSet<>0 And Datediff("d", Rs("DateAndTime"),Now())>iLockSet Then
			G_LockTopic = 1
			Rs("LockTopic")	= G_LockTopic
		End If
		On Error Resume Next
		Rs.Update
		If Err Then Err.Clear
		G_TopicTitle	= Rs("title")
		G_IsVote		= Rs("isvote")
		G_Childs		= Rs("child")
		G_PollID		= Rs("pollid")
		G_Expression	= Rs("Expression")
		T_GetMoneyType	= Rs("GetMoneyType")
		TotalUsetable	= Rs("PostTable")
		iTopicMode		= Rs("topicmode")
		AnnounceID		= Rs("TopicID")
		PostUserid		= Rs("PostUserid")
	End If
	Rs.Close
	Set Rs=Nothing
	If authorid>0 Then
	    Dim Rs1
		Set rs1=dvbbs.execute("select count(*) from "&TotalUsetable&" where boardid not in(444,777) and boardid="&dvbbs.boardid&" and postuserid="&authorid&" and rootid="&AnnounceID)
		G_Childs=Rs1(0):Rs1.Close:Set Rs=Nothing
	End If
	Dvbbs.Showerr()
	G_TopicTitle		= Dvbbs.ChkBadWords(G_TopicTitle)
	ReplyID			= Dvbbs.CheckNumeric(Request("ReplyID"))
	If 0=ReplyID Then ReplyID=AnnounceID
	If iTopicMode<>1 Then G_TopicTitle=replace(G_TopicTitle, "<", "&lt;")
	Dvbbs.Stats			= G_TopicTitle	
	G_Childs			= G_Childs+1
	Select Case iTopicMode
		Case 2	G_TopicTitle	= "<font color=""red"">"&G_TopicTitle&"</font>"
		Case 3	G_TopicTitle	= "<font color=""blue"">"&G_TopicTitle&"</font>"
		Case 4	G_TopicTitle	= "<font color=""green"">"&G_TopicTitle&"</font>"
		Case Else
	End Select
End Sub

Sub LoadBBSListData()
	On Error Resume next
	G_ItemsPerPage	= Dvbbs.CheckNumeric(Dvbbs.Board_Setting(27))
	G_Pages	= G_Childs \ G_ItemsPerPage
	If (G_Childs Mod G_ItemsPerPage)<>0 Then G_Pages = G_Pages + 1
	If G_Pages<=0 Then G_Pages = 1
	If G_CurrentPage > G_Pages Then G_CurrentPage = G_Pages
	G_Moved	= G_ItemsPerPage*(G_CurrentPage-1)
	Dim Rs, SQL, Cmd,sqlfields,sqlfieldswhere,authorwhere
	sqlfields="AnnounceID,UserName,Topic,dateandtime,body,Expression,ip,RootID,signflag,isbest,PostUserid,layer,isagree,GetMoneyType,IsUpload,Ubblist,LockTopic,GetMoney,UseTools,PostBuyUser,ParentID,FlashId"
	If authorid>0 Then
	    authorwhere=" and postuserid="&authorid
	Else
	    authorwhere=""
	End If
	If 1=Skin Then
		If ReplyID=AnnounceID Then
			sqlfieldswhere=" RootID="& AnnounceID &" and Boardid="& Dvbbs.Boardid&authorwhere
			SQL="Select Top 1 "&sqlfields&" From "& TotalUseTable & " where " &sqlfieldswhere
		Else
			sqlfieldswhere=" AnnounceID="&ReplyID&" and Boardid="& Dvbbs.Boardid&authorwhere
			SQL="Select "&sqlfields&" From "& TotalUseTable &" where "& sqlfieldswhere
		End If
	Else
		sqlfieldswhere=" RootID="& ReplyID &" and Boardid="& Dvbbs.Boardid&authorwhere
		SQL="Select "&sqlfields&" From "& TotalUsetable & " where "& sqlfieldswhere&" Order By Announceid" '0-AnnounceID,1-UserName,2-Topic,3-dateandtime,4-body,5-Expression,6-ip,7-RootID,8-signflag,9-isbest,10-PostUserid,11-layer,12-isagree,13-GetMoneyType,14-IsUpload,15-Ubblist,16-LockTopic,17-GetMoney,18-UseTools,19-PostBuyUser,20-ParentID
	End If
	'response.Write sqlfieldswhere
    

	If IsSqlDataBase=1 And IsBuss=1 And Skin=0 Then
		Dim mypage
		Set mypage=new Pager
		'If Not IsObject(Conn) Then ConnectionDatabase
		mypage.getconn=conn '得到数据库连接
		mypage.pagesize=G_ItemsPerPage '定义分页每一页的记录数
		mypage.TableName=TotalUsetable '要查询的表名
		mypage.Tablezd=sqlfields
		mypage.KeyName="announceid"
		mypage.OrderType=0
		mypage.PageWhere=sqlfieldswhere
		mypage.GetStyle =1
		Set Rs=mypage.getrs()
		If Not Rs.EoF Then
			G_ItemList=Rs.GetRows(-1)
		Else
			Rs.close():Set Rs=Nothing:Dvbbs.AddErrCode(29)
		End If
		Rs.close():Set Rs=Nothing
		'Set Cmd =  Nothing
	Else
		Set Rs=Dvbbs.Execute(SQL)
		If Rs.eof Or Rs.bof Then
			Dvbbs.AddErrCode(29)
		Else
			On Error Resume Next
			If 1<>Skin Then Rs.Move(G_Moved)
			If Err Then Err.clear
			If Not Rs.eof Then
				G_ItemList = Rs.GetRows(G_ItemsPerPage)
			Else
				Dvbbs.AddErrCode(29)
			End If
		End If
		Rs.Close
		Set Rs=Nothing
	End If
	Dvbbs.Showerr()
	G_CanReply=False '是否允许回复
	If Not Dvbbs.Board_Setting(0)="1"  And Cint(G_LockTopic)=0 Then
		If Dvbbs.GroupSetting(5)="1" Then
			G_CanReply=True
		ElseIf Dvbbs.UserID = PostUserid and  Dvbbs.GroupSetting(4)="1" Then
			G_CanReply=True
		ElseIf Dvbbs.master Or Dvbbs.superboardmaster Or Dvbbs.boardmaster Then
			G_CanReply=True
		End If
	End If
End Sub

Sub LoadUserListData()
	If IsArray(G_UserList) Or Not IsArray(G_ItemList) Then Exit Sub
	Dim Rs, i, j, iTempUserID, iUbd, sUserIDList
	iUbd		= UBound(G_ItemList,2)
	sUserIDList	= G_ItemList(10, 0)
	For i=0 To iUbd
		sUserIDList	= sUserIDList & ("," & G_ItemList(10,i))
	Next
	Set Rs		= Dvbbs.Execute("Select " & G_UserItemQuery & " From dv_user Where UserID IN ("& sUserIDList &")")
	If Rs.Eof Or Rs.Bof Then
		'全部是客人
	Else
		G_UserList	= Rs.GetRows(-1)
	End If
	Rs.Close
	Set Rs		= Nothing
	'处理用户资料
	For i=0 To iUbd
		iTempUserID			= G_ItemList(10, i)
		G_ItemList(10, i)	= 0	'初始为游客
		If IsArray(G_UserList)	Then
			For j=UBound(G_UserList,2) To 0 Step -1
				If G_UserList(0, j)=iTempUserID Then
					G_ItemList(10, i)	= j+1	'这里加了1，实际用时要减1
					Exit For
				End If
			Next
		End If
	Next
End Sub

Sub LoadAndParseVote(sTemplate)
	Dim Rs,aVote,s,a1,a2,u1,u2,i,j,t,sLoop
	Dim votetype,votchilds,votchilds_title,votchilds_ep
	Set Rs=Dvbbs.Execute("Select voteid,vote,votenum,votetype,lockvote,voters,timeout,uarticle,uwealth,uep,ucp,upower From Dv_Vote Where VoteID="&G_PollID)
	If Not Rs.eof Then
		aVote=Rs.GetRows(-1)
	Else
		Exit Sub
	End If
	Set Rs=Nothing
	s=sTemplate
	s=Replace(s,"{$showvote.voteid}",aVote(0,0))
	votetype=aVote(3,0)
	s=Replace(s,"{$showvote.lockvote}",aVote(4,0))
	s=Replace(s,"{$showvote.voters}",aVote(5,0))
	s=Replace(s,"{$showvote.timeout}",aVote(6,0))
	s=Replace(s,"{$showvote.uarticle}",aVote(7,0))
	s=Replace(s,"{$showvote.uwealth}",aVote(8,0))
	s=Replace(s,"{$showvote.uep}",aVote(9,0))
	s=Replace(s,"{$showvote.ucp}",aVote(10,0))
	s=Replace(s,"{$showvote.upower}",aVote(11,0))
	If 0=Dvbbs.userid Then
		s=Replace(s,"{$showvote.input}","您还未登录，不能参与。")
	Else
		If datediff("d",aVote(6,0),Now()) > 0 Then
			s=Replace(s,"{$showvote.input}","已过期，不能参与。")
		Else
			If G_LockTopic Then
				s=Replace(s,"{$showvote.input}","相关主题已经锁定，不能参与。")
			Else
				If Not Dvbbs.Execute("Select * From Dv_voteuser Where voteid="& G_PollID &" And userid="& Dvbbs.userid).EOF Then
					s=Replace(s,"{$showvote.input}","您已经投过票了，看结果吧！")
				Else
					s=Replace(s,"{$showvote.input}","<input type=""submit"" name=""VoteSubmit"" value=""投 票"" style=""margin:5px;""/>")
				End If
			End If
		End If
	End If
	a1=Split(Dvbbs.ChkBadWords(aVote(1,0)),"|")
	a2=Split(aVote(2,0),"|")
	u1=UBound(a1)
	t=0
	For i=0 To u1
		sLoop=sLoop&"<tr>"
		votchilds = Split(a1(i),"@@")
		If votetype = 2 and Ubound(votchilds)>=3 Then
			sLoop=sLoop&("<td width=""40%"">"&(i+1)&". "&votchilds(0)&"</td>") 'title
			votchilds_title = Split(votchilds(2),"$$")
			votchilds_ep = Split(votchilds(3),"$$")
			sLoop=sLoop&"<td>"
			For j=0 To UBound(votchilds_title)
				If ""<>votchilds_title(j) Then
					Select Case votchilds(1)
						Case "1"
							sLoop=sLoop&" <input type=""checkbox"" name=""postvote_"&i&""" value="""&j&""" class=""chkbox""/>"&votchilds_title(j)&" "
						Case "2"
							sLoop=sLoop&" 回答：<textarea name=""postvote_"&i&""" style=""width:70%;height:80px;""></textarea> "
						Case Else
							sLoop=sLoop&" <input type=""radio"" name=""postvote_"&i&""" value="""&j&""" style=""border:none;""/>"&votchilds_title(j)&" "
					End Select
				End If
			Next
			sLoop=sLoop&"</td>"
		Else
			sLoop=sLoop&"<td width=""20%"">"
			If 0=aVote(3,0) Then
				sLoop=sLoop&"<input type=""radio"" name=""postvote"" value="""&i&""" style=""border:none;""/>"
			Else
				sLoop=sLoop&"<input type=""checkbox"" name=""postvote_"&i&""" value="""&i&""" style=""border:none;""/>"
			End If
			sLoop=sLoop&(a1(i)&"</td><td width=""80%"" valign=""top"">") 'title
			sLoop=sLoop&"<script language=""javascript"">try{ShowVoteList("&a2(i)&",{$total});}catch(e){}</script>"
			sLoop=sLoop&("</td>")
			t=t+a2(i)
		End If
		sLoop=sLoop&"</tr>"
	Next
	sLoop=Replace(sLoop,"{$total}",t)
	s=Replace(s,"{$showvote.list}",sLoop)
	TPL_Scan s
End Sub

Sub ParsePageNode(sToken)
	Dim a,i,s
	Select Case sToken
		Case "topicid"
			TPL_Echo AnnounceID
		Case "announceid"
			TPL_Echo G_ItemList(0,0)
		Case "topic"
			TPL_Echo G_TopicTitle
		Case "hits"
			TPL_Echo G_Hits
		Case "currentpage"
			TPL_Echo G_CurrentPage
		Case "boardpage"
			TPL_Echo Dvbbs.CheckNumeric(request("page"))
		Case "bbstable"
			TPL_Echo TotalUsetable
		Case "postfacelist"
			a=split(Dvbbs.Forum_PostFace,"|||")
			s=s&"<div style=""float:left;""><input id=""face_1_gif"" type=""radio"" value=""face1.gif"" name=""Expression"" checked=""checked"" style=""border:none"" /><img src=""Skins/default/topicface/face1.gif"" alt="""" /></div>"
			For i=2 To UBound(a)-1
				s=s&"<div style=""float:left;""><input type=""radio"" value=""face"&i&".gif"" name=""Expression"" style=""border:none"" /><img src="""&a(0)&"face"&i&".gif"" alt="""" /></div>"
				If 1=(i-2) Mod 3 Then s=s&"<br style=""clear:both"" />"
			Next
			TPL_Echo s
		Case "modelink"
			If 1=Skin Then
				TPL_Echo "<a href=""dispbbs.asp?BoardID="&Dvbbs.Boardid&"&ID="&AnnounceID&"&skin=0"">平板</a>"
			Else
				TPL_Echo "<a href=""dispbbs.asp?BoardID="&Dvbbs.Boardid&"&replyID="&G_ItemList(0,0)&"&ID="&AnnounceID&"&skin=1"">树形</a>"
			End If
		Case "treemode"
			If 1=Skin Then TPL_Echo "<div id=""postlist"" style=""margin-top : 10px; margin-bottom : 10px; ""> </div><span id=""showpagelist""></span><iframe style=""border:0px;width:0px;height:0px;"" src=""loadtree.asp?boardid="&Dvbbs.Boardid&"&amp;star="&G_CurrentPage&"&amp;replyid="&ReplyID&"&amp;id="&AnnounceID&"&amp;openid="&ReplyID&""" name=""hiddenframe""></iframe>"
		Case "topicadminlist"
			s=""
			If Dvbbs.Boardmaster Then
				If 1=T_GetMoneyType Then
					s=s& "		<a href=""BuyPost.asp?Action=Close&amp;BoardID="&Dvbbs.boardid&"&amp;PostTable="&TotalUseTable&"&amp;ID="&AnnounceID&"&amp;ReplyID="&ReplyID&""" title=""结帖管理"">结帖管理</a><br />"
				End If
				s=s& "	<a href=""admin_postings.asp?action=专题管理&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""专题管理"">专题管理</a><br />"
				If 1=G_LockTopic Then
					s=s& "		<a href=""admin_postings.asp?action=解锁&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""将本主题解开锁定"">解除锁定</a><br />"
				Else
					s=s& "		<a href=""admin_postings.asp?action=锁定&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""锁定本主题"">锁定帖子</a><br />"
				End If
				s=s& "	<a href=""admin_postings.asp?action=提升&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""将本主题提升到帖子列表最前面"">提升帖子</a><br />"
				s=s& "	<a href=""admin_postings.asp?action=沉底&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""将本主题放到帖子列表较靠后位置"">沉底帖子</a><br />"
				s=s& "	<a href=""admin_postings.asp?action=跟帖管理&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""批量删除本主题的跟帖"">跟帖管理</a><br />"
				s=s& "	<a href=""admin_postings.asp?action=删除主题&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""注意：本操作将删除本主题所有帖子，不能恢复"">删除帖子</a><br />"
				s=s& "	<a href=""admin_postings.asp?action=移动&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&"&amp;replyID="&ReplyID&""" title=""移动主题"">移动帖子</a><br />"
				s=s& "	<a href=""admin_postings.asp?action=设置固顶&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""将本主题设置固顶"">设置固顶</a><br />"
				TPL_Echo "<div class=""m_li_top"" style=""display:inline;"" onmouseover=""showmenu1('Menu_0',0);""><a href=""#"">主题管理</a>"
				TPL_Echo "	<div class=""submenu submunu_popup"" id=""Menu_0"" onmouseout=""hidemenu1()"">"
				If ""<>s Then TPL_Echo s
				TPL_Echo "	</div></div>"
			ElseIf IsSelfPost() Then
				If 1=T_GetMoneyType Then
					s=s& "		<a href=""BuyPost.asp?Action=Close&amp;BoardID="&Dvbbs.boardid&"&amp;PostTable="&TotalUseTable&"&amp;ID="&AnnounceID&"&amp;ReplyID="&ReplyID&""" title=""结帖管理"">结帖管理</a><br />"
				End If
				If "1"=Dvbbs.GroupSetting(13) Then
					If 1=G_LockTopic Then
						s=s& "		<a href=""admin_postings.asp?action=解锁&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""将本主题解开锁定"">解除锁定</a><br />"
					Else
						s=s& "		<a href=""admin_postings.asp?action=锁定&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""锁定本主题"">锁定帖子</a><br />"
					End If
				End If
				If "1"=Dvbbs.GroupSetting(11) Then
					s=s& "<a href=""admin_postings.asp?action=跟帖管理&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""批量删除本主题的跟帖"">跟帖管理</a><br />"
					s=s& "<a href=""admin_postings.asp?action=删除主题&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""注意：本操作将删除本主题所有帖子，不能恢复"">删除帖子</a><br />"
				End If
				If "1"=Dvbbs.GroupSetting(12) Then
					s=s& "<a href=""admin_postings.asp?action=移动&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&"&amp;replyID="&ReplyID&""" title=""移动主题"">移动帖子</a><br />"
				End If
				If ""<>s Then
					TPL_Echo "<div class=""m_li_top"" style=""display:inline;"" onmouseover=""showmenu1('Menu_0',0);""><a href=""jascript:;"">主题管理</a><div class=""submenu submunu_popup"" id=""Menu_0"" onmouseout=""hidemenu1()"">"
					TPL_Echo s
					TPL_Echo "</div></div>"
				End If
			End If
		Case Else
	End Select
End Sub

Sub ParseUserNode(sToken)
	Dim i, p, s, s2, bShowAll
	i		= G_ItemList(10, G_Floor)
	bShowAll	= False
	If i>0 Then bShowAll = 2<>G_ItemList(8, G_Floor) Or Dvbbs.BoardMaster Or Dvbbs.UserID=G_UserList(0, G_ItemList(10, G_Floor)-1)
	'			  非客人帖 而且 (非匿名帖			或者	 是管理员	或者	是自己)
	If bShowAll Then
		Select Case sToken
			Case "userid"		TPL_Echo G_UserList(0, i-1)
			Case "username"		TPL_Echo UserName
			Case "richname"
				s	= G_UserList(26, i-1)
				Select Case s
					'Case 0	s2= Dvbbs.mainsetting(9)
					Case 1	s2= Dvbbs.mainsetting(9)
					Case 2	s2= Dvbbs.mainsetting(7)
					Case 3	s2= Dvbbs.mainsetting(7)
					Case 4	s2= Dvbbs.mainsetting(5)
					Case 5	s2= "gray"
					Case 6	s2= "gray"
					Case 7	s2= Dvbbs.mainsetting(5)
					Case 8	s2= Dvbbs.mainsetting(11)
					Case Else s2= Dvbbs.mainsetting(5)
				End Select
				s=Split(Split(Application(Dvbbs.CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='"& G_UserList(26, i-1) &"']/@groupsetting").text,",")(58),"§")
				TPL_Echo "<span style=""width:105px;filter:glow(color='"&s2&"',strength='2');"">"&(s(0)&replace(UserName,chr(255),"")&s(1))
				If 2=G_ItemList(8, G_Floor) Then TPL_Echo "&nbsp;&nbsp;[已匿名]"
				TPL_Echo "</span>"
			Case "useremail"	TPL_Echo G_UserList(2, i-1)
			Case "userpost"		TPL_Echo G_UserList(3, i-1)
			Case "usertopic"	TPL_Echo G_UserList(4, i-1)
			Case "usersign"
				s	= G_UserList(5, i-1)
				p	= G_UserList(26, i-1)
				If ""<>s And 1=G_ItemList(8, G_Floor) And Dvbbs.forum_setting(42)="1" Then
					Set s2	= Application(Dvbbs.CacheName &"_groupsetting")
					If Not s2 Is Nothing Then
						If Application(Dvbbs.CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='"& p &"']/@groupsetting") Is Nothing Then p	= 7
						If Split(Application(Dvbbs.CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='"& p &"']/@groupsetting").text,",")(55)	Then TPL_Echo	"<img src=""images/sigline.gif"" alt="""" /><br />" & Dvbbs.ChkBadWords(Dv_ubb.Dv_SignUbbCode(s, p))
					End If
					Set s2	= Nothing
				End If
			Case "usersex"
				s2=G_UserList(11, i-1)
				If "2"=G_UserList(22, i-1) Then
					If IsDate(s2) Then
						If DateDiff("s",s2,Now())>(cCur(dvbbs.Forum_Setting(8))*60) Then
							G_UserList(22, i-1)="1"
						End If
					Else
							G_UserList(22, i-1)="1"
					End If
				Else
					G_UserList(22, i-1)="1"
				End If
				s	= G_UserList(6, i-1)
				If "1"=G_UserList(22, i-1) Then
					Select Case s
						Case "1"
							TPL_Echo	"<img src=""Skins/Default/ofMale.gif"" alt=""帅哥哟，离线，有人找我吗？"" />"
						Case Else
							TPL_Echo	"<img src=""Skins/Default/ofFeMale.gif"" alt=""美女呀，离线，留言给我吧！"" />"
					End Select
				Else
					Select Case s
						Case "1"
							TPL_Echo	"<img src=""Skins/Default/Male.gif"" alt=""帅哥，在线噢！"" />"
						Case Else
							TPL_Echo	"<img src=""Skins/Default/FeMale.gif"" alt=""美女呀，在线，快来找我吧！"" />"
					End Select
				End If
			Case "userface"
				s2	= s2 &	"<img"
				s2	= s2 &	(" width="""&G_UserList(8, i-1)&"""")
				s2	= s2 &	(" height="""&G_UserList(9, i-1)&"""")
				s2	= s2 &	" src="""
				s	= Dv_FilterJS(G_UserList(7, i-1))
				p	= InStr(s, "|")
				If p>0 Then
					s2	= s2 &	Mid(s, p+1)
					s	= Left(s, p-1)
				Else
					s2	= s2 &	s
					s	= "0"
				End If
				s2	= s2 &	""" alt="""" />"
				If "0"<>s Then	s2	= s2 &	("<br/><div><a href=""javascript:DispMagicEmot('"&s&"',350,500)"">查看魔法头像</a></div>")
				TPL_Echo	s2
			Case "joindate"		TPL_Echo G_UserList(10, i-1)
			Case "lastlogin"	TPL_Echo G_UserList(11, i-1)
			Case "userlogins"	TPL_Echo G_UserList(12, i-1)
			Case "lockuser"		TPL_Echo G_UserList(13, i-1)
			Case "userclass"	TPL_Echo G_UserList(14, i-1)
			Case "userwealth"	TPL_Echo G_UserList(15, i-1)
			Case "userep"		TPL_Echo G_UserList(16, i-1)
			Case "usercp"		TPL_Echo G_UserList(17, i-1)
			Case "userpower"
				s	= G_UserList(18, i-1)
				If "0"<>s Then
					TPL_Echo "<b><font color=""red"">" & s & "</font></b>"
				Else
					TPL_Echo s
				End If
			Case "userdel"		TPL_Echo G_UserList(19, i-1)
			Case "userisbest"	TPL_Echo G_UserList(20, i-1)
			Case "usertitle"	TPL_Echo G_UserList(21, i-1)
			Case "usermoney"	TPL_Echo G_UserList(23, i-1)
			Case "userticket"	TPL_Echo G_UserList(24, i-1)
			Case "titlepic"
				s2	= s2 &	"<img src=""skins/Default/star/"
				s2	= s2 &	G_UserList(25, i-1)
				s2	= s2 &	""" alt="""" />"
				TPL_Echo	s2
			Case "qq"
				If Not IsArray(G_UserList(27, i-1)) Then	G_UserList(27, i-1)=Split(G_UserList(27, i-1), "|||")
				TPL_Echo	G_UserList(27, i-1)(1)
			Case "link_qq"
				If Not IsArray(G_UserList(27, i-1)) Then	G_UserList(27, i-1)=Split(G_UserList(27, i-1), "|||")
				If ""<>G_UserList(27, i-1)(1) Then TPL_Echo "	 | <a href=""tencent://message/?uin="&G_UserList(27, i-1)(1)&""" title=""点击发送QQ消息给"&UserName&""">QQ</a>"
			Case "email"
				TPL_Echo	G_UserList(28, i-1)
			Case "homepage"
				If Not IsArray(G_UserList(27, i-1)) Then	G_UserList(27, i-1)=Split(G_UserList(27, i-1), "|||")
				TPL_Echo	G_UserList(27, i-1)(0)
			Case "uc"
				If Not IsArray(G_UserList(27, i-1)) Then	G_UserList(27, i-1)=Split(G_UserList(27, i-1), "|||")
				TPL_Echo	G_UserList(27, i-1)(6)
			
			Rem 以下为荣誉勋章标签，fish 2010-2-19
			Case "medal"
				if Cint(dvbbs.Forum_Setting(104))=1 then
					Dim j
					If G_UserList(29, i-1) <> "" Then
						s = Split(G_UserList(29, i-1),",")
						For i = 0 To Ubound(s)
							For j = 0 To Ubound(G_MedalData,2)
								If Clng(s(i)) = G_MedalData(0,j) Then
									s2 = s2 & "<img src='dv_plus/medal/images/"&G_MedalData(2,j)&"' alt='"&G_MedalData(1,j)&"' /> "
									Exit For
								End If
							Next
						Next
					End If
					TPL_Echo	s2
				End If
		End Select
	Else	'游客 或 匿名帖
		Select Case sToken
			Case "username"
				s	= Split(G_ItemList(6, G_Floor), ".")
				If i>0 Then
					s2	= s2 &	"匿名"
				Else
					s2	= s2 &	"客人"
				End If
				s2	= s2 &	("(" & s(0) & "." & s(1) & ".*.*)")
				TPL_Echo	s2
			Case "richname"
				s	= Split(G_ItemList(6, G_Floor), ".")
				If i>0 Then
					s2	= s2 &	"匿名"
				Else
					s2	= s2 &	"客人"
				End If
				s2	= s2 &	("(" & s(0) & "." & s(1) & ".*.*)")
				s=Split(Split(Application(Dvbbs.CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='7']/@groupsetting").text,",")(58),"§")
				TPL_Echo "<span style=""width:130px;filter:glow(color='gray',strength='2');"">"&(s(0)&s2&s(1)&"</span>")
			Case "userface"
				If i>0 Then
					TPL_Echo	"<img src=""Skins/Default/anyon.gif"" width=""111"" height=""111"" border=""0"" />"
				Else
					TPL_Echo	"<img src=""Skins/Default/guest.gif"" width=""111"" height=""111"" border=""0"" />"
				End If
			Case Else
		End Select
	End If
End Sub

Function IsSelfPost()
	IsSelfPost=False
	If G_ItemList(10, G_Floor)>0 Then
		If G_UserList(0, G_ItemList(10, G_Floor)-1)=Dvbbs.UserID Then
			IsSelfPost=True
		End If
	End If
End Function

Function GetPostUserID()
	If G_ItemList(10, G_Floor)>0 Then
		GetPostUserID=G_UserList(0, G_ItemList(10, G_Floor)-1)
	Else
		GetPostUserID=0
	End If
End Function

Sub ParseBBSListNode(sToken)
	Dim i, a, postbuyusers, postbuyinfo, j 
	Select Case sToken
		Case "announceid"
			TPL_Echo G_ItemList(0, G_Floor)
		Case "title"
			TPL_Echo Dvbbs.Replacehtml(Dvbbs.ChkBadWords(G_ItemList(2, G_Floor)))
		Case "tyisbest"
			Rem 精华帖加盖章,小易
			If 1=G_ItemList(9, G_Floor) Then
		    TPL_Echo "<span class=""isbestcss""></span>"
			end If
		Case "url"
			If 0=G_Floor Then TPL_Echo "<INPUT TYPE=""hidden"" id=""url"" name=""url"" />"
		Case "body"
			i	= G_ItemList(10, G_Floor)
			If i>0 Then
				i	= G_UserList(26, i-1)
			Else
				i	= 7	'客人
			End If
			If 0=G_Floor And 1=G_CurrentPage And CLng(G_ItemList(20, G_Floor))=0 Then	'顶楼需要判断购买帖
				If G_LockTopic Then TPL_Echo "<div class=""limitinfo"">贴子已被锁定</div><br/>"
				If 3=T_GetMoneyType Then
					TPL_Echo "<div class=""limitinfo""><font color=""gray"">以下内容需要支付 <b><font color=""red"">"&G_ItemList(17, G_Floor)&"</font></b> 个金币方可查看，"
					If IsSelfPost() Then
						TPL_Echo "这是您发的帖子"
					ElseIf TrueMaster Then
						TPL_Echo "由于您是工作人员，你可以看到内容。"
					Else
						If Trim(PostBuyUser)="" Then PostBuyUser="0@@@-1@@@0@@@|||$PayMoney|||"
						postbuyusers=split(PostBuyUser,"|||")
						postbuyinfo=postbuyusers(0)
						postbuyinfo=Split(postbuyinfo,"@@@") 'Rem postbuyinfo(0) 收入money   postbuyinfo(1) 购买限制maxbuy   postbuyinfo(2) vip是否需要购买notvipbuy   postbuyinfo(3) 允许购买用户列表buyuser
						If UBound(postbuyinfo)<=2 Then Exit Sub
						a=False
						For j=2 to UBound(postbuyusers)
							If postbuyusers(j)<>"" And postbuyusers(j)=Dvbbs.MemberName Then
								a=True
							End If
						Next
						If a Then
							TPL_Echo "您已经购买。"
						ElseIf Dvbbs.VipGroupUser And "1"=postbuyinfo(2) Then
							TPL_Echo "由于您是vip用户，并且因为设置了vip用户可免购买查看，您可以直接查看。"
						ElseIf (""=postbuyinfo(3) Or InStr(","&postbuyinfo(3)&",", ","&Dvbbs.MemberName&",")>0) And Dvbbs.userid>0 Then
							TPL_Echo "您需要购买方可看到内容。</font><br /><input type=""button"" value=""我要查看内容，决定购买"" onclick=""location.href='BuyPost.asp?action=buy&amp;boardid="&Dvbbs.BoardID&"&amp;id="&AnnounceID&"&amp;ReplyID="&G_ItemList(0, G_Floor)&"&amp;PostTable="&TotalUsetable&"'""/>"
							TPL_Echo "<br/></div>"
							Exit Sub
						Else
							TPL_Echo "您需要购买方可看到内容。"
							If Dvbbs.userid>0 Then
								TPL_Echo "楼主设置了您不可以购买。"
							Else
								TPL_Echo "您还未登录，不能购买。"
							End If
							TPL_Echo "</font><br/></div>"
							Exit Sub
						End If
					End If
					TPL_Echo "</font><br/></div>"
				End If
			End If
			Ubblists=G_ItemList(15, G_Floor)
			'增加允许管理员发iframe功能 by 牛头
			if Dvbbs.userID<>0 then
			If dvbbs.checknumeric(G_ItemList(10, G_Floor))>0 Then Dv_ubb.ismanager1= G_UserList(26, G_ItemList(10, G_Floor)-1)
			End If 
			If InStr(Ubblists,",39,") > 0  Then
				
				TPL_Echo	Dvbbs.ChkBadWords( Dv_ubb.Dv_UbbCode(G_ItemList(4, G_Floor),i,1,0) )
			Else
				TPL_Echo	Dvbbs.ChkBadWords( Dv_ubb.Dv_UbbCode(G_ItemList(4, G_Floor),i,1,1) )
			End If
		Case "bodystyle"
			TPL_Echo ("font-size:"&Dvbbs.Board_setting(28)&"pt;text-indent:"&Dvbbs.Board_setting(69)&"px;")
		Case "floor"
			i	= G_Moved+G_Floor
			TPL_Echo i+1
		Case "dateandtime" TPL_Echo G_ItemList(3, G_Floor)
		Case "authorid"
            If authorid = 0 Then
                TPL_Echo "[<a href=""dispbbs.asp?boardid="&Dvbbs.BoardID&"&Id="&AnnounceID&"&authorid="&GetPostUserID()&""">只看该作者</a>]"
            Else
                TPL_Echo "[<a href=""dispbbs.asp?boardid="&Dvbbs.BoardID&"&Id="&AnnounceID&""">显示全部帖子</a>]"
            End If

		Case "showpage"
			TPL_ShowPage	G_CurrentPage, G_Childs, Dvbbs.CheckNumeric(Dvbbs.Board_Setting(27)), 10, "dispbbs.asp?boardid="&Dvbbs.BoardID&"&id="&AnnounceID&"&authorid="&authorid&"&page="&Dvbbs.CheckNumeric(request("page"))&"&star="
		Case "bestinfo"
			If 1=G_ItemList(9, G_Floor) Then
				TPL_Echo "<div class=""info""><img src="""&Dvbbs.Forum_PicUrl&"jing.gif"" border=""0"" title=""本帖被加为精华"" align=""absmiddle""/>[本帖被加为精华]</div>"
			End If
		Case "bestpic"
			If 1=G_ItemList(9, G_Floor) Then
				TPL_Echo "<img src=""images/best.gif"" border=""0"" style=""position:absolute;z-index:1;"" title=""精华帖子认证"" align=""absmiddle"" />"
			End If
		Case "appraise"
			If IsNull(G_ItemList(12, G_Floor)) Then Exit Sub
			SplitIsAgree
			a = G_ItemList(12, G_Floor)
			If a(1)>0 Then
				TPL_Echo	"<div class=""info"">版主评定：<img src="""&Dvbbs.Forum_PicUrl&"agree.gif"" border=""0"" alt=""好评，获得"&a(1)&"个金币奖励"" align=""absmiddle""/>好评，获得<font color=""red"">"&a(1)&"</font>个金币奖励</div>"
				If ""<>a(3) Then TPL_Echo "(" & a(3) & ")"
			ElseIf a(0)>0 Then
				TPL_Echo	"<div class=""info"">版主评定：<img src="""&Dvbbs.Forum_PicUrl&"disagree.gif"" border=""0"" alt=""差评，扣除"&a(0)&"个金币"" align=""absmiddle""/>差评，扣除<font color=""red"">"&a(0)&"</font>个金币</div>"
				If ""<>a(2) Then TPL_Echo "(" & a(2) & ")"
			End If
		
		Case "moneytype" 'modifty by reoaiq at 090922
			i = T_GetMoneyType
			If 0=i Then Exit Sub
			If 0=G_Floor And G_CurrentPage=1 Then
				Select Case i
					Case 1
					If Dvbbs.BoardMaster Then 
					TPL_Echo "<div class=""info"">悬赏金币帖，要悬赏 <font color=""red"">" & G_ItemList(17, G_Floor) & "</font> 个金币</div>"
					TPL_Echo "<div class=""info""><a href=""BuyPost.asp?Action=Cancel&PostTable="&TotalUsetable&"&BoardId="&Dvbbs.BoardID&"&ID="&AnnounceIor=""red"">"&a(1)&"</font>涓