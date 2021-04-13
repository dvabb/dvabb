<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/ubblist.asp"-->
<!--#include file="inc/Email_Cls.asp"-->
<!--#include file="inc/dv_ubbcode.asp"-->
<!--#include file="inc/code_encrypt.asp"-->
<%
Dim parameter

'以下定义的变量在Dv_ubbcode.asp页面会用到
Dim replyid_a,AnnounceID,TotalUseTable,AnnounceID_a,RootID_a
Dim T_GetMoneyType,EmotPath
Dim UserName:UserName=Dvbbs.Membername
Dim IsThisBoardMaster '确定当前用户是否本版版主，防止下面的操作影响到 Dvbbs.BoardMaster导致出错
IsThisBoardMaster = Dvbbs.BoardMaster
Dim PostStyle,ajaxPost
PostStyle = Request.Form("poststyle")
ajaxPost = CInt(Request.Form("ajaxPost"))

If ajaxPost=1 Then
	ajaxPost = True
Else
	ajaxPost = False
End If

If Dvbbs.BoardID < 1 Then
	Response.Write "参数错误"
	Dvbbs.PageEnd()
	Response.End
End If
Dim MyPost
Dim postbuyuser,bgcolor,abgcolor,FormID
Dvbbs.Loadtemplates("post")
Set MyPost = New Dvbbs_Post
Dvbbs.Stats = MyPost.ActionName
EmotPath=Split(Dvbbs.Forum_emot,"|||")(0)		'em心情路径
If Not ajaxPost Then
	If PostStyle = "1" Then
		Dvbbs.Head()
		Dvbbs.ErrType=1
	Else
		Dvbbs.Nav()
		Dvbbs.Head_var 1,Application(Dvbbs.CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"&Dvbbs.BoardID&"']/@depth").text,"",""
	End If
End if
MyPost.Save_CheckData
Set MyPost = Nothing
If Not ajaxPost Then Dvbbs.Footer
Dvbbs.PageEnd()
Class Dvbbs_Post
	Public Action,ActionName,Star,Page,IsAudit,ToAction,TopicMode,Reuser
	Private ReplyID,ParentID,RootID,Topic,Content,char_changed,signflag,mailflag,iLayer,iOrders
	Private TopTopic,IsTop,LastPost,LastPost_1,UpLoadPic_n,ihaveupfile,smsuserlist,upfileinfo
	Private UserName,UserPassWord,UserPost,GroupID,UserClass,DateAndTime,DateTimeStr,Expression,MyLastPostTime,LastPostTimes
	Private LockTopic,MyLockTopic,MyIsTop,MyIsTopAll,MyTopicMode,Child
	Private CanLockTopic,CanTopTopic,CanTopTopic_a,CanEditPost,Rs,SQL,i,IsAuditcheck
	Private vote,votetype,votenum,votetimeout,voteid,isvote,ErrCodes
	Private GetPostType,ToMoney,UseTools,ToolsBuyUser,GetMoneyType,Tools_UseTools,Tools_LastPostTime,ToolsInfo,ToolsSetting
	Private tMagicFace,iMagicFace,tMagicMoney,tMagicTicket,FoundUseMagic,isAlipayTopic
	Private Sub Class_Initialize()

		ErrCodes = ""
		Dvbbs.ErrCodes=""
		'管理员及该版版主允许在锁定论坛发帖
		If Dvbbs.Board_Setting(0)="1" And Not (Dvbbs.Master or Dvbbs.Boardmaster) Then
			If ajaxpost Then
				Call showAjaxMsg(0,"锁定论坛，只对管理员及该版版主开放！","","")
			Else
				parameter="showerr.asp?ShowErrType="&Dvbbs.ErrType&"&action=lock&boardid="&dvbbs.boardID&""
				Response.redirect parameter
			End If
		End If

		Rem 针对不同用户组设置只读权限，动网.小易 2010-2-3
		If Dvbbs.TyReadOnly Then
			If ajaxpost Then
				Call showAjaxMsg(0,"论坛只读，您所在用户组做了限制！","","")
			Else
				parameter="showerr.asp?ShowErrType="&Dvbbs.ErrType&"&action=readonly&boardid="&dvbbs.boardID&""
				Response.redirect parameter
			End If
		End If

		If Dvbbs.IsReadonly()  And Not Dvbbs.Master Then
			If ajaxpost Then
				Call showAjaxMsg(0,"论坛只读，只对管理员开放！","","")
			Else
				parameter="showerr.asp?ShowErrType="&Dvbbs.ErrType&"&action=readonly&boardid="&dvbbs.boardID&""
				Response.redirect parameter
			End If
		End If
		Action = Request("Action")
		TotalUseTable = Dvbbs.NowUseBBS
		Select Case Action
		Case "snew"
			Action = 5
			ActionName = template.Strings(1)
			If Dvbbs.GroupSetting(3)="0" Then Dvbbs.AddErrCode(70)
		Case "sre"
			Action = 6
			ActionName = template.Strings(3)
			'If Dvbbs.GroupSetting(5)="0" then Dvbbs.AddErrCode(71)
		Case "svote"
			Action = 7
			ActionName = template.Strings(5)
			If Dvbbs.GroupSetting(8)="0" then Dvbbs.AddErrCode(56)
		Case "sedit"
			Action = 8
			ActionName = template.Strings(7)
		Case Else
			Action = 1
			ActionName = template.Strings(0)
		End Select
		Star = Request("star")
		If Star = "" Or Not IsNumeric(Star) Then Star = 1
		Star = Clng(Star)
		Page = Request("page")
		If Page = "" Or Not IsNumeric(Page) Then Page = 1
		Page = Clng(Page)
		'IsAudit = Cint(Dvbbs.Board_Setting(3))
		IsAudit=0
		Reuser = False'此变量标识是否更名发贴
		FoundUseMagic = False
	End Sub

	Public Function CheckFormID(id)
		CheckFormID=false
		Dim i,Str
		For i=1 to Len(id)
			Str=Str & Asc(Mid(id,i,1))-97
		Next
		If Session.SessionID=Str Then
			CheckFormID=True
		End If
	End Function

	Rem 改写ShowErr过程 ajax模式下 不跳转 By Dv.唧唧  于 火星
	Sub ShowErr()
		If ajaxpost And Dvbbs.ErrCodes<>"" Then
			Call showAjaxMsg(0,getErrCodeMsg(),"","")
		Else
			Dvbbs.ShowErr()
		End If
	End Sub

	'通用判断
	Public Function Chk_Post()
		'FormID=Request("Dvbbs")
		'If FormID="" Then FormID=Request.Cookies("Dvbbs"):Response.Cookies("Dvbbs")=""
		'If Not CheckFormID(FormID) Then Response.redirect "showerr.asp?ShowErrType="&Dvbbs.ErrType&"&ErrCodes=您提交的参数错误&action=OtherErr"
		If Dvbbs.Board_Setting(43)="1" Then Dvbbs.AddErrCode(72)
		If Dvbbs.Board_Setting(1)="1" and Dvbbs.GroupSetting(37)="0" Then Dvbbs.AddErrCode(26)
		If Dvbbs.UserID>0 Then
			If Clng(Dvbbs.GroupSetting(52))>0 And DateDiff("s",Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@joindate").text,Now)<Clng(Dvbbs.GroupSetting(52))*60 Then
				If ajaxPost Then
					Call showAjaxMsg(0,Replace(template.Strings(21),"{$timelimited}",Dvbbs.GroupSetting(52)),"","")
				Else
					parameter="showerr.asp?ShowErrType="&Dvbbs.ErrType&"&ErrCodes=<li>"&Replace(template.Strings(21),"{$timelimited}",Dvbbs.GroupSetting(52))&"&action=OtherErr"
					Response.redirect parameter
				End if

			End If
			If Dvbbs.GroupSetting(62)<>"0" And Not Action = 8 Then
				If Clng(Dvbbs.GroupSetting(62))<=Clng(Dvbbs.UserToday(0)) Then
					If ajaxPost Then
						Call showAjaxMsg(0,Replace(template.Strings(27),"{$topiclimited}",Dvbbs.GroupSetting(62)),"","")
					Else
						parameter="showerr.asp?ShowErrType="&Dvbbs.ErrType&"&ErrCodes=<li>"&Replace(template.Strings(27),"{$topiclimited}",Dvbbs.GroupSetting(62))&"&action=OtherErr"
						Response.redirect parameter
					End if
				End If
			End If
		End If
		If Dvbbs.GroupSetting(3)="0" And (Action = 5 Or Action = 7) Then
			If ajaxPost Then
				Call showAjaxMsg(0,template.Strings(28),"","")
			Else
				Response.redirect "showerr.asp?ShowErrType="&Dvbbs.ErrType&"&ErrCodes=<li>"&template.Strings(28)&"&action=OtherErr"
			End if
		End if
		'If Dvbbs.GroupSetting(5)="0" And (Action = 6) Then Response.redirect "showerr.asp?ShowErrType="&Dvbbs.ErrType&"&ErrCodes=<li>"&template.Strings(29)&"&action=OtherErr"
	End Function

	'返回判断和参数
	Public Function Get_M_Request()
		AnnounceID = Request("ID")
		If AnnounceID = "" Or Not IsNumeric(AnnounceID) Then Dvbbs.AddErrCode(30)
		ShowErr()
		AnnounceID = cCur(AnnounceID)
	End Function

	'检查提交来源
	Public Sub CheckfromScript()
		If Not Dvbbs.ChkPost() Then Dvbbs.AddErrCode(16):ShowErr()
 		'If CStr(Request.Cookies("Dvbbs"))=CStr(Dvbbs.Boardid) Then Dvbbs.AddErrCode(30):ShowErr() '非法的贴子参数。
 		If (Not ChkUserLogin) And (Action = 5 Or Action = 6 Or Action = 7) And Dvbbs.UserID>0 Then Dvbbs.AddErrCode(12):ShowErr()
	End Sub

	'判断发贴时间间隔
	Private Sub CheckpostTime()
		If Dvbbs.Board_Setting(30)="1"  Then
			If IsDate(Session(Dvbbs.CacheName & "posttime"))  Then
				If DateDiff("s",Session(Dvbbs.CacheName & "posttime"),Now())<CLng(Dvbbs.Board_Setting(31)) Then
					If ajaxPost Then
						Call showAjaxMsg(0,Replace(template.Strings(33),"{$PostTimes}",Dvbbs.Board_Setting(31)),"","")
					Else
						parameter="showerr.asp?ShowErrType="&Dvbbs.ErrType&"&ErrCodes=<Br>"+"<li>"&Replace(template.Strings(33),"{$PostTimes}",Dvbbs.Board_Setting(31))&"&action=OtherErr"
						Response.redirect parameter
					End if
				End If
			End If
			Session(Dvbbs.CacheName & "posttime")=Now()
		End If
	End Sub
	'检查用户身份
	Public Function ChkUserLogin()
 		ChkUserLogin=False
 		'取得发贴用户名和密码
		If Dvbbs.UserID=0 Then
			UserName="客人"
		Else
			If ajaxPost Then
				UserName=Dvbbs.Checkstr(unescape(Request.Form("username")))
			Else
				UserName=Dvbbs.Checkstr(Request.Form("username"))
			End If
		End If
		'校验用户名和密码是否合法
		If UserName="" Or Dvbbs.strLength(userName)>Cint(Dvbbs.Forum_setting(41)) Or Dvbbs.strLength(userName) < Cint(Dvbbs.Forum_setting(40)) Then Dvbbs.AddErrCode(17)
		If Not IstrueName(UserName) Then Dvbbs.AddErrCode(18)
		ShowErr()
		If Action = 8 Then
			'编辑贴子，检查用户身份
			UserPassWord=Dvbbs.checkStr(Trim(Request.Cookies(Dvbbs.Forum_sn)("password")))
			SQL = "Select JoinDate,UserID,UserPost,UserGroupID,userclass,lockuser,TruePassWord From [Dv_User] Where UserID="&Dvbbs.UserID
		Else
			'检查用户是否当前用户
			If UserName<>Dvbbs.MemberName Then
				Reuser=True
				UserPassWord=Dvbbs.Checkstr(Trim(Request.Form("passwd")))
				UserPassWord=md5(UserPassWord,16)
				SQL = "Select JoinDate,UserID,UserPost,UserGroupID,userclass,lockuser,userpassword From [Dv_User] Where UserName='"&UserName&"' "
			Else
				UserPassWord=Dvbbs.checkStr(Trim(Request.Cookies(Dvbbs.Forum_sn)("password")))
				SQL = "Select JoinDate,UserID,UserPost,UserGroupID,userclass,lockuser,TruePassWord From [Dv_User] Where UserID="&Dvbbs.UserID
			End If
		End If
		If Len(UserPassWord)<>16 AND Len(UserPassWord)<>32 Then Dvbbs.AddErrCode(18)
 		Set Rs=Dvbbs.Execute(SQL)
 		If Not Rs.EOF Then
			If Not (UserPassWord<>rs(6) Or rs(5)=1 or rs(3)=5) Then
				'不允许使用马甲
				If Dvbbs.UserID<>Rs(1) Then
					ChkUserLogin = False
				Else
					Dvbbs.UserID=Rs(1)
 					UserPost=Rs(2)
 					GroupID=Rs(3)
 					userclass=Rs(4)
 					ChkUserLogin=True
				End If
				Response.cookies("upNum")=0
 			Else
  				Dvbbs.LetGuestSession()
			End If
 		End If
 		Set Rs = Nothing
 	End Function

	'判断发表类型及权限 GetPostType 0=赠送金币贴(求回复答案),1=获赠金币贴,2=金币购买贴
	Private Sub Chk_PostType()
		Dim ToolsID
		ToolsID = Trim(Request.Form("ToolsID"))
		GetPostType = Trim(Request.Form("GetPostType"))
		ToMoney = Trim(Request.Form("ToMoney"))
		If ToMoney="" or Not Isnumeric(ToMoney) Then ToMoney = 0
		If ToolsID="" or Not Isnumeric(ToolsID) Then
			ToolsID = ""
		Else
			ToolsID = Cint(ToolsID)
		End If
		ToMoney = cCur(ToMoney)
		UseTools = ""
		ToolsBuyUser = ""
		GetMoneyType = 0
		If Dvbbs.GroupSetting(59)<>1 Then Exit Sub
		If GetPostType<>"" and (Action = 5 or Action = 7) Then
			Select Case GetPostType
			Case "0"
				If ToMoney = 0 or ToMoney > CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text)  Or ToMoney < 0 Then
					If ajaxpost Then
						Call showAjaxMsg(0,"您设置的金币值为空或者多于您拥有的金币数量。","","")
					Else
						Response.redirect "showerr.asp?ShowErrType="&Dvbbs.ErrType&"&ErrCodes=<li>您设置的金币值为空或者多于您拥有的金币数量。&action=OtherErr"
					End If
				End if
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text  = CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text) -ToMoney
				'UseTools = "-1111"
				ToolsBuyUser = "0|||$SendMoney"
				GetMoneyType = 1
			Case "1"
				ToolsBuyUser = "0|||$GetMoney"
				GetMoneyType = 2
				'UseTools = ToolsInfo(4)
			Case "2"
				If ToMoney = 0 Or ToMoney < 0 Then
					If ajaxpost Then
						Call showAjaxMsg(0,"请正确填写购买帖的金币数量。","","")
					Else
						Response.redirect "showerr.asp?ShowErrType="&Dvbbs.ErrType&"&ErrCodes=<li>请正确填写购买帖的金币数量。&action=OtherErr"
					End If
				End if
				Dim Buy_Orders,Buy_VIPType,Buy_UserList
				Buy_Orders = Request.FORM("Buy_Orders")
				Buy_VIPType = Request.FORM("Buy_VIPType")
				Buy_UserList = Request.FORM("Buy_UserList")
				If Buy_Orders<>"" and IsNumeric(Buy_Orders) Then
					Buy_Orders = cCur(Buy_Orders)
				Else
					Buy_Orders = -1
				End If
				If Not IsNumeric(Buy_VIPType) Then Buy_VIPType = 0
				If Buy_UserList<>"" Then Buy_UserList = Replace(Replace(Replace(Buy_UserList,"|||",""),"@@@",""),"$PayMoney","")
				ToolsBuyUser = "0@@@"&Buy_Orders&"@@@"&Buy_VIPType&"@@@"&Buy_UserList&"|||$PayMoney|||"
				GetMoneyType = 3
				'UseTools = ToolsInfo(4)
			End Select
		End If
		'回复获赠金币帖判断
		If Action = 6 and GetPostType = "1" Then
			If ToMoney = 0 or ToMoney > CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text)  Or ToMoney < 0 Then
				If ajaxpost Then
					Call showAjaxMsg(0,"您设置的金币值为空或者多于您拥有的金币数量。","","")
				Else
					Response.redirect "showerr.asp?ShowErrType="&Dvbbs.ErrType&"&ErrCodes=<li>您设置的金币值为空或者多于您拥有的金币数量。&action=OtherErr"
				End If
			End If
		End If
	End Sub


	Rem ------------------------
	Rem 保存部分函数开始
	Rem ------------------------
	'检查数据,提取数据，获得贴子数据表名等。

	Public Sub Save_CheckData()
		Chk_Post()
		CheckfromScript()
		'把提交的数据保存到session
		Content = CheckAlipay()
		isAlipayTopic = 2
		If Content = "" Then
			Content = Dvbbs.Checkstr(unescape(Request.Form("body")))
			isAlipayTopic = 0
		End If
		'不再把内容保存到session
		'Dvbbs.UserSession.documentElement.selectSingleNode("userinfo").attributes.setNamedItem(Dvbbs.UserSession.createNode(2,"postdata","")).text= Request.Form("body")
        
		'验证码校验
		If Dvbbs.Board_Setting(4) <> "0" Then
			If Not Dvbbs.CodeIsTrue() Then
				If ajaxpost Then
					Call showAjaxMsg(0,"验证码校验失败,请点击验证码进行刷新！","","")
				Else
					Response.redirect "showerr.asp?ShowErrType="&Dvbbs.ErrType&"&ErrCodes=<li>验证码校验失败，2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
				End If
			End If
		End If
        If dvbbs.forum_setting(109)=1 Then
            If Application(Dvbbs.CacheName&"_loadBadlanguage")="" Or IsNull(Application(Dvbbs.CacheName&"_loadBadlanguage")) Then loadBadlanguage
			Dim Badlanguage
			Badlanguage=Split(Application(Dvbbs.CacheName&"_loadBadlanguage"),"|")
            Dim Badlanguage_i,splitstr
            For Badlanguage_i=0 To UBound(Badlanguage)
                splitstr=Replace(Trim(Badlanguage(Badlanguage_i)),Chr(13),"")
				If InStr(Dvbbs.Checkstr(unescape(Content)),splitstr)>0 Then
				    If ajaxpost Then
					    Call showAjaxMsg(0,"统检测您发表的话题内容中含有非法字符<b>"&splitstr&"</b>","","")
					Else
					    Response.redirect "showerr.asp?ErrCodes=<li>系统检测您发表的话题内容中含有非法字符<b>"&splitstr&"</b>禁止提交。&action=OtherErr"
					End If
				End If
                If InStr(Dvbbs.Checkstr(unescape(Request.Form("topic"))),splitstr)>0 Then
				    If ajaxpost Then
					    Call showAjaxMsg(0,"统检测您发表的话题内容中含有非法字符<b>"&splitstr&"</b>","","")
					Else
					    Response.redirect "showerr.asp?ErrCodes=<li>系统检测您发表的话题中含有非法字符<b>"&splitstr&"</b>禁止提交。&action=OtherErr"
					End If
				End If
            Next
		End If
		If InStr(Content,"[/payto]") > 0 And InStr(Content,"[payto]") > 0 And InStr(Content,"(/seller)") > 0 And InStr(Content,"(seller)") > 0 Then isAlipayTopic = 2
		Chk_PostType()

		'魔法表情检查部分
		tMagicFace = Request("tMagicFace")
		If tMagicFace = "" Or Not IsNumeric(tMagicFace) Then tMagicFace = 0
		tMagicFace = Cint(tMagicFace)
		iMagicFace = Request("iMagicFace")
		If iMagicFace = "" Or Not IsNumeric(iMagicFace) Then iMagicFace = 0
		iMagicFace = Clng(iMagicFace)
		Expression = Dvbbs.Checkstr(Request.Form("Expression"))
		If Expression = "" Then
			Expression = "face1.gif"
		Else
			Expression = Replace(Expression,"|","")
		End If
		If tMagicFace = 1 And iMagicFace > 0 And Dvbbs.Forum_Setting(98)="1" Then
			Set Rs = Dvbbs.Plus_Execute("Select tMoney,tTicket,MagicSetting From Dv_Plus_Tools_MagicFace Where MagicFace_s = " & iMagicFace)
			If Rs.Eof And Rs.Bof Then
				Expression = "0|" & Expression
				tMagicMoney = 0
				tMagicTicket = 0
			Else
				If cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text) < Rs(1) And cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text ) < Rs(0) Then
					If ajaxpost Then
						Call showAjaxMsg(0,"您没有足够的金币或点券使用魔法表情！","","")
					Else
						Response.redirect "showerr.asp?ShowErrType="&Dvbbs.ErrType&"&ErrCodes=<li>您没有足够的金币或点券使用魔法表情，2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
					End If
				End if
				Dim iMagicSetting, iMagicStr
				iMagicStr = ""
				iMagicSetting = Split(Rs(2),"|")
				If cCur(iMagicSetting(0)) > cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpost").text ) Then iMagicStr ="您的帖子数"
				If cCur(iMagicSetting(1)) > cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text) Then iMagicStr ="您的金钱数"
				If cCur(iMagicSetting(2)) > cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text) Then iMagicStr ="您的积分数"
				If cCur(iMagicSetting(3)) > cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text ) Then iMagicStr ="您的魅力数"
				If cCur(iMagicSetting(4)) > cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpower").text) Then iMagicStr ="您的威望数"

				If iMagicStr<>"" Then
					If ajaxpost Then
						Call showAjaxMsg(0,iMagicStr&"没有达到使用魔法表情的标准，","","")
					Else
						Response.redirect "showerr.asp?ShowErrType="&Dvbbs.ErrType&"&ErrCodes=<li>"&iMagicStr&"没有达到使用魔法表情的标准，2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
					End If
				End If

				Expression = iMagicFace & "|" & Expression
				tMagicMoney = Rs(0)
				tMagicTicket = Rs(1)
				FoundUseMagic = True
			End If
			Rs.Close
			Set Rs=Nothing
		Else
			Expression = "0|" & Expression
		End If
		Expression = Split(Expression,"|")
		Topic = Dvbbs.Checkstr(Trim(unescape(Request.Form("topic"))))
		signflag = Dvbbs.Checkstr(Trim(Request.Form("signflag")))
		mailflag = Dvbbs.Checkstr(Trim(Request.Form("emailflag")))
		MyTopicMode = Dvbbs.Checkstr(Trim(Request.Form("topicximoo")))
		MyLockTopic = Dvbbs.Checkstr(Trim(Request.Form("locktopic")))
		Myistop = Dvbbs.Checkstr(Trim(Request.Form("istop")))
		Myistopall = Dvbbs.Checkstr(Trim(Request.Form("istopall")))
		TopicMode = Request.Form("topicmode")

		If Dvbbs.strLength(topic)> CLng(Dvbbs.Board_Setting(45)) Then
			If ajaxPost Then
				Call showAjaxMsg(0,Replace(template.Strings(23),"{$topiclimited}",Dvbbs.Board_Setting(45)),"","")
			Else
				parameter="showerr.asp?ShowErrType="&Dvbbs.ErrType&"&ErrCodes=<li>"&Replace(template.Strings(23),"{$topiclimited}",Dvbbs.Board_Setting(45))&"<BR>2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
				Response.redirect parameter
			End If
		End If
		Rem 限制提交数据不能大于64K
		If Len(Content) > 64*1024*1024 Then
			If ajaxPost Then
				Call showAjaxMsg(0,"您提交的数据过大，提交数据不能大于64K.","","")
			Else
				parameter="showerr.asp?ShowErrType="&Dvbbs.ErrType&"&ErrCodes=<li>您提交的数据过大，提交数据不能大于64K&action=OtherErr&autoreload=1"
				Response.redirect parameter
			End if
		End If

		Rem 老迷增加xhtml格式限制
		Dim XMLPOST,XHTML
		XHTML=False
		If XHTML Then
			Set XMLPOST=Dvbbs.CreateXmlDoc("msxml2.DOMDocument"& MsxmlVersion)
			If XMLPOST.loadxml("<xhtml>" & replace(Content,"&","&amp;") &"</xhtml>") Then
				Content=replace(Mid(XMLPOST.documentElement.xml,8,Len(XMLPOST.documentElement.xml)-15),"&amp;","&")
			Else
				If ajaxPost Then
					Call showAjaxMsg(0,"您提交的数据不合法(必须提交XHTML格式).","","")
				Else
					parameter="showerr.asp?ShowErrType="&Dvbbs.ErrType&"&ErrCodes=<li>您提交的数据不合法(必须提交XHTML格式)&action=OtherErr&autoreload=1"
					Response.redirect parameter
				End if

			End If
			Set XMLPOST=Nothing
		End If
		If Dvbbs.strLength(Content) > CLng(Dvbbs.Board_Setting(16)) Then
			If ajaxPost Then
				Call showAjaxMsg(0,Replace(template.Strings(24),"{$bodylimited}",Dvbbs.Board_Setting(16)),"","")
			Else
				Response.redirect "showerr.asp?ShowErrType="&Dvbbs.ErrType&"&ErrCodes=<li>"&Replace(template.Strings(24),"{$bodylimited}",Dvbbs.Board_Setting(16))&"<BR>2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
			End If
		End If
		REM 2004-4-23添加限制帖子内容最小字节数,下次在模板中添加。Dvbbs.YangZheng
		If Dvbbs.strLength(Content) < CLng(Dvbbs.Board_Setting(52)) And Not CLng(Dvbbs.Board_Setting(52)) = 0 Then
			If ajaxPost Then
				Call showAjaxMsg(0,Replace(template.Strings(24),"大于{$bodylimited}","小于"&Dvbbs.Board_Setting(52)),"","")
			Else
				parameter="showerr.asp?ShowErrType="&Dvbbs.ErrType&"&ErrCodes=<li>"&Replace(template.Strings(24),"大于{$bodylimited}","小于"&Dvbbs.Board_Setting(52))&"<BR>2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
				Response.redirect parameter
			End if
		End If
		Dim testContent
        '如果使用了非HTML编辑器，则替换换行标志以实现换行 by 雨·漫步
        If Dvbbs.GroupSetting(67) = 0 Then
            Content = Replace(Content,vbCrLf,"<br />")
        End If
		testContent=Content
		testContent=Replace(testContent,vbNewLine,"")
		testContent=Replace(testContent," ","")
		testContent=Replace(testContent,"&nbsp;","")
		'testContent=Trim(Dvbbs.Replacehtml(testContent))
		If testContent="" and InStr(Content,"<img")=0 and InStr(Content,"<input")=0 and InStr(Content,"<object")=0  and InStr(Content,"<embed")=0 Then
			If ajaxpost Then
				Call showAjaxMsg(0,"您没有填写内容,请填写帖子内容","","")
			Else
				Response.redirect "showerr.asp?ShowErrType="&Dvbbs.ErrType&"&ErrCodes=<li>您没有填写内容.2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
			End If
		End If
		If Dvbbs.UserID=0 Then
			mailflag=0:signflag=0
		Else
			If Not IsNumeric(mailflag) Or mailflag="" Then mailflag=0
			mailflag=CInt(mailflag)
			If Not IsNumeric(signflag) Or signflag="" Then signflag=1
			signflag=CInt(signflag)
		End If
		If TopicMode<>"" and IsNumeric(TopicMode) Then TopicMode=Cint(TopicMode) Else TopicMode=0
		If Request.form("upfilerename")<>"" Then
			ihaveupfile=1
			upfileinfo=Replace(Request.form("upfilerename"),"'","")
			upfileinfo=Replace(upfileinfo,";","")
			upfileinfo=Replace(upfileinfo,"--","")
			upfileinfo=Replace(upfileinfo,")","")
			Dim fixid,upfilelen
			fixid=Replace(upfileinfo," ","")
			fixid=Replace(fixid,",","")
			If Not IsNumeric(fixid) Then ihaveupfile=0
			upfilelen=len(upfileinfo)
			upfileinfo=left(upfileinfo,upfilelen-1)
		Else
			ihaveupfile=0
		End If
		voteid=0
		isvote=0
		Dim VoteTemp,VoteTemp1,VoteTemp2
		If Action = 7 Then
			votetype=Dvbbs.Checkstr(request.Form("votetype"))
			If IsNumeric(votetype)=0 or votetype="" Then votetype=0
			vote=Dvbbs.Checkstr(trim(Replace(request.Form("vote"),"|","")))
			Dim j,k,vote_1,votelen,votenumlen
			If vote="" Then
				Dvbbs.AddErrCode(81)
			Else
				vote=split(vote,chr(13)&chr(10))
				j=0
				For i = 0 To ubound(vote)
					VoteTemp1 = ""
					VoteTemp2 = ""
					If Not (vote(i)="" Or vote(i)=" ") Then
						VoteTemp = Split(vote(i),"@@")
						If Ubound(VoteTemp) = 3 Then '判断是否调查
							If VoteTemp(1)="0" or VoteTemp(1)="1" Then
								VoteTemp1 = Split(VoteTemp(2),"$$")
								For k=0 to Ubound(VoteTemp1)-1
									VoteTemp2 = VoteTemp2 & "0$$"
								Next
							Else
								VoteTemp2 = 0
							End If
							votenum= votenum & VoteTemp2&"|"
						Else
							votenum=votenum&"0|"
						End If
						vote_1=""&vote_1&""&vote(i)&"|"
						j=j+1
					End If
					If i>cint(Dvbbs.Board_Setting(32))-2 Then Exit For
				Next
				'For k = 1 to j
				'	votenum=""&votenum&"0|"
				'Next
				votelen=len(vote_1)
				votenumlen=len(votenum)
				votenum=left(votenum,votenumlen-1)
				vote=left(vote_1,votelen-1)
				'Response.Write votenum
				'Response.End
			End If
			If Not IsNumeric(request("votetimeout")) Then
				Dvbbs.AddErrCode(82)
			Else
				If request("votetimeout")="0" Then
					votetimeout=dateadd("d",9999,Now())
				Else
					votetimeout=dateadd("d",CCur(request("votetimeout")),Now())
				End If
				votetimeout=Replace(Replace(CSTR(votetimeout+Dvbbs.Forum_Setting(0)/24),"上午",""),"下午","")
			End If
		End If
		If Action = 5 Or Action = 7 Then
			CanLockTopic=False
			CanTopTopic=False
			CanTopTopic_a=False
			If Topic="" OR Replace(Topic&"","