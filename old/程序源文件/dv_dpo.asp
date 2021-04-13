<!--#include file="Conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="inc/chan_const.asp"-->
<!--#include file="inc/chkinput.asp"-->
<!--#include file="inc/Email_Cls.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="dv_dpo/Api_Config.asp"-->
<%
Dim XMLDom,XmlDoc,Node,Status,Messenge
Dim UserName,Act,appid
Status = 1
Messenge = ""
If Not DvApi_Enable Then Response.end
If Request.QueryString<>"" Then
	SaveUserCookie()
Else
	Set XmlDoc = Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
	XmlDoc.ASYNC = False
	If Not XmlDoc.LOAD(Request) Then
		Status = 1
		Messenge = "数据非法，操作中止！"
		appid = "未知"
	Else
		If Not (XmlDoc.documentElement.selectSingleNode("userip") is nothing) Then
			Dvbbs.UserTrueIP = Dvbbs.CheckStr(XmlDoc.documentElement.selectSingleNode("userip").text)
		End If
		If CheckPost() Then
			Select Case Act
				Case "checkname"
					Checkname()
				Case "reguser"
					Reguser()
				Case "login"
					UesrLogin()
				Case "logout"
					LogoutUser()
				Case "update"
					UpdateUser()
				Case "delete"
					Deleteuser()
				Case "lock"
					Lockuser()
				Case "getinfo"
					GetUserinfo()
			End Select
		End If
	End If
	ReponseData()
	Set XmlDoc = Nothing
End If
Dvbbs.PageEnd()

Sub ReponseData()
	'XmlDoc.loadxml "<root><appid>powereasy</appid><status>0</status><body><message/><email/><question/><answer/><savecookie/><truename/><gender/><birthday/><qq/><msn/><mobile/><telephone/><address/><zipcode/><homepage/><userip/><jointime/><experience/><ticket/><valuation/><balance/><posts/><userstatus/></body></root>"
	If Act <> "getinfo" Then
		XmlDoc.loadxml "<root><appid>powereasy</appid><status>0</status><body><message/></body></root>"
	End If
	XmlDoc.documentElement.selectSingleNode("appid").text = "Dvbbs"
	XmlDoc.documentElement.selectSingleNode("status").text = status
	XmlDoc.documentElement.selectSingleNode("body/message").text = ""
	Set Node = XmlDoc.createCDATASection(Replace(Messenge,"]]>","]]&gt;"))
	XmlDoc.documentElement.selectSingleNode("body/message").appendChild(Node)
	Response.Clear
	Response.ContentType="text/xml"
	Response.CharSet="gb2312"
	Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>"&vbNewLine
	Response.Write XmlDoc.documentElement.XML
End Sub


Function CheckPost()
	CheckPost = False
	Dim Syskey
	If XmlDoc.documentElement.selectSingleNode("action") is Nothing or XmlDoc.documentElement.selectSingleNode("syskey") is Nothing or XmlDoc.documentElement.selectSingleNode("username")  is Nothing Then
		Status = 1
		Messenge = Messenge & "<li>非法请求。"
		Exit Function
	End If

	UserName = Dvbbs.checkstr(Trim(XmlDoc.documentElement.selectSingleNode("username").text))
	Syskey = XmlDoc.documentElement.selectSingleNode("syskey").text
	Act = XmlDoc.documentElement.selectSingleNode("action").text
	Appid = Dvbbs.CheckStr(XmlDoc.documentElement.selectSingleNode("appid").text)
	
	Dim NewMd5,OldMd5
	NewMd5 = Md5(UserName&DvApi_SysKey,16)
	Md5OLD = 1
	OldMd5 = Md5(UserName&DvApi_SysKey,16)
	Md5OLD = 0
	If DvApi_SysKey = "API_TEST" or Syskey = "Syskey" Then
		Status = 1
		Messenge = Messenge & "<li>默认非法请求。"
		Exit Function
	End If
	If Syskey=NewMd5 or Syskey=OldMd5 Then
		CheckPost = True
	Else
		Status = 1
		Messenge = Messenge & "<li>请求数据验证不通过，请与管理员联系。"
	End If
End Function


Sub GetUserinfo()
	Dim Rs,Sql
	Dim Userinfo,UserIM
	
	XmlDoc.loadxml "<root><appid>powereasy</appid><status>0</status><body><message/><email/><question/><answer/><savecookie/><truename/><gender/><birthday/><qq/><msn/><mobile/><telephone/><address/><zipcode/><homepage/><userip/><jointime/><experience/><ticket/><valuation/><balance/><posts/><userstatus/></body></root>"
	
	Sql = "Select Top 1 * From Dv_User Where UserName='"&Dvbbs.Checkstr(UserName)&"'"
	Set Rs = Dvbbs.Execute(Sql)
	If Not Rs.Eof And Not Rs.Bof Then
		Userinfo = Split(Rs("UserInfo"),"|||")
		UserIM = Split(Rs("UserIM"),"|||")
		XmlDoc.documentElement.selectSingleNode("body/email").text = Rs("UserEmail")&""
		XmlDoc.documentElement.selectSingleNode("body/question").text = Rs("UserQuesion")&""
		XmlDoc.documentElement.selectSingleNode("body/answer").text = Rs("UserAnswer")&""
		'XmlDoc.documentElement.selectSingleNode("body/savecookie").text = ""
		XmlDoc.documentElement.selectSingleNode("body/gender").text = Rs("Usersex")&""
		XmlDoc.documentElement.selectSingleNode("body/birthday").text = Rs("UserBirthday")&""
		XmlDoc.documentElement.selectSingleNode("body/mobile").text = Rs("UserMobile")&""
		'XmlDoc.documentElement.selectSingleNode("body/zipcode").text = ""
		XmlDoc.documentElement.selectSingleNode("body/userip").text = Rs("UserLastIP")&""
		XmlDoc.documentElement.selectSingleNode("body/jointime").text = Rs("JoinDate")&""
		XmlDoc.documentElement.selectSingleNode("body/experience").text = Rs("userEP")&""
		XmlDoc.documentElement.selectSingleNode("body/ticket").text = Rs("UserTicket")&""
		XmlDoc.documentElement.selectSingleNode("body/valuation").text = Rs("userCP")&""
		XmlDoc.documentElement.selectSingleNode("body/balance").text = Rs("UserMoney")&""
		XmlDoc.documentElement.selectSingleNode("body/posts").text = Rs("UserPost")&""
		XmlDoc.documentElement.selectSingleNode("body/userstatus").text = Rs("Lockuser")&""
		XmlDoc.documentElement.selectSingleNode("body/homepage").text = UserIM(0)
		XmlDoc.documentElement.selectSingleNode("body/qq").text = UserIM(1)
		XmlDoc.documentElement.selectSingleNode("body/msn").text = UserIM(3)
		XmlDoc.documentElement.selectSingleNode("body/truename").text = Userinfo(0)
		XmlDoc.documentElement.selectSingleNode("body/telephone").text = Userinfo(13)
		XmlDoc.documentElement.selectSingleNode("body/address").text = Userinfo(14)
		Status = 0
		Messenge = Messenge & "<li>读取用户资料成功。"
	Else
		Status = 1
		Messenge = Messenge & "<li>该用户不存在。"
	End If
	Rs.Close
	Set Rs = Nothing
End Sub


Sub Deleteuser()
	Dim D_Users,i,UserID,AllUserID
	Dim Rs
	D_Users = Split(UserName,",")
	AllUserID = ""
	For i=0 To UBound(D_Users)
		Set Rs=Dvbbs.Execute("Select UserName,UserID from [dv_User] where UserName='"&Dvbbs.Checkstr(D_Users(i))&"'")
		If not (rs.eof and rs.bof) then
			AllUserID = AllUserID & Rs(1) & ","
			Dvbbs.Execute("update dv_message set delR=1 where incept='"&Dvbbs.Checkstr(rs(0))&"' and delR=0")
			Dvbbs.Execute("update dv_message set delS=1 where sender='"&Dvbbs.Checkstr(rs(0))&"' and delS=0 and issend=0")
			Dvbbs.Execute("update dv_message set delS=1 where sender='"&Dvbbs.Checkstr(rs(0))&"' and delS=0 and issend=1")
			Dvbbs.Execute("delete from dv_message where incept='"&Dvbbs.Checkstr(rs(0))&"' and delR=1") 
			Dvbbs.Execute("update dv_message set delS=2 where sender='"&Dvbbs.Checkstr(rs(0))&"' and delS=1")
			Dvbbs.Execute("delete from dv_friend where F_username='"&Dvbbs.Checkstr(rs(0))&"'") 
			Dvbbs.Execute("delete from dv_bookmark where username='"&Dvbbs.Checkstr(rs(0))&"'")
			Messenge = Messenge & "<li>用户（"&D_Users(i)&"）删除成功。"
		End If
	Next
	If AllUserID<>"" Then
		If Right(AllUserID,1) = "," Then AllUserID = Left(AllUserID,Len(AllUserID)-1)
		'删除用户的帖子和精华
		Dvbbs.Execute("Delete From dv_topic where PostUserID in ("&AllUserID&")")
		Dim PostTable
		PostTable = AllPostTable
		For i=0 to ubound(PostTable,2)
			Dvbbs.Execute("Delete From "&PostTable(0,i)&" where PostUserID in ("&AllUserID&")")
		Next
		Dvbbs.Execute("Delete From dv_besttopic where PostUserID in ("&AllUserID&")")
		'删除用户上传表
		Dvbbs.Execute("Delete From dv_upfile where F_UserID in ("&AllUserID&")")
		Dvbbs.Execute("Delete From [dv_user] where userid in ("&AllUserID&")")
	End If
	Status = 0
End Sub

Function AllPostTable()
	Dim Trs
	Set Trs = Dvbbs.Execute("select TableName from [Dv_TableList]")
	AllPostTable = TRs.GetRows(-1)
	Trs.Close 
	Set Trs = Nothing
End Function

Sub SaveUserCookie()
	Dim S_syskey,Password,SaveCookie,TruePassWord,userclass,Userhidden
	S_syskey = Request.QueryString("syskey")
	UserName = Request.QueryString("UserName")
	Password = Request.QueryString("Password")
	SaveCookie = Request.QueryString("savecookie")
	If UserName="" or S_syskey="" Then Exit Sub
	Dim NewMd5,OldMd5
	NewMd5 = Md5(UserName&DvApi_SysKey,16)
	Md5OLD = 1
	OldMd5 = Md5(UserName&DvApi_SysKey,16)
	Md5OLD = 0
	If Not (S_syskey=NewMd5 or S_syskey=OldMd5) Then
		Exit Sub
	End If
	If EnabledSession Then
		Session(Dvbbs.CacheName & "UserID")=Empty
		Set Dvbbs.UserSession=Nothing
	End If
	If SaveCookie="" or Not IsNumeric(SaveCookie) Then SaveCookie = 0
	'用户退出
	If Password = "" Then
		Response.Cookies(Dvbbs.Forum_sn).path=Dvbbs.cookiepath
		Response.Cookies(Dvbbs.Forum_sn)("username")=""
		Response.Cookies(Dvbbs.Forum_sn)("password")=""
		Response.Cookies(Dvbbs.Forum_sn)("userclass")=""
		Response.Cookies(Dvbbs.Forum_sn)("userid")=""
		Response.Cookies(Dvbbs.Forum_sn)("userhidden")=""
		Response.Cookies(Dvbbs.Forum_sn)("usercookies")=""
		Session("flag")=Empty
		Exit Sub
	End If

	'用户登陆
	'Password = Md5(Password,16)
	TruePassWord = Dvbbs.Createpass
	Dim Rs,Sql
	If Not IsObject(Conn) Then ConnectionDatabase
	Set Rs = Dvbbs.iCreateObject("Adodb.RecordSet")
	Sql = "Select Top 1 UserID,UserName,UserPassword,Userclass,Userhidden,TruePassWord From Dv_User Where UserName='"&Dvbbs.Checkstr(UserName)&"'"
	Rs.Open Sql,Conn,1,3
	If Not Rs.Eof And Not Rs.Bof Then
		If Rs(2)<>Password Then
			Exit Sub
		End If
		Dvbbs.UserID = Rs(0)
		UserName = Rs(1)
		UserClass = Rs(3)
		Userhidden = Rs(4)
		Rs(5) = TruePassword
		Rs.Update
	Else
		Exit Sub
	End If
	Rs.Close
	Set Rs = Nothing
	'Response.Write "document.write("""&Dvbbs.cookiepath&""");"
	Select case SaveCookie
		case 0
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = SaveCookie
		case 1
			Response.Cookies(Dvbbs.Forum_sn).Expires=Date+1
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = SaveCookie
		case 2
			Response.Cookies(Dvbbs.Forum_sn).Expires=Date+31
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = SaveCookie
		case 3
			Response.Cookies(Dvbbs.Forum_sn).Expires=Date+365
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = SaveCookie
	End Select
	Response.Cookies(Dvbbs.Forum_sn).path = Dvbbs.cookiepath
	Response.Cookies(Dvbbs.Forum_sn)("username") = UserName
	Response.Cookies(Dvbbs.Forum_sn)("userid") = Dvbbs.UserID
	Response.Cookies(Dvbbs.Forum_sn)("password") = TruePassWord
	Response.Cookies(Dvbbs.Forum_sn)("userclass") = UserClass
	Response.Cookies(Dvbbs.Forum_sn)("userhidden") = Userhidden
	rem 清除图片上传数的限制
	Response.Cookies("upNum")=0
	'Response.Write "document.write(""OK"");"
End Sub


Sub Checkname()
	Dim UserEmail
	Dim Temp_tr,i,Rs,Sql
	UserEmail = Dvbbs.checkstr(Trim(XmlDoc.documentElement.selectSingleNode("email").text))
	Dvbbs.LoadTemplates("login")
	LoadRegSetting()
	'信息验证
	If strLength(UserName)>Cint(Dvbbs.Forum_Setting(41)) or strLength(UserName)<Cint(Dvbbs.Forum_Setting(40)) Then
		Temp_tr = Template.Strings(28)
		Temp_tr = Replace(Temp_tr,"{$RegMaxLength}",Dvbbs.Forum_Setting(41))
		Temp_tr = Replace(Temp_tr,"{$RegLimLength}",Dvbbs.Forum_Setting(40))
		Messenge = Messenge & "<li>"+Temp_tr
		Temp_tr = ""
	Else
		If XMLDom.documentElement.selectSingleNode("@checknumeric").text = "1" Then
			If IsNumeric(UserName) Then
				Messenge = Messenge & "<li>论坛不接受全数字的用户名注册."
			End If
		End If
		If Instr(UserName,"=")>0 or Instr(UserName,"%")>0 or Instr(UserName,chr(32))>0 or Instr(UserName,"?")>0 or Instr(UserName,"&")>0 or Instr(UserName,";")>0 or Instr(UserName,",")>0 or Instr(UserName,"'")>0 or Instr(UserName,",")>0 or Instr(UserName,chr(34))>0 or Instr(UserName,chr(9))>0 or Instr(UserName,"