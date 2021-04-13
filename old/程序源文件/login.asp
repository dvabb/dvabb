<!--#include file="Conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="inc/chan_const.asp"-->
<!--#include file="inc/chkinput.asp"-->
<!--#include file="inc/Email_Cls.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="dv_dpo/cls_dvapi.asp"-->
<%
Dim ajaxPro
Dim ajaxpostwin
ajaxPro=CInt(Request.Form("ajaxPost"))
ajaxpostwin=Request.Form("ajaxpostwin")
'Response.Write(ajaxPro):Response.End()

If ajaxPro=1 Then
	ajaxPro=True
Else
	ajaxPro=False Rem 非ajax登陆
End If

Dim comeurl
Dim TruePassWord
'-----------------'owz
Dim EnterCount
If Request.Cookies("count")<>"" Then
	EnterCount=Request.Cookies("count")
Else
	EnterCount=Request.QueryString("count")'地址栏带参数
End If
'----------------------
session("flag")=empty
Dvbbs.LoadTemplates("login")
Dvbbs.stats=template.Strings(1)
Dvbbs.Nav()
Dvbbs.Head_var 0,0,template.Strings(0),"login.asp"
TruePassWord=Dvbbs.Createpass

Select Case request("action")
	Case "chk"
		Dvbbs_ChkLogin
		Dvbbs.Showerr()
	Case "redir"
		redir		
		Dvbbs.Showerr()		
	Case "save_redir_reg"
		call save_redir_reg()	
		Dvbbs.Showerr()
	Case Else
		Main
End Select
Dvbbs.ActiveOnline
Dvbbs.Footer()
Dvbbs.PageEnd()

Sub Main()
	Dim TempStr
	TempStr = template.html(0)
		If Dvbbs.forum_setting(79)="0" Then
		TempStr = Replace(TempStr,"{$getcode}","")
		Else 
		TempStr = Replace(TempStr,"{$getcode}",Replace(template.html(23),"{$codestr}",Dvbbs.GetCode()))
		End If 
	TempStr = Replace(TempStr,"{$rayuserlogin}",template.html(1))
	Dim Comeurl,tmpstr
	If Request("f")<>"" Then
		Comeurl=Request("f")
	ElseIf Request.ServerVariables("HTTP_REFERER")<>"" Then 
		tmpstr=split(Request.ServerVariables("HTTP_REFERER"),"/")
		Comeurl=tmpstr(UBound(tmpstr))
	Else
		If isUrlreWrite = 1 Then
			Comeurl="index.html"
		Else
			Comeurl="index.asp"
		End If
	End If

	If UBound(Split(Comeurl,"."))>=2 Or Trim(Comeurl)="" Then
		Comeurl = "index.asp"
	End If

	TempStr = Replace(TempStr,"{$comeurl}",Comeurl)
	Response.Write TempStr
	TempStr=""	
End Sub

Function Dvbbs_ChkLogin	
	Dim UserIP
	Dim username
	Dim userclass
	Dim password
	Dim article
	Dim usercookies
	Dim mobile
	Dim chrs,i
	UserIP=Dvbbs.UserTrueIP
	mobile=trim(Dvbbs.CheckStr(request("passport")))
    If Dvbbs.forum_setting(79)="1" Then			
        If mobile="" And Not Dvbbs.CodeIsTrue() Then
			    If  ajaxPro Then
				    If ajaxpostwin="1" Then 'modifty by reoaiq at 091022
						If Dvbbs.forum_setting(120)="1" Then 
						strString("语音验证码校验失败@@@@2")
						Else 
						strString("验证码校验失败@@@@0")
						End If
					else 
						If Dvbbs.forum_setting(120)="1" Then 
						strString("语音验证码校验失败@@@@2")
						Else 
						strString("验证码校验失败@@@@0")
						End If
					end if  
				Else 
				    If ajaxpostwin="1" Then 
						If Dvbbs.forum_setting(120)="1" Then 
						strString("语音验证码校验失败@@@@2")
						Else 
						strString("验证码校验失败@@@@0")
						End If
					Else 
                    Response.redirect "showerr.asp?ErrCodes=<li>验证码校验失败，请返回刷新页面后再输入验证码。&action=OtherErr"
					End If 
                End If 			
        End If
    End If
	
	If Request("username")="" Then
		If Request("passport")="" Then	
			If Not ajaxPro Then
				Dvbbs.AddErrCode(10)
			Else			
				strString("用户名不能为空@@@@0") 'o
			End If
		End If
	Else
		username=trim(Dvbbs.CheckStr(request("username")))
		If ajaxPro Then username = Dvbbs.CheckStr(unescape(username))
	End If
	If request("password")="" and mobile="" Then
		If Not ajaxPro Then
			Dvbbs.AddErrCode(11)
		Else 
			strString("密码不能为空!@@@@0") 'o
		End If
	Else
		password=md5(trim(Dvbbs.CheckStr(request("password"))),16)
		If Request("password") = "" Then password = ""
	End If

	If Dvbbs.ErrCodes<>"" Then Exit Function

	'-----------------------------------------------------------------以上注释(o)
	'系统整合
	'-----------------------------------------------------------------
	Dim DvApi_Obj,DvApi_SaveCookie,SysKey
	If DvApi_Enable Then
		Set DvApi_Obj = New DvApi
			'DvApi_Obj.NodeValue "syskey",SysKey,0,False
			DvApi_Obj.NodeValue "action","login",0,False
			DvApi_Obj.NodeValue "username",UserName,1,False
			Md5OLD = 1
			SysKey = Md5(DvApi_Obj.XmlNode("username")&DvApi_SysKey,16)
			Md5OLD = 0
			DvApi_Obj.NodeValue "syskey",SysKey,0,False
			DvApi_Obj.NodeValue "password",Request("password"),0,False
			DvApi_Obj.SendHttpData
			'strString(UserName&"---"&SysKey&"---"&Md5(Request("password"),16))
			If DvApi_Obj.Status = "1" Then
				If Not ajaxPro Then
					Response.Redirect "showerr.asp?ErrCodes="& DvApi_Obj.Message &"&action=OtherErr"
				Else
					strString(DvApi_Obj.Message &".@@@@0")
				End If
			Else
				DvApi_SaveCookie = DvApi_Obj.SetCookie(SysKey,UserName,Password,request("CookieDate"))
			End If
		Set DvApi_Obj = Nothing
	End If
	'-----------------------------------------------------------------

	usercookies=request("CookieDate")
	'判断更新cookies目录
	Dim cookies_path_s,cookies_path_d,cookies_path
	cookies_path_s=split(Request.ServerVariables("PATH_INFO"),"/")
	cookies_path_d=ubound(cookies_path_s)
	cookies_path="/"
	For i=1 to cookies_path_d-1
		If not (cookies_path_s(i)="upload" or cookies_path_s(i)="admin") Then cookies_path=cookies_path&cookies_path_s(i)&"/"
	Next
	If dvbbs.cookiepath<>cookies_path Then
		cookies_path=replace(cookies_path,"'","")
		Dvbbs.execute("update dv_setup set Forum_Cookiespath='"&cookies_path&"'")
		Dim setupData 
		Dvbbs.CacheData(26,0)=cookies_path
		Dvbbs.Name="setup"
		Dvbbs.value=Dvbbs.CacheData
	End If
	
	If ChkUserLogin(username,password,mobile,usercookies,1)=false Then
		Set chrs=Dvbbs.Execute("select Passport,IsChallenge from [Dv_User] where username='"&username&"' and IsChallenge=1")
		If chrs.eof and chrs.bof Then
			If Not ajaxPro Then
				Dvbbs.AddErrCode(12)
			Else 
				strString("本论坛不存在该用户名.@@@@0")'o
			End If
			Exit Function
		End If
		set chrs=nothing
	End If

	Dim comeurlname
	If instr(lcase(request("comeurl")),"reg.asp")>0 or instr(lcase(request("comeurl")),"login.asp")>0 or trim(request("comeurl"))="" or instr(lcase(request("comeurl")),"index.asp")>0 Or instr(lcase(request("comeurl")),"showerr.asp")>0 Then
		comeurlname=""
		If isUrlreWrite = 1 Then
			Comeurl="index.html"
		Else
			comeurl="index.asp"
		End If
	Else
		comeurl=request("comeurl")		
		comeurlname="<li><a href="&request("comeurl")&">"&request("comeurl")&"</a></li>"
	End If
	
	Dim TempStr
	TempStr = template.html(2)
	'-----------------------------------------------------------------
	'系统整合
	'-----------------------------------------------------------------
	If DvApi_Enable Then
		Response.Write DvApi_SaveCookie
		Response.Flush
	End If
	'-----------------------------------------------------------------
	TempStr = Replace(TempStr,"{$ray_logininfo}","")
	TempStr = Replace(TempStr,"{$comeurl}",comeurl)
	TempStr = Replace(TempStr,"{$comeurlinfo}",comeurlname)
	TempStr = Replace(TempStr,"{$forumname}",Dvbbs.Forum_Info(0))

	Session.Contents.Remove("xcount")

	If Not ajaxPro And DvApi_Enable Then'非ajax
		Response.Write TempStr
	ElseIf Not ajaxPro And Not DvApi_Enable Then
		Response.Redirect(comeurl)
	Else
		Response.Cookies("count")=""'o(清空ajax里写入的cookies)
		strString(comeurl&"@@@@1")'o
	End If 
End Function

Function strAnsi2Unicode(asContents)
	Dim len1,i,varchar,varasc
	strAnsi2Unicode = ""
	len1=LenB(asContents)
	If len1=0 Then Exit Function
	  For i=1 to len1
	  	varchar=MidB(asContents,i,1)
	  	varasc=AscB(varchar)
	  	If varasc > 127  Then
	  		If MidB(asContents,i+1,1)<>"" Then
	  			strAnsi2Unicode = strAnsi2Unicode & chr(ascw(midb(asContents,i+1,1) & varchar))
	  		End If
	  		i=i+1
	     Else
	     	strAnsi2Unicode = strAnsi2Unicode & Chr(varasc)
	     End If	
	  Next
End Function

Sub save_redir_reg()
	If Session("re_challenge_reg_temp")="" Then
		Dvbbs.AddErrCode(14)
		Exit Sub
	End If

	Dim username,sex,pass1,pass2,password,ErrCodes
	Dim useremail,face,width,height
	Dim oicq,sign,showRe,birthday
	Dim mailbody,sendmsg,rndnum,num1
	Dim quesion,answer,topic
	Dim userinfo,usersetting
	Dim userclass,UserIM
	Dim re_challenge_reg_temp
	Dim rs,sql,i,namebadword,SplitWords
	Dim t
	Dim StatUserID,UserSessionID
	Dim TempStr
	t = Request("t")
	If t = "" Or Not IsNumeric(t) Then t = 1
	t = Cint(t)
	If t <> 1 And t <> 2 Then t = 1
	re_challenge_reg_temp=split(Session("re_challenge_reg_temp"),"|||")

	If Request("name")="" or strLength(Request("name"))>Cint(Dvbbs.Forum_Setting(41)) or strLength(Request("name"))<Cint(Dvbbs.Forum_Setting(40)) Then
		Dvbbs.AddErrCode(17)
	Else
		username=Dvbbs.CheckStr(Trim(Request("name")))
	End If

	If Instr(username,"=")>0 or Instr(username,"%")>0 or Instr(username,chr(32))>0 or Instr(username,"?")>0 or Instr(username,"&")>0 or Instr(username,";")>0 or Instr(username,",")>0 or Instr(username,"'")>0 or Instr(username,",")>0 or Instr(username,chr(34))>0 or Instr(username,chr(9))>0 or Instr(username,"