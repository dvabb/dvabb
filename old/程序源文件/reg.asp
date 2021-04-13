<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/chan_const.asp"-->
<!--#include file="inc/chkinput.asp"-->
<!--#include file="inc/Email_Cls.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="dv_dpo/cls_dvapi.asp"-->
<%
Const question_answer=False
Dim Selectinfo(5)
Dim XMLDom
Dim Stats,ErrCodes
session("flag")=empty
If request("action")="postipinfo" and Request.form("comfrom")<>"" Then
	saveipinfo
Else
	LoadRegSetting()
	If Request("t")="1" Then
		ChkReg_Main()
	Else
		Reg_Main()		
	End If
End If

Sub Reg_Main()
	Dim PageSid
	PageSid = Dvbbs.Skinid
	Dvbbs.LoadTemplates("usermanager")
	Dvbbs.Skinid = PageSid
	Selectinfo(0)=chk_select("",template.Strings(11))
	Selectinfo(1)=chk_select("",template.Strings(12))
	Selectinfo(2)=chk_select("",template.Strings(13))
	Selectinfo(3)=chk_select("",template.Strings(14))
	Selectinfo(4)=Chk_KidneyType("character","",template.Strings(15))
	Selectinfo(5)=chk_select("",template.Strings(16))
	Dvbbs.LoadTemplates("login")
	Stats=Split(template.Strings(25),"||")
	Dvbbs.Stats=Stats(0)
	Dvbbs.Nav()
	Dvbbs.ActiveOnline
	If request("action")<>"" and (not Request.servervariables("REQUEST_METHOD") = "POST"  ) Then
		'Response.Write("ERROR_1"):Response.End
		'Response.redirect "showerr.asp?ErrCodes=<li>您提交的参数错误.1</li>&action=OtherErr"
	ElseIf (Request.servervariables("REQUEST_METHOD") = "POST"  )  Then
		'If Not CheckFormID(Request.form(GetFormID())) Then
			'Response.Write("ERROR_2"):Response.End
			'Response.redirect "showerr.asp?ErrCodes=您提交的参数错误.2&action=OtherErr"			
		'End If
	End If

	If Cint(dvbbs.Forum_Setting(37))=0 Then
		ErrCodes=ErrCodes+"<li>"+template.Strings(26)
	Else	
		If request("action")="apply" Then
			Dvbbs.stats=Stats(2)
			Dvbbs.Head_var 0,0,Stats(0),"reg.asp"
			reg_2()
		ElseIf request("action")="save" Then
			Dvbbs.stats=Stats(3)
			Dvbbs.Head_var 0,0,Stats(0),"reg.asp"
			reg_3()
		ElseIf request("action")="redir" Then
			Dvbbs.stats=Stats(3)
			Dvbbs.Head_var 0,0,Stats(0),"reg.asp"
			redir()
		Else
			Dvbbs.stats=Stats(1)
			Dvbbs.Head_var 0,0,Stats(0),"reg.asp"
			reg_1()
		End If
	End If
	Dvbbs.Showerr()
	If ErrCodes<>"" Then 
		Dvbbs.PageEnd()
		Response.redirect "showerr.asp?ErrCodes="&ErrCodes&"&action=OtherErr"	
	End If
	Dvbbs.Footer()
	Dvbbs.PageEnd()
End Sub

Sub saveipinfo()
	Dim Node,rs
	Set XMLDom=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	If XMLDom.loadxml(Dvbbs.CacheData(27,0)) Then
		If XMLDom.documentElement.selectSingleNode("checkip/@use").text = 1 Then
			Set Node=XMLDom.documentElement.selectSingleNode("checkip/iplist1")
			If Not Node.selectNodes("ip").length =0 Then
				If Not IpInList(Node) Then
				Set Rs=Dvbbs.Execute("select Forum_BirthUser From Dv_setup")
				If Not XMLDom.loadxml(Rs(0)) Then
					XMLDom.LoadXML "<?xml version=""1.0""?><regpost/>"
				Else
					Set Node=XMLDom.documentElement.selectNodes("ip")
					If Node.length > 200 Then
						XMLDom.documentElement.removeChild(XMLDom.documentElement.firstChild)
					End If
				End If
				If XMLDom.documentElement.selectSingleNode("ip[.='"&Dvbbs.userTrueIP&"']") Is Nothing Then
				Set node=XMLDom.documentElement.appendChild(XMLDom.createNode(1,"ip",""))
					node.text=Dvbbs.userTrueIP
					Node.attributes.setNamedItem(XMLDom.createNode(2,"description","")).text=Request.form("comfrom")
					Node.attributes.setNamedItem(XMLDom.createNode(2,"dateandtime","")).text=Now()
					Dvbbs.Execute("update Dv_setup set Forum_BirthUser='"&Dvbbs.checkstr(XMLDom.xml)&"'")
				End If
			End If
			Dvbbs.LoadTemplates("")
				Dvbbs.Stats="提交注册允许请求"
				Dvbbs.Nav()
				Dvbbs.ActiveOnline
				Dvbbs.Head_var 0,0,"提交成功","reg.asp"
				Dvbbs.Dvbbs_Suc("<li>您提交的信息已经成功保存,管理员会尽快处理,请在一个工作日后再次尝试注册.</li>")
				Dvbbs.Footer()
			End If
		End If
	End If
	
End Sub

Sub reg_1()
	Dim TempLateStr
	TempLateStr=template.html(12)
	TempLateStr=Replace(TempLateStr,"{$Forum_Name}",Dvbbs.Forum_Info(0))
	TempLateStr=Replace(TempLateStr,"{$hidden}",GetFormID())

	Response.Write TempLateStr
End Sub

Sub reg_2()
	Dim grouploopinfo,TempLateStr,Rs,FormID,fname,template_html22,template_html24
	TempLateStr=template.html(13)
	If Dvbbs.forum_setting(78)="0" Then
		TempLateStr=Replace(TempLateStr,"{$getcode}","")
	Else
		template_html24=Replace(template.html(24),"{$codestr}",Dvbbs.GetCode())
		TempLateStr=Replace(TempLateStr,"{$getcode}",template_html24)
	End If

	If Dvbbs.forum_setting(107)="0" Then
		TempLateStr=Replace(TempLateStr,"{$regask}","")
	Else
		Dim regask,n
		regask=Split(Dvbbs.forum_setting(105),"!")
		If UBound(regask)>=0 And Trim(regask(0))<>"" Then
			Randomize()
			n = CInt(UBound(regask)*Rnd(now()))
			If n>UBound(regask) Then n=UBound(regask)
			template_html22=Replace(template.html(22),"{$regask_1}",regask(n))
			template_html22=Replace(template_html22,"{$regask_2}",md5(n,16))
			TempLateStr=Replace(TempLateStr,"{$regask}",template_html22)
		Else
			TempLateStr=Replace(TempLateStr,"{$regask}","")
		End If
	End If

	Dim userregface,i,Forum_userface,FaceDefault
	Forum_userface = split(Dvbbs.Forum_userface,"|||")
	FaceDefault=Forum_userface(0)&Forum_userface(1)
	For i = 1 to Ubound(Forum_userface)-1
		userregface = userregface & "<option value="""&Forum_userface(0)&Forum_userface(i)
		userregface = userregface & """>" & Forum_userface(i) & "</option>"
	Next
	TempLateStr=Replace(TempLateStr,"{$color}",Dvbbs.mainsetting(1))
	TempLateStr=Replace(TempLateStr,"{$FaceDefault}",FaceDefault)
	TempLateStr=Replace(TempLateStr,"{$Face_select}",userregface)
	TempLateStr=Replace(TempLateStr,"{$FaceMaxWidth}",Dvbbs.Forum_Setting(38))
	TempLateStr=Replace(TempLateStr,"{$FaceMaxHeight}",Dvbbs.Forum_Setting(39))
	TempLateStr=Replace(TempLateStr,"{$ForumFaceMax}",Dvbbs.Forum_Setting(57))
	TempLateStr=Replace(TempLateStr,"{$NameLimLength}",Dvbbs.Forum_Setting(40))
	TempLateStr=Replace(TempLateStr,"{$NameMaxLength}",Dvbbs.Forum_Setting(41))
	TempLateStr=Replace(TempLateStr,"{$Forum_Setting7}",Dvbbs.Forum_UploadSetting(0))
	TempLateStr=Replace(TempLateStr,"{$Forum_Setting23}",Dvbbs.Forum_Setting(23))
	TempLateStr=Replace(TempLateStr,"{$Forum_Setting32}",Dvbbs.Forum_Setting(32))
	TempLateStr=Replace(TempLateStr,"{$Forum_Setting54}",Dvbbs.Forum_Setting(54))
	TempLateStr=Replace(TempLateStr,"{$Forum_Setting42}",Dvbbs.Forum_Setting(42))
	TempLateStr=Replace(TempLateStr,"{$grouploopinfo}",grouploopinfo)
	TempLateStr=Replace(TempLateStr,"{$user_blood}",chk_select("","A,B,AB,O"))
	TempLateStr=Replace(TempLateStr,"{$user_shengxiao}",Selectinfo(0))
	TempLateStr=Replace(TempLateStr,"{$user_occupation}",Selectinfo(1))
	TempLateStr=Replace(TempLateStr,"{$user_marital}",Selectinfo(2))
	TempLateStr=Replace(TempLateStr,"{$user_education}",Selectinfo(3))
	TempLateStr=Replace(TempLateStr,"{$user_character}",Selectinfo(4))
	TempLateStr=Replace(TempLateStr,"{$user_belief}",Selectinfo(5))
	FormID=GetFormID()
	TempLateStr=Replace(TempLateStr,"{$hidden}",FormID)
	If XMLDom.documentElement.selectSingleNode("@usevarform").text = "1" Then
		fname="_"&Md5(FormID,16)
	End If
	TempLateStr=Replace(TempLateStr,"{$username}","username"&fname)
	TempLateStr=Replace(TempLateStr,"{$psw}","psw"&fname)
	TempLateStr=Replace(TempLateStr,"{$pswc}","pswc"&fname)
	If XMLDom.documentElement.selectSingleNode("@checktime").text = "1" Then
		TempLateStr=Replace(TempLateStr,"{$difference}",Replace(template.html(4),"{$options}",Getoptions()))
	Else
		TempLateStr=Replace(TempLateStr,"{$difference}","")
	End If
	Response.Write TempLateStr
End Sub

Function Getoptions()
	Dim xmltime_difference,node
	Set xmltime_difference=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
	xmltime_difference.load Server.MapPath(MyDbPath &"inc\Time_difference.xml")
	For each node in xmltime_difference.documentElement.selectnodes("time_difference")
		Getoptions=Getoptions& "<option value="""&node.selectSingleNode("@value").text&""">"&node.text&"</option>"&vbnewline
	Next
End Function

Function GetFormID()
	Dim i,sessionid
	sessionid = Session.SessionID
	For i=1 to Len(sessionid)
		GetFormID=GetFormID&Chr(Mid(sessionid,i,1)+97)
	Next
End Function

Function CheckFormID(id)
	CheckFormID=false
	Dim i,Str
	For i=1 to Len(id)
		Str=Str & Asc(Mid(id,i,1))-97
	Next
	If Session.SessionID=Str Then
		CheckFormID=True
	End If
	'Response.Write(Session.SessionID&"***"&Str):Response.End
End Function
'下拉菜单转换输出
Function Chk_select(str1,str2)
	Dim k
	str2=Split(str2,",")
	If  str1="" Then chk_select="<option value="""" selected=""selected"">...</option>"
	For k=0 to ubound(str2)
		chk_select=chk_select & "<option value=""" & str2(k)&""""
		If str2(k)=str1 Then chk_select=chk_select &" selected=""selected"" "
		chk_select=chk_select & " >" & str2(k) &"</option>"
	Next
End Function

'多项选取转换输出
Function Chk_KidneyType(str0,str1,str2)
	Dim k
	str2=split(str2,",")
	For k = 0 to ubound(str2)	
		chk_KidneyType=chk_KidneyType+"<input type=""checkbox"" class=""chkbox"" name="""&str0&""" value="""&trim(str2(k))&""" "	 
		If instr(str1,trim(str2(k)))>0 Then '如果有此项性格
		chk_KidneyType=chk_KidneyType + "checked" 
		End If 
		chk_KidneyType=chk_KidneyType + ">"&trim(str2(k))&" "
	If ((k+1) mod 5)=0 Then chk_KidneyType=chk_KidneyType +  "<br>"  '每行显示六个性格进行换行
	Next
End Function

Function checktime(time_difference,time)
	Dim GMT,YGMT
	GMT=DateAdd("s",-(8*3600),Now())
	YGMT=DateAdd("s",time_difference*3600,GMT)
	checktime=( Hour(YGMT)=CLng(time))
End Function

Sub reg_3()

	If Dvbbs.forum_setting(78)="1" Then
		If Not Dvbbs.CodeIsTrue() Then
			'Response.write "验证码校验失败，请返回刷新页面后再输入验证码"
			Response.redirect "showerr.asp?ErrCodes=<li>验证码校验失败，请返回刷新页面后再输入验证码。&action=OtherErr"
		End If
	End If

	If Dvbbs.forum_setting(107)="1" Then
		Dim regask,n,CanReg
		regask=Split(Dvbbs.forum_setting(106),"!")
		CanReg = False

		For n=0 To UBound(regask)
			If Request.Form(md5(n,16))>"" Then
				If Trim(LCase(Request.Form(md5(n,16)))) <> Trim(LCase(regask(n))) Then
					Response.redirect "showerr.asp?ErrCodes=<li>注册答案错误，请返回刷新页面后重新输入，或者联系管理员。&action=OtherErr"
				Else
					CanReg = True
				End If				
				Exit For
			End If
		Next
		If Not CanReg Then
			Response.redirect "showerr.asp?ErrCodes=<li>注册答案不能为空，请返回刷新页面后重新输入，或者联系管理员。&action=OtherErr"
		End If
	End If
	
	Dim username,sex,pass1,pass2,password,FormID,fname
	Dim useremail,face,width,height
	Dim sign,showRe,birthday,UserIM
	Dim mailbody,sendmsg,rndnum,num1
	Dim question,answer,topic
	Dim userinfo,usersetting
	Dim userclass,UserJoinTime
	Dim rs,sql,i,TempLateStr
	Dim Qq
	'判断同一IP注册间隔时间
	If Not Isnull(Session("regtime")) Or Clng(Dvbbs.Forum_Setting(22)) > 0 Then
		If DateDiff("s",Session("regtime"),Now()) < Clng(Dvbbs.Forum_Setting(22)) Then
			ErrCodes = ErrCodes + "<li>" + Replace(Template.Strings(27), "{$RegTime}", Dvbbs.Forum_Setting(22))
			Exit Sub
		End If
	End If
	If Not Dvbbs.ChkPost() Then
		Dvbbs.AddErrCode(16)
		Exit sub
	End If
	
	If XMLDom.documentElement.selectSingleNode("@checktime").text = "1" Then
		If Trim(Request.form("time_difference"))="" Or Trim(Request.form("time"))="" Or Not IsNumeric(Trim(Request.form("time_difference"))) or Not IsNumeric(Trim(Request.form("time")))Then
			Response.redirect "showerr.asp?ErrCodes=<li>您必须选择时区和时间&action=OtherErr"
			Exit sub
		Else
			If not  checktime(Trim(Request.form("time_difference")),Trim(Request.form("time"))) Then
					Response.redirect "showerr.asp?ErrCodes=<li>您选择时区和时间不正确&action=OtherErr"
			End If
		End If
	End If
	
	FormID=GetFormID()
	If XMLDom.documentElement.selectSingleNode("@usevarform").text = "1" Then
		fname="_"&Md5(FormID,16)
	End If
	username=Request.form("username"&fname)
	username=replace(UserName,chr(255),"")
	If Trim(username)="" or strLength(username)>Cint(Dvbbs.Forum_Setting(41)) or strLength(username)<Cint(Dvbbs.Forum_Setting(40)) Then
		TempLateStr=template.Strings(28)
		TempLateStr=Replace(TempLateStr,"{$RegMaxLength}",Dvbbs.Forum_Setting(41))
		TempLateStr=Replace(TempLateStr,"{$RegLimLength}",Dvbbs.Forum_Setting(40))
		ErrCodes=ErrCodes+"<li>"+TempLateStr
		TempLateStr=""
		Exit Sub
	End If
	If XMLDom.documentElement.selectSingleNode("@checknumeric").text = "1" Then
		If IsNumeric(username) Then
			Response.redirect "showerr.asp?ErrCodes=<li>本论坛不接受全数字的用户名注册.&action=OtherErr"
		End If
	End If
	username=Dvbbs.CheckStr(username)
	If Instr(username,"=")>0 or Instr(username,"%")>0 or Instr(username,"?")>0 or Instr(username,"&")>0 or Instr(username,";")>0 or Instr(username,",")>0 or Instr(username,"'")>0 or Instr(username,",")>0 or Instr(username,chr(34))>0  or Instr(username,"