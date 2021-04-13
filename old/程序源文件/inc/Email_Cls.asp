<%
'-----------------------------------------------------------------------
'--- EMAIL邮件处理类模块
'--- Copyright (c) 2004 Aspsky, Inc.
'--- Mail: Sunwin@artbbs.net   http://www.aspsky.net
'--- 2004-12-18
'-----------------------------------------------------------------------
'--- 设置项
'-----------------------------------------------------------------------
'--- ServerLoginName	设置您的邮件服务器登录名
'--- ServerLoginPass	设置登录密码
'--- SendSMTP			设置SMTP邮件服务器地址
'--- SendFromEmail		设置发件人的E-MAIL地址
'--- SendFromName		设置发送人名称
'--- ContentType		设置邮件类型 默认：text/html
'--- CharsetType		设置编码类型 默认：gb2312
'--- SendObject			设置选取组件 1=Jmail,2=Cdonts,3=Aspemail
'-----------------------------------------------------------------------
'--- 属性
'-----------------------------------------------------------------------
'--- SendMail Email, Topic, MailBody	收件人地址，标题，邮件内容
'-----------------------------------------------------------------------
'--- 获取信息
'-----------------------------------------------------------------------
'--- ErrCode			信息编号 0=正常
'--- Description		相应操作信息
'--- Count				发送邮件数
'-----------------------------------------------------------------------
Class Dv_SendMail
	Public Count,ErrCode,ErrMsg
	Private LoginName,LoginPass,SMTP,FromEmail,FromName,Object,Content_Type,Charset_Type
	Private Obj,cdoConfig

	Private Sub Class_Initialize()
		Object = 0
		Count = 0
		ErrCode = 0
		Content_Type = "text/html"
		Charset_Type = "gb2312"
	End Sub

	Private Sub Class_Terminate()
		If Isobject(Obj) Then
			Set Obj = Nothing
		End If
		If IsObject(cdoConfig) Then
			Set cdoConfig = Nothing
		End If
	End Sub

	'设置您的邮件服务器登录名
	Public Property Let ServerLoginName(Byval Value)
		LoginName = Value
	End Property

	'设置登录密码
	Public Property Let ServerLoginPass(Byval Value)
		LoginPass = Value
	End Property
	'设置SMTP邮件服务器地址
	Public Property Let SendSMTP(Byval Value)
		SMTP = Value
	End Property
	'设置发件人的E-MAIL地址
	Public Property Let SendFromEmail(Byval Value)
		FromEmail = Value
	End Property
	'设置发送人名称
	Public Property Let SendFromName(Byval Value)
		FromName = Value
	End Property
	'设置邮件类型
	Public Property Let ContentType(Byval Value)
		Content_Type = Value
	End Property
	'设置编码类型
	Public Property Let CharsetType(Byval Value)
		Charset_Type = Cstr(Value)
	End Property
	'获取错误信息
	Public Property Get Description()
		Description = ErrMsg
	End Property
	'设置选取组件 SendObject 0=Jmail,1=Cdonts,2=Aspemail
	Public Property Let SendObject(Byval Value)
		Object = Value
		On Error Resume Next
		Select Case Object
			Case 1
				Set Obj = Dvbbs.iCreateObject("JMail.Message")
			Case 2
				Set Obj = Dvbbs.iCreateObject("CDONTS.NewMail")
			Case 3
				Set Obj = Dvbbs.iCreateObject("Persits.MailSender")
			Case 4
				Set Obj = Dvbbs.iCreateObject("CDO.Message")	'window 2003 new SendMailCom Object
			Case Else
				ErrNumber = 2
		End Select
		If Err<>0 Then
			ErrNumber = 3
		End If
	End Property

	Private Property Let ErrNumber(Byval Value)
		ErrCode = Value
		ErrMsg = ErrMsg & Msg
	End Property
	Private Function Msg()
		Dim MsgValue
		Select Case ErrCode
		Case 1
			MsgValue = "未选取邮件组件或服务器不支持该组件！"
		Case 2
			MsgValue = "所选的组件不存在！"
		Case 3
			MsgValue = "错误：服务器不支持该组件!"
		Case 4
			MsgValue = "发送失败!"
		Case Else
			MsgValue = "正常。"
		End Select
		Msg = MsgValue
	End Function

	Public Sub SendMail(Byval Email,Byval Topic,Byval MailBody)
		If ErrCode <> 0 Then
			Exit Sub
		End If
		If Email="" or ISNull(Email) Then Exit Sub
		If Object>0 Then
			Select Case Object
				Case 1
					Jmail Email,Topic,MailBody
				Case 2
					Cdonts Email,Topic,Mailbody
				Case 3
					Aspemail Email,Topic,Mailbody
				Case 4
					CDOMessage Email,Topic,Mailbody
				Case Else
					ErrNumber = 2
			End Select
		Else
			ErrNumber = 1
		End If
	End Sub

	Private Sub Jmail(Email,Topic,Mailbody)
		On Error Resume Next
		Obj.Silent = True
		Obj.Logging = True
		Obj.Charset = Charset_Type
		If Not(LoginName = "" Or LoginPass = "") Then
			Obj.MailServerUserName = LoginName '您的邮件服务器登录名
			Obj.MailServerPassword = LoginPass '登录密码
		End If
		Obj.ContentType = Content_Type
		Obj.Priority = 1
		Obj.From = FromEmail
		Obj.FromName = FromName
		Obj.AddRecipient Email
		Obj.Subject = Topic
		Obj.Body = Mailbody
		If Err<>0 Then
			ErrMsg = ErrMsg & "发送失败!原因：" & Err.Description
			ErrNumber = 4
		Else
			Obj.Send (SMTP)
			Obj.ClearRecipients()
			If Err<>0 Then
				ErrMsg = ErrMsg & "发送失败!原因：" & Err.Description
				ErrNumber = 4
			Else
				Count = Count + 1
				ErrMsg = ErrMsg & "发送成功!"
			End If
		End If
	End Sub
		
	Private Sub Cdonts(Email,Topic,Mailbody)
		On Error Resume Next
		Obj.From = FromEmail
		Obj.To = Email
		Obj.Subject = Topic
		Obj.BodyFormat = 0 
		Obj.MailFormat = 0 
		Obj.Body = Mailbody
		If Err<>0 Then
			ErrMsg = ErrMsg & "发送失败!原因：" & Err.Description
			ErrNumber = 4
		Else
			Obj.Send
			If Err<>0 Then
				ErrMsg = ErrMsg & "发送失败!原因：" & Err.Description
				ErrNumber = 4
			Else
				Count = Count + 1
				ErrMsg = ErrMsg & "发送成功!"
			End If
		End If
	End Sub

	Private Sub Aspemail(Email,Topic,Mailbody)
		On Error Resume Next
		Obj.Charset = Charset_Type
		Obj.IsHTML = True
		Obj.username = LoginName	'服务器上有效的用户名
		Obj.password = LoginPass	'服务器上有效的密码
		Obj.Priority = 1
		Obj.Host = SMTP
		'Obj.Port = 25			' 该项可选.端口25是默认值
		Obj.From = FromEmail
		Obj.FromName = FromName	' 该项可选
		Obj.AddAddress Email,Email
		Obj.Subject = Topic
		Obj.Body = Mailbody
		If Err<>0 Then
			ErrMsg = ErrMsg & "发送失败!原因：" & Err.Description
			ErrNumber = 4
		Else
			Obj.Send
			If Err<>0 Then
				ErrMsg = ErrMsg & "发送失败!原因：" & Err.Description
				ErrNumber = 4
			Else
				Count = Count + 1
				ErrMsg = ErrMsg & "发送成功!"
			End If
		End If
	End Sub

	Private Sub CDOMessage(Email,Topic,Mailbody)
		On Error Resume Next
		If Not IsObject(cdoConfig) Then
			Call CreatCDOConfig()
		End If
		Set Obj = Dvbbs.iCreateObject("CDO.Message") 
		With Obj 
			Set .Configuration = cdoConfig 
			'.From = FromEmail
			.To = Email
			.Subject = Topic 
			.TextBody = Mailbody
			.Send
		End With
		If Err<>0 Then
			ErrMsg = ErrMsg & "发送失败!原因：" & Err.Description
			ErrNumber = 4
		Else
			Count = Count + 1
			ErrMsg = ErrMsg & "发送成功!"
		End If
	End Sub

	Private Sub CreatCDOConfig()
		On Error Resume Next
		Dim Sch
		sch = "http://schemas.microsoft.com/cdo/configuration/"
		Set cdoConfig = Dvbbs.iCreateObject("CDO.Configuration")
		With cdoConfig.Fields 
			.Item(sch & "smtpserver") = SMTP
			'.Item(sch & "smtpserverport") = 25
			.Item(sch & "sendusing") = 2					'cdoSendUsingPort CdoSendUsing enum value =  2
			.Item(sch & "smtpaccountname") = FromName		'"My Name"
			.Item(sch & "sendemailaddress") = FromEmail		'"""MySelf"" <example@example.com>"
			.Item(sch & "smtpuserreplyemailaddress") = 25	'"""Another"" <another@example.com>"
			'.Item(sch & "smtpauthenticate") = cdoBasic
			.Item(sch & "sendusername") = LoginName
			.Item(sch & "sendpassword") = LoginPass
			.update 
		End With
		If Err<>0 Then
			ErrMsg = ErrMsg & "发送失败!原因：" & Err.Description
			ErrNumber = 4
		End If
	End Sub
End Class
%>