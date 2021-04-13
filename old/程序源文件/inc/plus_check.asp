<%
Dim Dv_plus
Set Dv_plus=new Cls_plus
Class Cls_plus
	'VBS类说明：属性Name用于提取设置，调用方法:实例.namr=插件ID
	'变量Mian_setting数组，存储插件常规设置,
	'变量plus_Settings数组，存储插件自定义扩展设置
	'变量plus_Settingnames数组，存储插件自定义扩展设置的定义的
	'变量plus_Copyright存储插件版权信息。
	'变量Plus_Name存储插件在菜单中显示的名称。
	'plus_master存储是否是插件管理员的布尔值
	'方法checklogin()验证使用插件的权限，判断插件管理员的身份。
	'方法updateuser()更新使用插件后的用户的数据，如金钱，积分等的更新。
	Public Mian_settings,plus_Settings,plus_Copyright,Plus_Name,plus_master,plus_Settingnames
	Public Property Let Name(ByVal vNewValue)
		Call GetPlus_Setting(vNewValue)
	End Property
	'验证使用插件的权限，判断插件管理员的身份。
	Public Sub checklogin()
		plus_master=False
		If Dvbbs.UserID>0 Then 
			If Dvbbs.master Then
				plus_master=True
			Else
				Dim masterlist
				If Trim(Mian_settings(3))<>"" Then 
					masterlist="|"&Mian_settings(3)&"|"
					If InStr(masterlist,"|"&Dvbbs.MemberName&"|")>0 Then
						plus_master=True
					End If
				End If 
			End If
		End If 
		If Not plus_master Then
			If Mian_settings(0)="1" Then
				Dim Otime
				Otime=split(Mian_settings(1),"|")
				If UBound(Otime)=1 Then 
			   		If IsNumeric(Otime(0)) And IsNumeric(OTime(1)) Then
						If CInt(OTime(0))< CInt(OTime(1)) Then
							If Hour(Now) < Cint(Otime(0)) or Hour(Now) > Cint(Otime(1)) Then
								Response.redirect "showerr.asp?ErrCodes=<li>"&Plus_Name&"<B>"&OTime(0)&"</B>至<B>"&OTime(1)&"</B>点开放，请在规定时间内使用，谢谢。&action=plus"
							End If
						Else
							If Hour(Now)< Cint(OTime(0)) And Hour(Now) > Cint(OTime(1)) Then
								Response.redirect "showerr.asp?ErrCodes=<li>"&Plus_Name&"<B>"&OTime(0)&"</B>至<B>"&OTime(1)&"</B>点开放，请在规定时间内使用，谢谢。&action=plus"
							End If
						End If 
					End If
				End If 			
			End If
			Dim UserGroupIDlist
			UserGroupIDlist="@"&Mian_settings(2)&"@"
			If Not InStr(UserGroupIDlist,"@"&Dvbbs.UserGroupID&"@")>0 Then
				Response.redirect "showerr.asp?ErrCodes=<li>您没有权限进入"&Plus_Name&"&action=plus"
			End If
			'检查使用插件的限制,如果设置了客人组可使用，此判断忽略。
			If Not InStr(UserGroupIDlist,"@7@")>0 Then
				Dim Plus_UserPost,Plus_userWealth,Plus_UserEP,Plus_UserCP,Plus_UserPower
				Plus_UserPost=Mian_settings(4)
				Plus_userWealth=Mian_settings(5)
				Plus_UserEP=Mian_settings(6)
				Plus_UserCP=Mian_settings(7)
				Plus_UserPower=Mian_settings(8)	
				If IsNumeric(Plus_UserPost) Then 
					If CLng(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpost").text)< CLng(Plus_UserPost) Then
						Response.redirect "showerr.asp?ErrCodes=<li>使用"&Plus_Name&"的最少发贴数是"&Plus_UserPost&",您只有" & Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpost").text & "篇&action=plus"
					End If
				End If
				If IsNumeric(Plus_userWealth) Then 
					If CLng(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text)< CLng(Plus_userWealth) Then
						Response.redirect "showerr.asp?ErrCodes=<li>使用"&Plus_Name&"的最少金钱"&Plus_userWealth&",您只有"& Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text &"&action=plus"
					End If
				End If
				If IsNumeric(Plus_UserEP) Then 
					If CLng(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text)< CLng(Plus_UserEP) Then
						Response.redirect "showerr.asp?ErrCodes=<li>使用"&Plus_Name&"的最少积分是"&Plus_UserEP&",您只有"& Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text &"&action=plus"
					End If
				End If
				If IsNumeric(Plus_UserCP) Then 
					If CLng(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text)< CLng(Plus_UserCP) Then
						Response.redirect "showerr.asp?ErrCodes=<li>使用"&Plus_Name&"的最少魅力是"&Plus_UserCP&",您只有"& Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text &"&action=plus"
					End If
				End If
				If IsNumeric(Plus_UserPower) Then 
					If Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpower").text = "" Then Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpower").text = 0
					If CLng(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpower").text) < CLng(Plus_UserPower) Then
						Response.redirect "showerr.asp?ErrCodes=<li>使用"&Plus_Name&"的最少魅力是"&Plus_UserPower&",您只有"& Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpower").text &"&action=plus"
					End If
				End If
			End If
		End If
	End Sub
	'使用插件后更新用户数据部分
	Public Sub updateuser()
		If Dvbbs.UserID>0 Then
			Dim ADDuserWealth,ADDUserEP,ADDUserCP,ADDUserPower
			ADDuserWealth=Mian_settings(9)
			ADDUserEP=Mian_settings(10)
			ADDUserCP=Mian_settings(11)
			ADDUserPower=Mian_settings(12)
			If Not IsNumeric(ADDuserWealth) Then ADDuserWealth=0
			If Not IsNumeric(ADDUserEP) Then ADDUserEP=0
			If Not IsNumeric(ADDUserCP) Then ADDUserCP=0
			If Not IsNumeric(ADDUserPower) Then ADDUserPower=0
			ADDuserWealth=CLng(ADDuserWealth)
			ADDUserEP=CLng(ADDUserEP)
			ADDUserCP=CLng(ADDUserCP)
			ADDUserPower=CLng(ADDUserPower)
			If ADDuserWealth<>0 Or ADDUserEP<>0  Or ADDUserCP <>0 Or ADDUserPower<>0 Then
				Dvbbs.Execute("Update Dv_user Set userWealth=userWealth+("&ADDuserWealth&"),UserEP=UserEP+("&ADDUserEP&"),UserCP=UserCP+("&ADDUserCP&"),UserPower=UserPower+("&ADDUserPower&") Where userID="&Dvbbs.userID&"")
				Dvbbs.TrueCheckUserLogin
			End If
		End If 
	End Sub 
	'---------------取得插件设置数据
	Public Sub GetPlus_Setting(plus_ID)
	plus_Settingnames="设置字段1,设置字段2,设置字段3,设置字段4,设置字段5,设置字段6,设置字段7,设置字段8,设置字段9,设置字段10,设置字段11,设置字段12,设置字段13,设置字段14,设置字段15,设置字段16,设置字段17,设置字段18,设置字段19"
	plus_Settings="0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
	Mian_settings="0,0|24,1@2@3@4@5@6@7@8,,0,0,0,0,0,0,0,0,0"
	plus_Copyright="dvbbs"
	Plus_Name="未知插件"
	Dim SettingDatas
	SettingDatas=Plus_Setting()
		If IsArray(SettingDatas) Then		
			Dim i,SettingData
			For i=0 to UBound(SettingDatas,2)
				If CStr(LCase(SettingDatas(0,i)))=CStr(LCase(plus_ID)) Then 	
					SettingData=SettingDatas(1,i)
					Plus_Name=SettingDatas(2,i)
					plus_Copyright=SettingDatas(3,i)
					Dvbbs.Forum_Copyright=Dvbbs.Forum_Copyright&"<br><b>"&Plus_Name&"</b> 由"&plus_Copyright&"开发"
					Exit For
				End If
			Next
			If SettingData<>"" Then
				SettingData=Split(SettingData,"|||")
				If UBound(SettingData)>1 Then
					plus_Settings=SettingData(1)
					plus_Settingnames=SettingData(2)
					Mian_settings=SettingData(3)
				End If
			End If
		End If
		plus_Settingnames=Split(plus_Settingnames,",")
		plus_Settings=Split(plus_Settings,",")
		Mian_settings=Split(Mian_settings,",")
	End Sub
	Public Function Plus_Setting()
		Dvbbs.Name="Plus_Settingts"
		If Dvbbs.ObjIsEmpty() Then
			Dim Rs,SQL
			SQL = "select plus_ID,Plus_Setting,Plus_Name,plus_Copyright from [Dv_plus] Order By ID"
			Set Rs = Dvbbs.Execute(SQL)	
			If Not Rs.Eof Then 
				Dvbbs.value = Rs.GetRows(-1)
			Else
				Dvbbs.value=""
			End If
			Set Rs = Nothing
		End If
		Plus_Setting=Dvbbs.Value
	End Function 
	'--------------------
End Class
%>