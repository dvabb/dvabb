<!--#include file =../conn.asp-->
<!--#include file="inc/const.asp"-->
<%
Head()
Dim admin_flag
admin_flag=",19,"
CheckAdmin(admin_flag)
Main_head()
Select Case Request("action")
	Case "SendMoney" : SendMoney
	Case Else
		SendForm
End Select
If ErrMsg<>"" Then Dvbbs_Error
If founderr then call dvbbs_error()
footer()

'顶部说明及注意事项
Sub Main_head

End Sub

'相关设置
Sub SendForm
Dim Rs
%>
<form METHOD=POST ACTION="?action=SendMoney">
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr><th style="text-align:center;" colspan="4">奖励赠送设置</th></tr>
<tr>
<td class="td1" align=right width="15%"><U>赠送金币</U>：</td>
<td class="td1" width="40%">
<INPUT TYPE="text" NAME="SendMoney" size=10 onkeyup="CheckNumer(this.value,this,'')">
<INPUT TYPE="radio" class="radio" class= NAME="SendMoneyType" checked value="0">增加 <INPUT TYPE="radio" class="radio" NAME="SendMoneyType" value="1">减少 <INPUT TYPE="radio" class="radio" NAME="SendMoneyType" value="2">更新
</td>
<td class="td1" width="10%"><INPUT TYPE="checkbox" class="checkbox" NAME="SelectType" value="SendMoney">选取</td>
<td class="td2" width="*" rowspan="6" valign=top>
<li>请正确填写相关数值；
<li>选取后该项更新才能生效；
<li>若选取更新，则目标用户相关数据将更新为该设置；
</td>
</tr>
<tr>
<td class="td1" align=right><U>赠送点券</U>：</td>
<td class="td1"><INPUT TYPE="text" NAME="SendTicket" size=10 onkeyup="CheckNumer(this.value,this,'')">
<INPUT TYPE="radio" class="radio" NAME="SendTicketType" checked value="0">增加 <INPUT TYPE="radio" class="radio" NAME="SendTicketType" value="1">减少 <INPUT TYPE="radio" class="radio" NAME="SendTicketType" value="2">更新
</td>
<td class="td1"><INPUT TYPE="checkbox" class="checkbox" NAME="SelectType" value="SendTicket">选取</td>
</tr>
<tr>
<td class="td1" align=right><U>赠送积分</U>：</td>
<td class="td1"><INPUT TYPE="text" NAME="SendUserEP" size=10 onkeyup="CheckNumer(this.value,this,'')">
<INPUT TYPE="radio" class="radio" NAME="SendUserEPType" checked value="0">增加 <INPUT TYPE="radio" class="radio" NAME="SendUserEPType" value="1">减少 <INPUT TYPE="radio" class="radio" NAME="SendUserEPType" value="2">更新
</td>
<td class="td1"><INPUT TYPE="checkbox" class="checkbox" NAME="SelectType" value="SendUserEP">选取</td>
</tr>
<tr>
<td class="td1" align=right><U>赠送魅力</U>：</td>
<td class="td1"><INPUT TYPE="text" NAME="SendUserCP" size=10 onkeyup="CheckNumer(this.value,this,'')">
<INPUT TYPE="radio" class="radio" NAME="SendUserCPType" checked value="0">增加 <INPUT TYPE="radio" class="radio" NAME="SendUserCPType" value="1">减少 <INPUT TYPE="radio" class="radio" NAME="SendUserCPType" value="2">更新
</td>
<td class="td1"><INPUT TYPE="checkbox" class="checkbox" NAME="SelectType" value="SendUserCP">选取</td>
</tr>
<tr>
<td class="td1" align=right><U>赠送金钱</U>：</td>
<td class="td1"><INPUT TYPE="text" NAME="SendUserWealth" size=10 onkeyup="CheckNumer(this.value,this,'')">
<INPUT TYPE="radio" class="radio" NAME="SendUserWealthType" checked value="0">增加 <INPUT TYPE="radio" class="radio" NAME="SendUserWealthType" value="1">减少 <INPUT TYPE="radio" class="radio" NAME="SendUserWealthType" value="2">更新
</td>
<td class="td1"><INPUT TYPE="checkbox" class="checkbox" NAME="SelectType" value="SendUserWealth">选取</td>
</tr>
<tr>
<td class="td1" align=right><U>赠送威望</U>：</td>
<td class="td1"><INPUT TYPE="text" NAME="SendUserPower" size=10 onkeyup="CheckNumer(this.value,this,'')">
<INPUT TYPE="radio" class="radio" NAME="SendUserPowerType" checked value="0">增加 <INPUT TYPE="radio" class="radio" NAME="SendUserPowerType" value="1">减少 <INPUT TYPE="radio" class="radio" NAME="SendUserPowerType" value="2">更新
</td>
<td class="td1"><INPUT TYPE="checkbox" class="checkbox" NAME="SelectType" value="SendUserPower">选取</td>
</tr>
<tr><th style="text-align:center;" colspan="4">奖励赠送目标</th></tr>
<tr><td class="td2" height=20 colspan="4">
<INPUT TYPE="radio" class="radio" NAME="Sendtype" value="0" onclick="formstep(0)">按指定用户
<INPUT TYPE="radio" class="radio" NAME="Sendtype" value="1" onclick="formstep(1)">按指定用户组
<INPUT TYPE="radio" class="radio" NAME="Sendtype" value="2" onclick="formstep(2)">按所有用户
</td></tr>
</table>
<div id="ToUser" style="display:none;">
	<br>
	<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
	<tr><th style="text-align:center;" colspan="2">指定用户</th></tr>
	<tr><td height=20 colspan="2">用户名以英文逗号“,”分隔；为节省资源，每次更新限制10位用户。注意区分大小写。</td></tr>
	<td class="td1"><u>用户名单</u>：</td>
	<td class="td1"><INPUT TYPE="text" NAME="ToUserName" size="80"></td>
	</table>
</div>
<div id="ToUserGroup" style="display:none;">
	<br>
	<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
	<tr><th style="text-align:center;">指定用户组</th></tr>
	<tr><td>
	<li>请选取指定更新的用户组<LI>若只对某部分用户组更新，请不要选取所有用户。<br>
	<%
	Set Rs=DvBBS.Execute("Select UserGroupID,Title,UserTitle,parentgid From Dv_UserGroups where parentgid>0  Order By parentgid,UserGroupID")
	Do while not Rs.eof
		Response.Write "&nbsp;&nbsp;<INPUT TYPE=""checkbox"" class=""checkbox"" NAME=""GetGroupID"" value="""&Rs(0)&""">"
		Response.Write Rs(2)
	Rs.movenext
	Loop
	Rs.close
	Set Rs=Nothing
	%>
	</td></tr>
	<tr><td height=20 class="td2" ><input type="button" class="button" value="打开高级设置" NAME="OPENSET" onclick="openset(this,'UpSetting')"></td></tr>
	<tr><td height=20 ID="UpSetting" style="display:NONE" class="td2">
		<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
		<tr><th style="text-align:center;" colspan="4">符合条件设置</th></tr>
		<tr>
		<td class="td1" width="15%">最后登陆时间：</td>
		<td class="td1" width="35%">
		<input type="text" name="LoginTime" onkeyup="CheckNumer(this.value,this,'')" size=6>天 &nbsp;<INPUT TYPE="radio" class="radio" NAME="LoginTimeType" checked value="0">多于 <INPUT TYPE="radio" class="radio" NAME="LoginTimeType" value="1">少于
		</td>
		<td class="td1" width="15%">注册时间：</td>
		<td class="td1" width="35%">
		<input type="text" name="RegTime" onkeyup="CheckNumer(this.value,this,'')" size=6>天 &nbsp;<INPUT TYPE="radio" class="radio" NAME="RegTimeType" checked value="0">多于 <INPUT TYPE="radio" class="radio" NAME="RegTimeType" value="1">少于
		</td>
		</tr>
		<tr>
		<td class="td1">登陆次数：</td>
		<td class="td1"><input type="text" name="Logins" size=6 onkeyup="CheckNumer(this.value,this,'')">次 &nbsp;<INPUT TYPE="radio" class="radio" NAME="LoginsType" checked value="0">多于 <INPUT TYPE="radio" class="radio" NAME="LoginsType" value="1">少于
		</td>
		<td class="td1">发表文章：</td>
		<td class="td1"><input type="text" name="UserPost" size=6 onkeyup="CheckNumer(this.value,this,'')">篇 &nbsp;<INPUT TYPE="radio" class="radio" NAME="UserPostType" checked value="0">多于 <INPUT TYPE="radio" class="radio" NAME="UserPostType" value="1">少于</td>
		</tr>
		<tr>
		<td class="td1">主题文章：</td>
		<td class="td1"><input type="text" name="UserTopic" size=6 onkeyup="CheckNumer(this.value,this,'')">篇 &nbsp;<INPUT TYPE="radio" class="radio" NAME="UserTopicType" checked value="0">多于 <INPUT TYPE="radio" class="radio" NAME="UserTopicType" value="1">少于</td>
		<td class="td1">精华文章：</td>
		<td class="td1"><input type="text" name="UserBest" size=6 onkeyup="CheckNumer(this.value,this,'')">篇 &nbsp;<INPUT TYPE="radio" class="radio" NAME="UserBestType" checked value="0">多于 <INPUT TYPE="radio" class="radio" NAME="UserBestType" value="1">少于
		</td>
		</tr>
		</table>
	</td></tr>
	</table>
</div>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
	<tr><td height=20 align=center><input type="submit" class="button" value="执行更新"></td></tr>
</table>
<form>
<SCRIPT LANGUAGE="JavaScript">
<!--
function openset(v,s){
	if (v.value=='打开高级设置'){
		document.getElementById(s).style.display = "";
		v.value="关闭高级设置";
	}
	else{
		v.value="打开高级设置";
		document.getElementById(s).style.display = "none";
	}
}
//验证表单数值 n:number 表单值 | v:value object 表单对象 | n_max 最大值
function CheckNumer(n,v,n_max)
{
	if (isNaN(n)){
		v.value = "";
		alert("请填写正确的数值！");
	}
	else{
		n = parseInt(n);
		if (!isNaN(n_max)){
			n_max = parseInt(n_max);
			if (n>n_max){v.value = "";alert("该项数值不能高于："+n_max);}
		}
	}
}

function formstep(OpenID){
	var ToUser = document.getElementById("ToUser");
	var ToUserGroup = document.getElementById("ToUserGroup");
	if (OpenID==0){
	ToUser.style.display = "";
	ToUserGroup.style.display = "none";
	}
	else if (OpenID==1){
	ToUser.style.display = "none";
	ToUserGroup.style.display = "";
	}
	else{
	ToUser.style.display = "none";
	ToUserGroup.style.display = "none";
	}
}
//-->
</SCRIPT>
<%
End Sub

'保存更新设置
Sub SendMoney
	Dim SelectType,UPString,TempData
	SelectType = Replace(Request.Form("SelectType"),chr(32),"")
	If SelectType="" Then
		ErrMsg = "请选取奖励设置项!"
		Exit Sub
	End If
	SelectType = ","&SelectType&","
	UPString = ""
	'更新金币
	If Instr(SelectType,"SendMoney") Then
		UPString = GetUPString(Request.Form("SendMoney"),UPString,Request.Form("SendMoneyType"),"UserMoney")
	End If
	'更新点券
	If Instr(SelectType,"SendTicket") Then
		UPString = GetUPString(Request.Form("SendTicket"),UPString,Request.Form("SendTicketType"),"UserTicket")
	End If
	'更新积分
	If Instr(SelectType,"SendUserEP") Then
		UPString = GetUPString(Request.Form("SendUserEP"),UPString,Request.Form("SendUserEPType"),"UserEP")
	End If
	'更新魅力
	If Instr(SelectType,"SendUserCP") Then
		UPString = GetUPString(Request.Form("SendUserCP"),UPString,Request.Form("SendUserCPType"),"UserCP")
	End If
	'更新金钱
	If Instr(SelectType,"SendUserWealth") Then
		UPString = GetUPString(Request.Form("SendUserWealth"),UPString,Request.Form("SendUserWealthType"),"UserWealth")
	End If
	'更新威望
	If Instr(SelectType,"SendUserPower") Then
		UPString = GetUPString(Request.Form("SendUserPower"),UPString,Request.Form("SendUserPowerType"),"UserPower")
	End If
	'Response.Write UPString
	Select Case Request.Form("Sendtype")
		Case "0" : Call Sendtype_0(UPString)	'按指定用户
		Case "1" : Call Sendtype_1(UPString)	'按指定用户组
		Case "2" : Call Sendtype_2(UPString)	'按所有用户
		Case Else
			ErrMsg = "请选取奖励赠送目标!"
			Exit Sub
	End Select
End Sub

'按指定用户
Sub Sendtype_0(Str)
	Dim ToUserName,Rs,Sql,i,ToUserID
	ToUserName = Trim(Request.Form("ToUserName"))
	If ToUserName = "" Then ErrMsg = "请填写目标用户名，注意区分大小写。" : Exit Sub
	ToUserName = Replace(ToUserName,"'","")
	ToUserName = Split(ToUserName,",")
	If Ubound(ToUserName)>10 Then ErrMsg = "限制一次不能超过10位目标用户。" : Exit Sub
	For i=0 To Ubound(ToUserName)
		SQL = "Select UserID From [Dv_user] Where UserName = '"&ToUserName(i)&"'"
		SET Rs = Dvbbs.Execute(SQL)
		If Not Rs.eof Then
			If i=0 or ToUserID="" Then
				ToUserID = ToUserID & Rs(0)
			Else
				ToUserID = ToUserID &","& Rs(0)
			End If
		Else
			ErrMsg = "目标用户不存在，注意区分大小写。" : Exit Sub
		End If
	Next
	Rs.Close : Set Rs = Nothing
	If ToUserID<>"" Then
		SQL = "Update [Dv_user] Set "&Dvbbs.Checkstr(Str)&" where UserID in ("&ToUserID&") "
		Dvbbs.Execute(SQL)
		Dv_suc("共位"&Ubound(ToUserName)+1&"目标会员更新成功!")
	Else
		ErrMsg = "目标用户不存在，注意区分大小写。" : Exit Sub
	End If
End Sub

'按指定用户组
Sub Sendtype_1(Str)
	Dim GetGroupID
	Dim SearchStr,TempValue,DayStr
	GetGroupID = Replace(Request.Form("GetGroupID"),chr(32),"")
	If GetGroupID="" or Not Isnumeric(Replace(GetGroupID,",","")) Then
		ErrMsg = "请正确选取相应的用户组。" : Exit Sub
	Else
		GetGroupID = Dvbbs.Checkstr(GetGroupID)
	End If
	If IsSqlDataBase=1 Then
		DayStr = "d"
	Else
		DayStr = "'d'"
	End If
	If Instr(GetGroupID,"-1") Then
		SearchStr = ""
	Else
		If Instr(GetGroupID,",")=0 Then
			SearchStr = "UserGroupID = "&GetGroupID
		Else
			SearchStr = "UserGroupID in ("&GetGroupID&")"
		End If
	End If
	'登陆次数
	TempValue = Request.Form("Logins")
	If TempValue<>"" and IsNumeric(TempValue) Then
		SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("LoginsType"),"UserLogins")
	End If
	'发表文章
	TempValue = Request.Form("UserPost")
	If TempValue<>"" and IsNumeric(TempValue) Then
		SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("UserPostType"),"UserPost")
	End If
	'主题文章
	TempValue = Request.Form("UserTopic")
	If TempValue<>"" and IsNumeric(TempValue) Then
		SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("UserTopicType"),"UserTopic")
	End If
	'精华文章
	TempValue = Request.Form("UserBest")
	If TempValue<>"" and IsNumeric(TempValue) Then
		SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("UserBestType"),"UserIsBest")
	End If
	'最后登陆时间
	TempValue = Request.Form("LoginTime")
	If TempValue<>"" and IsNumeric(TempValue) Then
		SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("LoginTimeType"),"Datediff("&DayStr&",Lastlogin,"&SqlNowString&")")
	End If
	'注册时间
	TempValue = Request.Form("RegTime")
	If TempValue<>"" and IsNumeric(TempValue) Then
		SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("RegTimeType"),"Datediff("&DayStr&",JoinDate,"&SqlNowString&")")
	End If

	Dim SQL
	SQL = "Update [Dv_user] Set "&Dvbbs.Checkstr(Str)&" Where "&SearchStr
	Dvbbs.Execute(SQL)
	Dv_suc("目标会员更新成功!")
End Sub

'按所有用户
Sub Sendtype_2(Str)
	Dim sql
	SQL = "Update [Dv_user] Set "& Dvbbs.Checkstr(Str)
	Dvbbs.Execute(SQL)
	Dv_suc("所有会员更新成功!")
End Sub

Function GetSearchString(Get_Value,Get_SearchStr,UpType,UpColumn)
	Get_Value = Clng(Get_Value)
	If Get_SearchStr<>"" Then Get_SearchStr = Get_SearchStr & " and " 
	If UpType="1" Then
		Get_SearchStr = Get_SearchStr & UpColumn &" <= "&Get_Value
	Else
		Get_SearchStr = Get_SearchStr & UpColumn &" >= "&Get_Value
	End If
	GetSearchString = Get_SearchStr
End Function

Function GetUPString(TempData,UPString,UpType,UpColumn)
	If TempData<>"" and IsNumeric(TempData) Then
			If UPString<>"" Then UPString = UPString & ","
			Select Case UpType
				Case "2" : UPString = UPString &" "&UpColumn&" = "&cCur(TempData)
				Case "1" : UPString = UPString &" "&UpColumn&" = "&UpColumn&"-"&cCur(TempData)
				Case Else : UPString = UPString &" "&UpColumn&" = "&UpColumn&"+"&cCur(TempData)
			End Select
			GetUPString = UPString
	End If
End Function
%>