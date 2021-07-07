<!--#include file="../conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="../inc/Email_Cls.asp"-->
<%
Head()
Dim Admin_flag
Admin_flag=",21,"
CheckAdmin(admin_flag)
Founderr=False

Dim XmlDom
Dim FilePath
Dim EmailTopic,EmailBody
FilePath = MyDbPath & "data/SendMailLog.config"
FilePath = Server.MapPath(FilePath)

Call Main()
Footer()

Sub Main()
%>
<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">
<tr><th colspan="2" style="text-align:center;">用户邮件通知</th></tr>
<tr>
<td width="20%" class="td1" align="center">
<button Style="width:80;height:50;border: 1px outset ;" class="button">注意事项</button>
</td>
<td width="80%" class="td2">
	①发送邮件列表只会保留最新十条记录；
	<br>②每次发送邮件请不要设置过多，要根据服务器的情况而定；
	<br>③邮件列表将保留发送的记录，还未发送完的可以在下一次执行发送；
	<br>④批量发送邮件，将会占用服务器资源，请尽量在访问量少的时间进行批量操作。
<!-- <br>⑤
	<br>⑥ -->
</td>
</tr>
<tr><td colspan="2" class="td2">
<a href="?">系统群发邮件</a> | <a href="?Act=ShowLog">群发邮件任务记录</a>
</td></tr>
</table>
<%
Select Case Request("Act")
	Case "sendemail" : Call SendStep2()
	Case "ShowLog" : Call ShowLog()
	Case "DelSendLog" : Call DelSendLog()
	Case "SendLog"	: Call SendLog()
	Case Else
	Call SendStep1
End Select
End Sub

'删除记录
Sub DelSendLog()
	Dim DelNodes,DelChildNodes
	Set XmlDom = Dvbbs.iCreateObject("MSXML.DOMDocument")
	If Not XmlDom.load(FilePath) Then
		ErrMsg = "邮件列表中为空，请填写发邮件后再执行本操作!"
		Dvbbs_Error()
		Exit Sub
	End If
	'Response.Write Request.Form("DelNodes").count
	For Each DelNodes in Request.Form("DelNodes")
		Set DelChildNodes = XmlDom.DocumentElement.selectSingleNode("SendLog[@AddTime='"&DelNodes&"']")
		If Not (DelChildNodes is nothing) Then
			XmlDom.DocumentElement.RemoveChild(DelChildNodes)
		End If
	Next
	XmlDom.save FilePath
	Set XmlDom=Nothing
	Dv_suc("所选的记录已删除!")
End Sub

'根据记录发送邮件
Sub SendLog()
	Dim SelNodes,SelChildNodes,SendOrders
	SelNodes = Trim(Request.Form("DelNodes"))
	SendOrders = Trim(Request.Form("SendOrders"))
	If SendOrders="" or Not IsNumeric(SendOrders) Then
		ErrMsg = "请填写每次发送邮件的记录数!"
		Dvbbs_Error()
		Exit Sub
	Else
		SendOrders = Clng(SendOrders)
	End If
	Set XmlDom = Dvbbs.iCreateObject("MSXML.DOMDocument")
	If Not XmlDom.load(FilePath) Then
		ErrMsg = "邮件列表中为空，请填写发邮件后再执行本操作!"
		Dvbbs_Error()
		Exit Sub
	End If
	Set SelChildNodes = XmlDom.DocumentElement.selectSingleNode("SendLog[@AddTime='"&SelNodes&"']")
	If SelChildNodes is nothing Then
		ErrMsg = "发送的记录不存在，请填写发邮件后再执行本操作!"
		Dvbbs_Error()
		Exit Sub
	End If

	Dim EmailTopic,EmailBody,Total,SearchStr,LastUserID,Remain
	Dim Sql,Rs,i,ii
	Total = SelChildNodes.getAttribute("Total")
	Remain = SelChildNodes.getAttribute("Remain")
	EmailTopic = SelChildNodes.selectSingleNode("EmailTopic").text
	EmailBody = SelChildNodes.selectSingleNode("EmailBody").text
	EmailBody = Replace(EmailBody, CHR(10) & CHR(10), "</P><P> ")
	EmailBody = Replace(EmailBody, CHR(10), "<br /> ")
	SearchStr = SelChildNodes.selectSingleNode("Search").text
	LastUserID = Int(SelChildNodes.getAttribute("LasterUserID"))
	If Remain="0" Then
		ErrMsg = "已经发送完毕!"
		Dvbbs_Error()
		Exit Sub
	End If
	SQL = "Select Top "&SendOrders&" UserID,UserName,UserEmail From Dv_User where UserID>= " & LastUserID
	If SearchStr<>"" Then
		SQL = SQL &" and "& SearchStr
	End If
	SQL = SQL & " order by UserID "
	SET Rs = Dvbbs.Execute(SQL)
	If Not Rs.eof Then
		SQL=Rs.GetRows(-1)
		Rs.close:Set Rs = Nothing
	Else
		ErrMsg = "已经发送完毕!"
		Dvbbs_Error()
		Exit Sub
	End If
	%>
	<table cellpadding="0" cellspacing="0" border="0" width="95%" class="tableBorder" align=center>
	<tr><td colspan=2 class=td1>
	下面开始发送邮件给目标用户，总共发送<%=Total%>封，目前剩余发送<%=Remain%>封，每次发送最限为<%=SendOrders%>封。
	<table width="400" border="0" cellspacing="1" cellpadding="1">
	<tr> 
	<td bgcolor=000000>
	<table width="400" border="0" cellspacing="0" cellpadding="1">
	<tr><td bgcolor=ffffff height=9><img src="../skins/default/bar/bar3.gif" width=0 height=16 id=img2 name=img2 align=absmiddle></td></tr></table>
	</td></tr></table>
	<span id=txt2 name=txt2 style="font-size:9pt">0</span><span style="font-size:9pt">%</span></td></tr>
	</table>
	<table cellpadding="0" cellspacing="0" border="0" width="95%" class="tableBorder" align=center>
	<tr><td colspan=2 class=td1>
	<span id=txt3 name=txt3 style="font-size:9pt">
	</span>
	</td></tr></table>
	<%
	Dim DvEmail
	Set DvEmail = New Dv_SendMail
	DvEmail.SendObject = Cint(Dvbbs.Forum_Setting(2))	'设置选取组件 1=Jmail,2=Cdonts,3=Aspemail
	DvEmail.ServerLoginName = Dvbbs.Forum_info(12)	'您的邮件服务器登录名
	DvEmail.ServerLoginPass = Dvbbs.Forum_info(13)	'登录密码
	DvEmail.SendSMTP = Dvbbs.Forum_info(4)			'SMTP地址
	DvEmail.SendFromEmail = Dvbbs.Forum_info(5)		'发送来源地址
	DvEmail.SendFromName = Dvbbs.Forum_info(0)		'发送人信息
	For i=0 To Ubound(SQL,2)
		If DvEmail.ErrCode = 0 Then
			DvEmail.SendMail SQL(2,i),EmailTopic,EmailBody	'执行发送邮件
			If Not DvEmail.ErrCode = 0 Then
				ErrMsg = DvEmail.Description
				Dvbbs_Error()
				Exit Sub
			End If
		Else
			ErrMsg = DvEmail.Description
			Dvbbs_Error()
			Exit Sub
		End If
		ii=ii+1
		Response.Write "<script>img2.width=" & Fix((ii/Remain) * 400) & ";" & VbCrLf
		Response.Write "txt2.innerHTML=""发送给"&SQL(1,i)&"（"&SQL(2,i)&"）的邮件完成，正在发送下一个用户邮件，" & FormatNumber(ii/Remain*100,4,-1) & """;" & VbCrLf
		Response.Write "txt3.innerHTML+=""发送给"&SQL(1,i)&"（"&SQL(2,i)&"）的邮件完成<br>"";"
		Response.Write "</script>"
		Response.Flush
		LastUserID = SQL(0,i)
	Next
	Set DvEmail = Nothing
	Remain = Remain -ii
	If Remain<0 Then Remain = 0
	SelChildNodes.attributes.getNamedItem("Remain").text = Remain
	SelChildNodes.attributes.getNamedItem("LasterUserID").text = LastUserID
	SelChildNodes.attributes.getNamedItem("LastTime").text = now()
	XmlDom.documentElement.appendChild(SelChildNodes)
	XmlDom.save FilePath
	Set XmlDom=Nothing
	If Remain>0 Then
		'改继续发送方式 2005-10-6 Dv.Yz
		Response.Write "<form method=""POST"" name=""resend"" action=""?Act=SendLog"">"
		Response.Write "<input type=hidden name=""SendOrders"" value=""" & SendOrders & """>"
		Response.Write "<input type=hidden name=""DelNodes"" value=""" & SelNodes & """>"
		Response.Write "&nbsp;&nbsp;<input type=""submit"" class=""button"" value=继续发送></form>"
	End If
End Sub

'显示邮件记录列表
Sub ShowLog()
Set XmlDom = Dvbbs.iCreateObject("MSXML.DOMDocument")
If Not XmlDom.load(FilePath) Then
	ErrMsg = "邮件列表中为空，请填写发邮件后再执行本操作!"
	Dvbbs_Error()
	Exit Sub
End If
Dim Node,SendLogNode,Childs
Set SendLogNode = XmlDom.DocumentElement.SelectNodes("SendLog")
Childs = SendLogNode.Length	'列表数
If Childs>10 Then
	Dim objRemoveNode,i
	For i=0 To (Childs-11)
	XmlDom.documentElement.removeChild(SendLogNode.item(i))
	Next
	XmlDom.save FilePath
End If
%>
<br>
<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">
<tr><th colspan="9" style="text-align:center;">发送邮件列表</th></tr>
<tr>
<td width="1%" class=bodytitle align=center nowrap>选取</td>
<td width="20%" class=bodytitle align=center>标题</td>
<td width="10%" class=bodytitle align=center nowrap>总共发送数目</td>
<td width="10%" class=bodytitle align=center nowrap>剩余发送数目</td>
<td width="10%" class=bodytitle align=center>操作者</td>
<td width="10%" class=bodytitle align=center>操作者IP</td>
<td width="10%" class=bodytitle align=center>添加时间</td>
<td width="10%" class=bodytitle align=center>更新时间</td>
<td width="10%" class=bodytitle align=center>操作</td>
</tr>
<form action="?" method=post name="TheForm">
<tr><td colspan="9" class="td2" height="23">
每次发送邮件<INPUT TYPE="text" NAME="SendOrders" value="10" size="4">封
</td></tr>
<%
Dim SearchStr,Topic
i=0
For Each Node in SendLogNode
	'SearchStr = Node.selectSingleNode("Search").text
	Topic = Node.selectSingleNode("EmailTopic").text
	'Node.getAttribute("MasterName")
%>
<tr>
<td class="td2" align=center><INPUT TYPE="checkbox" class="checkbox" NAME="DelNodes" value="<%=Node.getAttribute("AddTime")%>"></td>
<td class="td1" align=center><%=Topic%></td>
<td class="td1"><%=Node.getAttribute("Total")%></td>
<td class="td1"><%=Node.getAttribute("Remain")%></td>
<td class="td1" align=center><%=Node.getAttribute("MasterName")%></td>
<td class="td1"><%=Node.getAttribute("MasterIP")%></td>
<td class="td1"><%=Node.getAttribute("AddTime")%></td>
<td class="td1"><%=Node.getAttribute("LastTime")%></td>
<td class="td1" align=center><input type="submit" class="button" onclick="this.form.Act.value='SendLog';Selchecked(this.form.DelNodes,<%=i%>);" value="发送"></td>
</tr>
<%
i=i+1
Next
%>
<tr>
	<td colspan="9" class="td2">
	<input type=hidden name=Act value="DelSendLog">
	<input type=submit class="button" name=Submit value="删除记录"  onclick="{if(confirm('注意：所删除的模版将不能恢复！')){this.form.submit();return true;}return false;}">  <input type=checkbox class="checkbox" name=chkall value=on onclick="CheckAll(this.form)">全选</td>
</tr>
</form>
</table>
<SCRIPT LANGUAGE="JavaScript">
<!--
function Selchecked(obj,n){
if (obj[n]){
	obj[n].checked=true;
}else{
	obj.checked=true;
}
}
//-->
</SCRIPT>
<%
Set XmlDom = Nothing
End Sub

'填写发送邮件信息
Sub SendStep1()
%>
<br>
<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">
<form METHOD=POST ACTION="?" name="TheForm">
<tr><th colspan="2" style="text-align:center;">用户邮件通知</th></tr>
<tr>
<td width="15%" class="td2" align="right">
选择用户：
</td>
<td width="85%" class="td1">
<INPUT TYPE="text" NAME="UserName" size="40">(多个用户名请以英文逗号“,”分隔，注意区分大小写)
</td>
</tr>
<tr>
<td class="td2" align="right">
用户类别：
</td>
<td class="td1">
<INPUT TYPE="radio" class="radio" NAME="UserType" value="0" checked onclick="UType(this.value)">用户名单
<INPUT TYPE="radio" class="radio" NAME="UserType" value="1" onclick="UType(this.value)">用户组
<INPUT TYPE="radio" class="radio" NAME="UserType" value="2" onclick="UType(this.value)">所有用户
<div id="ToUserGroup" style="display:none;">
	<br>
	<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
	<tr><td height=20 class="td2">指定用户组</td></tr>
	<tr><td>
	<%
	'Response.Write "<INPUT TYPE=""checkbox"" NAME=""GetGroupID"" value=""-1"" checked>所有用户"
	Dim Rs
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
	<tr><td height=20 class="td2"><input type="button" class="button" value="打开高级设置" NAME="OPENSET" onclick="openset(this,'UpSetting')"></td></tr>
	<tr><td height=20 ID="UpSetting" style="display:NONE" class="td2">
		<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
		<tr><td height=20 colspan="4">符合条件设置(若不选取用户组，则以下条件将对所有用户生效)</td></tr>
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
</td>
</tr>
<tr>
<td class="td2" align="right">
邮件标题：
</td>
<td class="td1">
<INPUT TYPE="text" NAME="EmailTopic" size="80">
</td>
</tr>
<tr>
<td class="td2" align="right">
邮件内容：
</td>
<td class="td1">
<TEXTAREA NAME="EmailBody" Style="width:100%;height:250;"></TEXTAREA>
</td>
</tr>
<tr>
<td class="td2" align="right">&nbsp;
</td>
<td class="td2" align="center">
<INPUT TYPE="hidden" name="Act" value="sendemail">
<INPUT TYPE="submit" class="button" value="提交">&nbsp;&nbsp;&nbsp;<INPUT TYPE="reset" class="button" value="重填">
</td>
</tr>
</form>
</table>
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
function UType(n){
	var ToUserGroup = document.getElementById("ToUserGroup");
	if (n==0&&TheForm.UserName.disabled==true){
		TheForm.UserName.disabled = false;
		ToUserGroup.style.display = "none";
	}
	else{
		TheForm.UserName.disabled=true;
		if (n==1){
			ToUserGroup.style.display = "";
		}else{
			ToUserGroup.style.display = "none";
		}
	}
}
//-->
</SCRIPT>
<%
End Sub

Sub SendStep2()
	Server.ScriptTimeout=999999
	Dim UserType
	UserType = Request.Form("UserType")
	EmailTopic = Request.Form("EmailTopic")
	EmailBody = Request.Form("EmailBody")
	If EmailTopic="" or EmailBody="" Then
		ErrMsg = "请填写邮件的标题和内容!"
		Dvbbs_Error()
		Exit Sub
	End If
	Select Case UserType
		Case "0" : Call Sendtype_0()	'按指定用户
		Case "1" : Call Sendtype_1()	'按指定用户组
		Case "2" : Call Sendtype_2()	'按所有用户
		Case Else
			ErrMsg = "请选收信的用户!"
			Dvbbs_Error()
			Exit Sub
	End Select
	Dv_suc("已经成功将发送事件存入列表，请在发送列表中选取发送!")
End Sub

'按指定用户
Sub Sendtype_0()
	Dim Searchstr
	Dim ToUserName,Rs,Sql,i,ToUserID,FirstUserID
	ToUserName = Trim(Request.Form("UserName"))
	If ToUserName = "" Then
		ErrMsg = "请填写目标用户名，注意区分大小写。"
		Dvbbs_Error()
		Exit Sub
	End If
	ToUserName = Replace(ToUserName,"'","")
	ToUserName = Split(ToUserName,",")
	If Ubound(ToUserName)>100 Then
		ErrMsg = "限制一次不能超过100位目标用户。"
		Dvbbs_Error()
		Exit Sub
	End If
	For i=0 To Ubound(ToUserName)
		SQL = "Select UserID From [Dv_user] Where UserName = '"&ToUserName(i)&"' order by userid"
		SET Rs = Dvbbs.Execute(SQL)
		If Not Rs.eof Then
			If i=0 or ToUserID="" Then
				ToUserID = ToUserID & Rs(0)
				FirstUserID = Rs(0)
			Else
				ToUserID = ToUserID &","& Rs(0)
			End If
		End If
	Next
	Rs.Close : Set Rs = Nothing
	Dim Total
	Total = Ubound(Split(ToUserID,","))+1
	If Total = 0 Then
		ErrMsg = "系统找不到相应目标用户名，注意区分大小写。"
		Dvbbs_Error()
		Exit Sub
	Else
		SearchStr = "UserID in ("&ToUserID&")"
		Call CreateXmlLog(Total,SearchStr,FirstUserID)
	End If
End Sub

'按指定用户组及条件发送
Sub Sendtype_1()
	Dim GetGroupID
	Dim SearchStr,TempValue,DayStr
	GetGroupID = Replace(Request.Form("GetGroupID"),chr(32),"")
	If GetGroupID<>"" and Not Isnumeric(Replace(GetGroupID,",","")) Then
		ErrMsg = "请正确选取相应的用户组。"
'	Else
'		GetGroupID = Dvbbs.Checkstr(GetGroupID)
	End If
	If IsSqlDataBase=1 Then
		DayStr = "d"
	Else
		DayStr = "'d'"
	End If
	If GetGroupID<>"" Then
		If InStr(GetGroupID,",")=0 Then
			SearchStr = "UserGroupID = "&Dvbbs.CheckNumeric(GetGroupID)
		Else
			SearchStr = "UserGroupID in ("&Replace(GetGroupID,"'","")&")"
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
	If SearchStr="" Then
		ErrMsg = "请填写发送的条件选项。"
	End If
	If ErrMsg<>"" Then Dvbbs_Error() : Exit Sub
	Dim Rs,Sql,Total,FirstUserID
	Sql = "Select Count(UserID) From Dv_user Where "& SearchStr
	Total = Dvbbs.Execute(Sql)(0)
	If Total>0 Then
		Sql = "Select Top 1 UserID From Dv_user Where "& SearchStr & " order by userid"
		FirstUserID = Dvbbs.Execute(Sql)(0)
		Call CreateXmlLog(Total,SearchStr,FirstUserID)
	Else
		ErrMsg = "发送目标用户为空，请更改发送条件再进行发送。"
		Dvbbs_Error()
		Exit Sub
	End If
End Sub

'按所有用户
Sub Sendtype_2()
	Dim SearchStr
	Dim Rs,Sql,Total,FirstUserID
	Sql = "Select Count(UserID) From Dv_user"
	Total = Dvbbs.Execute(Sql)(0)
	If Total>0 Then
		Sql = "Select Top 1 UserID From Dv_user order by userid"
		FirstUserID = Dvbbs.Execute(Sql)(0)
		Call CreateXmlLog(Total,SearchStr,FirstUserID)
	Else
		ErrMsg = "发送目标用户为空，请更改发送条件再进行发送。"
		Dvbbs_Error()
		Exit Sub
	End If
End Sub

'添加发送记录
Sub CreateXmlLog(SendTotal,Search,LasterUserID)
	Dim node,attributes,createCDATASection,ChildNode
	Set XmlDom = Dvbbs.iCreateObject("MSXML.DOMDocument")
	If Not XmlDom.load(FilePath) Then
		XmlDom.loadxml "<?xml version=""1.0"" encoding=""gb2312""?><EmailLog/>"
	End If
	Set node=XmlDom.createNode(1,"SendLog","")
	Set attributes=XmlDom.createAttribute("Total")
	attributes.text = SendTotal
	node.attributes.setNamedItem(attributes)
	Set attributes=XmlDom.createAttribute("Remain")
	attributes.text = SendTotal
	node.attributes.setNamedItem(attributes)
	Set attributes=XmlDom.createAttribute("LasterUserID")
	attributes.text = LasterUserID
	node.attributes.setNamedItem(attributes)
	Set attributes=XmlDom.createAttribute("MasterName")
	attributes.text = Dvbbs.Membername
	node.attributes.setNamedItem(attributes)
	Set attributes=XmlDom.createAttribute("MasterUserID")
	attributes.text = Dvbbs.UserID
	node.attributes.setNamedItem(attributes)
	Set attributes=XmlDom.createAttribute("MasterIP")
	attributes.text = Dvbbs.UserTrueIP
	node.attributes.setNamedItem(attributes)
	Set attributes=XmlDom.createAttribute("AddTime")
	attributes.text = Now()
	node.attributes.setNamedItem(attributes)
	Set attributes=XmlDom.createAttribute("LastTime")
	attributes.text = Now()
	node.attributes.setNamedItem(attributes)
	Set ChildNode = XmlDom.createNode(1,"Search","")
	Set createCDATASection=XmlDom.createCDATASection(replace(Search,"]]>","]]&gt;"))
	ChildNode.appendChild(createCDATASection)
	node.appendChild(ChildNode)
	Set ChildNode = XmlDom.createNode(1,"EmailTopic","")
	Set createCDATASection=XmlDom.createCDATASection(replace(EmailTopic,"]]>","]]&gt;"))
	ChildNode.appendChild(createCDATASection)
	node.appendChild(ChildNode)
	Set ChildNode = XmlDom.createNode(1,"EmailBody","")
	Set createCDATASection=XmlDom.createCDATASection(replace(EmailBody,"]]>","]]&gt;"))
	ChildNode.appendChild(createCDATASection)
	node.appendChild(ChildNode)
	XmlDom.documentElement.appendChild(node)
	XmlDom.save FilePath
	Set XmlDom = Nothing
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
%>