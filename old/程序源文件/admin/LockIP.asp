<!--#include file="../conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
Head()
Dim admin_flag
admin_flag=",30,"
CheckAdmin(admin_flag)
Call main()

Sub main()
Dim userip,ips,GetIp1,GetIp2
If request("userip")<>"" then
	userip=request("userip")
	ips=Split(userIP,".")
	If Ubound(ips)=3 Then GetIp1=ips(0)&"."&ips(1)&"."&ips(2)&".*"
End if
If request("action")="add" or request("userip")<>"" then
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th width="100%" colspan=2>IP限制管理－－添加</th>
</tr>
<%
dim sip,str1,str2,str3,str4,num_1,num_2,i
if request.querystring("reaction")="save" then
	sip=cstr(request.form("ip1"))
	If sip<>"" Then
		If Trim(Dvbbs.cachedata(25,0))<>"" Then
			sip=Trim(Dvbbs.cachedata(25,0)) & "|" & Replace(sip,"|","")
		End If
	End If
	if sip<>"" then
		dvbbs.execute("update dv_setup set Forum_LockIP='"&replace(sip,"'","''")&"'")
		Dvbbs.loadSetup()
	end if
%>
<tr>
<td width="100%" colspan=2 class=td1>添加成功！</td>
</tr>
<%
else
%>
<form action="LockIP.asp?action=add&reaction=save" method="post">
<tr>
<td width="100%" class=td1 colspan=2><B>说明</B>：您可以添加多个限制IP，每个IP用|号分隔，限制IP的书写方式如202.152.12.1就限制了202.152.12.1这个IP的访问，如202.152.12.*就限制了以202.152.12开头的IP访问，同理*.*.*.*则限制了所有IP的访问。在添加多个IP的时候，请注意最后一个IP的后面不要加|这个符号</td>
</tr>
<tr>
<td width="20%" class=td1>限制I&nbsp;P</td>
<td width="80%" class=td1><input type="text" name="ip1" size="30" value="<%=GetIp1%>">&nbsp;如202.152.12.*</td>
</tr>
<tr>
<td width="20%" class=td1>&nbsp;</td>
<td width="80%" class=td1>
<input type="submit" class="button" name="Submit" value="添 加">
</td>
</tr>
</form>
<%
end if
elseif request("action")="delip" then
	userip=request("ips")
	'userip=split(userip,chr(10))
	userip=split(userip,vbCrLf)
	for i = 0 to ubound(userip)
		if not (userip(i)="" or userip(i)=" ") then
			If i=0 Then
				getip1 = userip(i)
			Else
				getip1 = getip1 & "|" & userip(i)
			End If
		End If
	next
	dvbbs.execute("update dv_setup set forum_lockip='"&replace(getip1,"'","''")&"'")
	Dvbbs.loadSetup()
	Dv_suc("更新限制IP成功！")
else
%>
<FORM METHOD=POST ACTION="?action=delip">
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th width="100%" colspan=2>IP限制管理－－管理</th>
</tr>
<tr>
<td width="100%" class=td1 colspan=2>
<B>说明</B>：您可以添加多个限制IP，每个IP用回车分隔，限制IP的书写方式如202.152.12.1就限制了202.152.12.1这个IP的访问，如202.152.12.*就限制了以202.152.12开头的IP访问，同理*.*.*.*则限制了所有IP的访问。在添加多个IP的时候，请注意最后一个IP的后面不要加回车。
</td>
</tr>
<tr>
<td width="100%" class=td1 colspan=2>
<textarea name="ips" cols="80" rows="8">
<%
userip=split(Trim(Dvbbs.cachedata(25,0)),"|")
For i = 0 To Ubound(userip)
	Response.Write userip(i)
	If i < Ubound(Userip) Then Response.Write Chr(10)
Next
%></textarea>
</td>
</tr>
<tr>
<td width="20%" class=td1>&nbsp;</td>
<td width="80%" class=td1>
<input type="submit" class="button" name="Submit" value="修 改">
</td>
</tr>
</FORM>
</table>
<%
End If
Footer()
end Sub
%>
