<!--#include file=../conn.asp-->
<!-- #include file="inc/const.asp" -->
<%
dim admin_flag
Dim XMLDom,node,nodename,checkedstr
Head()
admin_flag=Split("28,29",",")
CheckAdmin(","&admin_flag(0)&",")
CheckAdmin(","&admin_flag(1)&",")
call main()
Footer()

Sub main()
dim sel

If request("action") = "savebadword" Then
	Call savebadword()
Else
%>

<form action="badword.asp?action=savebadword" method=post>

<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">

<%if request("reaction")="badword" then%>
<tr>
<th colspan=2>帖子过滤字符</th>
</tr>
<tr>
<td class=td1 width="100%" colspan=2><B>说明</B>：过滤字符设定规则为  <B>要过滤的字符=过滤后的字符</B> ，每个过滤字符用回车分割开。</td>
</tr>
<tr>
<td class=td1 width="100%" colspan=2>
<textarea name="badwords" cols="80" rows="8"><%
Dim i
For i=0 To Ubound(Dvbbs.BadWords)
	If i > UBound(Dvbbs.rBadWord) Then
		Response.Write Dvbbs.BadWords(i) & "=*"
	Else
		Response.Write Dvbbs.BadWords(i) & "=" & Dvbbs.rBadWord(i)
	End If
	If i<Ubound(Dvbbs.BadWords) Then Response.Write chr(10)
Next
%></textarea>
</td>
</tr>
<%elseif request("reaction")="splitreg" Then
LoadSetting
	%>
<tr>
<th colspan=2>注册过滤字符</th>
</tr>
<tr>
<td class=td1 width="20%">说明：</td>
<td class=td1 width="80%">注册过滤字符将不允许用户注册包含以下字符的内容，请您将要过滤的字符串添入，如果有多个字符串，请用“,”分隔开，例如：沙滩,quest,木鸟</td>
</tr>
<tr>
<td class=td1 width="20%">请输入过滤字符</td>
<td class=td1 width="80%"><input type="text" name="splitwords" value="<%=split(Dvbbs.cachedata(1,0),"|||")(4)%>" size="80"></td>
</tr>
</table>
<br>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr>
<th colspan=2>注册限制设置</th>
</tr>
<tr>
<td class=td1 width="20%">说明：</td>
<td class=td1 width="80%">扩展注册设置,请根据自己需要设置.</td>
</tr>
<tr>
<%
nodename="checkip"
Set Node=XMLDom.documentElement.selectSingleNode(nodename)
If Node Is Nothing Then
	checkedstr=" checked=""checked"""
Else
If Node.selectSingleNode("@use").text="1" Then
	checkedstr=" checked=""checked"""
Else
	checkedstr=""
End If
End If
%>
<td class=td1 width="20%">采用IP策略</td>
<td class=td1 width="80%"><input type="checkbox" class="checkbox" value="1" name="<%=nodename%>" <%=checkedstr %>  /> 如果不选择,所有和IP地址有关的设置都不起作用.</td>
</tr>
<tr>
<td class=td1 width="20%">允许注册IP(IP白名单)<br>
填写可以注册的IP地址,格式是IP地址=说明,每个用换行分开
支持通配符,如192.168.*.* =内网IP 如不采用IP白名单,请留空
</td>
<td class=td1 width="80%"><textarea name="iplist1" cols="80" rows="8"><%
For Each  Node In XMLDom.documentElement.selectNodes("checkip/iplist1/ip")
	Response.Write Node.text &" = "&Node.selectSingleNode("@description").text &Chr(10)
Next
%></textarea></td>
</tr>
<tr>
<td class=td1 width="20%">禁止注册IP(ip黑名单)<br>
填写可以注册的IP地址,格式是IP地址=说明,每个用换行分开
支持通配符,如192.168.*.* =内网IP 如不采用IP黑名单,请留空
</td>
<td class=td1 width="80%"><textarea name="iplist2" cols="80" rows="8"><%
For Each  Node In XMLDom.documentElement.selectNodes("checkip/iplist2/ip")
	Response.Write Node.text &" = "&Node.selectSingleNode("@description").text &Chr(10)
Next
%></textarea></td>
</tr>
<%
nodename="postipinfo"
Set Node=XMLDom.documentElement.selectSingleNode("@"&nodename)
If Node Is Nothing Then
	checkedstr=" checked=""checked"""
Else
If Node.text="1" Then
	checkedstr=" checked=""checked"""
Else
	checkedstr=""
End If
End If
%>
<tr>
<td class=td1 width="20%">如IP受限制提交IP来源信息</td>
<td class=td1 width="80%">
<input type="checkbox" class="checkbox" value="1" name="<%=nodename%>" <%=checkedstr %> /><br> 如果注册用户所在IP不在允许注册之列,可以引导注册者进入提交当前IP信息的页面,以便管理员可以增加该段IP地址的许可.
</td>
</tr>
<%
nodename="checkregcount"
Set Node=XMLDom.documentElement.selectSingleNode("@"&nodename)
If Node Is Nothing Then
	checkedstr="1"
Else
		checkedstr=Node.text
End If
%>
<tr>
<td class=td1 width="20%">一个IP地址一天可以注册次数</td>
<td class=td1 width="80%"><input type="text" Size="4" value="<%=checkedstr %>" name="<%=nodename%>"  />请填写数字否则会出错,如不想限制,请填写0</td>
</tr>
<%
nodename="checknumeric"
Set Node=XMLDom.documentElement.selectSingleNode("@"&nodename)
If Node Is Nothing Then
	checkedstr=" checked=""checked"""
Else
If Node.text="1" Then
	checkedstr=" checked=""checked"""
Else
	checkedstr=""
End If
End If
%>
<tr>
<td class=td1 width="20%">禁止纯数字ID注册</td>
<td class=td1 width="80%"><input type="checkbox" class="checkbox" value="1" name="<%=nodename%>" <%=checkedstr %> />是否允许采用纯数的用户名注册</td>
</tr>
<%
nodename="checktime"
Set Node=XMLDom.documentElement.selectSingleNode("@"&nodename)
If Node Is Nothing Then
	checkedstr=" checked=""checked"""
Else
If Node.text="1" Then
	checkedstr=" checked=""checked"""
Else
	checkedstr=""
End If
End If
%>
<tr>
<td class=td1 width="20%">要求输入当前时间</td>
<td class=td1 width="80%"><Input type="checkbox" class="checkbox" value="1" name="<%=nodename%>" <%=checkedstr %> /><br>如果启用,要求用户选择自己所在时区和输入他所在地的时间(以小时为单位)</td>
</tr>
<%
nodename="usevarform"
Set Node=XMLDom.documentElement.selectSingleNode("@"&nodename)
If Node Is Nothing Then
	checkedstr=" checked=""checked"""
Else
If Node.text="1" Then
	checkedstr=" checked=""checked"""
Else
	checkedstr=""
End If
End If
%>
<tr>
<td class=td1 width="20%">采用动态的表单项目名称</td>
<td class=td1 width="80%"><input type="checkbox" class="checkbox" value="1" name="<%=nodename%>" <%=checkedstr %> />采用不定名的表单项目名称,增加机器人注册的难度.</td>
</tr>
<%end if%>
<input type=hidden value="<%=request("reaction")%>" name="reaction">
<tr> 
<td class=td1 width="20%">&nbsp;</td>
<td width="80%" class=td1><input type="submit" class="button" name="Submit" value="提 交"></td>
</tr>
</table>

</form>
<%end if%>
<%
end sub

sub savebadword()
dim iforum_setting,forum_setting,i,sql,rs
If request("reaction")="badword" then
dim badwords,badwords_1,badwords_2,badwords_3
badwords=request("badwords")
badwords=split(badwords,vbCrlf)
for i = 0 to ubound(badwords)
	if not (badwords(i)="" or badwords(i)=" ") then
		badwords_1 = split(badwords(i),"=")
		If ubound(badwords_1)=1 Then
			If i=0 Then
				badwords_2 = badwords_1(0)
				badwords_3 = badwords_1(1)
			Else
				badwords_2 = badwords_2 & "|" & badwords_1(0)
				badwords_3 = badwords_3 & "|" & badwords_1(1)
			End If
		End If
	End If
next

sql = "update dv_setup set Forum_Badwords='"&replace(badwords_2,"'","''")&"',Forum_rBadword='"&replace(badwords_3,"'","''")&"'"
dvbbs.execute(sql)
elseif request("reaction")="splitreg" then
Call LoadXML
'forum_info|||forum_setting|||forum_user|||copyright|||splitword|||stopreadme
Set rs=Dvbbs.execute("select forum_setting from dv_setup")
iforum_setting=split(rs(0),"|||")
forum_setting=iforum_setting(0) & "|||" & iforum_setting(1) & "|||" & iforum_setting(2) & "|||" & iforum_setting(3) & "|||" & request("splitwords") & "|||" & iforum_setting(5)
sql = "update dv_setup set forum_setting='"&replace(forum_setting,"'","''")&"',Forum_Boards='"&Dvbbs.checkstr(XMLDom.XML)&"'"
dvbbs.execute(sql)
End If
Dvbbs.loadSetup()
Dv_suc("更新成功")

End Sub
Sub LoadSetting()
	Dim Rs,Nedupdate
	Set Rs=Dvbbs.Execute("Select Forum_Boards From Dv_setup")
	Set XMLDom=Dvbbs.CreateXmlDoc("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	If Not (xmldom.LoadXML(Rs(0))) Then
		Nedupdate=True
	ElseIf xmldom.documentElement.nodeName<>"regsetting" Then
		Nedupdate=True
	End If
	If Nedupdate Then
		XMLDom.LoadXML "<?xml version=""1.0""?><regsetting/>"
		Dvbbs.Execute"update Dv_setup Set Forum_Boards='"&Dvbbs.checkstr(XMLDom.XML)&"'"
	End If
End Sub
Sub LoadXML()
	Dim Node,node1,i,iplist,node2,iplist1
	Set XMLDom=Dvbbs.CreateXmlDoc("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	XMLDom.appendChild(XMLDom.createElement("regsetting"))
	Set Node=xmldom.documentElement.appendChild(XMLDom.createNode(1,"checkip",""))
	If Request.form("checkip")="1" Then
			Node.attributes.setNamedItem(XMLDom.createNode(2,"use","")).text="1"
	Else
		Node.attributes.setNamedItem(XMLDom.createNode(2,"use","")).text="0"
	End If
	Set node1=Node.appendChild(XMLDom.createElement("iplist1"))
	For each iplist in split(Request.form("iplist1"),vbnewline)
		If iplist<>"" Then
			iplist1=Split(iplist,"=")
			If UBound(iplist1)>0 Then
			Set node2=node1.appendChild(XMLDom.createNode(1,"ip",""))
			node2.text=Trim(iplist1(0))
			Node2.attributes.setNamedItem(XMLDom.createNode(2,"description","")).text=Trim(iplist1(1))
			End If
		End If
	Next
	Set node1=Node.appendChild(XMLDom.createElement("iplist2"))
	For each iplist in split(Request.form("iplist2"),vbnewline)
		If iplist<>"" Then
			iplist1=Split(iplist,"=")
			If UBound(iplist1)>0 Then
			Set node2=node1.appendChild(XMLDom.createNode(1,"ip",""))
			node2.text=Trim(iplist1(0))
			Node2.attributes.setNamedItem(XMLDom.createNode(2,"description","")).text=Trim(iplist1(1))
			End If
		End If
	Next
	If Request.form("postipinfo")="1" Then
			xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"postipinfo","")).text="1"
	Else
		xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"postipinfo","")).text="0"
	End If
	'If Request.form("checkproxy")="1" Then
	'		xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"checkproxy","")).text="1"
	'Else
	'	xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"checkproxy","")).text="0"
	'End If
	If Request.form("checknumeric")="1" Then
			xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"checknumeric","")).text="1"
	Else
		xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"checknumeric","")).text="0"
	End If
	If Request.form("checktime")="1" Then
			xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"checktime","")).text="1"
	Else
		xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"checktime","")).text="0"
	End If
	If Request.form("usevarform")="1" Then
			xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"usevarform","")).text="1"
	Else
		xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"usevarform","")).text="0"
	End If
	If Request.form("checkregcount")<>"" and IsNumeric(Request.form("checkregcount")) Then
			xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"checkregcount","")).text=Request.form("checkregcount")
	Else
		xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"checkregcount","")).text="0"
	End If
	
End Sub
%>