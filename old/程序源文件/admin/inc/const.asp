<!--#Include File="../../inc/Dv_ClsMain.asp"-->
<%
If Session("flag")<> "" Then Dvbbs.Master = True
'UpUserFaceFolder 
'如是独立的虚拟目录,则要写成"/uploadFace"如果是论坛目录下的普通目录,则写成""
'Const UpUserFaceFolder=""
Const UpUserFaceFolder="/uploadFace/"
MyDbPath = "../"
If IsNumeric(Dvbbs.UserHidden) = 0 or Dvbbs.Userhidden = "" Then Dvbbs.UserHidden = 2
If IsNumeric(Dvbbs.UserID) = 0 Or Dvbbs.UserID="" Then Dvbbs.UserID=0
Dvbbs.UserID = Clng(Dvbbs.UserID)
Dvbbs.MemberClass = Dvbbs.checkStr(Request.Cookies(Dvbbs.Forum_sn)("userclass"))

Set MyBoardOnline=new Cls_UserOnlne 
'获得论坛基本信息和检测用户登陆状态
Dvbbs.GetForum_Setting
Dvbbs.CheckUserLogin
'重新赋予用户是否可进入后台权限		已移至Admin_login.asp验证，轻飘飘
'If Dvbbs.GroupSetting(70)="1" Then
'	Dvbbs.Master = True
'Else
'	Dvbbs.Master = False
'End If

'后台信息和函数部分
Dim AllPostTable
Dim AllPostTableName
Dim FoundErr
FoundErr=False 
Dim ErrMsg
'Dim Rs,sql
template.ChildFolder="Admin"
template.Folder="../"&template.Folder
Dvbbs.LoadTemplates("Admin")

'Set Rs=Dvbbs.Execute("Select H_Content From Dv_Help Where H_ID=1")
'template.value = Rs(0)

'检测管理权限
Sub CheckAdmin(flag)
	If Not Dvbbs.Master Or Session("flag")="" Then
		Response.Redirect "../showerr.asp?action=OtherErr&ErrCodes=<li>本页面为管理员专用，请<a href=admin_login.asp target=_top>登录</a>后进入。"
	End If

	If Instr(","&session("flag")&",",flag)=0 and flag<>"" then
		Errmsg=ErrMsg +	"<br /><li>本页面为管理员专用，请<a href=../admin_login.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		Dvbbs_error()
	End If
End Sub



Sub AllPostTable1()
	Dim Trs
	Set Trs=Dvbbs.Execute("select * from [Dv_TableList]")
	AllPostTable=""
	Do While Not TRs.EOF
		If AllPostTable=""  Then 
			AllPostTable=TRs("TableName")
			AllPostTableName=TRs("TableType")
		Else
			AllPostTable=AllPostTable&"|"&TRs("TableName")
			AllPostTableName=AllPostTableName&"|"&TRs("TableType")
		End If
	TRs.MoveNext
	Loop 
	Trs.Close
End Sub

AllPostTable1
AllPostTableName=Split(AllPostTableName,"|")
AllPostTable=Split(AllPostTable,"|")
Dim NowUseBbs
NowUseBbs=Dvbbs.NowUseBbs

Sub Footer()
	Response.Write "</html>"
	SaveLog()
	Set Dvbbs=Nothing 
End Sub

Sub Head()
%>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312" />
<TITLE><%=Dvbbs.Forum_info(0)%>-管理页面</TITLE>
<link href="skins/css/main.css" rel="stylesheet" type="text/css" />
<script src="../inc/main82.js" type="text/javascript"></script>
<script src="inc/admin.js" type="text/javascript"></script>
<script src="inc/Dvbbs_CheckForm.js" type="text/javascript"></script>
</head>
<%
	Dim XMLDOM
	Set XMLDOM=Application(Dvbbs.CacheName&"_boardlist").cloneNode(True)
	Response.Write "<script language=""javascript"" type=""text/javascript"">"
	Response.Write "var boardxml='';var ISAPI_ReWrite="&IsUrlreWrite&";"
	'Response.Write "var boardxml='<?xml version=""1.0"" encoding=""gb2312""?>"& replace(XMLDom.documentElement.XML ,"'","\'")&"';var ISAPI_ReWrite="&IsUrlreWrite&";"
	Response.Write "</script>"
	Response.Write "<body>"
	Response.Write Chr(10)

End Sub

Sub Dv_suc(info)
	Dim UrlArr,InfoArr
	InfoArr = Split(info,"##")
	If UBound(InfoArr)>0 Then
		UrlArr = InfoArr(1)
	Else
		UrlArr = Request.ServerVariables("HTTP_REFERER")
		If UrlArr="" Or InStr(UrlArr,Dvbbs.CacheData(33,0)&"index.asp")>0 Then UrlArr="javascript:window.history.go(-1)"
	End If
	Response.Write"<br>"
	Response.Write"<table cellpadding=0 cellspacing=0 align=center width=""100%"">"
	Response.Write"<tr align=center>"
	Response.Write"<th width=""100%"" style=""text-align:center;"" colspan=2>成功信息"
	Response.Write"</td>"
	Response.Write"</tr>"
	Response.Write"<tr>"
	Response.Write"<td width=""100%"" class=""td2"" colspan=2>"
	Response.Write InfoArr(0)
	Response.Write"</td></tr>"
	Response.Write"<tr>"
	Response.Write"<td class=""td2"" valign=middle colspan=2 style=""text-align:center;""><a href="&UrlArr&" ><<返回上一页</a></td></tr>"
	Response.Write"</table>"
End Sub
'页面错误提示信息
Sub dvbbs_error()
	Response.Write"<br>"
	Response.Write"<table cellpadding=3 cellspacing=1 align=center width=""100%"">"
	Response.Write"<tr align=center>"
	Response.Write"<th width=""100%"" style=""text-align:center;"" colspan=2>错误信息"
	Response.Write"</td>"
	Response.Write"</tr>"
	Response.Write"<tr>"
	Response.Write"<td width=""100%"" class=""td1"" colspan=2>"
	Response.Write ErrMsg
	Response.Write"</td></tr>"
	Response.Write"<tr>"
	Response.Write"<td class=""td1"" valign=middle colspan=2 style=""text-align:center;""><a href=""javascript:history.go(-1)""><<返回上一页</a></td></tr>"
	Response.Write"</table>"
	footer()
	Response.End 
End Sub
Function fixjs(Str)
	If Str <>"" Then
		str = replace(str,"\", "\\")
		Str = replace(str, chr(34), "\""")
		Str = replace(str, chr(39),"\'")
		Str = Replace(str, chr(13), "\n")
		Str = Replace(str, chr(10), "\r")
		str = replace(str,"'", "&#39;")
	End If
	fixjs=Str
End Function
Function enfixjs(Str)
	If Str <>"" Then
		Str = replace(str,"&#39;", "'")
		Str = replace(str,"\""" , chr(34))
		Str = replace(str, "\'",chr(39))
		Str = Replace(str, "\r", chr(10))
		Str = Replace(str, "\n", chr(13))
		Str = replace(str,"\\", "\")
	End If
	enfixjs=Str
End Function

Function Reload_All_Board_Cache()
	'更新版面列表缓存
	ReloadBoardListAll
	'更新单个版面缓存（循环）
	Dim BoardListAll,BoardListNum,myBoardID
	Dim i,Rs
	BoardListAll=myCache.value
	BoardListNum=Ubound(BoardListAll,2)
	For i=0 To BoardListNum
		myBoardID=BoardListAll(0,i)
		ReloadBoardInfo(myBoardID)
		Set rs=Dvbbs.Execute("Select ParentStr from board where boardid="&myBoardID)
		If not rs.eof Then
			Dvbbs.ReloadBoardParentStr(rs(0))
		End If
		Rs.close
		Set Rs=nothing
	Next
End Function
Sub SaveLog()
	On Error Resume Next
	Dim RequestStr
	Dim Sql
	RequestStr= Request("action")
	If RequestStr<>"" Then 
		RequestStr="action="&RequestStr
		RequestStr=Dvbbs.checkStr(RequestExp(RequestStr))
		RequestStr=Left(RequestStr,250)
		sql="insert into [Dv_log] (l_touser,l_username,l_content,l_ip,l_type) values ('"&Dvbbs.ScriptName&"','"&Dvbbs.membername&"','"&RequestStr&"','"&Dvbbs.UserTrueIP&"',0)"		
		Dvbbs.Execute(sql)
	End If
	If request.form<>"" Then
		RequestStr=Dvbbs.checkStr(request.form)
		RequestStr=Left(RequestExp(RequestStr),250)
		sql="insert into [Dv_log] (l_touser,l_username,l_content,l_ip,l_type) values ('"&Dvbbs.ScriptName&"','"&Dvbbs.membername&"','"&RequestStr&"','"&Dvbbs.UserTrueIP&"',1)"		
		Dvbbs.Execute(sql)
	End If
End Sub

Public Function RequestExp(Textstr)
	Dim Str,re
	Str = Textstr
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
		re.Pattern = "(password|answer)([^=]*)=([^(&|&amp;)]*)"
		str = re.Replace(str,"$1$2=******")
	Set Re = Nothing
	RequestExp = Str
End Function
%>