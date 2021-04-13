<!--#include file="conn.asp"-->
<!--#include file="inc/Const.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="boke/config.asp"-->
<!--#include file="boke/checkinput.asp"-->
<%
Dim ErrMsg
If not Dvbbs.master or  Instr(","&session("flag")&",",",38,")=0 Then Response.Redirect "index.asp"
'Set MyBoardOnline=new Cls_UserOnlne 
'Dvbbs.GetForum_Setting
'Dvbbs.CheckUserLogin
'Response.Write "test"
'DvBoke.Execute("update Dv_Boke_user set SysCatID=1 where SysCatID=0")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name=keywords content="动网先锋,动网论坛,dvbbs,博客,blog,boke">
<title>动网博客系统管理页面</title>
<link rel="stylesheet" type="text/css" href="<%=Dvbbs.CacheData(33,0)%>skins/css/main.css" />

<SCRIPT LANGUAGE="JavaScript">
<!--
function alertreadme(str,url){
{if(confirm(str)){
location.href=url;
return true;
}return false;}
}
//-->
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=2 height=25>动网博客系统管理
</th>
</tr>
<tr>
<td class="td1" colspan=2>
<p><B>注意</B>：
<BR>① 删除博客系统栏目或话题前请先将其中文章、评论和用户转移到其他栏目或分栏中，不需要的文章或评论可在信息管理中批量删除
<BR>② 删除用户博客栏目前请先将其中文章和评论转移到该用户的其它栏目后再执行删除操作
</td>
</tr>
<tr>
<td class="td2" height=25>
<B>管理操作选项</B></td>
<td class="td2"><a href="?s=8">设置</a> | 
<a href="?s=1&t=1">栏目</a> | <a href="?s=1&t=2">话题</a> | <a href="?s=2">用户管理</a> | <a href="?s=3">用户栏目</a> | <a href="?s=4">公告管理</a> | <a href="?s=5">上传管理</a> | <a href="?s=6">关键字</a> | <a href="?s=7">模板</a> | <a href="?s=9">数据更新</a>
</td>
</tr>
</table>
<p></p>
<%
Select Case Request("s")
Case "1"
	Boke_SysCat()
Case "2"
	Boke_User()
Case "3"
	Boke_UserCat()
Case "4"
	Boke_SysNews()
Case "5"
	Boke_UploadFile()
Case "6"
	Boke_KeyWord()
Case "7"
	Boke_Skins()
Case "8"
	Boke_Setting()
Case "9"
	Boke_Update()
Case Else
	Boke_Setting()
End Select
Dvbbs.PageEnd()
Sub Boke_UploadFile()
	Dim FID,Sql,Rs
	If Request.QueryString("act")="del" Then
		Dim FileSize,SpaceSize,objFSO,FilePath,ViewPath
		FID = DvBoke.CheckNumeric(Request("fid"))
		If FID = 0 Then
			ErrMsg = "文件参数错误，请重新选取正确的文件再进行操作!"
			Dvbbs_error()
			Exit Sub
		End If
		Set Rs = DvBoke.Execute("Select ID,BokeUserID,UserID,UserName,CatID,sType,TopicID,PostID,IsTopic,Title,FileName,sFileName,FileType,FileSize,FileNote,DownNum,ViewNum,DateAndTime,PreviewImage,IsLock From Dv_Boke_Upfile where id="&FID)
		If Not Rs.Eof THen
			FileSize = Formatnumber((Rs("FileSize")/1024)/1024,2)
			ViewPath = Rs("PreviewImage")
			FilePath = Rs("FileName")
			If Not FilePath = "" Then
				FilePath = DvBoke.System_UpSetting(19)&FilePath
			End If
			SpaceSize = DvBoke.Execute("Select SpaceSize From Dv_Boke_User where UserID="&Rs("BokeUserID"))(0)
			If SpaceSize>0 Then
				SpaceSize = SpaceSize - FileSize
				If SpaceSize<0 Then SpaceSize = 0
				DvBoke.Execute("Update Dv_Boke_User set SpaceSize = "&SpaceSize&" where UserID="&Rs("BokeUserID"))
			End If
			DvBoke.Execute("delete from Dv_Boke_Upfile where id="&FID)
			Set objFSO = Dvbbs.iCreateObject("Scripting.FileSystemObject")
			If ViewPath<>"" Then
				If objFSO.FileExists(Server.MapPath(ViewPath)) Then
					objFSO.DeleteFile(Server.MapPath(ViewPath))
				End If
			End If
			If objFSO.FileExists(Server.MapPath(FilePath)) Then
				objFSO.DeleteFile(Server.MapPath(FilePath))
			End If
			Set objFSO = Nothing
		End If
		Dv_suc("文件已成功删除！")
		Exit Sub
	End If
%>
	<table width="100%" border="0" cellspacing="1" cellpadding="5"  align="center" class="tableBorder">
	<tr> 
	<th width="100%" height="25" align="left" ID="TableTitleLink">
	博客上传文件管理
	</th>
	</tr>
	<tr><td class="td1">说明：
	<ul>
	<li><font color="red">未知</font>文件:是指作者上传后未发表或未使用的文件。</li>
	</ul></td></tr>
	</table>

	<br/>
	<table width="100%" border="0" cellspacing="1" cellpadding="5"  align="center" class="tableBorder">
	<tr> 
	<th width="100%" colspan="5" height="25" align="left" ID="TableTitleLink">
	上传信息列表
	</th>
	</tr>
	<tr>
	<td class="bodytitle" height="24" width="15%">
	演示
	</td>
	<td class="bodytitle" height="24" width="35%">
	名称/ 路径
	</td>
	<td class="bodytitle" height="24" width="15%">
	作者
	</td>
	<td class="bodytitle" height="24" width="15%">
	上传时间
	</td>
	<td class="bodytitle" height="24" width="20%">
	操作
	</td>
	</tr>
<%
	
	Dim CurrentPage,Page_Count,Pcount,i
	Dim TotalRec,EndPage
	Dim ViewFile
	CurrentPage=Request("page")
	If CurrentPage="" Or Not IsNumeric(CurrentPage) Then
		CurrentPage=1
	Else
		CurrentPage=Clng(CurrentPage)
		If Err Then
			CurrentPage=1
			Err.Clear
		End If
	End If
	'ID=0 ,BokeUserID=1 ,UserID=2 ,UserName=3 ,CatID=4 ,sType=5 ,TopicID=6 ,PostID=7 ,IsTopic=8 ,Title=9 ,FileName=10 ,sFileName=11 ,FileType=12 ,FileSize=13 ,FileNote=14 ,DownNum=15 ,ViewNum=16 ,DateAndTime=17 ,PreviewImage=18 ,IsLock=19
	Sql = "Select ID,BokeUserID,UserID,UserName,CatID,sType,TopicID,PostID,IsTopic,Title,FileName,sFileName,FileType,FileSize,FileNote,DownNum,ViewNum,DateAndTime,PreviewImage,IsLock From Dv_Boke_Upfile order by ID Desc"
	If Not IsObject(Boke_Conn) Then Boke_ConnectionDatabase
	Set Rs=Dvbbs.iCreateObject("Adodb.RecordSet")
	Rs.Open Sql,Boke_Conn,1,1
	If Not (Rs.Eof And Rs.Bof) Then
		Rs.PageSize = 30
		Rs.AbsolutePage=CurrentPage
		Page_Count=0
		TotalRec=Rs.RecordCount
		While (Not Rs.Eof) And (Not Page_Count = 30)
		ViewFile = Rs(18)
		If ViewFile = "" Then
			ViewFile = DvBoke.System_UpSetting(19) & Rs(10)
		End If
%>
<tr>
<td class="td1">
<%
If Rs(12)=1 Then
	'修改图片路径为非父路径 2005-10-6 Dv.Yz
	Response.Write "<a href="""&DvBoke.System_UpSetting(19)&Rs(10)&""" target=""_blank""><img src="""&ViewFile&""" border=""1"" width=""80"" height=""60""></a>"
Else
	Response.Write "其它"
End If
%>
</td>
<td class="td1">
《<b><%=Rs(9)%></b>》
<br/><u><%=Rs(10)%></u>
</td>
<td class="td2">
<%=Rs(3)%>
</td>
<td class="td1">
<%=Rs(17)%>
</td>
<td class="td2">
<%If Rs(19)=4 Then%>
<font color="red">未知</font>
<%Else%>
<a href="Userid_<%=Rs(1)%>.showtopic.<%=Rs(6)%>.html" target="_blank">查看</a>
<%End If%>
| <a href="?s=5&act=del&fid=<%=Rs(0)%>">删除</a>
</td>
</tr>
<%
			Page_Count = Page_Count + 1
		Rs.MoveNext
		Wend
		Pcount=Rs.PageCount
%>
<tr><td colspan=5 class=td1>共有<%=TotalRec%>条记录，分页：
<%
	Dim Searchstr
	Searchstr = "?s=5"
	if currentpage > 4 then
		response.write "<a href="""&Searchstr&"&page=1"">[1]</a> ..."
	end if
	if Pcount>currentpage+3 then
		endpage=currentpage+3
	else
		endpage=Pcount
	end if
	for i=currentpage-3 to endpage
	if not i<1 then
		if i = clng(currentpage) then
        response.write " <font color=""red"">["&i&"]</font>"
		else
        response.write " <a href="""&Searchstr&"&page="&i&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
		response.write "... <a href="""&Searchstr&"&page="&Pcount&""">["&Pcount&"]</a>"
	end if
%>
</td>
</tr>
<%
End If
Rs.Close
Set Rs = Nothing
%>
</table>
<%
End Sub


Sub Boke_Update()
	If Request.QueryString("t")<>"" Then
		Select Case Request.QueryString("t")
			Case "1"
				Boke_Update_Users()
			Case "2"
				Boke_Update_SysCats()
			Case "3"
				Boke_Update_ChatCats()
			Case "4"
				Boke_Update_System()
			Case "5"
				Boke_Update_UserInfo()
		End Select
		Exit Sub
	End If
%>
	<table width="100%" border="0" cellspacing="1" cellpadding="5"  align="center" class="tableBorder">
	<tr> 
	<th width="100%" colspan="2" height="25" align="left" ID="TableTitleLink">
	博客信息更新
	</th>
	</tr>
	<tr>
	<td class="td2" colspan="2">
	说明:
	</td>
	</tr>
	<tr>
	<td width="10%" class="td2">
	<input type="button" name="act" value="博客用户统计" class="button" onclick="location.href='?s=9&t=1'"/>
	</td>
	<td align="left" width="90%" class="td1">重新统计当前博客用户总数</td>
	</tr>
	<tr>
	<td width="10%" class="td2">
	<input type="button" name="act" value="博客索引数据统计" class="button" onclick="location.href='?s=9&t=2'"/>
	</td>
	<td align="left" width="90%" class="td1">重新统计当前博客索引用户数，帖子数信息</td>
	</tr>
	<tr>
	<td width="10%" class="td2">
	<input type="button" name="act" value="博客话题数据统计" class="button" onclick="location.href='?s=9&t=3'"/>
	</td>
	<td align="left" width="90%" class="td1">重新统计当前博客话题帖子数信息</td>
	</tr>
	<tr>
	<td width="10%" class="td2">
	<input type="button" name="act" value="博客总数据统计" class="button" onclick="location.href='?s=9&t=4'"/>
	</td>
	<td align="left" width="90%" class="td1">重新统计当前博客帖子数信息</td>
	</tr>
	<tr>
	<td width="10%" class="td2">
	<input type="button" name="act" value="博客用户数据总更新" class="button" onclick="location.href='?s=9&t=5'"/>
	</td>
	<td align="left" width="90%" class="td1">更新所有博客用户的相关数据,包括文章数、评论数以及博客用户首页缓存数据等</td>
	</tr>
	</table>
<%

End Sub

Sub Boke_Update_Users()
	Dim AllUsers
	AllUsers = DvBoke.Execute("Select Count(*) From Dv_Boke_User")(0)
	DvBoke.Execute("update Dv_Boke_System set S_UserNum = "&AllUsers)
	DvBoke.LoadSetup(1)
	Dv_suc("博客用户统计完成,当前共有"&AllUsers&"位博客用户!")
End Sub

Sub Boke_Update_SysCats()
	Dim SucMsg,Rs
	Dim uCatNum,TopicNum,PostNum,TodayNum,LastUpTime
	Dim Nodes,ChildNode
	Set Nodes = DvBoke.SysCat.selectNodes("rs:data/z:row")
	If Nodes.Length>0 Then
		For Each ChildNode in Nodes
			uCatNum = DvBoke.Execute("Select Count(*) From Dv_Boke_User where SysCatID="&ChildNode.getAttribute("scatid"))(0)
			TopicNum= DvBoke.Execute("Select Sum(TopicNum) From Dv_Boke_User where SysCatID="&ChildNode.getAttribute("scatid"))(0)
			PostNum= DvBoke.Execute("Select Sum(PostNum) From Dv_Boke_User where SysCatID="&ChildNode.getAttribute("scatid"))(0)
			TodayNum= DvBoke.Execute("Select Sum(TodayNum) From Dv_Boke_User where SysCatID="&ChildNode.getAttribute("scatid"))(0)
			Set Rs = DvBoke.Execute("Select top 1 LastUpTime From Dv_Boke_User where SysCatID="&ChildNode.getAttribute("scatid")&" order by LastUpTime desc")
			If Rs.Eof Then
				LastUpTime = Now()
			Else
				LastUpTime = Rs(0)
			End If
			Rs.Close
			If IsNull(TopicNum) Then
				TopicNum = 0
			End If
			If IsNull(PostNum) Then
				PostNum = 0
			End If
			If IsNull(TodayNum) Then
				TodayNum = 0
			End If
			DvBoke.Execute("update Dv_Boke_SysCat set uCatNum="&uCatNum&",TopicNum="&TopicNum&",PostNum="&PostNum&",TodayNum="&TodayNum&",LastUpTime='"&LastUpTime&"' where sCatID="&ChildNode.getAttribute("scatid"))
			SucMsg = SucMsg &"<li>"&ChildNode.getAttribute("scattitle")&" :共有"&uCatNum&"用户，"&TopicNum&"篇文章，"&PostNum&"篇评论，今日发表共"&TodayNum&"篇，最后更新时间："&LastUpTime&"</li>"
		Next
	End If
	DvBoke.LoadSetup(1)
	Dv_suc(SucMsg)
End Sub

Sub Boke_Update_ChatCats()
	Dim SucMsg,Rs,DayStr
	Dim TopicNum,PostNum,TodayNum,LastUpTime
	Dim Nodes,ChildNode
	If Dv_Boke_DataBase = 1 Then
		DayStr = "d"
	Else
		DayStr = "'d'"
	End If
	Set Nodes = DvBoke.SysChatCat.selectNodes("rs:data/z:row")
	If Nodes.Length>0 Then
		For Each ChildNode in Nodes
			TopicNum= DvBoke.Execute("Select Count(TopicID) From Dv_Boke_Topic where sCatID="&ChildNode.getAttribute("scatid"))(0)
			PostNum= DvBoke.Execute("Select Count(PostID) From Dv_Boke_Post where ParentID>0 and  sCatID="&ChildNode.getAttribute("scatid"))(0)
			TodayNum= DvBoke.Execute("Select Count(PostID) From Dv_Boke_Post where sCatID="&ChildNode.getAttribute("scatid")&" and DateDiff("&DayStr&",JoinTime,"&bSqlNowString&") = 0")(0)

			Set Rs = DvBoke.Execute("Select top 1 JoinTime From Dv_Boke_Post where sCatID="&ChildNode.getAttribute("scatid")&" order by JoinTime desc")
			If Rs.Eof Then
				LastUpTime = Now()
			Else
				LastUpTime = Rs(0)
			End If
			Rs.Close

			DvBoke.Execute("update Dv_Boke_SysCat set  TopicNum="&TopicNum&",PostNum="&PostNum&",TodayNum="&TodayNum&",LastUpTime='"&LastUpTime&"' where sCatID="&ChildNode.getAttribute("scatid"))
			SucMsg = SucMsg &"<li>"&ChildNode.getAttribute("scattitle")&" :"&TopicNum&"篇文章，"&PostNum&"篇评论，今日发表共"&TodayNum&"篇，最后更新时间："&LastUpTime&"</li>"
		Next
	End If
	DvBoke.LoadSetup(1)
	Dv_suc(SucMsg)
End Sub

Sub Boke_Update_System()
	Dim SucMsg,Rs,DayStr
	Dim S_LastPostTime,S_TopicNum,S_PhotoNum,S_FavNum,S_TodayNum,S_PostNum
	If Dv_Boke_DataBase = 1 Then
		DayStr = "d"
	Else
		DayStr = "'d'"
	End If
	S_TopicNum = DvBoke.Execute("Select Count(*) From [Dv_Boke_Topic] Where sType=0")(0)
	S_PhotoNum = DvBoke.Execute("Select Count(*) From [Dv_Boke_Topic] Where sType=4")(0)
	S_FavNum = DvBoke.Execute("Select Count(*) From [Dv_Boke_Topic] Where sType=1")(0)
	S_PostNum = DvBoke.Execute("Select Count(*) From [Dv_Boke_Post] Where ParentID>0")(0)
	S_TodayNum = DvBoke.Execute("Select Count(*) From [Dv_Boke_Post] Where DateDiff("&DayStr&",JoinTime,"&bSqlNowString&") = 0")(0)
	Set Rs = DvBoke.Execute("Select Top 1 JoinTime From [Dv_Boke_Post] order by JoinTime desc")
	If Rs.Eof Then
		S_LastPostTime = Now()
	Else
		S_LastPostTime = Rs(0)
	End If
	DvBoke.Execute("update Dv_Boke_System set S_LastPostTime='"&S_LastPostTime&"',S_TopicNum="&S_TopicNum&",S_PhotoNum="&S_PhotoNum&",S_FavNum="&S_FavNum&",S_TodayNum="&S_TodayNum&",S_PostNum="&S_PostNum)
	
	SucMsg = "<li>博客系统总信息： :文章共"&S_TopicNum&"篇，相册共"&S_PhotoNum&"篇，收藏共"&S_FavNum&"篇，评论共"&S_PostNum&"篇，今日发表共"&S_TodayNum&"篇，最后更新时间："&S_LastPostTime&"</li>"
	
	DvBoke.LoadSetup(1)
	Dv_suc(SucMsg)
End Sub

Sub Boke_Update_UserInfo()
	Dim BokeUserCount,Rs,i
	BokeUserCount = DvBoke.Execute("Select Count(*) From [Dv_Boke_User]")(0)
	If BokeUserCount = "" Or IsNull(BokeUserCount) Then Exit Sub
%>
<table cellpadding="0" cellspacing="0" border="0" width="100%" class="tableBorder" align=center>
<tr><td colspan=2 class=td1>
下面开始更新论坛用户资料，预计本次共有<%=BokeUserCount%>个用户需要更新
<table width="100%" border="0" cellspacing="1" cellpadding="1">
<tr> 
<td bgcolor=000000>
<table width="100%" border="0" cellspacing="0" cellpadding="1">
<tr> 
<td bgcolor=ffffff height=9><img src="skins/default/bar/bar3.gif" width=0 height=16 id=img2 name=img2 align=absmiddle></td></tr></table>
</td></tr></table> <span id=txt2 name=txt2 style="font-size:9pt">0</span><span style="font-size:9pt">%</span></td></tr>
</table>
<%
	Dim uTopicNum,uFavNum,uPostNum,uTodayNum,uPhotoNum,uXmlData,DayStr,SucMsg,iBokeCat
	Dim Node,XmlDoc,NodeList,ChildNode,BokeBody
	Dim tRs,Sql
	Dim DvCode
	Set DvCode = New DvBoke_UbbCode
	If Dv_Boke_DataBase = 1 Then
		DayStr = "d"
	Else
		DayStr = "'d'"
	End If
	i = 0
	Set Rs = DvBoke.Execute("Select UserID,BokeName,XmlData From [Dv_Boke_User]")
	Do While Not Rs.Eof
		i = i + 1
		uTopicNum = DvBoke.Execute("Select Count(*) From Dv_Boke_Topic Where sType=0 And UserID = " & Rs(0))(0)
		uFavNum = DvBoke.Execute("Select Count(*) From Dv_Boke_Topic Where sType=1 And UserID = " & Rs(0))(0)
		uPhotoNum = DvBoke.Execute("Select Count(*) From Dv_Boke_Topic Where sType=4 And UserID = " & Rs(0))(0)
		uTodayNum = DvBoke.Execute("Select Count(*) From Dv_Boke_Post Where BokeUserID = " & Rs(0) & " And DateDiff("&DayStr&",JoinTime,"&bSqlNowString&") = 0")(0)
		uPostNum = DvBoke.Execute("Select Count(*) From Dv_Boke_Post Where ParentID>0 And BokeUserID = " & Rs(0))(0)
		'目前仅更新首页主题列表数据
		Set iBokeCat = Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument")
		If Rs(2)="" Or IsNull(Rs(2)) Then
			iBokeCat.Load(Server.MapPath(DvBoke.Cache_Path &"usercat.config"))
		Else
			If Not iBokeCat.LoadXml(Rs(2)) Then
				iBokeCat.Load(Server.MapPath(DvBoke.Cache_Path &"usercat.config"))
			End If
		End If
		Set Node = iBokeCat.selectNodes("xml/boketopic")
		If Not (Node Is Nothing) Then
			For Each NodeList in Node
				iBokeCat.DocumentElement.RemoveChild(NodeList)
			Next
		End If
		Set Node=iBokeCat.createNode(1,"boketopic","")
		Set XmlDoc=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument")
		If Not IsNumeric(DvBoke.BokeSetting(6)) Then DvBoke.BokeSetting(6) = "10"
		Sql = "Select Top "&DvBoke.BokeSetting(6)&" TopicID,CatID,sCatID,UserID,UserName,Title,TitleNote,PostTime,Child,Hits,IsView,IsLock,sType,LastPostTime,IsBest,S_Key,Weather From [Dv_Boke_Topic] Where UserID="&Rs(0)&" and sType <>2 order by PostTime desc"
		Set tRs = DvBoke.Execute(LCase(Sql))
		If Not tRs.Eof Then
			tRs.Save XmlDoc,1
			XmlDoc.documentElement.RemoveChild(XmlDoc.documentElement.selectSingleNode("s:Schema"))
			Set ChildNode = XmlDoc.documentElement.selectNodes("rs:data/z:row")
			For Each NodeList in ChildNode
				If tRs("TitleNote")="" Or IsNull(tRs("TitleNote")) Then
					BokeBody = DvBoke.Execute("Select Content From Dv_Boke_Post Where ParentID=0 and Rootid="&tRs(0))(0)
					If Len(BokeBody) > 250 Then
						BokeBody = SplitLines(BokeBody,DvBoke.BokeSetting(2))
					End If
				Else
					BokeBody = tRs("TitleNote")
				End If
				BokeBody = DvCode.UbbCode(BokeBody) & "...<br/>[<a href=""boke.asp?"&Rs(1)&".showtopic."&tRs("TopicID")&".html"">阅读全文</a>]"
				NodeList.attributes.getNamedItem("titlenote").text = BokeBody
				NodeList.attributes.getNamedItem("posttime").text = tRs("PostTime")
				NodeList.attributes.getNamedItem("lastposttime").text = tRs("LastPostTime")
				tRs.MoveNext
			Next
			Set ChildNode=XmlDoc.documentElement.selectSingleNode("rs:data")
			Node.appendChild(ChildNode)
		End If
		tRs.Close
		Set tRs = Nothing
		iBokeCat.documentElement.appendChild(Node)
		'End
		DvBoke.Execute("Update Dv_Boke_User set XmlData = '"&Replace(iBokeCat.documentElement.xml,"'","''")&"',TopicNum="&uTopicNum&",FavNum="&uFavNum&",PhotoNum="&uPhotoNum&",TodayNum="&uTodayNum&",PostNum="&uPostNum&" where UserID="&Rs(0))
		Response.Write "<script>img2.width=" & Fix((i/BokeUserCount) * 400) & ";" & VbCrLf
		Response.Write "txt2.innerHTML=""更新完"&Rs(1)&"的数据，正在更新下一个用户数据，" & FormatNumber(i/BokeUserCount*100,4,-1) & """;" & VbCrLf
		Response.Write "img2.title=""" & Rs(1) & "(" & i & ")"";</script>" & VbCrLf
		Response.Flush
	Rs.MoveNext
	Loop
	Rs.Close
	Set Rs=Nothing
	Response.Write "<script>img2.width=400;txt2.innerHTML=""100"";</script>"

	SucMsg = "<li>更新所有用户数据成功！</li>"	
	Dv_suc(SucMsg)
End Sub

Sub Boke_SysNews()
	Dim Bodystr,Bodystr1,Node,Node1,createCDATASection
	Set Node = DvBoke.SystemDoc.documentElement.selectSingleNode("/bokesystem/topnews")
	If Node Is Nothing Then
		Set Node = DvBoke.SystemDoc.createNode(1,"topnews","")
		DvBoke.SystemDoc.documentElement.appendChild(Node)
	End If
	Bodystr = Node.text
	Set Node1 = DvBoke.SystemDoc.documentElement.selectSingleNode("/bokesystem/managenews")
	If Node1 Is Nothing Then
		Set Node1 = DvBoke.SystemDoc.createNode(1,"managenews","")
		DvBoke.SystemDoc.documentElement.appendChild(Node1)
	End If
	Bodystr1 = Node1.text
	'Response.Write Bodystr1
	%>
	<table width="100%" border="0" cellspacing="1" cellpadding="5"  align="center" class="tableBorder">
	<tr> 
	<th width="100%" colspan="2" height="25" align="left" ID="TableTitleLink">
	首页公告信息
	</th>
	</tr>
	<%
	If Request.Form("act") = "save" Then
		Node.text = Request.Form("boketopnews")
		DvBoke.SaveSystemCache()
		Manage_Suc "您成功编辑了博客首页公告信息","2","?s=4"
	ElseIf Request.Form("act") = "save1" Then
		Node1.text = Request.Form("bokemanagenews")
		DvBoke.SaveSystemCache()
		Manage_Suc "您成功编辑了个人博客管理首页系统通知信息","2","?s=4"
	Else
	%>
	<form method="post" action="?s=4">
	<tr>
	<td class="td2" width="20%">
	编辑信息内容:
	</td>
	<td class="td2"width="80%">
	<input type="hidden" name="act" value="save">
	<textarea name="boketopnews" rows="0" cols="0" style="display:none;"><%=Server.Htmlencode(Bodystr)%></textarea>
	<iframe id="EditFrame" src="boke/edit_plus/FCKeditor/editor/fckeditor.html?InstanceName=boketopnews&Toolbar=Default" width="100%" height="200" frameborder="no" scrolling="no"></iframe>
	<input type="submit" value="提交更改" class="button">
	</td>
	</tr>
	</form>
	<tr> 
	<td width="100%" class="td1" colspan="2" height="25" align="left">&nbsp;
	</td>
	</tr>
	<tr> 
	<th width="100%" colspan="2" height="25" align="left" ID="TableTitleLink">
	个人博客管理首页系统通知
	</th>
	</tr>
	<form method="post" action="?s=4">
	<tr>
	<td class="td2" width="20%">
	编辑信息内容:
	</td>
	<td class="td2"width="80%">
	<input type="hidden" name="act" value="save1">
	<textarea name="bokemanagenews" rows="0" cols="0" style="display:none;"><%=Server.Htmlencode(Bodystr1)%></textarea>
	<iframe id="EditFrame" src="boke/edit_plus/FCKeditor/editor/fckeditor.html?InstanceName=bokemanagenews&Toolbar=Default" width="100%" height="200" frameborder="no" scrolling="no"></iframe>
	<input type="submit" value="提交更改" class="button">
	</td>
	</tr>
	</form>
	<%
	End If
	%>
	</table>
	<%
End Sub

Sub Boke_Skins()
Dim Rs,Sql
Dim S_ID,S_Name,S_Builder,S_Path,S_ViewPic,S_Info
S_ID = 0

If Request("act")="save" Then
	S_ID = DvBoke.CheckNumeric(Request.Form("S_ID"))
	If Request.Form("S_Name") = "" or Len(Request.Form("S_Name"))>50 Then
		ErrMsg = "模板名称不能为空或超出50个字符!"
		Dvbbs_error()
		Exit Sub
	End If
	If Request.Form("S_Path")="" or Len(Request.Form("S_Path"))>150 Then
		ErrMsg = "模板路径不能为空或超出150个字符!"
		Dvbbs_error()
		Exit Sub
	End If
	If Len(Request.Form("S_Info"))>250 Then
		ErrMsg = "模板信息及说明不能超出250个字符!"
		Dvbbs_error()
		Exit Sub
	End If
	Sql  = "Select S_ID,S_SkinName,S_Path,S_ViewPic,S_Info,S_Builder From Dv_Boke_Skins where S_ID="&S_ID
	If Not IsObject(Boke_Conn) Then Boke_ConnectionDatabase
	Set Rs=Dvbbs.iCreateObject("Adodb.RecordSet")
	Rs.Open Sql,Boke_Conn,1,3
	If Rs.Eof and Rs.Bof Then
		Rs.AddNew
	End If
	Rs("S_SkinName") = Request.Form("S_Name")
	Rs("S_Path") = Request.Form("S_Path")
	Rs("S_ViewPic") = Request.Form("S_ViewPic")
	Rs("S_Info") = Request.Form("S_Info")
	Rs("S_Builder") = Request.Form("S_Builder")
	Rs.Update
	Rs.Close
	Set Rs = Nothing
	Dv_suc("模板数据保存成功")
	Exit Sub
ElseIf Request("act") = "edit" Then
	S_ID = DvBoke.CheckNumeric(Request("S_ID"))
	If S_ID>0 Then
		Sql  = "Select S_ID,S_SkinName,S_Path,S_ViewPic,S_Info,S_Builder From Dv_Boke_Skins where S_ID="&S_ID
		Set Rs = DvBoke.Execute(Sql)
		If Not Rs.Eof Then
			S_ID = Rs(0)
			S_Name = Rs(1)
			S_Builder = Rs(5)
			S_Path = Rs(2)
			S_ViewPic = Rs(3)
			S_Info = Rs(4)&""
		End If
		Rs.Close
		Set Rs = Nothing
	End If
ElseIf Request("act") = "addsys" Then
	S_ID = DvBoke.CheckNumeric(Request("S_ID"))
	If S_ID>0 Then
		Sql  = "Select S_ID,S_SkinName From Dv_Boke_Skins where S_ID="&S_ID
		Set Rs = DvBoke.Execute(Sql)
		If Not Rs.Eof Then
			S_Name = Rs(1)
			DvBoke.Execute("Update Dv_Boke_System Set SkinID = "&S_ID)
			DvBoke.LoadSetup(1)
			Dv_suc("已将模板["& S_Name &"]设为系统默认模板!")
		End If
		Rs.Close
		Set Rs = Nothing
		Exit Sub
	End If
ElseIf Request("act")="del" Then
	Dim NewS_ID
	S_ID = DvBoke.CheckNumeric(Request("S_ID"))
	If S_ID>0 Then
		If Clng(DvBoke.System_Node.getAttribute("skinid")) = S_ID Then
			ErrMsg = "不能删除系统默认模板，请重新选取!"
			Dvbbs_error()
			Exit Sub
		End If
		Sql  = "Select S_ID,S_SkinName From Dv_Boke_Skins where S_ID="&S_ID
		Set Rs = DvBoke.Execute(Sql)
		If Not Rs.Eof Then
			S_Name = Rs(1)
			NewS_ID = DvBoke.Execute("Select Top 1 S_ID From Dv_Boke_Skins Order by S_ID")(0)
			If NewS_ID>0 Then
				DvBoke.Execute("Update Dv_Boke_User Set SkinID = "&NewS_ID&" where SkinID="&S_ID)
				DvBoke.Execute("Delete from Dv_Boke_Skins where S_ID="&S_ID)
				Dv_suc("模板["& S_Name &"]删除成功!")
			Else
				ErrMsg = "请添加可用模板后再进行删除操作!"
				Dvbbs_error()
			End If
		Else
			ErrMsg = "模板的不存在，删除失败!"
			Dvbbs_error()
		End If
		Rs.Close
		Set Rs = Nothing
	Else
		ErrMsg = "模板的参数错误，删除失败!"
		Dvbbs_error()
	End If
	Exit Sub
End If
%>
<table width="100%" border="0" cellspacing="1" cellpadding="5"  align="center" class="tableBorder">
<tr> 
<th width="100%" colspan="2" height="25" align="left" ID="TableTitleLink">
模板信息管理
</th>
</tr>
<form method="post" action="?s=7&act=save">
<tr>
<td class="td2" width="20%">模板名称</td>
<td class="td2"width="80%">
<input type="text" name="S_Name" value="<%=S_Name%>">
</td>
</tr>
<tr>
<td class="td2">提供者</td>
<td class="td2">
<input type="text" name="S_Builder" value="<%=S_Builder%>">
</td>
</tr>
<tr>
<td class="td2">模板路径</td>
<td class="td2">
<input type="text" name="S_Path" size="50" value="<%=S_Path%>">
</td>
</tr>
<tr>
<td class="td2">演示图片</td>
<td class="td2">
<input type="text" name="S_ViewPic" size="50" value="<%=S_ViewPic%>">
</td>
</tr>
<tr>
<td class="td2">信息及说明</td>
<td class="td2">
<textarea name="S_Info" rows="5" cols="50"><%=Server.Htmlencode(S_Info)%></textarea>
</td>
</tr>
<tr>
<td class="td2" colspan="2" align="center">
<input type="hidden" name="S_ID" value="<%=S_ID%>">
<input type="submit" value="保存" class="button">
</td>
</tr>
</form>
</table>
<br/>
<table width="100%" border="0" cellspacing="1" cellpadding="5"  align="center" class="tableBorder">
<tr> 
<th width="100%" colspan="5" height="25" align="left" ID="TableTitleLink">
模板信息列表
</th>
</tr>
<tr>
<td class="bodytitle" height="24" width="15%">
演示
</td>
<td class="bodytitle" height="24" width="20%">
名称/ 路径
</td>
<td class="bodytitle" height="24" width="15%">
提供者
</td>
<td class="bodytitle" height="24" width="30%">
信息及说明
</td>
<td class="bodytitle" height="24" width="20%">
操作
</td>
</tr>
<%
	
	Dim CurrentPage,Page_Count,Pcount,i
	Dim TotalRec,EndPage

	CurrentPage=Request("page")
	If CurrentPage="" Or Not IsNumeric(CurrentPage) Then
		CurrentPage=1
	Else
		CurrentPage=Clng(CurrentPage)
		If Err Then
			CurrentPage=1
			Err.Clear
		End If
	End If
	Sql = "Select S_ID,S_SkinName,S_Path,S_ViewPic,S_Info,S_Builder From Dv_Boke_Skins order by S_id Desc"
	If Not IsObject(Boke_Conn) Then Boke_ConnectionDatabase
	Set Rs=Dvbbs.iCreateObject("Adodb.RecordSet")
	Rs.Open Sql,Boke_Conn,1,1
	If Not (Rs.Eof And Rs.Bof) Then
		Rs.PageSize = 30
		Rs.AbsolutePage=CurrentPage
		Page_Count=0
		TotalRec=Rs.RecordCount
		While (Not Rs.Eof) And (Not Page_Count = 30)

%>
<tr>
<td class="td1">
<%
If Rs(3)<>"" Then
	Response.Write "<img src=""../"&Rs(3)&""" border=""1"" width=""80"" height=""60"">"
Else
	Response.Write "<img src=""../boke/images/viewskins_bck.png"" border=""1"" width=""80"" height=""60"">"
End If
%>
</td>
<td class="td1">
<b><%=Rs(1)%></b>
<br/><u><%=Rs(2)%></u>
</td>
<td class="td2">
<%=Rs(5)%>&nbsp;
</td>
<td class="td1">
<%=Rs(4)%>&nbsp;
</td>
<td class="td2">
<a href="?s=7&act=edit&s_id=<%=Rs(0)%>">编辑</a> | <a href="?s=7&act=del&s_id=<%=Rs(0)%>">删除</a>
 | 
<%If Clng(DvBoke.System_Node.getAttribute("skinid")) = Rs(0) Then%>
 <font color="red">系统默认</font>
<%Else%>
 <a href="?s=7&act=addsys&s_id=<%=Rs(0)%>">设为默认</a>
<%End If%>
</td>
</tr>
<%
			Page_Count = Page_Count + 1
		Rs.MoveNext
		Wend
		Pcount=Rs.PageCount
%>
<tr><td colspan=5 class=td1>共有<%=TotalRec%>条记录，分页：
<%
	Dim Searchstr
	Searchstr = "?s=7"
	if currentpage > 4 then
		response.write "<a href="""&Searchstr&"&page=1"">[1]</a> ..."
	end if
	if Pcount>currentpage+3 then
		endpage=currentpage+3
	else
		endpage=Pcount
	end if
	for i=currentpage-3 to endpage
	if not i<1 then
		if i = clng(currentpage) then
        response.write " <font color=""red"">["&i&"]</font>"
		else
        response.write " <a href="""&Searchstr&"&page="&i&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
		response.write "... <a href="""&Searchstr&"&page="&Pcount&""">["&Pcount&"]</a>"
	end if
%>
</td>
</tr>
<%
End If
Rs.Close
Set Rs = Nothing
%>
</table>
<%
End Sub

'博客系统栏目管理
Sub Boke_SysCat()
	Dim Rs,i,TableClass,t,tStr
	t = Request("t")
	If t = "" Or Not IsNumeric(t) Then t = 1
	t = Cint(t)
	If t = 1 Then
		tStr = "栏目"
	Else
		tStr = "话题"
	End If
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=6 height=25 align=left ID="TableTitleLink">&nbsp;&nbsp;博客系统<%=tStr%>管理 | <a href="?s=1&t=<%=t%>&Action=Add">添加<%=tStr%></a>
</th>
</tr>
<%
If Request("Action")="Add" Then
%>
<FORM METHOD=POST ACTION="?s=1&t=<%=t%>&Action=Save">
<tr align=center>
<td class="td2" height=24 colspan=6>
<B>添加博客系统<%=tStr%></B>
</td>
</tr>
<tr>
<td class="td1" height=24 width="35%" align=right>
<B><%=tStr%>名称</B>：
</td>
<td class="td1" width="65%">
<input type="text" name="Title" size="50">
</td>
</tr>
<tr>
<td class="td2" height=24 width="35%" align=right>
<B><%=tStr%>说明</B>：
</td>
<td class="td2" width="65%">
<textarea name="Note" cols="50" rows="5"></textarea>
</td>
</tr>
<tr align=center>
<td class="td1" height=24 colspan=6>
<input type=submit name=submit value="添加博客<%=tStr%>" class="button">
</td>
</tr>
</FORM>
<%
ElseIf Request("Action")="Save" Then
	If Request("Title")="" Then
		Manage_Err "请填写"&tStr&"的名称","6","?s=1&t="&t&""
		Exit Sub
	End If
	If t=1 Then
		DvBoke.Execute("Insert Into Dv_Boke_SysCat (sCatTitle,sCatNote) Values ('"&Replace(Request("Title"),"'","''")&"','"&Replace(Request("Note"),"'","''")&"')")
	Else
		DvBoke.Execute("Insert Into Dv_Boke_SysCat (sCatTitle,sCatNote,stype) Values ('"&Replace(Request("Title"),"'","''")&"','"&Replace(Request("Note"),"'","''")&"',1)")
	End If
	
	Manage_Suc "您成功添加了博客"&tStr&"","6","?s=1&t="&t&""
	DvBoke.LoadSetup(1)
ElseIf Request("Action")="Edit" Then
	If Request("ID") = "" Or Not IsNumeric(Request("ID")) Then
		Manage_Err "非法的"&tStr&"参数","6","?s=1&t="&t&""
		Exit Sub
	End If
	Set Rs = DvBoke.Execute("Select * From Dv_Boke_SysCat Where sCatID = " & Request("ID"))
	If Rs.Eof And Rs.Bof Then
		Manage_Err "非法的"&tStr&"参数","6","?s=1&t="&t&""
		Rs.Close
		Set Rs=Nothing
		Exit Sub
	End If
%>
<FORM METHOD=POST ACTION="?s=1&t=<%=t%>&Action=SaveEdit">
<input type=hidden value="<%=Request("ID")%>" name="ID">
<tr align=center>
<td class="td2" height=24 colspan=6>
<B>编辑博客系统<%=tStr%></B>
</td>
</tr>
<tr>
<td class="td1" height=24 width="35%" align=right>
<B><%=tStr%>名称</B>：
</td>
<td class="td1" width="65%">
<input type="text" name="Title" size="50" value="<%=Server.HtmlEncode(Rs("sCatTitle"))%>">
</td>
</tr>
<tr>
<td class="td2" height=24 width="35%" align=right>
<B><%=tStr%>说明</B>：
</td>
<td class="td2" width="65%">
<textarea name="Note" cols="50" rows="5"><%=Server.HtmlEncode(Rs("sCatNote")&"")%></textarea>
</td>
</tr>
<tr align=center>
<td class="td1" height=24 colspan=6>
<input type=submit name=submit value="编辑博客<%=tStr%>" class="button">
</td>
</tr>
</FORM>
<%
ElseIf Request("Action")="SaveEdit" Then
	If Request("Title")="" Then
		Manage_Err "请填写"&tStr&"的名称","6","?s=1&t="&t&""
		Exit Sub
	End If
	If Request("ID") = "" Or Not IsNumeric(Request("ID")) Then
		Manage_Err "非法的"&tStr&"参数","6","?s=1&t="&t&""
		Exit Sub
	End If
	DvBoke.Execute("Update Dv_Boke_SysCat Set sCatTitle='"&Replace(Request("Title"),"'","''")&"',sCatNote='"&Replace(Request("Note"),"'","''")&"' Where sCatID = " & Request("ID"))
	Manage_Suc "您成功编辑了博客"&tStr&"","6","?s=1&t="&t&""
	DvBoke.LoadSetup(1)
ElseIf Request("Action")="Del" Then
	If Request("ID") = "" Or Not IsNumeric(Request("ID")) Then
		Manage_Err "非法的"&tStr&"参数","6","?s=1&t="&t&""
		Exit Sub
	End If
	DvBoke.Execute("Delete From Dv_Boke_SysCat Where sCatID = " & Request("ID"))
	Manage_Suc "您成功删除了博客"&tStr&"","6","?s=1&t="&t&""
	DvBoke.LoadSetup(1)
Else
%>
<tr>
<td class="td1" colspan=6 height=25>
<B>说明</B>：点击用户数可查看此分栏的用户博客列表
</td>
</tr>
<tr align=center>
<td class="bodytitle" height=24>
<B><%=tStr%></B>
</td>
<td class="bodytitle">
<B>今日</B>
</td>
<td class="bodytitle">
<B>文章</B>
</td>
<td class="bodytitle">
<B>回复</B>
</td>
<td class="bodytitle">
<B>用户数</B>
</td>
<td class="bodytitle" align=left>
<B>操作</B>
</td>
</tr>
<%
	i = 0
	'TableClass = "td1"
	Set Rs=DvBoke.Execute("Select * From Dv_Bokequest("Note"),"'","''")&"' Where sCatID = " & Request("ID"))
	Manage_Suc "鎮ㄦ垚鍔熺紪杈戜簡鍗氬