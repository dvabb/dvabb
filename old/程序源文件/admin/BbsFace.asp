<!--#include file="../conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
Head()
Dim admin_flag
Dim StyleID,StyleName,Style_Pic,Stype
Dim Forum_emotNum,Forum_userfaceNum,Forum_PostFaceNum
Dim Forum_PostFace,Forum_userface,Forum_emot
Dim face_id,Count
Dim newnum,newfilename
Dim bbspicmun,bbspicurl,picfilename,actname,connfile,upconfig
Dim TempForum_PostFace,TempForum_userface,TempForum_emot
admin_flag=",38,"
CheckAdmin(admin_flag)
Stype = Dvbbs.CheckNumeric(Request("Stype"))
'Stype:1=表情，2=心情em，3=头像
If Stype=0 Then Stype=4
If 2=Stype Then
	response.redirect "BbsFaceX.asp"
	response.End 
End If 
StyleID = Dvbbs.CheckNumeric(Request("StyleID"))
If StyleID=0 Or StyleID="" Then StyleID=Dvbbs.CacheData(17,0)

If Application(Dvbbs.CacheName &"_style").documentElement.selectSingleNode("style[@id='"& StyleID &"']") Is Nothing Then
	If Not Application(Dvbbs.CacheName &"_style").documentElement.selectSingleNode("style/@id") Is Nothing Then
		StyleID = Application(Dvbbs.CacheName &"_style").documentElement.selectSingleNode("style/@id").text
		StyleName = Application(Dvbbs.CacheName &"_style").documentElement.selectSingleNode("style/@type").text
	Else
		Response.Write "模板数据无法提取,请检查或重新导入"
		Response.End
	End If
Else
	StyleID = Application(Dvbbs.CacheName &"_style").documentElement.selectSingleNode("style[@id='"& StyleID &"']/@id").text
	StyleName = Application(Dvbbs.CacheName &"_style").documentElement.selectSingleNode("style[@id='"& StyleID &"']/@type").text
End If

Dim StyleFolder,FilePath,FileName
StyleFolder = Application(Dvbbs.CacheName &"_style").documentElement.selectSingleNode("style[@id='"& StyleID &"']/@folder").text
FilePath = "../Resource/"& StyleFolder &"/"
FileName = FilePath &"pub_FaceEmot.htm"

'Response.Write StyleID
GetNum()

If Founderr=false Then
	Select Case Stype
	case 1
	'skins/default/topicface/face1.gif
		bbspicmun=Forum_PostFaceNum-1
		If not isarray(Forum_PostFace) Then
			bbspicurl="../Skins/default/topicface/"
		Else
			bbspicurl="../" & Forum_PostFace(0)
		End If
		connfile=Forum_PostFace
		actname="发贴表情图片"
		picfilename="face"
	case 2
	'Skins/Default/emot/em01.gif	'Forum_emot
        bbspicmun=Forum_emotNum-1
		If not isarray(Forum_emot) Then
			bbspicurl="../Skins/Default/emot/"
		Else
			bbspicurl="../" & Forum_emot(0)
		End If
		connfile=Forum_emot
		actname="发贴心情图片"
		picfilename="em"
	case 3
	'Images/userface/image1.gif
		bbspicmun=Forum_userfaceNum-1
		If not isarray(Forum_userface) Then
			bbspicurl="../Images/userface/"
		Else
			bbspicurl="../" & Forum_userface(0)
		End If
		connfile=Forum_userface
		actname="注册头像"
		picfilename="image"
	case else
	'Images/userface/image1.gif
		bbspicmun=Forum_userfaceNum-1
		If not isarray(Forum_userface) Then
			bbspicurl="../Images/userface/"
		Else
			bbspicurl="../" & Forum_userface(0)
		End If
		connfile=Forum_userface
		actname=""
		picfilename="image"
	End Select

	if trim(Request("newfilename"))<>"" then
		newfilename=trim(request("newfilename"))
	else
		newfilename=picfilename
	end if 

	if bbspicmun<0 then 
	count=1
	else
	count=bbspicmun+1
	end if

	if REQUEST("Newnum")<>"" and request("Newnum")<>0 then
		newnum=REQUEST("Newnum")
	else
		newnum=0
	end if

	if request("Submit")="保存设置" then
		call saveconst()
	elseif request("Submit")="恢复默认设置" then
		call savedefault()
	ElseIf request("Submit")="恢复默认总设置" then
		Stype=4
		call savedefault()
	else
		call consted()
	end if
End If
if Founderr then dvbbs_error()
Footer()
sub consted()
dim sel
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
<tr> 
<td height="23" colspan="4" ><B>说明</B>：<br>①、以下图片均保存于论坛<%=bbspicurl%>目录中，如要更换也请将图片放于该目录<br>②、右边复选框为删除选项，如果选择后点保存设置，则删除相应图片<BR>③、如仅仅修改文件名，可在修改相应选项后直接点击保存设置而不用选择右边复选框
</td>
</tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th colspan="4"><%=actname%>管理设置 （目前共有<%=count%>个<%=actname%>图片在文件夹：<%=bbspicurl%>）</th>
</tr>
<tr>
<td width="20%" align="left" class="forumrow">请选择相关模板： </td>
<form method="post" action="">
<td width="80%" align="left" class="forumrow" colspan="3">
	<%
	Response.Write ""
	Response.Write ""
	'利用系统缓存数据取得所有模板名称和ID
	Dim Templateslist
	Response.Write "<select name=""StyleID"" size=""1"">"
	For Each Templateslist in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style")
		Response.Write "<option value="""& Templateslist.selectSingleNode("@id").text &""""
		If CLng(Templateslist.selectSingleNode("@id").text) = CLng(StyleID) Then 
			Response.Write " selected"
		End If 
		Response.Write ">"&Templateslist.selectSingleNode("@type").text &"</option>"
	Next 
	Response.Write "</select>"
	Response.Write "&nbsp;&nbsp;"
	Response.Write "<input type=submit class=""button"" value=""编 辑"" name=""mostyle"">"
	%>
</td>
</form>
</tr>
<!--主表单-->
<form method="POST" action="?Stype=<%=request("Stype")%>" name="bbspic" >
<tr> 
<td width="20%" align=left class=forumrow>当前模版名称：</td>
<td width="80%" align=left class=forumrow colspan="3"><%=StyleName%>
</td>
</tr>
<tr> 
<td width="20%" align=left class=forumrow>增加的文件名：</td>
<td width="80%" align=left class=forumrow colspan="3"><input  type="text" name="NEWFILENAME" value="<%=newfilename%>">（<font color=red>建议采用默认，增加后把相应的文件名上传到该目录下。</font>）
</td>
</tr>
<tr> 
<td width="20%" align=left class=forumrow>批量增加数目：</td>
<td width="80%" align=left class=forumrow colspan="3"><input  type="text" name="NEWNUM" value="<%=newnum%>">
<input type="submit" class="button" name="Submit" value="增加"><input type="hidden" name="StyleId" value="<%=StyleId%>" />
</td>
</tr>
<tr> 
<td width="20%" align="left" class=>覆盖所有模板：</td>
<td width="80%" align="left" class="forumrow" colspan="3">是<input type="radio" class="radio" name="coverall" value="1" >否<input type="radio" class="radio" name="coverall" value="0" checked>
</td>
</tr>
<%
Dim TempName,i
IF request("Submit")="增加" and request("Newnum")<>"" and request("Newnum")<>0 then
	newnum=REQUEST("Newnum")
	for i=count to count+newnum-1
		if stype=2 and i<10 Then
			TempName = newfilename&"0"&i
		Else
			TempName = newfilename&i
		End If
		%>
		<tr>
		<td width="20%" class=forumRowHighlight><%=actname%>ID：<input type=hidden name="face_id<%=i%>" size="10" value="<%=i%>"><%=i%></td>
		<td width="75%" class=forumRowHighlight colspan="2">新增加的文件：<input  type="text" name="userface<%=i%>" value="<%=TempName%>.gif"></td>
		<td width="5%" class=forumRowHighlight> 
		<input type="checkbox" class="checkbox" name="delid<%=i%>" value="<%=i%>">
		</td>
		</tr>
	<%next 
End If
%>
<tr>
<th width="20%" class=forumrow>文件</th>
<th width="45%" class=forumrow>文件名</th>
<th width="30%" class=forumrow>图片</th>
<th width="5%" class=forumrow>删除</th>
</tr>
<tr>
<td width="20%" class=forumrow>文件目录：<input type=hidden name="face_id0" size="10" ></td>
<td width="45%" class=forumrow>&nbsp;<input  type="text" name="userface0" value="<%=Replace(bbspicurl,"../","")%>"></td>
<td width="30%" class=forumrow>&nbsp;</td>
<td width="5%" class=forumrow>&nbsp;</td>
</tr>
<% for i=1 to bbspicmun %>
<tr>
<td width="20%" class=forumrow>文件名：<input type=hidden name="face_id<%=i%>" size="10" value="<%=i%>"></td>
<td width="45%" class=forumrow>&nbsp;<input type="text" name="userface<%=i%>" value="<%=connfile(i)%>"></td>
<td width="30%" class=forumrow> 
&nbsp;&nbsp;<img src=<%=bbspicurl%><%=connfile(i)%>>
<td width="5%" class=forumrow> 
<input type="checkbox" class="checkbox" name="delid<%=i%>" value="<%=i+1%>">
</td>
</tr>
<% next %>
<tr> 
<td  colspan="4" class=forumrow> 
<B>注意</B>：右边复选框为删除选项，如果选择后点保存设置，则删除相应图片<BR>如仅仅修改文件名，可在修改相应选项后直接点击保存设置而不用选择右边复选框
</td>
</tr>
<tr> 
<td  colspan="4" class=forumrow> 
<div align="center"> 
 删除选项：删除所选的实际文件（<font color=red>需要FSO支持功能</font>）：是<input type=radio class="radio" name=setfso value=1 >否<input type=radio class="radio" name=setfso value=0 checked> 请选择要删除的文件，<input type="checkbox" class="checkbox" name=chkall value=on onclick="CheckAll(this.form)">全选 <BR>
<input type="submit" class="button" name="Submit" value="保存设置">
<input type="submit" class="button" name="Submit" value="恢复默认设置">
<input type="submit" class="button" name="Submit" value="恢复默认总设置">
</div>
</td>
</tr>
</form>
<!--主表单结束-->
</table><BR><BR>

<%
end sub

sub saveconst()
	dim f_userface,formname,d_elid
	dim filepaths,objFSO,upface,Rs,sql,i
	For i=0 to count+newnum-1
		d_elid="delid"&i
		formname="userface"&i
		If CInt(request.Form(d_elid))=0 Then
			f_userface=f_userface&request.Form(formname)&"|||"
		Else
			upface=bbspicurl&Request.Form(formname)
			upface=replace(upface,"..","")
			upface=replace(upface,"\","")
			If request("setfso")=1 Then
				filepaths=Server.MapPath(""&upface&"")
				Set objFSO = Dvbbs.iCreateObject("Scripting.FileSystemObject")
				If objFSO.fileExists(filepaths) Then
					'objFSO.DeleteFile(filepaths)
					response.write "删除"&filepaths
				Else
					response.write "未找到"&filepaths
				End If
			End If
		End If
	Next
	Set objFSO=Nothing
	''1=表情，2=心情em，3=头像
	'Style_Pic=TempForum_userface+"@@@"+TempForum_PostFace+"@@@"+TempForum_emot
	f_userface=replace(f_userface,"@@@","")
	Select Case Stype
	Case 1
		upconfig=TempForum_userface+"@@@"+f_userface+"@@@"+TempForum_emot
	Case 2
		upconfig=TempForum_userface+"@@@"+TempForum_PostFace+"@@@"+f_userface
	Case 3
		upconfig=f_userface+"@@@"+TempForum_PostFace+"@@@"+TempForum_emot
    End Select
	If CInt(Request.form("coverall"))=1 Then
		Set Rs = Dvbbs.Execute("Select Id,Type,Folder From Dv_Templates")
		If Not (Rs.Eof And Rs.Bof) Then
			Do While Not Rs.Eof
				'设置全部风格为统一的表情/头像/发贴心情 By Dv.唧唧 2007-10-12
				Dvbbs.writeToFile "../Resource/"&Rs(2)&"/pub_FaceEmot.htm",Dvbbs.checkstr(upconfig)
				Dvbbs.Name = "Style_Pic"&Rs(0)
				Dvbbs.value=upconfig
				Rs.MoveNext
			Loop
		End If
		Rs.Close
		Set Rs=Nothing
	Else
		Dvbbs.writeToFile FileName,Dvbbs.checkstr(upconfig)
		Dvbbs.Name = "Style_Pic"&StyleID
		Dvbbs.value=upconfig
	End If
	Dv_suc(actname&"设置成功。")
End Sub

sub savedefault()
	dim userface,upconfig,sql,rs,i
	userface=""
	select case Stype
	case 1     
			for i=1 to 18
			userface=userface&"face"&i&".gif|||"
			next
			userface="Skins/default/topicface/|||"+userface
			upconfig=TempForum_userface+"@@@"+userface+"@@@"+TempForum_emot
	case 2
			for i=1 to 9
			userface=userface&"em0"&i&".gif|||"
			next
			for i=10 to 49
			userface=userface&"em"&i&".gif|||"
			next
			userface="Skins/Default/emot/|||"+userface
			upconfig=TempForum_userface+"@@@"+TempForum_PostFace+"@@@"+userface
	case 3
			for i=1 to 60
			userface=userface&"image"&i&".gif|||"
			next
			userface="Images/userface/|||"+userface
			upconfig=userface+"@@@"+TempForum_PostFace+"@@@"+TempForum_emot
	case else
			''头像---------------------------------------
			for i=1 to 60
			userface=userface&"image"&i&".gif|||"
			next
			userface="Images/userface/|||"+userface
			upconfig=userface+"@@@"
			''表情---------------------------------------
			userface=""
			for i=1 to 18
			userface=userface&"face"&i&".gif|||"
			next
			userface="Skins/default/topicface/|||"+userface
			upconfig=upconfig+userface+"@@@"
			''心情---------------------------------------
			userface=""
			for i=1 to 9
			userface=userface&"em0"&i&".gif|||"
			next
			for i=10 to 49
			userface=userface&"em"&i&".gif|||"
			next
			userface="Skins/Default/emot/|||"+userface
			upconfig=upconfig+userface
	end select
	If CInt(Request.form("coverall"))=1 Then
		Set Rs = Dvbbs.Execute("Select Id,Type,Folder From Dv_Templates")
		If Not (Rs.Eof And Rs.Bof) Then
			Do While Not Rs.Eof
				'设置全部风格为统一的表情/头像/发贴心情 By Dv.唧唧 2007-10-12
				Dvbbs.writeToFile "../Resource/"&Rs(2)&"/pub_FaceEmot.htm",Dvbbs.checkstr(upconfig)
				Dvbbs.Name = "Style_Pic"&Rs(0)
				Dvbbs.value=upconfig
				Rs.MoveNext
			Loop
		End If
		Rs.Close
		Set Rs=Nothing
	Else
		Dvbbs.writeToFile FileName,Dvbbs.checkstr(upconfig)
		Dvbbs.Name = "Style_Pic"&StyleID
		Dvbbs.value=upconfig
	End If
	Dv_suc(actname&"恢复设置成功。")
end sub

Rem  文件名称 pub_FaceEmot.html
Rem  1=表情，2=心情em，3=头像
Rem  模版大类以@@@分割；小类以|||分割;
Rem  第一个子项为文件保存的目录
Rem  eg.: 表情目录|||表情|||表情...@@@心情目录|||心情|||心情...@@@头像目录|||头像|||头像...

Sub GetNum()
	Dim NRs,sql
	If Application(Dvbbs.CacheName &"_style").documentElement.selectSingleNode("style[@id='"& StyleID &"']") Is Nothing Then
		SQL=" Select Id,Type,Folder From Dv_Templates where Id="&styleId
		Set NRs=Dvbbs.Execute (SQL)
		If not NRs.eof Then
			StyleId=NRs(0)
			StyleName=NRs(1)
		Else
			Rem 继续查找模板数据 防止缓存出错 By Dv.唧唧 2007-10-12
			SQL=" Select Id,Type,Folder From Dv_Templates"
			Set NRs=Dvbbs.Execute (SQL)
			If Not NRs.eof Then
				StyleId=NRs(0)
				StyleName=NRs(1)
			Else
				Errmsg=ErrMsg + "<li>"+"模块未找到，可能已被删除，请重新选取正确模版！"
				Founderr=True
				Exit Sub
			End If
		End if
		FilePath = "../Resource/"& NRs(2) &"/"
		FileName = FilePath &"pub_FaceEmot.htm"

		Dvbbs.Loadstyle()
		Response.Write "Style Application ReLoad"
		NRs.close:Set NRs=Nothing
	End If

	Style_Pic = Dvbbs.ReadTextFile(FileName)

	Style_Pic=Split(Style_Pic,"@@@")	'模版大类以@@@分割；小类以|||分割;
	TempForum_userface=Style_Pic(0)			'用户头像
	TempForum_PostFace=Style_Pic(1)			'发贴表情
	TempForum_emot=Style_Pic(2)				'发贴心情 EM

	Forum_PostFace=split(TempForum_PostFace,"|||")
	Forum_userface=split(TempForum_userface,"|||")
	Forum_emot=split(TempForum_emot,"|||")
	Forum_emotNum=UBound(Forum_emot)
	Forum_userfaceNum=UBound(Forum_userface)
	Forum_PostFaceNum=UBound(Forum_PostFace)

End Sub 
%>

