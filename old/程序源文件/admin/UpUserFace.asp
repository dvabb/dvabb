<!--#include file = "../conn.asp"-->
<!-- #include file = "inc/const.asp" -->
<%
Head()
dim admin_flag
dim objFSO
dim uploadfolder
dim uploadfiles
dim upname
dim uid,faceid
dim usernames
dim userface,dnum
dim upfilename
dim pagesize, page,filenum, pagenum
admin_flag = ",36,"
CheckAdmin(admin_flag)
Call main()
If Errmsg<>"" Then Dvbbs_Error()
Footer()

sub main()
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
<tr>
<td valign=top>
注意：本功能需要主机开放FSO权限，FSO相关帮助请看微软帮助文档<BR>
在这里您可以管理论坛所有用户自定义头像上传文件，搜索用户头像请用用户ID进行搜索<BR>
用户ID的获得可以通过用户信息管理中搜索相关用户，然后将鼠标移到用户名连接上，查看连接属性，参数UserID = 后面既是用户的ID
</td>
</tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
<tr align=center><th width="*">文件名</th><th width="100">所属用户</th><th width="50">大小</th><th width="120">最后访问</th><th width="120">上传日期</th><th width="35">管理</th></tr>
<form method="POST" action="?action=delall">
<%
pagesize = 20
page = request.querystring("page")
If page = "" or not isnumeric(page) Then
	page = 1
Else
	page = int(page)
End If

If trim(request("action"))<>"" Then
	If trim(request("action")) = "delall" Then
		call delface()
	Else
		call maininfo()
	End If
Else
	call maininfo()
End If
call foot()
End Sub

sub maininfo()
Dim rs,filename
filename = Replace(Replace(Request("filename"),"/",""),"..","")
On Error Resume Next
Set objFSO = Dvbbs.iCreateObject("Scripting.FileSystemObject")
If Err Then
	ErrMsg = "<li>您的系统不支持FSO文件读写，不能使用此功能。"
	Exit Sub
End If
If filename<>"" Then
	If UpUserFaceFolder="" Then 
		objFSO.DeleteFile(Server.MapPath("../uploadFace/"&filename))
	Else
		objFSO.DeleteFile(Server.MapPath(".."&UpUserFaceFolder&filename))
	End If
End If
If UpUserFaceFolder="" Then 
	Set uploadFolder = objFSO.GetFolder(Server.MapPath("../uploadFace/"))
Else
	Set uploadFolder = objFSO.GetFolder(Server.MapPath(".."&UpUserFaceFolder))
End If
If Err Then
	ErrMsg = "<li>您使用的上传头像目录不是系统默认目录，不能进行管理。"
	Exit Sub
End If
Set uploadFiles = uploadFolder.Files
filenum = uploadfiles.count
pagenum = int(filenum/pagesize)
If filenum mod pagesize>0 Then
	pagenum = pagenum+1
End If
If page> pagenum Then
	page = 1
End If
i = 0
For Each Upname In uploadFiles
	i = i+1
	If i>(page-1)*pagesize and i <= page*pagesize Then
	upfilename = "../uploadFace/"&upname.name
		If instr(upname.name,"_") Then    '取出头像的用户名
			uid = split(upname.name,"_")
			faceid = uid(0)
			If  IsNumeric(faceid)	then
				set rs = Dvbbs.Execute("select username from [dv_user] where userid = "&faceid&"")
				If not rs.eof  Then
					usernames = rs(0)
				End If
				rs.close
				Set rs = Nothing
			End If		
		End If
		response.write "<tr><td class=td1 height=23><a href=""../uploadface/"&upname.name&""" target=_blank>"&upname.name&"</a></td>"
		response.write "<td align=right class=td2>"&usernames&"</td>"
		response.write "<td align=right class=td1>"& upname.size &"</td>"
		response.write "<td align=center class=f>"& upname.datelastaccessed &"</td>"
		response.write "<td align=center class=td1>"& upname.datecreated &"</td>"
		response.write "<td align=center class=td2><a href='?filename="&upname.name&"'>删除</a></td></tr>"
	ElseIf i>page*pagesize Then
		Exit For
	End If
	usernames = ""
Next

End Sub 

'清理头像
Sub Delface()
	Server.ScriptTimeout = 999999
	On Error Resume Next
	Dim DllUserFace,i,rs,sql
	Dim Upfacepath, Newfilename
	If UpUserFaceFolder = "" Then
		Upfacepath = "../uploadFace/"
	Else
		Upfacepath = ".."&UpUserFaceFolder
	End If
	Dnum = 0
	DllUserFace = Request("filename")
	Set objFSO = Dvbbs.iCreateObject("Scripting.FileSystemObject")
	'删除返还头像参数的文件;
	If DllUserFace <> "" Then
		DllUserFace = Replace(DllUserFace,"..","")
		objFSO.DeleteFile(Server.MapPath(Upfacepath&DllUserFace))
		If Err Then
			Response.Write Err.Description
			Exit Sub
		End If
	End If
	Set uploadFolder = objFSO.GetFolder(Server.MapPath(Upfacepath))
	Set uploadFiles = uploadFolder.Files
	Filenum = uploadfiles.count
	i = 0
	For Each Upname In uploadFiles
		i = i + 1
		If i > 0 And i <= Filenum Then
			Upfilename = Lcase(Upfacepath&upname.name)
			'取出头像的用户名
			If Instr(upname.name,"_") Then
				Uid = Split(upname.name,"_")
				Faceid = Uid(0)
				If IsNumeric(Faceid) Then
					Set Rs = Dvbbs.Execute("SELECT Username, Userface FROM [Dv_User] WHERE Userid = " & Faceid)
					If Not (Rs.Eof And Rs.Bof) Then
						Usernames = Rs(0)
						Userface = Lcase(Trim(Rs(1)))
						If Instr(Userface,"|") > 0 Then Userface = Split(Userface,"|")(1)
						If Instr(Replace(Upfilename,"../",""),Userface) = 0 Then
							objFSO.DeleteFile(Server.MapPath(upfilename))
							Response.Write "头像已更改,用户" & Usernames & "旧头像文件："& upfilename &"已删除<br>"
							Response.Flush
							If Err Then
								Response.Write Err.Description
								Exit For
							End If
							Dnum = Dnum + 1
						End If
					Else
						objFSO.DeleteFile(Server.MapPath(upfilename))
						Response.Write "用户ID：" & Faceid & "已注销,文件：" & Upfilename &"已删除<br>"
						Response.Flush
						If Err Then
							Response.Write Err.Description
							Exit For
						End If
						Dnum = Dnum + 1
					End If
					Set Rs = Nothing
				End If
			Else
			'清理没有用户ID的头像文件
				Sql = "SELECT Top 1 Userid From [Dv_User] WHERE Userface = '" & Upfilename & "'"
				Set Rs = Dvbbs.Execute(Sql)
				If Rs.Eof And Rs.Bof Then
					objFSO.DeleteFile(Server.MapPath(upfilename))
					Response.Write "已清查删除文件：" & upfilename & "<br>"
					Response.Flush
					If Err Then
						Response.Write Err.Description
						Exit For
					End If
					Dnum = Dnum + 1
				Else
				'改为带ID的头像 2005-1-15 Dv.Yz
					Faceid = Rs(0)
					Newfilename = Upfacepath & Faceid & "_" & Upname.Name
					objFSO.Movefile ""&Server.MapPath(Upfilename)&"",""&Server.MapPath(Newfilename)&""
					If Not Err Then
						Dvbbs.Execute("UPDATE [Dv_User] Set UserFace = '"& Replace(Newfilename,"'","") & "' WHERE Userid = " & Faceid)
						Response.Write "旧头像：" & Upfilename & " 已改为：" & Newfilename & "<br>"
						Response.Flush
						Dnum = Dnum + 1
					Else
						Response.Write Err.Description
						Exit For
					End If
				End If
				Set Rs = Nothing
			End If
		End If
	Next
	Response.Write " 共清理 "& dnum &" 个文件  "
End Sub

Sub foot()
Dim i
Set uploadFolder = Nothing
Set uploadFiles = Nothing
%>
<tr><td colspan=6 class=td1 height=30>
<%
If page>1 Then
	response.write "<a href=?page=1>首页</a>&nbsp;&nbsp;<a href=""?page="& page-1 &""">上一页</a>&nbsp;&nbsp;"
Else
	response.write "首页&nbsp;&nbsp;上一页&nbsp;&nbsp;"
End If
If page<i/pagesize Then
	response.write "<a href=""?page="& page+1 &""">下一页</a>&nbsp;&nbsp;<a href=""?page="& pagenum &""">尾页</a>"
Else
	response.write "下一页&nbsp;&nbsp;尾页"
End If
%>
<input type="submit" class="button" value="清理"></td><tr></form></table><br>
<% End Sub %>
