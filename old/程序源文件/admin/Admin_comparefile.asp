<!--#include file =../conn.asp-->
<!-- #include file="inc/const.asp" -->
<%	
Head()
Dim admin_flag,rs_c
Dim DimFileExt,Sun,SumFiles,SumFolders
Dim timer1,TmpPath,fso
Set fso = createObject("Scripting.FileSystemObject")




admin_flag=",1,"
CheckAdmin(admin_flag)

if request.QueryString("act")<>"scan" then
Response.Write "<form action='?act=scan' method='post'>"
Response.Write "<b>填入你要检查的路径：网站根目录的相对路径，填 \ 即检查整个网站；. 为程序所在目录</b>"
Response.Write "<input name='path' type='text' style='border:1px solid #999' value='.' size='30' />"
Response.Write " <br>"
Response.Write "<!--网站根目录的相对路径，填\即检查整个网站；.为程序所在目录-->"
Response.Write "<br>"
Response.Write "<br>"
Response.Write " <input type='submit' value=' 开始扫描 ' style='background:#fff;border:1px solid #999;padding:2px 2px 0px 2px;margin:4px;border-width:1px 3px 1px 3px' /> "
Response.Write "</form>"

else
server.ScriptTimeout = 600
DimFileExt = "asp,cer,asa,cdx"
Sun = 0
SumFiles = 0
SumFolders = 1
if request.Form("path")="" then
response.Write("No Hack")
response.End()
end if
timer1 = timer
if request.Form("path")="\" then
TmpPath = Server.MapPath("\")
elseif request.Form("path")="." then
TmpPath = Server.MapPath(".")
else
TmpPath = Server.MapPath("\")&"\"&request.Form("path")
end if
Call ShowAllFile(TmpPath)
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
<tr>
<th>ASPSecurity For Hacking
</tr>
<tr>
<td class="CPanel" style="padding:5px;line-height:170%;clear:both;font-size:12px">
<div id="updateInfo" style="background:ffffe1;border:1px solid #89441f;padding:4px;display:none"></div>
扫描完毕！一共检查文件夹<font color="#FF0000"><%=SumFolders%></font>个，文件<font color="#FF0000"><%=SumFiles%></font>个，发现可疑点<font color="#FF0000"><%=Sun%></font>个
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
<td valign="top">
<table width="100%" border="1" cellpadding="0" cellspacing="0" style="padding:5px;line-height:170%;clear:both;font-size:12px">
<tr>
<td width="20%">文件相对路径</td>
<td width="20%">特征码</td>
<td width="40%">描述</td>
<td width="20%">创建/修改时间</td>
</tr>
<p>
<%=Report%>
<br/></p>
</table></td>
</tr>
</table>
</td></tr></table>
<%
timer2 = timer
thetime=cstr(int(((timer2-timer1)*10000 )+0.5)/10)
response.write "<br><font size=""2"">本页执行共用了"&thetime&"毫秒</font>"
end if
Footer()
%>

<%

'遍历处理path及其子目录所有文件
Sub ShowAllFile(Path)
Dim f,fc2,myfile,temp,fc,f1,fso
Set fso = createObject("Scripting.FileSystemObject")
if not fso.FolderExists(path) then 
exit Sub
End if
Set f = FSO.GetFolder(Path)
Set fc2 = f.files
For Each myfile in fc2
If CheckExt(FSO.GetExtensionName(path&"\"&myfile.name)) Then
Call ScanFile(Path&Temp&"\"&myfile.name, "")
SumFiles = SumFiles + 1
End If
Next
Set fc = f.SubFolders
For Each f1 in fc
ShowAllFile path&"\"&f1.name
SumFolders = SumFolders + 1
Next
Set FSO = Nothing
End Sub

'检测文件
Sub ScanFile(FilePath, InFile)
If InFile <> "" Then
Infiles = "该文件被<a href=""http://"&Request.Servervariables("server_name")&"\"&InFile&""" target=_blank>"& InFile & "</a>文件包含执行"
End If
on error resume next
set ofile = fso.OpenTextFile(FilePath)
filetxt = Lcase(ofile.readall())
If err Then Exit Sub end if
if len(filetxt)>0 then
'特征码检查
temp = "<a href=""http://"&Request.Servervariables("server_name")&"\"&replace(FilePath,server.MapPath("\")&"\","",1,1,1)&""" target=_blank>"&replace(FilePath,server.MapPath("\")&"\","",1,1,1)&"</a>"
'Check "WScr"&DoMyBest&"ipt.Shell"
If instr( filetxt, Lcase("WScr"&DoMyBest&"ipt.Shell") ) or Instr( filetxt, Lcase("clsid:72C24DD5-D70A"&DoMyBest&"-438B-8A42-98424B88AFB8") ) then
Report = Report&"<tr><td>"&temp&"</td><td>WScr"&DoMyBest&"ipt.Shell 或者 clsid:72C24DD5-D70A"&DoMyBest&"-438B-8A42-98424B88AFB8</td><td>危险组件，一般被ASP木马利用。"&infiles&"</td><td>"&GetDatecreate(filepath)&"<br>"&GetDatemodify(filepath)&"</td></tr>"
Sun = Sun + 1
End if
'Check "She"&DoMyBest&"ll.Application"
If instr( filetxt, Lcase("She"&DoMyBest&"ll.Application") ) or Instr( filetxt, Lcase("clsid:13709620-C27"&DoMyBest&"9-11CE-A49E-444553540000") ) then
Report = Report&"<tr><td>"&temp&"</td><td>She"&DoMyBest&"ll.Application 或者 clsid:13709620-C27"&DoMyBest&"9-11CE-A49E-444553540000</td><td>危险组件，一般被ASP木马利用。"&infiles&"</td><td>"&GetDatecreate(filepath)&"<br>"&GetDatemodify(filepath)&"</td></tr>"
Sun = Sun + 1
End If
'Check .Encode
Set regEx = New RegExp
regEx.IgnoreCase = True
regEx.Global = True
regEx.Pattern = "@\s*LANGUAGE\s*=\s*[""]?\s*(vbscript|jscript|javascript).encode\b"
If regEx.Test(filetxt) Then
Report = Report&"<tr><td>"&temp&"</td><td>(vbscript|jscript|javascript).Encode</td><td>似乎脚本被加密了，一般ASP文件是不会加密的。"&infiles&"</td><td>"&GetDatecreate(filepath)&"<br>"&GetDatemodify(filepath)&"</td></tr>"
Sun = Sun + 1
End If
'Check my ASP backdoor :(
regEx.Pattern = "\bEv"&"al\b"
If regEx.Test(filetxt) Then
Report = Report&"<tr><td>"&temp&"</td><td>Ev"&"al</td><td>e"&"val()函数可以执行任意ASP代码，被一些后门利用。其形式一般是：ev"&"al(X)<br>但是javascript代码中也可以使用，有可能是误报。"&infiles&"</td><td>"&GetDatecreate(filepath)&"<br>"&GetDatemodify(filepath)&"</td></tr>"
Sun = Sun + 1
End If
'Check exe&cute backdoor
regEx.Pattern = "[^.]\bExe"&"cute\b"
If regEx.Test(filetxt) Then
Report = Report&"<tr><td>"&temp&"</td><td>Exec"&"ute</td><td>e"&"xecute()函数可以执行任意ASP代码，被一些后门利用。其形式一般是：ex"&"ecute(X)。<br>"&infiles&"</td><td>"&GetDatecreate(filepath)&"<br>"&GetDatemodify(filepath)&"</td></tr>"
Sun = Sun + 1
End If
Set regEx = Nothing

'Check include file
Set regEx = New RegExp
regEx.IgnoreCase = True
regEx.Global = True
regEx.Pattern = "<!--\s*#include\s*file\s*=\s*"".*"""
Set Matches = regEx.Execute(filetxt)
For Each Match in Matches
tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")
If Not CheckExt(FSOs.GetExtensionName(tFile)) Then
Call ScanFile( Mid(FilePath,1,InStrRev(FilePath,"\"))&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1) )
SumFiles = SumFiles + 1
End If
Next
Set Matches = Nothing
Set regEx = Nothing

'Check include virtual
Set regEx = New RegExp
regEx.IgnoreCase = True
regEx.Global = True
regEx.Pattern = "<!--\s*#include\s*virtual\s*=\s*"".*"""
Set Matches = regEx.Execute(filetxt)
For Each Match in Matches
tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")
If Not CheckExt(FSOs.GetExtensionName(tFile)) Then
Call ScanFile( Server.MapPath("\")&"\"&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1) )
SumFiles = SumFiles + 1
End If
Next
Set Matches = Nothing
Set regEx = Nothing

'Check Server&.Execute|Transfer
Set regEx = New RegExp
regEx.IgnoreCase = True
regEx.Global = True
regEx.Pattern = "Server.(Exec"&"ute|Transfer)([ \t]*|\()"".*"""
Set Matches = regEx.Execute(filetxt)
For Each Match in Matches
tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")
If Not CheckExt(FSOs.GetExtensionName(tFile)) Then
Call ScanFile( Mid(FilePath,1,InStrRev(FilePath,"\"))&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1) )
SumFiles = SumFiles + 1
End If
Next
Set Matches = Nothing
Set regEx = Nothing

'Check Server&.Execute|Transfer
Set regEx = New RegExp
regEx.IgnoreCase = True
regEx.Global = True
regEx.Pattern = "Server.(Exec"&"ute|Transfer)([ \t]*|\()[^""]\)"
If regEx.Test(filetxt) Then
Report = Report&"<tr><td>"&temp&"</td><td>Server.Exec"&"ute</td><td>不能跟踪检查Server.e"&"xecute()函数执行的文件。请管理员自行检查。<br>"&infiles&"</td><td>"&GetDatecreate(filepath)&"<br>"&GetDatemodify(filepath)&"</td></tr>"
Sun = Sun + 1
End If
Set Matches = Nothing
Set regEx = Nothing

'Check Crea"&"teObject
Set regEx = New RegExp
regEx.IgnoreCase = True
regEx.Global = True
regEx.Pattern = "createO"&"bject[ |\t]*\(.*\)"
Set Matches = regEx.Execute(filetxt)
For Each Match in Matches
If Instr(Match.Value, "&") or Instr(Match.Value, "+") or Instr(Match.Value, """") = 0 or Instr(Match.Value, "(") <> InStrRev(Match.Value, "(") Then
Report = Report&"<tr><td>"&temp&"</td><td>Creat"&"eObject</td><td>Crea"&"teObject函数使用了变形技术，仔细复查。"&infiles&"</td><td>"&GetDatecreate(filepath)&"<br>"&GetDatemodify(filepath)&"</td></tr>"
Sun = Sun + 1
exit sub
End If
Next
Set Matches = Nothing
Set regEx = Nothing
end if
set ofile = nothing
set fsos = nothing
End Sub

'检查文件后缀，如果与预定的匹配即返回TRUE
Function CheckExt(FileExt)
Dim Ext,i 

If DimFileExt = "*" Then CheckExt = True
Ext = Split(DimFileExt,",")
For i = 0 To Ubound(Ext)
If Lcase(FileExt) = Ext(i) Then 
CheckExt = True
Exit Function
End If
Next
End Function

Function GetDatemodify(filepath)
Set fso = createObject("Scripting.FileSystemObject")
Set f = fso.GetFile(filepath) 
s = f.DateLastModified 
set f = nothing
set fso = nothing
GetDatemodify = s
End Function

Function GetDatecreate(filepath)
Set fso = createObject("Scripting.FileSystemObject")
Set f = fso.GetFile(filepath) 
s = f.Datecreated 
set f = nothing
set fso = nothing
GetDatecreate = s
End Function

%> 