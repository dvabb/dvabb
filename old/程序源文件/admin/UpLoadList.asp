<!--#include file="../conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
Head()
Dim admin_flag
admin_flag=",37,"
CheckAdmin(admin_flag)

Dim path
Dim objFSO
Dim uploadfolder
Dim uploadfiles
Dim upname
Dim UpFolder
Dim upfilename
Dim sfor(30,2)
Dim seachstr,sqlstr,delsql
Dim currentpage,page_count,Pcount
Dim totalrec,endpage
Dim SysFilePath
If Dvbbs.Forum_Setting(76)="0" Or  Dvbbs.Forum_Setting(76)="" Then
		Dvbbs.Forum_Setting(76)="../UploadFile/"
Else
		Dvbbs.Forum_Setting(76) = "../"& Replace(Dvbbs.Forum_Setting(76),".","")
End If
SysFilePath = Dvbbs.Forum_Setting(76)
if Request("path")<>"" then
	path = Replace(Request("path"),".","")
	If ".."=Left(Request("path"),2) Then path=".."&path
else 
	path = SysFilePath
end If

currentPage=Request("currentpage")
if currentpage="" or not IsNumeric(currentpage) then
	currentpage=1
else
	currentpage=clng(currentpage)
	if err then
		currentpage=1
		err.clear
	end if
end if

if Request("filesearch")<>"" and IsNumeric(Request("filesearch")) then
seachstr="&filesearch="&Request("filesearch")
end if

'----------------------------------
'多条件查询表单处理开始
'----------------------------------
if Request("filesearch")=7 and IsNumeric(Request("filesearch")) then

	'所属版块条件
	if Request("class")<>"" and IsNumeric(Request("class")) and Request("class")<>0 then
	seachstr=seachstr+"&class="&cint(Request("class"))
	sqlstr=" and F_BoardID="&cint(Request("class"))
	end if

	'附件分类条件
	if Request("f_type")<>"" and IsNumeric(Request("f_type")) then
	seachstr=seachstr+"&f_type="&cint(Request("f_type"))
	sqlstr=sqlstr+" and f_type="&cint(Request("f_type"))
	end if

	'附件类型条件
	if Request("f_filetype")<>"" then
	seachstr=seachstr+"&f_filetype="&Request("f_filetype")
	sqlstr=sqlstr+" and f_filetype='"&dvbbs.checkstr(Request("f_filetype"))&"'"
	end if

	'下载次数条件f_downnum
	if Request("f_downnum")<>"" and IsNumeric(Request("f_downnum")) then
		if Request("downtype")="more" then
		sqlstr=sqlstr+" and f_downnum>="&clng(Request("f_downnum"))
		else
		sqlstr=sqlstr+" and f_downnum<="&clng(Request("f_downnum"))
		end if
		seachstr=seachstr+"&f_downnum="&cint(Request("f_downnum"))&"&downtype="&Request("downtype")
	end if

	'浏览次数条件f_viewnum
	if Request("f_viewnum")<>"" and IsNumeric(Request("f_viewnum")) then
		if Request("viewtype")="more" then
		sqlstr=sqlstr+" and f_viewnum>="&clng(Request("f_viewnum"))
		else
		sqlstr=sqlstr+" and f_viewnum<="&clng(Request("f_viewnum"))
		end if
		seachstr=seachstr+"&f_viewnum="&cint(Request("f_viewnum"))&"&viewtype="&Request("viewtype")
	end if

	'附件大小条件f_size
	if Request("f_size")<>"" and IsNumeric(Request("f_size")) then
		if Request("sizetype")="more" then
		sqlstr=sqlstr+" and F_FileSize>="&clng(Request("f_size"))*1024
		else
		sqlstr=sqlstr+" and F_FileSize<="&clng(Request("f_size"))*1024
		end if
		seachstr=seachstr+"&f_size="&cint(Request("f_size"))&"&sizetype="&Request("sizetype")
	end if

	'多少天内发布条件f_adddatenum
	if Request("f_adddatenum")<>"" and IsNumeric(Request("f_adddatenum")) then
		If IsSqlDataBase=1 Then
			if Request("timetype")="more" then
			sqlstr=sqlstr+" and datediff(day,F_AddTime,"&SqlNowString&") >= "&clng(Request("f_adddatenum"))
			else
			sqlstr=sqlstr+" and datediff(day,F_AddTime,"&SqlNowString&") <= "&clng(Request("f_adddatenum"))
			end if
		Else
			if Request("timetype")="more" then
			sqlstr=sqlstr+" and datediff('d',F_AddTime,"&SqlNowString&") >= "&clng(Request("f_adddatenum"))
			else
			sqlstr=sqlstr+" and datediff('d',F_AddTime,"&SqlNowString&") <= "&clng(Request("f_adddatenum"))
			end if
		End If
		seachstr=seachstr+"&f_adddatenum="&cint(Request("f_adddatenum"))&"&timetype="&Request("timetype")
	end if

	'附件作者：
	if Request("f_username")<>"" then
		if Request("usernamechk")="yes" then
		sqlstr=sqlstr+" and f_username='"&dvbbs.checkstr(Request("f_username"))&"'"
		else
		sqlstr=sqlstr+" and f_username like '%"&dvbbs.checkstr(Request("f_username"))&"%'"
		end if
		seachstr=seachstr+"&f_username="&Request("f_username")&"&usernamechk="&Request("usernamechk")
	end if
	'附件说明：
	if Request("f_readme")<>"" then
		if Request("f_readmechk")="yes" then
		sqlstr=sqlstr+" and f_readme='"&dvbbs.checkstr(Request("f_readme"))&"'"
		else
		sqlstr=sqlstr+" and f_readme like '%"&dvbbs.checkstr(Request("f_readme"))&"%'"
		end if
		seachstr=seachstr+"&f_readme="&Request("f_readme")&"&f_readmechk="&Request("f_readmechk")
	end if
end if
'----------------------------------
'多条件查询表单处理结束
'----------------------------------
%>
  <table border="0" cellpadding="3" cellspacing="1" width="100%" align=center>
    <tr>
      <th style="text-align:center;" colspan="2">论坛上传附件管理</th>
    </tr>
    <tr>
      <td width="20%" height="23" class="td2">注意事项：</td>
      <td width="80%" class=td1>
	 ①、本功能必须服务器支持FSO权限方能使用，FSO使用帮助请浏览微软网站。如果您服务器不支持FSO请手动管理。	<BR>②、新版（ＤＶ６）之后的版本上传目录强制定义为UploadFile，只有该目录下文件可进行文件自动清理工作，新版之前的版本上传文件只能手动清除垃圾上传文件；（ＤＶ６．１）版后所有上传附件会自动存放到新自定义的文件夹中，文件目录以当年月明名。（需要空间支持ＦＳＯ读写权限）
	 <br>③、自动清理文件：将对所有上传文件进行核实，如发现文件没有被相关帖子所使用，将执行自动清除命令
	  </td>
    </tr>
	<tr>
	<form action="?action=FileSearch" method=post>
      <td width="20%" height="23" class="td2">快速查询：</td>
      <td width="80%" class=td1>
	  <select size=1 name="FileSearch" onchange="javascript:submit()">
	<option value="0">请选择查询条件</option>
	<option value="1" <%if Request("FileSearch")=1 then%>selected<%end if%>>列出所有上传附件</option>
	<option value="2" <%if Request("FileSearch")=2 then%>selected<%end if%>>最近	２４小时内上传的附件</option>
	<option value="3" <%if Request("FileSearch")=3 then%>selected<%end if%>>最近１个月内上传的附件</option>
	<option value="4" <%if Request("FileSearch")=4 then%>selected<%end if%>>最近３个月内上传的附件</option>
	<option value="5" <%if Request("FileSearch")=5 then%>selected<%end if%>>下载前１００名的附件</option>
	<option value="6" <%if Request("FileSearch")=6 then%>selected<%end if%>>点击前１００名的附件</option>
	</select>
	  </td>
	 </FORM>
    </tr>
  </table>
<%
	if Request("Submit")="清理所有上传记录" then
		call delall()
	elseif Request("Submit")="清除未记录文件" then
		call delall1()
	elseif Request("Submit")="清理当前列表记录" then
		call delall()
	elseif Request("action")="FileSearch" then
		call FileSearch()
	elseif Request("action")="delfiles" then
		call delfiles()
	else
		call main()
	end if
	Footer()

sub main()
%>
<br><table border="0" cellpadding="3" cellspacing="1" width="100%" align=center>
<form action="?action=FileSearch" method=post>
<tr>
	<th colspan="2">高级查询</th>
</tr>
<tr>
<td width=20% class="td2">注意事项</td>
<td width=80% class=td1 colspan=5>在记录很多的情况下搜索条件越多查询越慢，请尽量减少查询条件；</td>
</tr>
<tr>
	<td width="20%" height="23" class="td2">所属版块：</td>
	<td width="80%" class=td1>
	<select name=class>
	<option value="0">所有论坛版块</option>
<%
Dim rs_c,sql,i
set rs_c= Dvbbs.iCreateObject ("adodb.recordset")
sql = "select * from dv_board order by rootid,orders"
rs_c.open sql,conn,1,1
do while not rs_c.EOF%>
<option value="<%=rs_c("boardid")%>" <%if Request("editid")<>"" and clng(Request("editid"))=rs_c("boardid") then%>selected<%end if%>>
<%if rs_c("depth")>0 then%>
<%for i=1 to rs_c("depth")%>
－
<%next%>
<%end if%><%=rs_c("boardtype")%></option>
<%
rs_c.MoveNext 
loop
rs_c.Close
set rs_c=nothing
%>
	</select>
	</td>
</tr>
<tr>
	<td width="20%" height="23" class="td2">文件下载次数：</td>
	<td width="80%" class=td1><input size=45 name="f_downnum" type=text>
	<input type=radio class="radio" value=more name="downtype" checked >&nbsp;多于&nbsp;
	<input type=radio class="radio" value=less name="downtype" >&nbsp;少于
	</td>
</tr>
<tr>
	<td width="20%" height="23" class="td2">附件浏览次数：</td>
	<td width="80%" class=td1><input size=45 name="f_viewnum" type=text>
	<input type=radio class="radio" value=more name="viewtype" checked >&nbsp;多于&nbsp;
	<input type=radio class="radio" value=less name="viewtype" >&nbsp;少于
	</td>
</tr>
<tr>
	<td width="20%" height="23" class="td2">上传天数：</td>
	<td width="80%" class=td1><input size=45 name="f_adddatenum" type=text>
	<input type=radio class="radio" value=more name="timetype" checked >&nbsp;多于&nbsp;
	<input type=radio class="radio" value=less name="timetype" >&nbsp;少于
	</td>
</tr>
<tr>
	<td width="20%" height="23" class="td2">附件作者：</td>
	<td width="80%" class=td1><input size=45 name="f_username" type=text>
	&nbsp;<input type=checkbox class="checkbox" name="usernamechk" value="yes" checked>用户名完整匹配
	</td>
</tr>
<tr>
	<td width="20%" height="23" class="td2">附件说明：</td>
	<td width="80%" class=td1><input size=45 name="f_readme" type=text>
	&nbsp;<input type=checkbox class="checkbox" name="f_readmechk" value="yes" checked>说明内容完整匹配
	</td>
</tr>
<tr>
	<td width="20%" height="23" class="td2">附件大小：</td>
	<td width="80%" class=td1><input size=45 name="f_size" type=text>&nbsp;(单位：K)
	<input type=radio class="radio" value=more name="sizetype" checked >&nbsp;大于&nbsp;
	<input type=radio class="radio" value=less name="sizetype" >&nbsp;小于
	</td>
</tr>
<tr>
	<td width="20%" height="23" class="td2">附件分类：</td>
	<td width="80%" class=td1>
	<select name="f_type">
	<option value="all">所有分类</option>
	<option value="1">图片集分类</option>
	<option value="2">FLASH集分类</option>
	<option value="3">音乐集分类</option>
	<option value="4">电影集分类</option>
	<option value="0">文件集分类</option>
	</select>
	</td>
</tr>
<tr>
	<td width="20%" height="23" class="td2">附件类型：</td>
	<td width="80%" class=td1>
	<select name="f_filetype">
	<option value="">所有文件类型</option>
	<option value="gif">gif</option><option value="jpg">jpg</option>
	<option value="bmp">bmp</option><option value="zip">zip</option>
	<option value="rar">rar</option><option value="exe">exe</option>
	<option value="swf">swf</option><option value="swi">swi</option>
	<option value="mid">mid</option><option value="mp3">mp3</option>
	<option value="rm">rm</option><option value="txt">txt</option>
	<option value="doc">doc</option><option value="exl">exl</option>
	</select>
	</td>
</tr>
<tr>
<th style="text-align:center;" colspan="2"><input name="submit" type=submit class="button" value="开始搜索"></th>
</tr>
<input type=hidden value="7" name="FileSearch">
</form>
</table>
<%
end sub

sub FileSearch()
%>
<form method=post action="?action=delfiles" name="formpost">
<table cellpadding="2" cellspacing="1" border="0" width="100%" align=center>
<tr>
<th colspan=8 ID=TableTitleLink><a href=uploadlist.asp>上传文件管理</a> -->搜索结果</th>
</tr>
<tr>
<td class=td2 align=center><B>类型</B></td>
<td class=td2 height=23 align=center><B>用户名</B></td>
<td class=td2 align=center><B>文 件 名</B></td>
<td class=td2 align=center><B>所属版块</B></td>
<td class=td2 align=center><B>大小</B></td>
<td class=td2 align=center><B>时间/点击/下载</B></td>
<td class=td2 align=center><B>分类</B></td>
<td class=td2 align=center><B>删除</B></td>
</tr>
<%
	Dim rs,sql
	Set rs= Dvbbs.iCreateObject("ADODB.Recordset")
	sql="select F_ID,F_AnnounceID,F_BoardID,F_Filename,F_Username,F_FileType,F_Type,F_FileSize,F_DownNum,F_ViewNum,F_AddTime ,B.Boardtype from [DV_Upfile] U inner join dv_Board B on B.boardid=U.F_BoardID where F_Flag=0 "
	'条件查询
	select case Request("FileSearch")
	case 1
		sql=sql+" order by F_ID desc"
	case 2
		If IsSqlDataBase=1 Then
		sql=sql+" and datediff(hour,F_AddTime,"&SqlNowString&")<25"
		else
		sql=sql+" and datediff('h',F_AddTime,"&SqlNowString&")<25"
		end if
		sql=sql+" order by F_ID desc"
	case 3
		If IsSqlDataBase=1 Then
		sql=sql+" and datediff(month,F_AddTime,"&SqlNowString&")<1"
		else
		sql=sql+" and datediff('m',F_AddTime,"&SqlNowString&")<1"
		end if
		sql=sql+" order by F_ID desc"
	case 4
		If IsSqlDataBase=1 Then
		sql=sql+" and datediff(month,F_AddTime,"&SqlNowString&")<3"
		else
		sql=sql+" and datediff('m',F_AddTime,"&SqlNowString&")<3"
		end if
		sql=sql+" order by F_ID desc"
	case 5
		sql="select top 100 F_ID,F_AnnounceID,F_BoardID,F_Filename,F_Username,F_FileType,F_Type,F_FileSize,F_DownNum,F_ViewNum,F_AddTime ,B.Boardtype from [DV_Upfile] U inner join dv_Board B on B.boardid=U.F_BoardID where F_Flag=0 and F_BoardID<>0"
		sql=sql+" order by F_DownNum Desc,F_ID desc"
	case 6
		sql="select top 100 F_ID,F_AnnounceID,F_BoardID,F_Filename,F_Username,F_FileType,F_Type,F_FileSize,F_DownNum,F_ViewNum,F_AddTime ,B.Boardtype from [DV_Upfile] U inner join dv_Board B on B.boardid=U.F_BoardID where F_Flag=0 and F_BoardID<>0"
		sql=sql+" order by F_ViewNum Desc,F_ID desc"
	case 7
		sql=sql+sqlstr
		sql=sql+" order by F_ID desc"
	case else
		sql=sql+" order by F_ID desc"
	end select
	'response.write SQL
	rs.open sql,conn,1
	if rs.eof and rs.bof then
		response.write "<tr><td colspan=8 class=td1>没有找到相关记录。</td></tr>"
	else
		rs.PageSize = Cint(Dvbbs.Forum_Setting(11))
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		while (not rs.eof) and (not page_count = Cint(Dvbbs.Forum_Setting(11)))
		'列表内容'''''''''''''''''''''
%>
<tr>
<td class=td2 align=center width=20>
	<img src="../skins/default/filetype/<%=rs("F_FileType")%>.gif" border=0>
</td>
<td class=td1 height=23 align=center><%=rs("F_Username")%></td>
<td class="td2">
	<a href="<%=path%><%=rs("F_Filename")%>" target=_blank><%=rs("F_Filename")%></a>
</td>
<td class=td1><%=rs("Boardtype")%></td>
<td class=td2><%=getsize(rs("F_FileSize"))%></td>
<td class=td1>
	<%=formatdatetime(rs("F_AddTime"),1)%>/
	<FONT COLOR=RED><%=rs("F_ViewNum")%></FONT>/
	<%=rs("F_DownNum")%>
</td>
<td class="td2" align=center><%=filetypename(rs("F_Type"))%></td>
<td class="td1" width=20><input type="checkbox" class="checkbox" name="delid" value="<%=rs("F_ID")%>" ></td>
</tr>
<%		page_count = page_count + 1
		rs.movenext
		wend
		Pcount=rs.PageCount
	end if 
	rs.close
	if Request("FileSearch")=1 then sql=""
	if Request("FileSearch")=7 and sqlstr="" then sql=""
%>
<input type=hidden value="<%=sql%>" name="delsql">
<tr><th colspan=8>文件记录库清理操作</th></tr>
<tr>
<td colspan=5 height=25 class="td2"><LI>请选取要删除的文件，然后执行删除操作，<font color=red>附件将直接从服务器上删除并不能恢复！</font></td>
<td colspan=3 height=25 class="td2"><input type="submit" class="button" name="Submit" value="执行删除所选文件"></td></tr>
<tr>
<td colspan=5 height=25 class="td1"><LI>清理同时是否直接从服务器上删除文件，<font color=red>删除的文件将不能恢复 ！</font></td>
<td colspan=3 height=25 class="td1">
<input type=radio class="radio" name=delfile value=1 >是&nbsp;
<input type=radio class="radio" name=delfile value=2 checked>否
</td></tr>
<tr>
<td colspan=5 height=25 class="td2"><li>根据当前列表数据进行清理，清除其中所属的帖子已删改的附件。</td>
<td colspan=3 height=25 class="td2">
<input type="submit" class="button" name="Submit" value="清理当前列表记录">
</td></tr>
<tr>
<td colspan=5 height=25 class="td1"><li>从上传记录中，根据相关发表的帖子内容进行清除所有已删改的附件。</td>
<td colspan=3 height=25 class="td1">
<input type="submit" class="button" name="Submit" value="清理所有上传记录">
</td></tr>
<tr><th colspan=8>空间附件清理操作</th></tr>
<tr><td style="text-align:center;" colspan=8 class="td2">
<li>清除存在服务器空间而没有记录到上传库中的所有上传附件。
<li>请填写清理的上传目录，默认根目录为：“<%=SysFilePath%>”。
<li>目录格式规定：年－月（如：2003-8)。
</td></tr>
<tr><td colspan=5 height=25 class="td1">需要清理的上传目录：
<INPUT TYPE="text" NAME="path" Id="path" value="<%=path%>">
<select onchange="Changepath(this.options[this.selectedIndex].value)">
<option value="<%=SysFilePath%>">选取需要清理的目录</option>
<%
Dim uploadpath,ii
for ii=0 to datediff("m","2003-8",now())
uploadpath=DateAdd("m",-ii,now())
uploadpath=year(uploadpath)&"-"&month(uploadpath)
response.write "<option value="""&uploadpath&""">"&uploadpath&"</option>"
next
%>
</select>
</td>
<td colspan=3 height=25 class="td1">
<input type="submit" class="button" name="Submit" value="清除未记录文件" onclick="{if(confirm('您确定执行的操作吗?将删除所以未有记录的上传文件,并不能恢复。')){this.document.formpost.submit();return true;}return false;}">
</td></tr>
</form>
<SCRIPT LANGUAGE="JavaScript">
<!--
function Changepath(addTitle) {
document.getElementById("path").value=addTitle; 
document.getElementById("path").focus(); 
return; }
//-->
</SCRIPT>
<%
Response.Write "<tr><td class=""td2"" align=center colspan=8>"
call list()
Response.Write "</td></tr></table>"

end sub

SUB LIST()
Dim i
'分页代码
If totalrec="" Then totalrec=0:Pcount=0
response.write "<table cellspacing=0 cellpadding=0 align=center width=""100%""><form method=post action=""?action=FileSearch"&seachstr&""" ><tr><td width=35% class=""td2"">共<b>"&totalrec&"</b>个文件，共分<b><font color=red>"&Pcount&"</font></b>页：</td><td width=* valign=middle align=right nowrap class=""td2"">"

if currentpage > 4 then
	response.write "<a href=""?action=FileSearch&currentpage=1"&seachstr&""">[1]</a> ..."
end if
if Pcount>currentpage+3 then
	endpage=currentpage+3
else
	endpage=Pcount
end if
for i=currentpage-3 to endpage
	if not i<1 then
		if i = clng(currentpage) then
        response.write " <font color=red><b>["&i&"]</b></font>"
		else
        response.write " <a href=""?action=FileSearch&currentpage="&i&seachstr&""">["&i&"]</a>"
		end if
	end if
next
if currentpage+3 < Pcount then 
	response.write "... <a href=""?action=FileSearch&currentpage="&Pcount&seachstr&""">["&Pcount&"]</a>"
end if
response.write " 转到:<input type=text name=currentpage size=3 maxlength=10  value='"& currentpage &"'><input type=submit class=button value=Go  id=button1 name=button1 >"     
response.write "</td></tr></form></table>"
END SUB

SUB delfiles()
Dim delid,F_filename
Dim Rs,sql
if instrRev(path,"/")=0 then path=path&"/"
response.write "<table cellspacing=1 cellpadding=3 align=center width=""100%""><tr><td>"
delid=replace(Request.form("delid"),"'","")
if delid="" then 
response.write "请选择要删除的文件！"
else
Set objFSO = Dvbbs.iCreateObject("Scripting.FileSystemObject")
Set rs= Dvbbs.iCreateObject("ADODB.Recordset")
	sql="select F_id,F_Filename from DV_Upfile where F_ID in ("&delid&")"
	rs.open sql,conn,1
	if not rs.eof then
	response.write "总共删除记录和文件"&rs.recordcount&"个。<br>"
	do while not rs.eof
		if InStr(rs(1),":")=0 or InStr(rs(1),"//")=0 then '判断文件是否本论坛，若不是则采用表中的记录．
			F_filename=path&rs(1)
		else
			F_filename=rs(1)
		end if
		if objFSO.fileExists(Server.MapPath(F_filename)) then
		objFSO.DeleteFile(Server.MapPath(F_filename))
		end if
		Dvbbs.Execute("delete from DV_Upfile where F_ID="&rs(0))
		response.write "已经删除文件"&F_filename&" ！<br>"
	rs.movenext
	loop
	end if
	rs.close
	set rs=nothing
set objFSO=nothing
end if
response.write "</td></tr></table>"
END SUB

'清理所有记录
sub delall()
Server.ScriptTimeout=9999999
response.write "<table cellspacing=1 cellpadding=3 align=center width=""100%""><tr><td>"
Dim TempFileName
Dim F_ID,F_AnnounceID,F_boardid,F_filename
Dim S_AnnounceID,s_Rootid
Dim drs,delfile
Dim delinfo,i,rs
delfile=trim(Request.form("delfile"))
if cint(delfile)=1 then
delinfo="已被删除！"
else
delinfo="未被删除！"
end if

if Request.form("delsql")<>"" then
	If Dvbbs.chkpost=False Then
		Dvbbs.AddErrmsg "您提交的数据不合法，请不要从外部提交发言。"
		exit sub
		else
		delsql=Request.form("delsql")
	End If
end if
i=0
Set objFSO = Dvbbs.iCreateObject("Scripting.FileSystemObject")
If Delsql = "" Then
	Set Rs = Dvbbs.Execute("SELECT F_ID, F_AnnounceID, F_BoardID, F_Filename, F_Type FROM [DV_Upfile] WHERE F_Flag = 4 ORDER BY F_ID DESC ")
Else
	Set Rs = Dvbbs.Execute(Delsql)
End If
'response.write delsql
if rs.eof then
	response.write "还未有"
else
	do while not rs.eof
	F_ID=rs(0)
	F_boardid=rs(2)
	if InStr(rs(3),":")=0 or InStr(rs(3),"//")=0 then '判断文件是否本论坛，若不是则采用表中的记录．
		F_filename=path&rs(3)
	else
		F_filename=rs(3)
	end if
	'Response.Write Rs("F_Type")&"<br>"
	If Rs("F_Type")<>1 Then		'除图片文件外
		TempFileName="viewfile.asp?ID="&F_ID
	Else
		TempFileName=F_filename
	End If
	TempFileName=Lcase(TempFileName)
	if rs(1)="" or isnull(rs(1)) then
		if InStr(rs(3),":")=0 or InStr(rs(3),"//")=0 then '判断文件是否本论坛，若不是则采用表中的记录．
			if objFSO.fileExists(Server.MapPath(F_filename)) then
				if delfile=1 then
					Dvbbs.Execute("delete from DV_Upfile where F_ID="&F_ID)
					objFSO.DeleteFile(Server.MapPath(F_filename))
				end if
				response.write "文件未写帖子,<a href="&F_filename&" target=""_blank"">"&F_filename&"</a> "&delinfo&"<br>"
			else
				response.write "文件未写帖子,<a href="&F_filename&" target=""_blank"">"&F_filename&"</a> 已不存在！<br>"
			end if
		else
			response.write "外部文件<a href="&F_filename&" target=""_blank"">"&F_filename&"</a> "&delinfo&"<br>"
		end if
		i=i+1
	else
		if isnumeric(rs(1)) then
			S_AnnounceID=rs(1)
		else
			F_AnnounceID=split(rs(1),"|")
			s_Rootid=F_AnnounceID(0)
			S_AnnounceID=F_AnnounceID(1)
		end if
		'Response.Write rs(1)&"<br>"
		If S_AnnounceID="" Then
			Response.Write F_filename &"文件数据有问题<br>"
		Else
		'取出所属帖子表名
		Dim PostTablename
		set drs=Dvbbs.Execute("select PostTable from dv_topic where TopicID="&s_Rootid)
			if not drs.eof then
			PostTablename=drs(0)
			else
			PostTablename=AllPostTable(0)
			end if
		drs.close

		'找出相应的帖子进行判断文件是否存在帖子内容
		'Response.Write "select body from "&PostTablename&" where AnnounceID="&S_AnnounceID&"<br>"
		set drs=Dvbbs.Execute("select body from "&PostTablename&" where AnnounceID="&S_AnnounceID)
		if drs.eof then
			if delfile=1 then
			Dvbbs.Execute("delete from DV_Upfile where F_ID="&F_ID)
			end if
			if objFSO.fileExists(Server.MapPath(F_filename)) then
				if delfile=1 then
				objFSO.DeleteFile(Server.MapPath(F_filename))
				end if
				response.write "帖子未找到,<a href="&F_filename&" target=""_blank"">"&F_filename&"</a> "&delinfo&"<br>"
			else
				response.write "帖子未找到,<a href="&F_filename&" target=""_blank"">"&F_filename&"</a> 已不存在！<br>"
			end if
			i=i+1
		else
			'Response.Write TempFileName&"<br>"
			If Instr(Lcase(drs(0)),TempFileName)=0 Then
				if objFSO.fileExists(Server.MapPath(F_filename)) then
					if delfile=1 then
						objFSO.DeleteFile(Server.MapPath(F_filename))
						Dvbbs.Execute("delete from DV_Upfile where F_ID="&F_ID)
					end if
					response.write "帖子内容不符,<a href="&F_filename&" target=""_blank"">"&F_filename&"</a> "&delinfo&"[<a href=""dispbbs.asp?Boardid="&F_boardid&"&ID="&s_Rootid&"&replyID="&S_AnnounceID&"&skin=1"" target=""_blank"" title=""浏览相关帖子""><font color=red>查看相关讨论</font></a> | <a href=myfile.asp?action=edit&editid="&F_ID&" target=""_blank"" title=""编辑文件""><font color=red>编辑</font></a>]<br>"
				else
					response.write "帖子内容不符,<a href="&F_filename&" target=""_blank"">"&F_filename&"</a> 已不存在！[<a href=""dispbbs.asp?Boardid="&F_boardid&"&ID="&s_Rootid&"&replyID="&S_AnnounceID&"&skin=1"" target=""_blank"" title=""浏览相关帖子""><font color=red>查看相关讨论</font></a> | <a href=myfile.asp?action=edit&editid="&F_ID&" target=""_blank"" title=""编辑文件""><font color=red>编辑</font></a>]<br>"
				end if
				i=i+1
			end if
		end if
		drs.close
		End If
	End If
rs.movenext
loop
end if
rs.close
set drs=nothing
set rs=nothing
set objFSO=nothing

response.write"共清理　"&i&"　个无用文件 ［<a href=?path="&path&" >返回</a>］"
response.write "</td></tr></table>"
end sub


'删除所有未记录到上传库中的文件
Sub Delall1()
	REM 防脚本超时 2004-8-26.Dv.Yz
	Server.ScriptTimeout = 9999999
response.write "<table cellspacing=1 cellpadding=3 align=center width=""100%""><tr><td>"
Dim delfile,delinfo,datepath,i,rs
delfile=dvbbs.checkStr(trim(Request.form("delfile")))
if cint(delfile)=1 then
	delinfo="目前已被删除！"
else
	delinfo="目前未被删除！"
end if

if instrRev(path,"/")=0 then path=path&"/"
If instr(path,SysFilePath)=0 Then
	datepath=path
	path=SysFilePath&path
End If

Set objFSO = Dvbbs.iCreateObject("Scripting.FileSystemObject")
if objFSO.FolderExists(Server.MapPath(path))=false then
	response.write "路径："&Path&"不存在！"
else
	Set uploadFolder=objFSO.GetFolder(Server.MapPath(path))
	Set uploadFiles=uploadFolder.Files
	i=0
	For Each Upname In uploadFiles
		upfilename=path&upname.name
		'Response.Write "select top 1 F_ID from DV_Upfile where F_Filename = '"&datepath&upname.name&"'<br>"
		set rs=Dvbbs.Execute("SELECT TOP 1 F_ID FROM Dv_Upfile WHERE F_Filename = '"&datepath&upname.name&"'")
		if rs.eof then
			i=i+1
			if delfile=1 then
			objFSO.DeleteFile(Server.MapPath(upfilename))
			end if
			response.write "<a href="&upfilename&" target=""_blank"">"
			response.write upfilename&"</a>在库中没有记录！"&delinfo&"<br>"
		end if
		rs.close
		set rs=nothing
	next
	response.write"共删除　"&i&"　个无用文件 ［<a href=?path="&path&" >返回</a>］"
	set uploadFolder=nothing
	set uploadFiles=nothing
end if
set objFSO=nothing
response.write "</td></tr></table>"
end sub

function folder(path)
on error resume  next
       Set objFSO = Dvbbs.iCreateObject("Scripting.FileSystemObject")
          Set uploadFolder=objFSO.GetFolder(Server.MapPath(path))
		  if err.number<>"0" then
		  response.write Err.Description
		  response.end
		  end if
          For Each UpFolder In uploadFolder.SubFolders
            response.write "『<A HREF=?path="&path&"/"&upfolder.name&" >"&upfolder.name&"</a>』 | "
next
set uploadFolder=nothing
end function

function procGetFormat(sName)
 Dim str
 procGetFormat=0
 if instrRev(sName,".")=0 then exit function
 str=lcase(mid(sName,instrRev(sName,".")+1))
 for i=0 to uBound(sFor,1)
  if str=sFor(i,0) then 
    procGetFormat=sFor(i,1)
    exit for
  end if
 next
end function

function filetypename(stype)
if isempty(stype) or not isnumeric(stype) then exit function
select case cint(stype)
case 1
filetypename="图片集"
case 2
filetypename="FLASH集"
case 3
filetypename="音乐集"
case 4
filetypename="电影集"
case else
filetypename="文件集"
end select 
end function

function getsize(size)
if isEmpty(size) then exit function
	if size>1024 then
 		   size=(size\1024)
 		   getsize=size & "&nbsp;KB"
	else
		   getsize=size & "&nbsp;B"
 	end if
 	if size>1024 then
 		   size=(size/1024)
 		   getsize=formatnumber(size,2) & "&nbsp;MB"		
 	end if
 	if size>1024 then
 		   size=(size/1024)
 		   getsize=formatnumber(size,2) & "&nbsp;GB"	   
 	end if   
end function
%>