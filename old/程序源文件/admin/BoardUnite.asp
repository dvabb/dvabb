<!--#include file="../conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
Head()
Dim Str
Dim admin_flag
admin_flag=",13,"
CheckAdmin(admin_flag)
If Request("action") = "unite" Then
	Call unite()
Else
	Call boardinfo()
End if
Footer()

Sub boardinfo()
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
	<tr>
	<th style="text-align:center;">合并论坛数据
	</th>
	</tr>
	<form action=boardunite.asp?action=unite method=post>
	<tr>
	<td class=td1>
	<B>合并论坛选项</B>：<br />
<B>将本论坛及其下属版面的帖子都转移至目标论坛，并删除本论坛及其下属版面</B><br /><br />
<%
	Dim rs,sql,i
	set rs = Dvbbs.iCreateObject ("Adodb.recordset")
	sql="select boardid,boardtype,depth from dv_board order by rootid,orders"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "没有论坛"
	else
		response.write " 将论坛 "
		response.write "<select name=oldboard size=1>"
		do while not rs.eof
%>
<option value="<%=rs("boardid")%>"><%if rs("depth")>0 then%>
<%for i=1 to rs("depth")%>
－
<%next%>
<%end if%><%=rs("boardtype")%></option>
<%
		rs.movenext
		loop
		response.write "</select>"
	end if
	rs.close
	sql="select boardid,boardtype,depth from dv_board order by rootid,orders"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "没有论坛"
	else
		response.write " 合并到 "
		response.write "<select name=newboard size=1>"
		do while not rs.eof
%>
<option value="<%=rs("boardid")%>"><%if rs("depth")>0 then%>
<%for i=1 to rs("depth")%>
－
<%next%>
<%end if%><%=rs("boardtype")%></option>
<%
		rs.movenext
		loop
		response.write "</select>"
	end if
	rs.close
	set rs=nothing
	response.write " <br /><br /><input type=submit class=button name=Submit value=合并论坛><br /><br />"
%>
	</td>
	</tr>
	<tr>
	<td class=td1><B>注意事项</B>：<br /><FONT COLOR="red">所有操作不可逆，请慎重操作</FONT><br /> 不能在同一个版面内进行操作、不能将一个版面合并到其下属论坛中。<br />合并后您所指定的论坛（或者包括其下属论坛）将被删除，所有帖子将转移到您所指定的目标论坛中
	</td>
	</tr></form>
	</table>
<%
end sub

Sub Unite()
	Dim Newboard
	Dim Oldboard
	Dim ParentStr, iParentStr
	Dim Depth, iParentID, Child
	Dim ParentID, RootID
	Dim rs,i
	If Clng(Request("newboard")) = Clng(Request("oldboard")) Then
		Errmsg = "请不要在相同版面内进行操作。"
		dvbbs_error()
		Exit Sub
	End If
	Newboard = Clng(Request("newboard"))
	Oldboard = Clng(Request("oldboard"))
	'将本论坛及其下属版面的帖子都转移至目标论坛，并删除本论坛及其下属版面
	'得到当前版面下属论坛
set rs=Dvbbs.Execute("select ParentStr,Boardid,depth,ParentID,child,RootID from dv_board where boardid="&oldboard)
if rs(0)="0" then
	ParentStr=rs(1)
	iParentID=rs(1)
	ParentID=0
else
	ParentStr=rs(0) & "," & Rs(1)
	iParentID=rs(3)
	ParentID=rs(3)
end if
iParentStr=rs(1)
depth=rs(2)
child=rs(4)+1
RootID=rs(5)
i=0
If ParentID=0 Then
set rs=Dvbbs.Execute("select Boardid from dv_board where boardid="&newboard&" and RootID="&RootID)
Else
set rs=Dvbbs.Execute("select Boardid from dv_board where boardid="&newboard&" and ParentStr like '%"&ParentStr&"%'")
End If
if not (rs.eof and rs.bof) then
	response.write "不能将一个版面合并到其下属论坛中"
	exit sub
end if
'得到当前版面下属论坛ID
i=0
set rs=Dvbbs.Execute("select Boardid from dv_board where RootID="&RootID&" And ParentStr like '%"&ParentStr&"%'")
if not (rs.eof and rs.bof) then
do while not rs.eof
	if i=0 then
		iParentStr=rs(0)
	else
		iParentStr=iParentStr & "," & rs(0)
	end if
	i=i+1
rs.movenext
loop
end if
if i>0 then
	ParentStr=iParentStr & "," & oldboard
else
	ParentStr=oldboard
end if
'更新其原来所属论坛版面数
if depth>0 then
Dvbbs.Execute("update dv_board set child=child-"&child&" where boardid="&iparentid)
'更新其原来所属论坛数据，排序相当于剪枝而不需考虑
for i=1 to depth
	'得到其父类的父类的版面ID
	set rs=Dvbbs.Execute("select parentid from dv_board where boardid="&iparentid)
	if not (rs.eof and rs.bof) then
		iparentid=rs(0)
		Dvbbs.Execute("update dv_board set child=child-"&child&" where boardid="&iparentid)
	end if
next
end if
Conn.CommandTimeOut = 0
'更新论坛帖子数据
For i=0 to ubound(AllPostTable)
	Dvbbs.Execute("update "&AllPostTable(i)&" set boardid="&newboard&" where boardid in ("&ParentStr&")")
	'更新回收站部分内容
	Dvbbs.Execute("Update "&AllPostTable(i)&" Set LockTopic="&newboard&" Where BoardID=444 And LockTopic In ("&ParentStr&")")
	Response.Write "更新论坛帖子表" & AllPostTable(i) & "数据成功！<br>"
	Response.Flush
Next
Dvbbs.Execute("update dv_topic set boardid="&newboard&",mode=0 where boardid in ("&ParentStr&")")
Response.Write "更新主题表数据成功！<br>"
Response.Flush
Dvbbs.Execute("update dv_besttopic set boardid="&newboard&" where boardid in ("&ParentStr&")")
Response.Write "更新精华帖数据成功！<br>"
Response.Flush
'更新回收站部分内容
Dvbbs.Execute("Update Dv_Topic Set LockTopic="&newboard&" Where BoardID=444 And LockTopic In ("&ParentStr&")")
Response.Write "更新回收站数据成功！<br>"
Response.Flush
'shinzeal加入更新上传文件数据
Dvbbs.Execute("update DV_Upfile set F_boardid="&newboard&" where F_boardid in ("&ParentStr&")")
Response.Write "更新上传表数据成功！<br>"
Response.Flush
'删除被合并论坛
set rs=Dvbbs.Execute("select sum(postnum),sum(topicnum),sum(todayNum) from dv_board where RootID="&RootID&" And boardid in ("&ParentStr&")")
Dvbbs.Execute("delete from dv_board where RootID="&RootID&" And boardid in ("&ParentStr&")")
	'删除被合并论坛的自定义用户权限
	Dvbbs.Execute("DELETE FROM Dv_UserAccess WHERE NOT Uc_BoardID IN (SELECT BoardID FROM Dv_Board)")
'更新新论坛帖子计数
dim trs
set trs=Dvbbs.Execute("select ParentStr,boardid from dv_board where boardid="&newboard)
if trs(0)="0" then
ParentStr=trs(1)
else
ParentStr=trs(0)
end if
	'更新合并后版面帖子数信息
	Dvbbs.Execute("UPDATE Dv_Board SET Postnum = Postnum + " & Rs(0) & ", Topicnum = Topicnum + " & Rs(1) & ", Todaynum = Todaynum + " & Rs(2) & " WHERE Boardid = " & NewBoard)
	Response.Write "合并成功，已经将被合并论坛所有数据转入您所合并的论坛，请更新一下论坛数据。"
set rs=nothing
set trs=nothing
RestoreBoardCache()
End Sub
Sub RestoreBoardCache()
	Dim Board
	Dvbbs. LoadBoardList()
	For Each board in Application(Dvbbs.CacheName&"_boardlist").documentElement.selectNodes("board/@boardid")
		Dvbbs.LoadBoardData(board.text)
	Next
	If Request("action")="RestoreBoardCache" Then dv_suc("重建所有版面缓存成功！")
End Sub
%>