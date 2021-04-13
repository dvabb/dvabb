<!--#include file="../conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="../inc/dv_clsother.asp"-->
<!--#include file="../inc/GroupPermission.asp"-->
<!--#include file=../inc/md5.asp-->
<%
Head()
Server.ScriptTimeout=999999
dim Str
dim admin_flag
admin_flag=Split("11,12",",")
founderr=False 
CheckAdmin(admin_flag(0))
CheckAdmin(admin_flag(1))
call main()
footer()

Sub main()
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<th width="100%" colspan=2 style="text-align:center;">论坛管理
</th>
</tr>
<tr>
<td class="td2" colspan=2>
<p><B>注意</B>：<BR>①删除论坛同时将删除该论坛下所有帖子！删除分类同时删除下属论坛和其中帖子！ 操作时请完整填写表单信息；<BR>②如果选择<B>复位所有版面</B>，则所有版面都将作为一级论坛（分类），这时您需要重新对各个版面进行归属的基本设置，<B>不要轻易使用该功能</B>，仅在做出了错误的设置而无法复原版面之间的关系和排序的时候使用，在这里您也可以只针对某个分类进行复位操作(见分类的更多操作下拉菜单)，具体请看操作说明。<BR><font color=blue>每个版面的更多操作请见下拉菜单，操作前请仔细阅读说明，分类下拉菜单中比别的版面增加了分类排序和分类复位功能</font><BR>
<font color=red>如果您希望某个版面需要会员付出一定代价（货币）才能进入，可以在版面高级设置中设置相应版面进入所需的金币或点券数以及能访问的时间是多少</font>
</td>
</tr>
<tr>
<td class="td2" height=25>
<B>论坛操作选项</B></td>
<td class="td2"><a href="board.asp">论坛管理首页</a> | <a href="board.asp?action=add">新建论坛版面</a> | <a href="?action=settemplates">模板风格批量设置</a> | <a href="?action=orders">一级分类排序</a> | <a href="?action=boardorders">N级分类排序</a> | <a href="?action=RestoreBoard" onclick="{if(confirm('复位所有版面将把所有版面恢复成为一级大分类，复位后要对所有版面重新进行归属的基本设置，请慎重操作，确定复位吗?')){return true;}return false;}">复位所有版面</a> | <a href="?action=RestoreBoardCache" onclick="{if(confirm('有时候您对论坛版面的修改在前台看不出修改效果，这很可能是相应版面的缓存没有生效所致，在这里将重建所有版面的缓存，如果您的版面很多，这将消耗您一定的时间，确定吗?')){return true;}return false;}">重建版面缓存</a>
</td>
</tr>
</table>
<p></p>
<%
select case Request("action")
	case "add"
		call add()
	case "edit"
		call edit()
	case "savenew"'新增论坛
		call savenew()
	Case "savedit"
		call savedit()
	Case "del"
		call del()
	Case "orders"
		call orders()
	case "updatorders"
		call updateorders()
	Case "boardorders"
		call boardorders()
	case "updatboardorders"
		call updateboardorders()
	Case "mode"
		call mode()
	case "savemod"
		call savemod()
	Case "permission"
		call boardpermission()
	case "editpermission"
		call editpermission()
	case "RestoreBoard"
		call RestoreBoard()
	Case "RestoreBoardCache"
		Call RestoreBoardCache()
	Case "clearDate"
		Call clearDate
	Case "delDate"
		Call delDate
	Case "RestoreClass"
		Call RestoreClass
	Case "handorders"
		Call handorders
	Case "savehandorders"
		Call savehandorders
	Case "savesid"
		Call savesid
	Case "upallsid"
		Call upallsid
	Case "settemplates"
		Call Settemplates
	Case Else
		Call boardinfo()
end select
end Sub
Sub upallsid()
	Dim Sid,cid,board
	SID= Request("Sid")
'	cid=Request("CID")
	If SID="" Then SID=1
'	If CID="" Then CID=1
	SID=CLng(SID)
'	CID=CLng(CID)
	Dvbbs.Execute("Update Dv_Setup Set Forum_Sid="& SID )
	Dvbbs.Execute("Update Dv_Board Set Sid="& SID )
	Dvbbs.loadSetup()
	For Each board in Application(Dvbbs.CacheName&"_boardlist").documentElement.selectNodes("board/@boardid")
	Dvbbs.LoadBoardData(board.text)
	Next
	Dv_suc("论坛版面风格样式统一设置成功!")
End Sub
Sub savesid
	Dim i,boardid,TempStr
	Dim Templateslist,sid,j,bid,cid
	sid=""
	For Each TempStr in Request.form("upboardid")
		If Bid="" Then
			Bid=TempStr
		Else
			Bid=Bid&","&TempStr
		End If 
	Next
	Bid=split(Bid,",")
	For i=0 to UBound(bid)
		If sid="" Then
			sid=Request("sid"&bid(i))
		Else
			sid=sid&","&Request("sid"&bid(i))
		End If
'		If cid="" Then
'			cid=Request("cid"&bid(i))
'		Else
'			cid=cid&","&Request("cid"&bid(i))
'		End If
	Next
	sid=split(sid,",")
	For i=0 to UBound(bid)	
		Dvbbs.Execute("Update Dv_board set sid="&CLng(sid(i))&" Where BoardId="&Clng(bid(i))&" ")
		Dvbbs.LoadBoardData(Clng(bid(i)))
	Next 
	Dv_suc("论坛模板批量设置成功!")
End Sub

Sub Settemplates
'Application(Dvbbs.CacheName &"_style")
Dim reBoard_Setting,MoreMenu,i
Dim Templateslist,rs,SQL

%>

<form action ="board.asp?action=upallsid" method=post name="dv">
<table cellspacing="0" cellpadding="0" align="center" width="100%">
<tr> 
<th colspan="2" style="text-align:center;">模 板 统 一 设 置
</th>
</tr>
<tr>
<td width="70%" align=Left class="td2"><B>所有论坛版面风格模板设置为：</b>&nbsp; 
<%
	Dim forum_sid,iCssName,iCssID,iStyleName
	Dim Forum_cid
	set rs=dvbbs.execute("select forum_sid,Forum_CID from dv_setup")
	Forum_sid=rs(0)
'	Forum_CID=Rs(1)
	Rs.close:Set Rs=Nothing
%>
<Select Size=1 Name="sid">
<%
If Not TypeName(Application(Dvbbs.CacheName & "_style"))="DOMDocument" Then Dvbbs.Loadstyle()
For Each Templateslist in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style")
	Response.write Templateslist.getAttribute("id")&"--"&Templateslist.getAttribute("type")
	Response.Write "<Option value="""& Templateslist.selectSingleNode("@id").text &""""
	If Forum_sid = CLng(Templateslist.selectSingleNode("@id").text) Then Response.Write " selected "
	Response.Write ">"& Templateslist.selectSingleNode("@type").text &" </Option>"
Next
%>
</Select>
</td>
<td width="30%" align=Left  class="td2" >
<Input type="submit" name="Submit" value="设 定" class="button"></td>
</tr>
</table><BR>
</form>
<form action ="board.asp?action=savesid" method=post name="dv1">
<table cellspacing="0" cellpadding="0" align="center" width="100%">
<tr> 
<th width="70%">&nbsp;论坛版面
</th>
<th width="30%">采用风格样式
</th>
</tr>
<%
dim classrow
sql="select * from dv_board order by rootid,orders"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
do while not rs.eof
reBoard_Setting=split(rs("Board_setting"),",")
if classrow="td2" then
	classrow="td1"
else
	classrow="td2"
end if
%>
<tr> 
<td height="25"  class="<%=classrow%>">
<%if rs("depth")>0 then%>
<%for i=1 to rs("depth")%>
&nbsp;
<%next%>
<%end if%>
<%if rs("child")>0 then%><img src="../skins/default/plus.gif"><%else%><img src="../skins/default/nofollow.gif"><%end if%>
<%if rs("parentid")=0 then%><b><%end if%><%=rs("boardtype")%><%if rs("child")>0 then%>(<%=rs("child")%>)<%end if%>
<%if rs("parentid")=0 then%></b><%end if%>
</td>
<td align=Left  class="<%=classrow%>">
<Select Size=1 Name="sid<%=Rs("BoardID")%>">
<%
For Each Templateslist in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style")
	Response.Write "<Option value="""& Templateslist.selectSingleNode("@id").text &""""
	If Rs("SID") = CLng(Templateslist.selectSingleNode("@id").text) Then Response.Write " selected "
	Response.Write ">"& Templateslist.selectSingleNode("@type").text &" </Option>"
Next
%>
</Select>
<Input type="hidden" name="upboardid" value="<%=rs("boardid")%>">
</td></tr>
<%
Rs.movenext
loop
set rs=nothing
%>
<tr>
<td width=300 align=Left class="td2" >&nbsp;</td>
<td width=300 align=Left class="td2" ><input type="submit" name="Submit" value="设 定" class="button"></td>
</tr>
</table><BR><BR>
</form>
<%
End Sub 

sub boardinfo()
Dim reBoard_Setting,MoreMenu
Dim Rs,classrow,iii,SQL,i
%>
<div class=menuskin id=popmenu onmouseover="clearhidemenu();" onmouseout="dynamichide(event)" style="Z-index:100"></div>
<table width="100%" cellspacing="0" cellpadding="0" align="center">
<tr> 
<th width="35%">论坛版面
</th>
<th width="35%">操作
</th>
</tr>
<%
SQL="select boardid,boardtype,parentid,depth,child,Board_setting from dv_board order by rootid,orders"
SET Rs = Conn.Execute(SQL)
If Rs.eof Then
	Rs.close:Set Rs = Nothing
Else
SQL=Rs.GetRows(-1)
Rs.close:Set Rs = Nothing
For iii=0 To Ubound(SQL,2)
	reBoard_Setting=split(SQL(5,iii),",")
	if classrow="td2" then
		classrow="td1"
	else
		classrow="td2"
	end if
	Response.Write "<tr>"
	Response.Write "<td height=""25"" width=""55%"" class="
	Response.Write classrow 
	Response.Write ">"
	if SQL(3,iii)>0 then
		for i=1 to SQL(3,iii)
			Response.Write "&nbsp;&nbsp;"
		next
	end if
	if SQL(4,iii)>0 then
		Response.Write "<img src=""../skins/default/plus.gif"">"
	else
		Response.Write "<img src=""../skins/default/nofollow.gif"">"
	end if
	if SQL(2,iii)=0 then
		Response.Write "<b>"
	end if
	Response.Write SQL(1,iii)
	if SQL(4,iii)>0 then
		Response.Write "("
		Response.Write SQL(4,iii)
		Response.Write ")"
	end if
%>
</td>
<td width="45%" class="<%=classrow%>">
<a href="board.asp?action=add&editid=<%=SQL(0,iii)%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>添加版面</U></font></a> | <a href="board.asp?action=edit&editid=<%=SQL(0,iii)%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>基本设置</U></font></a> | <a href="BoardSetting.asp?editid=<%=SQL(0,iii)%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>高级设置</U></font></a>
<%

MoreMenu=MoreMenu & "<div class=menuitems><a href=update.asp?action=updat&submit=更新论坛数据&boardid="&SQL(0,iii)&" title=更新最后回复、帖子数、回复数><font color="&Dvbbs.mainsetting(3)&"><U>更新数据</U></font></a></div><div class=menuitems><a href=# onclick=alertreadme(\'清空将包括该论坛所有帖子置于回收站，确定清空吗?\',\'update.asp?action=delboard&boardid="&SQL(0,iii)&"\')><font color="&Dvbbs.mainsetting(3)&"><U>清空版面数据</U></font></a></div>"

if SQL(4,iii)=0 then
MoreMenu=MoreMenu & "<div class=menuitems><a href=# onclick=alertreadme(\'删除将包括该论坛的所有帖子，确定删除吗?\',\'board.asp?action=del&editid="&SQL(0,iii)&"\')><font color="&Dvbbs.mainsetting(3)&"><U>删除版面</U></font></a></div>"
else
MoreMenu=MoreMenu & "<div class=menuitems><a href=# onclick=alertreadme(\'该论坛含有下属论坛，必须先删除其下属论坛方能删除本论坛！\',\'#\')><font color="&Dvbbs.mainsetting(3)&"><U>删除版面</U></font></a></div>"
end if
MoreMenu=MoreMenu & "<div class=menuitems><a href=Board.asp?action=clearDate&boardid="&SQL(0,iii)&"><font color="&Dvbbs.mainsetting(3)&"><u>清理数据</u></font></a></div>"
If SQL(2,iii)=0 Then
	MoreMenu=MoreMenu & "<div class=menuitems><a href=# onclick=alertreadme(\'复位该分类将会把该分类下的所有版面都复位成二级版面，包括原来的多级分类都将复位成二级版面，请慎重操作，确定复位吗?\',\'?action=RestoreClass&classid="&SQL(0,iii)&"\')><font color="&Dvbbs.mainsetting(3)&"><u>复位该分类</u></font></a></div><div class=menuitems><a href=?action=handorders&classid="&SQL(0,iii)&"><font color="&Dvbbs.mainsetting(3)&"><u>分类排序(手动)</u></font></a></div>"
End If
%>
 | <a href="#" onMouseOver="showmenu(event,'<%=MoreMenu%>')" style="CURSOR:hand"><font color=<%=Dvbbs.mainsetting(3)%>><u>更多操作</u></font></a>
<%
if reBoard_Setting(2)=1 then
	Response.Write "<a href=board.asp?action=mode&boardid="&SQL(0,iii)&"><font color="&Dvbbs.mainsetting(3)&"><U>认证用户</U></font></a>"
end if
%>
</td></tr>
<%
MoreMenu=""
Next
End If
%>
</table><BR><BR>
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
<%
end sub

sub add()
dim rs_c,sql,Rs
Dim forum_sid,forum_cid,Style_Option,TempOption
Dim iCssName,iCssID,iStyleName
set rs_c= server.CreateObject ("adodb.recordset")
sql = "select * from dv_board order by rootid,orders"
rs_c.open sql,conn,1,1
	dim boardnum,i
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select Max(boardid) from dv_board"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
	boardnum=1
	else
	boardnum=rs(0)+1
	end if
	if isnull(boardnum) then boardnum=1
	if boardnum=444 then boardnum=445
	if boardnum=777 then boardnum=778
	rs.close
%>
<form action ="board.asp?action=savenew" method=post name=theform>
<input type="hidden" name="newboardid" value=<%=boardnum%>>
<table width="100%" border="0" cellspacing="1" cellpadding="0" align="center">
<tr> 
<th colspan=2 style="text-align:center;"><B>添加新论坛</th>
</tr>
<tr> 
<td width="100%" height=30 class="td2" colspan=2>
说明：<BR>1、添加论坛版面后，相关的设置均为默认设置，请返回论坛版面管理首页版面列表的高级设置中设置该论坛的相应属性，如果您想对该论坛做更具体的权限设置，请到<A HREF="board.asp?action=permission"><font color=blue>论坛权限管理</font></A>中设置相应用户组在该版面的权限。<BR>
2、<font color=blue>如果您添加的是论坛分类</font>，只需要在所属分类中选择作为论坛分类即可；<font color=blue>如果您添加的是论坛版面</font>，则要在所属分类中确定并选择该论坛版面的上级版面。
</td>
</tr>
<tr> 
<td width="40%" height=30 class="td1">论坛名称</td>
<td width="60%" class="td1"> 
<input type="text" name="boardtype" size="35">
</td>
</tr>
<tr> 
<td width="40%" height=24 class="td1">版面说明<BR>可以使用HTML代码</td>
<td width="60%" class="td1"> 
<textarea name="Readme" cols="50" rows="5"></textarea>
</td>
</tr>
<tr> 
<td width="40%" height=24 class="td1">版面规则<BR>可以使用HTML代码</td>
<td width="60%" class="td1"> 
<textarea name="Rules" cols="50" rows="5"></textarea>
</td>
</tr>
<tr> 
<td width="40%" height=30 class="td1"><U>所属类别</U></td>
<td width="60%" class="td1"> 
<select name="class" id="Boardid"></select>
<SCRIPT LANGUAGE="JavaScript">
<!--
BoardJumpListSelect('<%=Dvbbs.CheckNumeric(Request("editid"))%>','Boardid','做为论坛分类','0','0');
//-->
</SCRIPT>
</td>
</tr>
<tr> 
<td width="40%" height=30 class="td1"><U>使用样式风格</U><BR>相关样式风格中包含论坛颜色、图片<BR>等信息</td>
<td width="60%" class="td1">
<%
	set rs_c=dvbbs.execute("select forum_sid,forum_cid from dv_setup")
	Forum_sid=rs_c(0)
'	Forum_cid=rs_c(1)
	rs_c.close:Set rs_c=Nothing
%>
<Select Size=1 Name="sid">
<%
Dim Templateslist
For Each Templateslist in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style")
	Response.Write "<Option value="""& Templateslist.selectSingleNode("@id").text &""""
	If Forum_sid = CLng(Templateslist.selectSingleNode("@id").text) Then Response.Write " selected "
	Response.Write ">"& Templateslist.selectSingleNode("@type").text &" </Option>"
Next
%>
</td>
</tr>
<tr> 
<td width="40%" height=30 class="td1"><U>论坛版主</U><BR>多版主添加请用|分隔，如：沙滩小子|wodeail</td>
<td width="60%" class="td1"> 
<input type="text" name="boardmaster" size="35">
</td>
</tr>
<tr> 
<td width="40%" height=30 class="td1"><U>首页显示论坛图片</U><BR>出现在首页论坛版面介绍左边<BR>请直接填写图片URL</td>
<td width="60%" class="td1">
<input type="text" name="indexIMG" size="35">
</td>
</tr>
<tr> 
<td width="40%" height=24 class="td1">&nbsp;</td>
<td width="60%" class="td1"> 
<input type="submit" name="Submit" value="添加论坛" class="button">
</td>
</tr>
</table>
</form>
<%
set rs_c=nothing
set rs=nothing
end sub

sub edit()
dim rs_c,reBoard_Setting,rs,sql
Dim forum_sid,forum_cid,Style_Option,TempOption
Dim iCssName,iCssID,iStyleName,i
sql = "select * from dv_board order by rootid,orders"
set rs_c=Dvbbs.Execute(sql)
sql = "select * from dv_board where boardid="&Dvbbs.CheckNumeric(request("editid"))
set rs=Dvbbs.Execute(sql)
reBoard_Setting=split(rs("Board_setting"),",")

forum_sid=rs("sid")
'forum_cid=rs("cid")
%>
<form action ="board.asp?action=savedit" method=post name=theform>
<input type="hidden" name=editid value="<%=Request("editid")%>">
<table width="100%" border="0" cellspacing="1" cellpadding="0" align="center">
<tr> 
<th colspan=2 style="text-align:center;">编辑论坛：<%=rs("boardtype")%></th>
</tr>
<tr> 
<td width="100%" height=30 class="td2" colspan=2>
说明：<BR>1、添加论坛版面后，相关的设置均为默认设置，请返回论坛版面管理首页版面列表的高级设置中设置该论坛的相应属性，如果您想对该论坛做更具体的权限设置，请到<A HREF="board.asp?action=permission"><font color=blue>论坛权限管理</font></A>中设置相应用户组在该版面的权限。<BR>
2、<font color=blue>如果您添加的是论坛分类</font>，只需要在所属分类中选择作为论坛分类即可；<font color=blue>如果您添加的是论坛版面</font>，则要在所属分类中确定并选择该论坛版面的上级版面
</td>
</tr>
<tr> 
<td width="40%" height=30 class="td1">论坛名称</td>
<td width="60%" class="td1"> 
<input type="text" name="boardtype" size="35" value="<%=Server.htmlencode(rs("boardtype"))%>" >
</td>
</tr>
<tr> 
<td width="40%" height=24 class="td1">版面说明<BR>可以使用HTML代码</td>
<td width="60%" class="td1"> 
<textarea name="Readme" cols="50" rows="5"><%=server.HTMLEncode(Rs("readme")&"")%></textarea>
</td>
</tr>
<tr> 
<td width="40%" height=24 class="td1">版面规则<BR>可以使用HTML代码</td>
<td width="60%" class="td1"> 
<textarea name="Rules" cols="50" rows="5"><%=server.HTMLEncode(Rs("Rules")&"")%></textarea>
</td>
</tr>
<tr> 
<td width="40%" height=30 class="td1"><U>所属类别</U><BR>所属论坛不能指定为当前版面<BR>所属论坛不能指定为当前版面的下属论坛</td>
<td width="60%" class="td1"> 
<select name=class>
<option value="0">做为论坛分类</option>
<% do while not rs_c.EOF%>
<option value="<%=rs_c("boardid")%>" <% if cint(rs("parentid")) = rs_c("boardid") then%> selected <%end if%>>
<%if rs_c("depth")>0 then%>
<%for i=1 to rs_c("depth")%>&nbsp;&nbsp;|<%next%>
<%end if%>&nbsp;├&nbsp;<%=rs_c("boardtype")%></option>
<%
rs_c.MoveNext 
loop
rs_c.Close 
%>
</select>
</td>
</tr>
<tr> 
<td width="40%" height=30 class="td1"><U>使用样式风格</U><BR>相关样式风格中包含论坛颜色、图片<BR>等信息</td>
<td width="60%" class="td1">
<%
	set rs_c=dvbbs.execute("select forum_sid,forum_cid from dv_setup")
	Forum_sid=rs_c(0)
	Forum_cid=rs_c(1)
	rs_c.close:Set rs_c=Nothing
%>
<Select Size=1 Name="sid">
<%
Dim Templateslist
For Each Templateslist in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style")
	Response.Write "<Option value="""& Templateslist.selectSingleNode("@id").text &""""
	If rs("sid") = CLng(Templateslist.selectSingleNode("@id").text) Then Response.Write " selected "
	Response.Write ">"& Templateslist.selectSingleNode("@type").text &" </Option>"
Next
%>
</td>
</tr>
<tr> 
<td width="40%" height=30 class="td1"><U>论坛版主</U><BR>多斑竹添加请用|分隔，如：沙滩小子|wodeail</td>
<td width="60%" class="td1"> 
<input type="text" name="boardmaster" size="35" value='<%=rs("boardmaster")%>'>
<input type="hidden" name="oldboardmaster" value='<%=rs("boardmaster")%>'>
</td>
</tr>
<tr> 
<td width="40%" height=30 class="td1"><U>首页显示论坛图片</U><BR>出现在首页论坛版面介绍左边<BR>请直接填写图片URL</td>
<td width="60%" class="td1">
<input type="text" name="indexIMG" size="35" value="<%=rs("indexIMG")%>">
</td>
</tr>
<tr> 
<td width="40%" height=24 class="td1">&nbsp;</td>
<td width="60%" class="td1"> 
<input type="submit" name="Submit" value="提交修改" class="button">
</td>
</tr>
<tr> 
<td width="100%" height=30 class="td2" colspan=2 align=right>
<a href="board.asp?action=add&editid=<%=Request("editid")%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>添加版面</U></font></a> | <a href="board.asp?action=edit&editid=<%=Request("editid")%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>基本设置</U></font></a> | <a href="BoardSetting.asp?editid=<%=Request("editid")%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>高级设置</U></font></a>
<%if reBoard_Setting(2)=1 then%>
| <a href="board.asp?action=mode&boardid=<%=Request("editid")%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>认证用户</U></font></a>
<%end if%>
| <a href="update.asp?action=updat&submit=更新论坛数据&boardid=<%=Request("editid")%>" title="更新最后回复、帖子数、回复数"><font color="<%=Dvbbs.mainsetting(3)%>"><U>更新数据</U></font></a> | <a href="update.asp?action=delboard&boardid=<%=Request("editid")%>" onclick="{if(confirm('清空将包括该论坛所有帖子置于回收站，确定清空吗?')){return true;}return false;}"><font color="<%=Dvbbs.mainsetting(3)%>"><U>清空</U></font></a> | <%if rs("child")=0 then%><a href="board.asp?action=del&editid=<%=Request("editid")%>" onclick="{if(confirm('删除将包括该论坛的所有帖子，确定删除吗?')){return true;}return false;}"><font color="<%=Dvbbs.mainsetting(3)%>"><U>删除</U></a><%else%><a href="#" onclick="{if(confirm('该论坛含有下属论坛，必须先删除其下属论坛方能删除本论坛！')){return true;}return false;}"><font color="<%=Dvbbs.mainsetting(3)%>"><U>删除</U></a><%end if%>
| <a href="Board.asp?action=clearDate&boardid=<%=Request("editid")%>"> <font color="<%=Dvbbs.mainsetting(3)%>"><u>清理数据</u></a>
</td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
set rs_c=nothing
end sub

Sub Mode()
	Dim Boarduser
	Dim BoarduserNum
	Dim Rs,Sql
%>
<form action ="board.asp?action=savemod" method=post>
<table width="100%" cellspacing="1" cellpadding="1" align="center">
<tr> 
<th width="52%">说明：</th>
<th width="48%">操作：</th>
</tr>
<tr> 
<td width="52%" height=22 class=td1><B>论坛名称</B></td>
<td width="48%" class=td1> 
<%
Sql = "SELECT Boardid, Boardtype, Boarduser FROM Dv_Board WHERE Boardid = " & Dvbbs.CheckNumeric(Request("boardid"))
Set Rs = Dvbbs.Execute(Sql)
If Rs.Eof And Rs.Bof Then
	Response.Write "该版面并不存在或者该版面不是加密版面。"
Else
	Response.Write Rs(1)
	Response.Write "<input type=hidden value=" & Rs(0) & " name=boardid>"
	Boarduser = Rs(2)
End If
Set Rs = Nothing
%>
</td>
</tr>
<tr> 
<td width="52%" class=td1 valign=top><B>认证用户</B>：
<%
	If Not Isnull(Boarduser) Or Boarduser <> "" Then
		BoarduserNum = Split(Boarduser,",")
		Response.Write "（本版共有<font color=red>" & Ubound(BoarduserNum)+1 & "</font>位认证用户）"
	Else
		Response.Write "（本版暂时没有认证用户）"
	End If
%>
<br>
只有设定为认证论坛的论坛需要填写能够进入该版面的用户，每输入一个用户请确认用户名在论坛中存在，每个用户名用<B>回车</B>分开</font>
<%
If Clng(Dvbbs.Board_Setting(62))>0 Or Clng(Dvbbs.Board_Setting(63))>0 Then Response.Write "<BR><font color=blue>此版面设置了支付金币或点券方能进入，有效期为<font color=red>" & Clng(Dvbbs.Board_Setting(64)) & "</font>个月，请在每个用户名后面加上：=当前时间，每行效果如：admin="&Now&"</font>"
%>
</td>
<td width="48%" class=td1> 
<textarea cols="50" rows="3" name="vipuser" id="vipuser">
<%if not isnull(boarduser) or boarduser<>"" then
	response.write Replace(boarduser,",",Chr(10))
end if%></textarea>
<br><a href="javascript:admin_Size(-3,'vipuser')"><img src="skins/images/minus.gif" unselectable="on" border='0'></a> <a href="javascript:admin_Size(3,'vipuser')"><img src="skins/images/plus.gif" unselectable="on" border='0'></a>
</td>
</tr>
<tr> 
<td width="52%" height=22 class=td1>&nbsp;</td>
<td width="48%" class=td1> 
<input type="submit" name="Submit" value="设 定" class="button">
</td>
</tr>
</table>
</form>
<%
End Sub 

'保存编辑论坛认证用户信息
'入口：用户列表字符串
sub savemod()
	dim boarduser
	dim boarduser_1
	dim userlen
	dim updateinfo,i
	'清理付费论坛过期的认证用户 2005-3-10 Dv.Yz
	Dim Get_BoardUser_Money, BoardUser_Money
	Get_BoardUser_Money = False
	If Clng(Dvbbs.Board_Setting(62))>0 Or Clng(Dvbbs.Board_Setting(63))>0 Then Get_BoardUser_Money = True
	
	If trim(request("vipuser"))<>"" then
		boarduser=Replace(request("vipuser"),"'","")
		boarduser=split(boarduser,chr(13)&chr(10))
		For i = 0 To Ubound(Boarduser)
			If Not (Boarduser(i) = "" Or Boarduser(i) = " ") Then
				If Get_BoardUser_Money Then
					BoardUser_Money = Split(Boarduser(i),"=")
					If Not DateDiff("d",BoardUser_Money(1),Now()) > Cint(Dvbbs.Board_Setting(64))*30 Then
						Boarduser_1 = "" & Boarduser_1 & "" & Boarduser(i) & ","
					End If
				Else
					Boarduser_1 = "" & Boarduser_1 & "" & Boarduser(i) & ","
				End If
			End If
		Next
		userlen=len(boarduser_1)
		if boarduser_1<>"" then
			boarduser=left(boarduser_1,userlen-1)
			updateinfo=" boarduser='"&boarduser&"' "
			Dvbbs.Execute("update dv_board set "&updateinfo&" where boardid="&Dvbbs.CheckNumeric(request("boardid")))
			Dv_suc("论坛设置成功!<LI>成功添加认证用户："&boarduser&"<LI><a href=""?action=RestoreBoardCache"" >请执行重建版面缓存才能生效</a><br>")
			RestoreBoardCache()
		else
			Errmsg = errmsg + "你没有添加认证用户！"'response.write "你没有添加认证用户！"
			Dvbbs_Error()
			Exit Sub
		end if
	Else
		Errmsg = errmsg + "你没有添加认证用户！"'response.write "<p><font color=red>你没有添加认证用户</font><br><br>"
		Dvbbs_Error()
	End If
	
End Sub

'保存添加论坛信息
Sub savenew()
	If request("boardtype")="" Then
		Errmsg=Errmsg+"<br>"+"<li>请输入论坛名称。"
		founderr=true
	End If
	If request("class")="" Then
		Errmsg=Errmsg+"<br>"+"<li>请选择论坛分类。"
		founderr=true
	End If
	If request("readme")="" Then
		Errmsg=Errmsg+"<br>"+"<li>请输入论坛说明。"
		founderr=true
	End If
	If founderr=true Then
		dvbbs_error()
		exit sub
	End If
	Dim boardid,rootid,parentid,depth,orders,Fboardmaster,maxrootid,parentstr,rs,SQL
	If request("class")<>"0" Then
		Set rs=Dvbbs.Execute("select rootid,boardid,depth,orders,boardmaster,ParentStr from dv_board where boardid="&Dvbbs.CheckNumeric(request("class")))
		rootid=rs(0)
		parentid=rs(1)
		depth=rs(2)
		orders=rs(3)
		If depth+1>20 Then
			Errmsg="本论坛限制最多只能有20级分类"
		  dvbbs_error()
		  Exit Sub
		 End If 
		parentstr=rs(5)
	Else
		Set rs=Dvbbs.Execute("select max(rootid) from dv_board")
	  maxrootid=rs(0)+1
		If IsNull(MaxRootID) Then MaxRootID=1
	End If
	sql="select boardid from dv_board where boardid="&Dvbbs.CheckNumeric(request("newboardid"))
	Set rs=Dvbbs.Execute(sql)
	If not (rs.eof and rs.bof) then
		Errmsg="您不能指定和别的论坛一样的序号。"
		dvbbs_error()
		exit sub
	Else
		boardid=request("newboardid")
	End If
	Dim trs,forumuser,setting
	Set trs=Dvbbs.Execute("select * from dv_setup")
	Setting=Split(trs("Forum_Setting"),"|||")
	forumuser=Setting(2)
	set rs = server.CreateObject ("adodb.recordset")
	sql = "select * from dv_board"
	rs.Open sql,conn,1,3
	rs.AddNew
	If request("class")<>"0" Then
		rs("depth")=depth+1
		rs("rootid")=rootid
		rs("orders") = Request.form("newboardid")
		rs("parentid") = Request.Form("class")
		if ParentStr="0" then
		rs("ParentStr")=Request.Form("class")
	Else
	 rs("ParentStr")=ParentStr & "," & Request.Form("class")
	End If
	Else
		rs("depth")=0
		rs("rootid")=maxrootid
		rs("orders")=0
		rs("parentid")=0
		rs("parentstr")=0
		end if
		rs("boardid") = Request.form("newboardid")
		rs("boardtype") = request.form("boardtype")
		rs("readme") = Request.form("readme")
		rs("Rules") = Request.form("Rules")
		rs("TopicNum") = 0
		rs("PostNum") = 0
		rs("todaynum") = 0
		rs("child")=0
		rs("LastPost")="$0$"&Now()&"$$$$$"
		rs("Board_Setting")="0,0,0,0,0,1,1,1,1,1,1,1,1,1,1,1,16240,3,0,gif|jpg|jpeg|bmp|png|rar|txt|zip|mid,0,0,1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1,0,1,100,20,10,9,normal,1,10,10,0,0,0,0,1,0,0,1,4,0,0,0,200,0,0,,$$,0,0,0,1,0|0|0|0|0|0|0|0|0,0|0|0|0|0|0|0|0|0,0,0,0,0,0,0,0,0,0,灌水|广告|奖励|惩罚|好文章|内容不符|重复发帖,0,1,0,24,0,0"
		rs("sid")=Dvbbs.CheckNumeric(request.form("sid"))
'		rs("cid")=Dvbbs.CheckNumeric(request.form("cid"))
		rs("board_ads")=trs("forum_ads")
		rs("board_user")=forumuser
		If Request("boardmaster")<>"" Then 
			rs("boardmaster") = Request.form("boardmaster")
		End If
	If request.form("indexIMG")<>"" Then
		rs("indexIMG")=request.form("indexIMG")
	End If
	rs.Update 
	rs.Close
	If Request("boardmaster")<>"" Then Call addmaster(Request("boardmaster"),"none",0)
	dv_suc("论坛添加成功！<br>该论坛目前高级设置为默认选项，建议您返回论坛管理中心重新设置该论坛的高级选项，<A HREF=BoardSetting.asp?editid="&Request.form("newboardid")&">点击此处进入该版面高级设置</A><br>" & str)
	set rs=nothing
	trs.close
	set trs=nothing
	CheckAndFixBoard 0,1
	RestoreBoardCache()
End Sub

'保存编辑论坛信息
Sub savedit()
	if clng(request("editid"))=clng(request("class")) then
		Errmsg="所属论坛不能指定自己"
		dvbbs_error()
		exit sub
	end if
	dim newboardid,maxrootid,readme,Rules
	dim parentid,boardmaster,depth,child,ParentStr,rootid,iparentid,iParentStr
	dim trs,brs,mrs
	Dim iii,rs,sql
	set rs = server.CreateObject ("adodb.recordset")
	sql = "select * from dv_board where boardid="&request("editid")
	rs.Open sql,conn,1,3
	newboardid=rs("boardid")
	parentid=rs("parentid")
	iparentid=rs("parentid")
	boardmaster=rs("boardmaster")
	ParentStr=rs("ParentStr")
	depth=rs("depth")
	child=rs("child")
	rootid=rs("rootid")
	'判断所指定的论坛是否其下属论坛
	if ParentID=0 then
		if clng(request("class"))<>0 then
		set trs=Dvbbs.Execute("select rootid from dv_board where boardid="&request("class"))
		if rootid=trs(0) then
			errmsg="您不能指定该版面的下属论坛作为所属论坛1"
			dvbbs_error()
			exit sub
		end if
		end if
	else
		set trs=Dvbbs.Execute("select boardid from dv_board where ParentStr like '%"&ParentStr&","&newboardid&"%' and boardid="&request("class"))
		if not (trs.eof and trs.bof) then
			errmsg="您不能指定该版面的下属论坛作为所属论坛2"
			dvbbs_error()
			exit sub
		end if
	end if

	If parentid=0 then
		parentid=rs("boardid")
		iparentid=0
	Else
		Set mrs=Dvbbs.Execute("select max(rootid) from dv_board")
			Maxrootid=mrs(0)+1
		mrs.close:Set mrs=Nothing
		rs("rootid")=Maxrootid
	End if
	rs("boardtype") = Request.Form("boardtype")	'取消JS过滤。
	rs("parentid") = Request.Form("class")
	rs("boardmaster") = Request("boardmaster")
	rs("readme") = Request("readme")
	rs("Rules") = Request.form("Rules")
	rs("indexIMG")=request.form("indexIMG")
	rs("sid")=Cint(request.form("sid"))
'	rs("cid")=Cint(request.form("cid"))
	rs.Update 
	rs.Close:set rs=nothing
	if request("oldboardmaster")<>Request("boardmaster") then call addmaster(Request("boardmaster"),request("oldboardmaster"),1)
	
	dv_suc("论坛修改成功！<br>" & str)
	CheckAndFixBoard 0,1
	Boardchild()
	RestoreBoardCache()
End sub

'删除版面，删除版面帖子，入口：版面ID
Sub Del()
	Dim Trs,EditId
	EditId = Dvbbs.CheckNumeric(Request("editid"))
	'更新其上级版面论坛数，如果该论坛含有下级论坛则不允许删除
	Set tRs = Dvbbs.Execute("SELECT RootID FROM Dv_Board WHERE BoardID = " & EditId)
	Dim UpdateRootID,Rs,sql,i
	UpdateRootID = tRs(0)
	Set Rs = Dvbbs.Execute("SELECT ParentStr, Child, Depth FROM Dv_Board WHERE BoardID = " &  EditId)
	If Not (Rs.Eof And Rs.Bof) Then
		If Rs(1) > 0 Then
			Response.Write "该论坛含有下属论坛，请删除其下属论坛后再进行删除本论坛的操作"
			Exit Sub
		End If
		'如果有上级版面，则更新数据
		If Rs(2) > 0 Then
			Dvbbs.Execute("UPDATE Dv_Board SET Child = Child - 1 WHERE BoardID IN (" & Rs(0) & ")")
		End If
		Sql = "DELETE FROM Dv_Board WHERE Boardid = " & EditId
		Dvbbs.Execute(Sql)
		For i = 0 To Ubound(AllPostTable)
			Sql = "DELETE FROM " & AllPostTable(i) & " WHERE BoardID = " & EditId
			Dvbbs.Execute(Sql)
		Next
		Dvbbs.Execute("DELETE FROM Dv_Topic WHERE BoardID = " & EditId)
		Dvbbs.Execute("DELETE FROM Dv_BestTopic WHERE BoardID = " & EditId)
		Dvbbs.Execute("DELETE FROM Dv_Upfile WHERE F_BoardID = " & EditId)
		Dvbbs.Execute("DELETE FROM Dv_Appraise WHERE BoardID = " & EditId)
		'删除被删除论坛的自定义用户权限 2004-11-15 Dv.Yz
		Dvbbs.Execute("DELETE FROM Dv_UserAccess WHERE NOT Uc_BoardID IN (SELECT BoardID FROM Dv_Board)")
	End If
	Set Rs = Nothing
	CheckAndFixBoard 0,1
	RestoreBoardCache()
	Dv_suc("论坛删除成功！")
End Sub

sub orders()
	Dim rs,SQL
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" aligs") = Request.form("Rules")
	rs("indexIMG")=request.form("indexIMG")
	rs("sid")=Cint(request.form("sid"))
'	rs("cid")=Cint(request.form("cid"))
	rs.Update 
	rs.Close:set rs=nothing
	if request("oldboardmaster")<>Request("boardmaster") then call addmaster(Request("boardmaster"),request("oldboardmaster"),1)
	
	dv_suc("璁哄潧淇