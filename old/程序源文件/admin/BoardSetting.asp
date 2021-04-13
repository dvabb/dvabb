<!--#include file =../conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/ArrayList.asp" -->
<!--#include file="../inc/md5.asp"-->
<%
Head()
Dim Board_Setting,admin_flag
admin_flag = ",9,"
CheckAdmin(admin_flag)
If request("action")="save" Then
	Call saveconst()
Else
	Call consted()
End if
Footer()

Sub consted()
Dim rs,i,j
if not isnumeric(request("editid")) then
	Errmsg=ErrMsg + "<BR><li>错误的版面信息"
	dvbbs_error()
	exit sub
end if
set rs=Dvbbs.Execute("select * from dv_board where boardid="&request("editid"))
Board_Setting=split(rs("board_setting"),",")
%>
<table width="100%" cellspacing="1" cellpadding="1" align="center">
<tr><th colspan="7">论坛高级设置 → <%=rs("boardtype")%></th></tr>
<tr> 
<td width="100%" class=td1 colspan=7 height=25>
说明：<BR>
1、请仔细设置下面的高级选项，Flash标签如果打开，对安全有一定影响，请根据您的具体情况考虑。<BR>
2、您可以将高级设置的某项设置（选择该行设置右边的复选框）保存到所有版面、相同分类下所有版面（不包括分类）、相同分类下所有版面（包括分类）、同分类同级别版面，该项设置请慎重操作。<BR>
3、<font color=red>注意，选择批量更新包括主题将会使用相同设置</font>。
</td>
</tr>
<form method="POST" action="boardsetting.asp?action=save">
<!-- <form method="POST" action="boardsetting.asp?action=save"> -->


<input type="hidden" value="<%=request("editid")%>" name="editid">

<tr> 
<td width="100%" class=td1 colspan=7 height=25>
<font color=blue>
这里指的分类仅指一级分类，而不是该版面的上级版面</font>，比如您目前设置的是一个五级版面，选择了相同分类下所有版面都更新，那么这里将更新包括该分类的一级、二级、三级、四级所有版面，如果您担心更新范围太大，可以选择更新同分类同级别版面。
</td>
</tr>
<!-- 版面列表 -->
<tr>
<td rowspan=200 valign=top>
版面设置保存选项<br />请按 CTRL 键多选
<select name="getboardid" size="28" style="width:200px" multiple>
<%
set rs=Dvbbs.Execute("select boardid,boardtype,depth from dv_board order by rootid,orders")
do while not rs.eof
Response.Write "<option "
if rs(0)=CLng(request("editid")) then
Response.Write " selected"
end if
Response.Write " value="&rs(0)&">"
Select Case rs(2)
	Case 0
		Response.Write "╋"
	Case 1
		Response.Write "&nbsp;&nbsp;├"
End Select
If rs(2)>1 Then
	For ii=2 To rs(2)
		Response.Write "&nbsp;&nbsp;│"
	Next
	Response.Write "&nbsp;&nbsp;├"
End If
Response.Write rs(1)
Response.Write "</option>"
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
</td>
</tr>
<!-- 版面列表 -->

<!-- 高级设置 -->

<tr><td height="25" colspan="6" align=center><INPUT TYPE="checkbox" class="checkbox" NAME="chkall" onclick="CheckAll(this.form);">[全选]
编辑版块信息(选中则设置)</td></tr>
<tr><th height="25" colspan="6" align=left> &nbsp;功能设置导航</th></tr>
<tr> 
<td width="90%" class=td1 colspan=6 height=25>
[<a href="#setting1">基本属性</a>]
[<a href="#setting2">访问权限</a>]
[<a href="#setting3">前台管理权限</a>]
[<a href="#setting4">发贴相关</a>]
[<a href="#setting5">帖子列表显示</a>]
[<a href="#setting6">帖子内容显示</a>]
[<a href="#setting7">附件限制设置</a>]
[<a href="#setting8">论坛专题设置</a>]
[<a href="#setting9">论坛虚拟形象设置</a>]
</td>
</tr>

<tr><th height="25" colspan="6" id=tabletitlelink align=left> &nbsp;<a name="setting1"></a>基本属性[<a href="#top">顶部</a>]</th></tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(50)"></td>
<td colspan=2 class=td1>
<U>外部连接</U><BR>填写本内容后，在论坛列表点击此版面将自动切换到该网址<BR>请填写URL绝对路径</td>
<td colspan=2 class=td1>
<input type=text name="Board_Setting(50)" value="<%=Board_Setting(50)%>" size=50>
</td>
<input type="hidden" id="b0" value="<b>外部连接</b><br><li>填写本内容后，在论坛列表点击此版面将自动切换到该网址<br><li>请填写URL绝对路径">
<td class=td1><a href=# onclick="helpscript(b0);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(51)"></td>
<td colspan=2 class=td2>
<U>分论坛LOGO</U><BR>填写图片的相对或绝对路径，不填写则当前版面LOGO为论坛设置中LOGO</td>
<td colspan=2 class=td2>
<input type=text name="Board_Setting(51)" value="<%=Board_Setting(51)%>" size=50>
</td>
<input type="hidden" id="ba1" value="<b>分论坛LOGO</b><br><li>填写图片的相对或绝对路径，不填写则当前版面LOGO为论坛设置中LOGO">
<td class=td2><a href=# onclick="helpscript(ba1);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(40)"></td>
<td colspan=2 class=td1>
<U>是否采用版主继承制度</U></td>
<td colspan=2 class=td1>
<input type=radio class="radio" name="Board_Setting(40)" value=0 <%if Board_Setting(40)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio class="radio" name="Board_Setting(40)" value=1 <%if Board_Setting(40)="1" then%>checked<%end if%>>开放&nbsp;
</td>
<input type="hidden" id="b6" value="<b>是否采用版主继承制度</b><br><li>如果采用该制度，则上级论坛版主可管理下级论坛相关信息">
<td class=td1><a href=# onclick="helpscript(b6);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(39)"></td>
<td colspan=2 class=td2>
<U>论坛列表显示下属论坛风格</U><BR></td>
<td colspan=2 class=td2>
<input type=radio class="radio" name="Board_Setting(39)" value=0 <%if Board_Setting(39)="0" then%>checked<%end if%>>列表&nbsp;
<input type=radio class="radio" name="Board_Setting(39)" value=1 <%if Board_Setting(39)="1" then%>checked<%end if%>>简洁&nbsp;
</td>
<input type="hidden" id="b7" value="<b>论坛列表显示下属论坛风格</b><br><li>当该论坛有下属论坛的时候生效">
<td class=td2><a href=# onclick="helpscript(b7);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(41)"></td>
<td colspan=2 class=td1>
<U>论坛列表简洁风格一行版面数</U></td>
<td colspan=2 class=td1>
<input type=text size=10 name="Board_Setting(41)" value="<%=Board_Setting(41)%>"> 个
</td>
<input type="hidden" id="b8" value="<b>论坛列表简洁风格一行版面数</b><br><li>当论坛列表开启了下属论坛风格为简洁，此选项有效，此选项为设置简洁论坛列表风格一行排列版面数">
<td class=td1><a href=# onclick="helpscript(b8);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(36)"></td>
<td colspan=2 class=td2>
<U>是否公开论坛事件中的操作者</U></td>
<td colspan=2 class=td2>
<input type=radio class="radio" name="Board_Setting(36)" value=0 <%if Board_Setting(36)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio class="radio" name="Board_Setting(36)" value=1 <%if Board_Setting(36)="1" then%>checked<%end if%>>是&nbsp;
</td>
<input type="hidden" id="b12" value="<b>是否公开论坛事件中的操作者</b><br><li>论坛中对帖子的删除、固顶、设置精华等操作都是要记录操作者和操作内容的，管理员默认可看到这些操作内容，一般用户如果打开了此选项，他们将能看到操作者">
<td class=td1><a href=# onclick="helpscript(b12);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr><th height="25" colspan="6" id=tabletitlelink align=left>  &nbsp;<a name="setting2"></a>访问权限相关[<a href="#top">顶部</a>]</th></tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(43)"></td>
<td colspan=2 class=td1>
<U>本论坛作为分类论坛不允许发贴</U></td>
<td colspan=2 class=td1>
<input type=radio class="radio" name="Board_Setting(43)" value=0 <%if Board_Setting(43)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio class="radio" name="Board_Setting(43)" value=1 <%if Board_Setting(43)="1" then%>checked<%end if%>>是&nbsp;
</td>
<input type="hidden" id="b1" value="<b>本论坛作为分类论坛不允许发贴</b><br><li>如果已经有贴则显示或者您可以转移到别的论坛<br><li>选择了该项后所有会员均不能在本版发贴/回帖等操作">
<td class=td1><a href=# onclick="helpscript(b1);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(0)"></td>
<td colspan=2 class=td2>
<U>是否锁定论坛</U></td>
<td colspan=2 class=td2>
<input type=radio class="radio" name="Board_Setting(0)" value=0 <%if Board_Setting(0)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio class="radio" name="Board_Setting(0)" value=1 <%If Board_Setting(0)="1" then%>checked<%end if%>>是&nbsp;
</td>
<input type="hidden" id="b2" value="<b>是否锁定论坛</b><br><li>锁定论坛只有管理员和该版面版主可进">
<td class=td2><a href=# onclick="helpscript(b2);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(1)"></td>
<td colspan=2 class=td1>
<U>是否隐藏论坛</U></td>
<td colspan=2 class=td1>
<input type=radio class="radio" name="Board_Setting(1)" value=0 <%If Board_Setting(1)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio class="radio" name="Board_Setting(1)" value=1 <%if Board_Setting(1)="1" then%>checked<%end if%>>是&nbsp;
</td>
<input type="hidden" id="b3" value="<b>是否隐藏论坛</b><br><li>隐藏论坛只有管理员和该版面版主可见和进入<br><li>如果用户组或论坛权限管理或用户权限管理中允许则用户可见和进入<br><li>本限制对一级论坛不生效">
<td class=td1><a href=# onclick="helpscript(b3);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(2)"></td>
<td colspan=2 class=td2>
<U>是否认证论坛</U></td>
<td colspan=2 class=td2>
<input type=radio class="radio" name="Board_Setting(2)" value=0 <%if Board_Setting(2)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio class="radio" name="Board_Setting(2)" value=1 <%if Board_Setting(2)="1" then%>checked<%end if%>>是&nbsp;
</td>
<input type="hidden" id="b4" value="<b>是否认证论坛</b><br><li>认证论坛只有管理员和该版面版主可见和进入<br><li>认证论坛对认证用户的添加和管理在版面管理中有连接<br><li>设置了本选项后只有认证用户可进入">
<td class=td2><a href=# onclick="helpscript(b4);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>

<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(3)"></td>
<td colspan=2 class=td1>
<U>帖子审核制度</U></td>
<td colspan=2 class=td1>
<input type=radio class="radio" name="Board_Setting(3)" value=0 <%if Board_Setting(3)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio class="radio" name="Board_Setting(3)" value=1 <%if Board_Setting(3)="1" then%>checked<%end if%>>开放&nbsp;
</td>
<input type="hidden" id="b5" value="<b>帖子审核制度</b><br><li>版主、管理员和开放权限用户可进行审核帖子<br><li>版主、管理员和开放权限用户可直接发贴<br><li>一般用户需审核后帖子方可见">
<td class=td1><a href=# onclick="helpscript(b5);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(57)"></td>
<td colspan=2 class=td2>
<U>扩展审核制度</U></td>
<td colspan=2 class=td2>
<input type=radio class="radio" name="Board_Setting(57)" value=0 <%if Board_Setting(57)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio class="radio" name="Board_Setting(57)" value=1 <%if Board_Setting(57)="1" then%>checked<%end if%>>开放&nbsp;
<input type="hidden" id="bnew" value="<b>扩展帖子审核制度</b><br><li>版主、管理员和开放权限用户可进行审核帖子<br><li>版主、管理员和开放权限用户可直接发贴<br><li>一般用户如发贴内容如果有被过滤的敏感字需审核后帖子方可见,<br>如果无被过滤的内容，则可免审核发贴。">
</td>
<td class=td2><a href=# onclick="helpscript(bnew);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(58)"></td>
<td colspan=2 class=td1>
<U>敏感字设置</U></td>
<td colspan=2 class=td1>
<input type="text" Name=Board_Setting(58) Value="<%=Board_Setting(58)%>" Size=50><br>可设置多个敏感字中间用"|"分隔如不填写可以填0
<input type="hidden" id="bnewS" value="<b>敏感字设置</b><br><li>可设置多个敏感字中间用 | 分隔">
</td>
<td class=td1><a href=# onclick="helpscript(bnewS);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(18)"></td>
<td colspan=2 class=td2>
<U>允许同时在线数</U><BR>不限制则设置为0</td>
<td colspan=2 class=td2>
<input type=text size=10 name="Board_Setting(18)" value="<%=Board_Setting(18)%>"> 人
</td>
<input type="hidden" id="b9" value="<b>允许同时在线数</b><br><li>不限制则设置为0，如设置了允许同时在线数，则当论坛在线人数超过此数字的时候未登录用户将不能访问该版面">
<td class=td2><a href=# onclick="helpscript(b9);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(21)"></td>
<td colspan=2 class=td1>
<U>论坛定时设置</U></td>
<td colspan=2 class=td1>
<input type=radio class="radio" name="Board_Setting(21)" value="0" <%If Board_Setting(21)="0" Then %>checked <%End If%>>关 闭</option>
<input type=radio class="radio" name="Board_Setting(21)" value="1" <%If Board_Setting(21)="1" Then %>checked <%End If%>>定时关闭</option>
<input type=radio class="radio" name="Board_Setting(21)" value="2" <%If Board_Setting(21)="2" Then %>checked <%End If%>>定时只读</option>
</td>
<input type="hidden" id="b10" value="<b>定时设置选择:</b><br><li>在这里您可以设置是否起用定时的各种功能，如果开启了本功能，请设置好下面选项中的论坛设置时间，论坛该版面将在您规定的时间内有指定的设置">
<td class=td1><a href=# onclick="helpscript(b10);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(22)(22)"></td>
<td colspan=2 class=td2>
<U>定时设置</U><BR>请根据需要选择开或关</td></td>
<td colspan=2 class=td2>
<%
Board_Setting(22)=split(Board_Setting(22),"|")
If UBound(Board_Setting(22))<2 Then 
	Board_Setting(22)="1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1"
	Board_Setting(22)=split(Board_Setting(22),"|")
End If
For i= 0 to UBound(Board_Setting(22))
If i<10 Then Response.Write "&nbsp;"
%>
 <%=i%>点：<input type="checkbox" class="checkbox" name="Board_Setting(22)<%=i%>" value="<%=Board_Setting(22)(i)%>" <%If Board_Setting(22)(i)="1" Then %>checked<%End If%>>开
   
 <%
 If (i+1) mod 4 = 0 Then Response.Write "<br>"
 Next
 %>
</td>
<input type="hidden" id="b11" value="<b>论坛开放时间</b><br><li>设置了本选项必须同时打开是否起用定时开关论坛设置才有效，设置了此选项，论坛该版面将在您规定的时间内给用户开放">
<td class=td2><a href=# onclick="helpscript(b11);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<%
Dim VisitConfirm
VisitConfirm=Split(Board_Setting(54),"|")
IF Ubound(VisitConfirm)<>8 Then
	Redim VisitConfirm(8)
	For i=0 To 8
	VisitConfirm(i)=0
	Next
End If
%>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(54)(0)"></td>
<td colspan=2 class=td1>
<U>用户至少文章数</U></td>
<td colspan=2 class=td1>
<input type=text size=10 name="Board_Setting(54)(0)" value="<%=VisitConfirm(0)%>">
</td>
<input type="hidden" id="VisitConfirm1" value="<b>用户至少文章数</b><br><li>当用户发表的文章达到此设置时，才能拥有访问权限！<li>不限制设置为0">
<td class=td1><a href=# onclick="helpscript(VisitConfirm1);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(54)(1)"></td>
<td colspan=2 class=td2>
<U>用户至少积分</U></td>
<td colspan=2 class=td2>
<input type=text size=10 name="Board_Setting(54)(1)" value="<%=VisitConfirm(1)%>">
</td>
<input type="hidden" id="VisitConfirm2" value="<b>用户至少积分值</b><br><li>当用户的积分值达到此设置时，才能拥有访问权限！<li>不限制设置为0">
<td class=td2><a href=# onclick="helpscript(VisitConfirm2);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 

<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(54)(2)"></td>
<td colspan=2 class=td1>
<U>用户至少金钱</U></td>
<td colspan=2 class=td1>
<input type=text size=10 name="Board_Setting(54)(2)" value="<%=VisitConfirm(2)%>">
</td>
<input type="hidden" id="VisitConfirm3" value="<b>用户至少金钱数</b><br><li>当用户的金钱达到此设置时，才能拥有访问权限！<li>不限制设置为0">
<td class=td1><a href=# onclick="helpscript(VisitConfirm3);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(54)(3)"></td>
<td colspan=2 class=td2>
<U>用户至少魅力</U></td>
<td colspan=2 class=td2>
<input type=text size=10 name="Board_Setting(54)(3)" value="<%=VisitConfirm(3)%>">
</td>
<input type="hidden" id="VisitConfirm4" value="<b>用户至少魅力</b><br><li>当用户的魅力值达到此设置时，才能拥有访问权限！<li>不限制设置为0">
<td class=td2><a href=# onclick="helpscript(VisitConfirm4);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(54)(4)"></td>
<td colspan=2 class=td1>
<U>用户至少威望</U></td>
<td colspan=2 class=td1>
<input type=text size=10 name="Board_Setting(54)(4)" value="<%=VisitConfirm(4)%>">
</td>
<input type="hidden" id="VisitConfirm5" value="<b>用户至少威望</b><br><li>当用户威望达到此设置时，才能拥有访问权限！<li>不限制设置为0">
<td class=td1><a href=# onclick="helpscript(VisitConfirm5);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(54)(5)"></td>
<td colspan=2 class=td2>
<U>用户至少精华文章</U></td>
<td colspan=2 class=td2>
<input type=text size=10 name="Board_Setting(54)(5)" value="<%=VisitConfirm(5)%>">
</td>
<input type="hidden" id="VisitConfirm6" value="<b>用户至少精华文章数</b><br><li>当用户发表的精华文章达到此设置时，才能拥有访问权限！<li>不限制设置为0">
<td class=td2><a href=# onclick="helpscript(VisitConfirm6);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(54)(6)"></td>
<td colspan=2 class=td1>
<U>用户被删帖子数上限</U></td>
<td colspan=2 class=td1>
<input type=text size=10 name="Board_Setting(54)(6)" value="<%=VisitConfirm(6)%>">
</td>
<input type="hidden" id="VisitConfirm7" value="<b>用户被删帖子数上限</b><br><li>当用户被删帖子数超过此设置时，不能访问该分版！<li>不限制设置为0">
<td class=td1><a href=# onclick="helpscript(VisitConfirm7);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(54)(7)"></td>
<td colspan=2 class=td2>
<U>至少注册时间（单位为分钟）</U></td>
<td colspan=2 class=td2>
<input type=text size=10 name="Board_Setting(54)(7)" value="<%=VisitConfirm(7)%>">
</td>
<input type="hidden" id="VisitConfirm8" value="<b>用户至少注册时间</b><br><li>注册时间是指用户注册多少分钟后可进入论坛。<li>单位为分钟。<li>不限制设置为0">
<td class=td2><a href=# onclick="helpscript(VisitConfirm8);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(54)(8)"></td>
<td colspan=2 class=td1>
<U>至少上传文件个数</U></td>
<td colspan=2 class=td1>
<input type=text size=10 name="Board_Setting(54)(8)" value="<%=VisitConfirm(8)%>">
</td>
<input type="hidden" id="VisitConfirm9" value="<b>用户至少上传文件个数</b><br><li>当用户至少上传文件个数达到此设置时，才能拥有访问权限！<li>不限制设置为0">
<td class=td1><a href=# onclick="helpscript(VisitConfirm9);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>

<tr><th height="25" colspan="6" id=tabletitlelink align=left>  &nbsp;认证版块高级设置[<a href="#top">顶部</a>]</th></tr>
<tr><td height="25" colspan="6" class=td2>
<b>注</b>：当本版块设为认证版面时，以下设置才能生效。
</td></tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(62)"></td>
<td colspan=2 class=td1>
<U>用户进入需要金钱数</U></td>
<td colspan=2 class=td1>
<input type=text size=10 name="Board_Setting(62)" value="<%=Board_Setting(62)%>">
设置后进入该版面将需要支付一定量的金币
</td>
<input type="hidden" id="b62" value="<b>用户进入需要金钱数</b><br><li>设置后进入该版面将需要支付一定量的金币！<li>不限制设置为0">
<td class=td1><a href=# onclick="helpscript(b62);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(63)"></td>
<td colspan=2 class=td2>
<U>用户进入需要点券数</U></td>
<td colspan=2 class=td2>
<input type=text size=10 name="Board_Setting(63)" value="<%=Board_Setting(63)%>">
设置后进入该版面将需要支付一定量的点券
</td>
<input type="hidden" id="b63" value="<b>用户进入需要点券数</b><br><li>设置后进入该版面将需要支付一定量的点券！<li>不限制设置为0">
<td class=td2><a href=# onclick="helpscript(b63);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(66)"></td>
<td colspan=2 class=td1>
<U>VIP用户组进入收取金币点券标准</U><BR>请以小数设置，设置为0则不需要支持金币和点券，否则以上面两项设置的标准：VIP用户进入需要金币数或点券数 = VIP用户组收取金币点券标准 X 用户进入需要金币数或点券数</td>
<td colspan=2 class=td1>
<input type=text size=10 name="Board_Setting(66)" value="<%=Board_Setting(66)%>">
设置后进入该版面将需要支付一定量的点券
</td>
<input type="hidden" id="b66" value="<b>VIP用户进入需要金币数/点券数 = VIP用户组收取金币点券标准 X 用户进入需要金币数/点券数！<li>不限制设置为0">
<td class=td1><a href=# onclick="helpscript(b66);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(64)"></td>
<td colspan=2 class=td2>
<U>支付金币或点券进入版面的有效期</U></td>
<td colspan=2 class=td2>
<input type=text size=10 name="Board_Setting(64)" value="<%=Board_Setting(64)%>">
填写数字1－999，代表有效期为多少个月
</td>
<input type="hidden" id="b64" value="<b>支付金币或点券进入版面的有效期</b><br><li>填写数字1－12，代表有效期为多少个月！<li>不限制设置为0">
<td class=td2><a href=# onclick="helpscript(b64);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>

<tr><th height="25" colspan="6" id=tabletitlelink align=left>  &nbsp;<a name="setting3"></a>前台管理权限[<a href="#top">顶部</a>]</th></tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(33)"></td>
<td colspan=2 class=td1>
<U>主版主可以增删副版主</U></td>
<td colspan=2 class=td1>
<input type=radio class="radio" name="Board_Setting(33)" value=0 <%if Board_Setting(33)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio class="radio" name="Board_Setting(33)" value=1 <%if Board_Setting(33)="1" then%>checked<%end if%>>是&nbsp;
</td>
<td class=td1>&nbsp;</td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(34)"></td>
<td colspan=2 class=td2>
<U>主版主可以修改广告设置</U></td>
<td colspan=2 class=td2>
<input type=radio class="radio" name="Board_Setting(34)" value=0 <%if Board_Setting(34)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio class="radio" name="Board_Setting(34)" value=1 <%if Board_Setting(34)="1" then%>checked<%end if%>>是&nbsp;
</td>
<td class=td2>&nbsp;</td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(35)"></td>
<td colspan=2 class=td1>
<U>所有版主可以修改广告设置</U></td>
<td colspan=2 class=td1>
<input type=radio class="radio" name="Board_Setting(35)" value=0 <%if Board_Setting(35)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio class="radio" name="Board_Setting(35)" value=1 <%if Board_Setting(35)="1" then%>checked<%end if%>>是&nbsp;
</td>
<td class=td1>&nbsp;</td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(65)"></td>
<td colspan=2 class=td2>
<U>管理操作及评分理由选项</U><BR>每个理由用“|”分割</td>
<td colspan=2 class=td2>
<input type="text" Name="Board_Setting(65)" Value="<%=Board_Setting(65)%>" Size=50>
</td>
<td class=td2>&nbsp;</td>
</tr>
<tr> 
<th height="25" colspan="6" id=tabletitlelink align=left>  &nbsp;<a name="setting4"></a>发贴相关[<a href="#top">顶部</a>]</th>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(4)"></td>
<td colspan=2 class=td1>
<U>发贴是否采用验证码</U></td>
<td colspan=2 class=td1>
<input type=radio class="radio" name="Board_Setting(4)" value=0 <%if Board_Setting(4)="0" then%>checked<%end if%>>不采用&nbsp;
<input type=radio class="radio" name="Board_Setting(4)" value=1 <%if Board_Setting(4)="1" then%>checked<%end if%>>简单验证码&nbsp;
<input type=radio class="radio" name="Board_Setting(4)" value=2 <%if Board_Setting(4)="2" then%>checked<%end if%>>阵列验证码&nbsp;
</td>
<td class=td1>&nbsp;</td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(45)"></td>
<td colspan=2 class=td2>
<U>主题限制长度</U></td>
<td colspan=2 class=td2>
<input type=text size=10 name="Board_Setting(45)" value="<%=Board_Setting(45)%>"> Byte
</td>
<td class=td2>&nbsp;</td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(17)"></td>
<td colspan=2 class=td1>
<U>发贴后返回</U></td>
<td colspan=2 class=td1>
<input type=radio class="radio" name="Board_Setting(17)" value=1 <%if Board_Setting(17)="1" then%>checked<%end if%>>首页&nbsp;
<input type=radio class="radio" name="Board_Setting(17)" value=2 <%if Board_Setting(17)="2" then%>checked<%end if%>>论坛&nbsp;
<input type=radio class="radio" name="Board_Setting(17)" value=3 <%if Board_Setting(17)="3" then%>checked<%end if%>>帖子&nbsp;
<input type=radio class="radio" name="Board_Setting(17)" value=4 <%if Board_Setting(17)="4" then%>checked<%end if%>>快速返回到帖子&nbsp;
</td>
<td class=td1>&nbsp;</td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(16)"></td>
<td colspan=2 class=td2>
<U>帖子内容最大字节数</U><BR>1024字节等于1K</td>
<td colspan=2 class=td2>
<input type=text size=10 name="Board_Setting(16)" value="<%=Board_Setting(16)%>"> 字节
</td>
<td class=td2>&nbsp;</td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(52)"></td>
<td colspan=2 class=td1>
<U>帖子内容最小字节数</U><BR>1024字节等于1K</td>
<td colspan=2 class=td1>
<input type=text size=10 name="Board_Setting(52)" value="<%=Board_Setting(52)%>"> 字节
</td>
<td class=td1>&nbsp;</td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(53)"></td>
<td colspan=2 class=td2>
<U>投票后是否将投票贴提升到帖子列表顶部</U></td>
<td colspan=2 class=td2>
<input type=radio class="radio" name="Board_Setting(53)" value=0 <%if Board_Setting(53)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio class="radio" name="Board_Setting(53)" value=1 <%if Board_Setting(53)="1" then%>checked<%end if%>>是&nbsp;
</td>
<td class=td2>&nbsp;</td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(19)"></td>
<td colspan=2 class=td1>
<U>上传文件类型</U><BR>每种文件类型用“|”号分开</td>
<td colspan=2 class=td1>
<input type=text size=50 name="Board_Setting(19)" value="<%=Board_Setting(19)%>">
</td>
<td class=td1>&nbsp;</td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(30)"></td>
<td colspan=2 class=td2>
<U>是否起用防灌水机制</U></td>
<td colspan=2 class=td2>
<input type=radio class="radio" name="Board_Setting(30)" value=0 <%if Board_Setting(30)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio class="radio" name="Board_Setting(30)" value=1 <%if Board_Setting(30)="1" then%>checked<%end if%>>是&nbsp;
</td>
<td class=td2>&nbsp;</td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(31)"></td>
<td colspan=2 class=td1>
<U>每次发贴间隔</U></td>
<td colspan=2 class=td1>
<input type=text size=10 name="Board_Setting(31)" value="<%=Board_Setting(31)%>"> 秒
</td>
<td class=td1>&nbsp;</td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckBoardSetting(32)"></td>
<td colspan=2 class=td2>
<U>最多投票项目</U></td>
<td colspan=2 class=td2>
<input type=text size=10 name="Board_Setting%>>甯栧瓙&nbsp;
<input type=radio class="radio" name="Board_Setting(17)" value=4 <%if Board_Setting(17)="4" then%>checked<%end if%>>蹇