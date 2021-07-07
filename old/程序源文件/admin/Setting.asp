<!--#include file =../conn.asp-->
<!-- #include file="inc/const.asp" -->
<%	
Head()
Dim admin_flag,rs_c
admin_flag=",1,"
CheckAdmin(admin_flag)
If request("action")="save" Then
	Call saveconst()
ElseIf request("action")="restore" Then
	Call restore()
Else
	Call consted()
end if
Footer()

Sub consted()
Dim  sel
%>
<iframe width="260" height="165" id="colourPalette" src="../images/post/nc_selcolor.htm" style="visibility:hidden; position: absolute; left: 0px; top: 0px;border:1px gray solid" frameborder="0" scrolling="no" ></iframe>
<table border="0" cellspacing="1" cellpadding="3"  align="center" width="100%">
<form method="POST" action="setting.asp?action=save" name="theform" onsubmit="return checkForm(this)">
<tr> 
<th width="100%" colspan="3" style="text-align:center;">论坛基本设置（目前只提供一种设置)
</th></tr>
<tr> 
<td width="100%" colspan=3>
<a href="#setting3">[基本信息]</a>&nbsp;<a href="#setting21">[论坛系统数据设置]</a>&nbsp;<a href="#setting6">[悄悄话选项]</a>&nbsp;<a href="#setting7">[论坛首页选项]</a>&nbsp;<a href="#setting8">[用户与注册选项]</a>&nbsp;<a href="#setting10">[系统设置]</a>&nbsp;<a href="#setting12">[在线和用户来源]</a>&nbsp;<a href="#setting_seo">[搜索引擎优化设置(SEO)]</a>
</td>
</tr>
<tr> 
<td width="100%" colspan="3">
<a href="#setting13">[邮件选项]</a>&nbsp;<a href="#setting14">[上传设置]</a>&nbsp;<a href="#setting15">[用户选项(签名、头衔、排行等)]</a>&nbsp;<a href="#setting16">[帖子选项]</a>&nbsp;<a href="#setting17">[防刷新机制]</a>&nbsp;<a href="#setting18">[论坛分页设置]</a>
</td>
</tr>
<tr> 
<td width="100%" colspan="3">
<a href="#setting20">[搜索选项]</a>&nbsp;<a href="#settingxu">[<font color=blue>官方插件设置</font>]</a>&nbsp;<a href="#admin">[<font color=red>安全设置</font>]</a>&nbsp;<a href="challenge.asp">[<font color=blue>RSS/手机短信/在线支付</font>]</a>
<a href="#SettingVIP">[VIP用户组设置]</a>
</td>
</tr>
<tr> 
<td width="93%" colspan="2">
如果您的论坛的设置搞乱了，可以使用<a href="?action=restore"><B>还原论坛默认设置</B></a>
</td>
<input type="hidden" id="forum_return" value="<b>还原论坛默认设置:</b><br><li>如果您把论坛设置搞乱了，可以点击还原论坛默认设置进行还原操作。<br><li>使用此操作将使您原来的设置无效而还原到论坛的默认设置，请确认您做了论坛备份或者记得还原后该做哪些针对您论坛所需要的设置">
<td><a href=# onclick="helpscript(forum_return);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td width="50%">
<U>论坛默认使用风格</U></td>
<td width="43%">
<%
	Dim forum_sid,iforum_setting,stopreadme,forum_pack,iCssName,iCssID,iStyleName
	Dim rs,Style_Option,css_Option,Forum_cid,TempOption,i
	set rs=dvbbs.execute("select forum_sid,forum_setting,forum_pack,forum_cid from dv_setup")
	Forum_sid=rs(0)
	Forum_pack=Split(rs(2),"|||")
	Iforum_setting=split(rs(1),"|||")
	Forum_cid=rs(3)
	Rs.close:Set Rs=Nothing
	stopreadme=iforum_setting(5)
%>
<Select Size=1 Name="sid">
<%
Dim Templateslist
For Each Templateslist in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style")
	Response.Write "<option value="""& Templateslist.selectSingleNode("@id").text &""""
	If Forum_sid = CLng(Templateslist.selectSingleNode("@id").text) Then Response.Write " selected "
	Response.Write ">"& Templateslist.selectSingleNode("@type").text &" </option>"
Next
%>
</select> 
</td>
<input type="hidden" id="forum_skin" value="<b>论坛默认使用风格:</b><br><li>在这里您可以选择您论坛的默认使用风格。<br><li>如果想改变论坛风格请到论坛风格模板管理中进行相关设置">
<td><a href=# onclick="helpscript(forum_skin);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class="td2"><U>论坛当前状态</U><br />维护期间可设置关闭论坛</td>
<td class="td2"> 
<input type=radio name="forum_setting(21)" value=0 <%if Dvbbs.forum_setting(21)="0" then%>checked<%end if%> class="radio">打开&nbsp;
<input type=radio name="forum_setting(21)" value=1 <%if Dvbbs.forum_setting(21)="1" then%>checked<%end if%> class="radio">关闭&nbsp;
</td>
<input type="hidden" id="forum_open" value="<b>论坛当前状态:</b><br><li>如果您需要做更改程序、更新数据或者转移站点等需要暂时关闭论坛的操作，可在此处选择关闭论坛。<br><li>关闭论坛后，可直接使用论坛地址＋login.asp登录论坛，然后使用论坛地址＋admin_login.asp登录后台管理进行打开论坛的操作">
<td class="td2"><a href=# onclick="helpscript(forum_open);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td><U>维护说明</U><br />在论坛关闭情况下显示，支持html语法</td>
<td> 
<textarea name="StopReadme" cols="50" rows="3" ID="TDStopReadme"><%=Stopreadme%></textarea><br><a href="javascript:admin_Size(-3,'TDStopReadme')"><img src="skins/images/minus.gif" unselectable="on" border='0'></a> <a href="javascript:admin_Size(3,'TDStopReadme')"><img src="skins/images/plus.gif" unselectable="on" border='0'></a>
</td>
<input type="hidden" id="forum_opens" value="<b>论坛维护说明:</b><br><li>如果您在论坛当前状态中关闭了论坛，请在此输入维护说明，他将显示在论坛的前台给会员浏览，告知论坛关闭的原因，在这里可以使用HTML语法。">
<td><a href=# onclick="helpscript(forum_opens);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class="td2">
<U>论坛定时设置</U></td>
<td class="td2"> 
<input type=radio name="forum_setting(69)" value="0" <%If Dvbbs.forum_setting(69)="0" Then %>checked <%End If%> class="radio">关 闭</option>
<input type=radio name="forum_setting(69)" value="1" <%If Dvbbs.forum_setting(69)="1" Then %>checked <%End If%> class="radio">定时关闭
<input type=radio name="forum_setting(69)" value="2" <%If Dvbbs.forum_setting(69)="2" Then %>checked <%End If%> class="radio">定时只读
</td>
<input type="hidden" id="forum_isopentime" value="<b>定时设置选择:</b><br><li>在这里您可以设置是否起用定时的各种功能，如果开启了本功能，请设置好下面选项中的论坛设置时间。<br><li>如果在非开放时间内需要更改本设置，可直接使用论坛地址＋login.asp登录论坛，然后使用论坛地址＋admin_login.asp登录后台管理进行打开论坛的操作">
<td class="td2"><a href=# onclick="helpscript(forum_isopentime);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td>
<U>定时设置</U><br />请根据需要选择开或关</td>
<td> 
<%
Dvbbs.forum_setting(70)=split(Dvbbs.forum_setting(70),"|")
If UBound(Dvbbs.forum_setting(70))<2 Then 
	Dvbbs.forum_setting(70)="1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1"
	Dvbbs.forum_setting(70)=split(Dvbbs.forum_setting(70),"|")
End If
For i= 0 to UBound(Dvbbs.forum_setting(70))
If i<10 Then Response.Write "&nbsp;"
%>
  <%=i%>点：<input type="checkbox" name="forum_setting(70)<%=i%>" value="1" <%If Dvbbs.forum_setting(70)(i)="1" Then %>checked<%End If%> class="checkbox">开
 <%
 If (i+1) mod 4 = 0 Then Response.Write "<br>"
 Next
 %>
</td>
<input type="hidden" id="forum_opentime" value="<b>论坛开放时间:</b><br><li>设置本选项请确认您打开了定时开放论坛功能。<br><li>本设置以小时为单位，请务必按规定正确填写<br><li>如果在非开放时间内需要更改本设置，可直接使用论坛地址＋login.asp登录论坛，然后使用论坛地址＋admin_login.asp登录后台管理进行打开论坛的操作">
<td><a href=# onclick="helpscript(forum_opentime);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
</table><a name="admin"></a><br />
<table border="0" cellspacing="1" cellpadding="3" align="center" width="100%">
<tr> 
<th width="100%" colspan="3" align="Left" id="tabletitlelink"><b>安全设置</b>[<a href="#top">顶部</a>]
</th></tr>
<tr> 
<td class="td1" width="50%">
<U>后台管理目录的设定</U><br>缺省目录为admin为安全起见，不让其他人知道目录，请修改</td>
<td class="td1" width="43%"> 
<input title="值不能为空 $!" type="text" name="Forum_AdminFolder" size="35" value="<%=Dvbbs.CacheData(33,0)%>"><br><br /><b>注意：</b>目录名称后面要有"/"，如"admin/"
<input type="hidden" id="AdminFolder" value="<b>后台管理目录的设定:</b><br><li>在FTP上修改您的论坛的管理,目录名称。(缺省目录为admin)<br><li>然后重新修改管理目录,管理员登录后台后就可以自动被引导到您设定的目录.<br><li>除管理员外,其他人无法知道管理的地址.">
</td>
<td class="td1"><a href=# onclick="helpscript(AdminFolder);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class="td2" width="50%">
<U>是否禁止代理服务器访问</U><br />禁止代理服务器访问能避免恶意的CC攻击，但开放后影响站点排名，建议在受到明显的攻击的时候开启</td>
<td class="td2" width="43%"> 
<input type="radio" name="forum_setting(100)" value="0" <%if Dvbbs.forum_setting(100)="0" then%>checked<%end if%> class="radio">否&nbsp;
<input type="radio" name="forum_setting(100)" value="1" <%if Dvbbs.forum_setting(100)="1" then%>checked<%end if%> class="radio">是&nbsp;
</td>
<input type="hidden" id="killcc" value="<b>是否禁止代理服务器访问:</b><br><li>禁止代理服务器访问能避免恶意的CC攻击，但开放后影响站点排名，建议在受到明显的攻击的时候开启，平时则关闭。">
<td class="td2"><a href=# onclick="helpscript(killcc);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class="td1" width="50%">
<U>限制同一IP连接数为</U><br />限制同一IP连接数，可以减少恶意的CC攻击的影响，但会造成用户访问不便，建议设置为0关闭此功能，在受到攻击的时候才开放</td>
<td class="td1" width="43%"> 
<% Dim IP_MAX_value
If UBound(Dvbbs.forum_setting) > 101 Then
	IP_MAX_value=Dvbbs.forum_setting(101)
Else
	IP_MAX_value=0
End If
%>
<input title="请输入整数 $!cint" type="text" name="forum_setting(101)" size="5" value="<%=IP_MAX_value%>">
</td>
<input type="hidden" id="IP_MAX" value="<b>限制同一IP连接数:</b><br><li>限制同一IP连接数，可以减少恶意的CC攻击的影响，但会造成用户访问不便，建议设置为0关闭此功能，在受到攻击的时候才开放，平时则关闭。">
<td class="td1"><a href=# onclick="helpscript(IP_MAX);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
</table><br />
<table border="0" cellspacing="1" cellpadding="3" align="center" width="100%">
<tr> 
<th width="100%" colspan="3">动网官方自动通讯设置
</th></tr>
<tr> 
<td class="td1" width="50%">
<U>是否起用动网官方自动通讯系统</U><br />开启后可在论坛后台收到动网官方最新通知以及直接参与官方讨论区讨论和发贴</td>
<td class="td1" width="43%"> 
<input type=radio name="forum_pack(0)" value=0 <%if cint(forum_pack(0))=0 then%>checked<%end if%> class="radio">否&nbsp;
<input type=radio name="forum_pack(0)" value=1 <%if cint(forum_pack(0))=1 then%>checked<%end if%> class="radio">是&nbsp;
</td>
<input type="hidden" id="forum_pack1" value="<b>是否起用动网自动更新通知系统:</b><br><li>开启后管理后台顶部会提示动网的最新程序、补丁、通知等。">
<td class="td1"><a href=# onclick="helpscript(forum_pack1);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class="td2">
<U>开启通讯系统用户名与密码</U><br />用户名与密码用符号“|||”分开<br />用户名和密码请到 <a href="http://bbs.dvbbs.net/Union_GetUserInfo.asp" target=_blank><font color=blue>动网官方</font></a> 获取</td>
<td class="td2">
<%
If UBound(forum_pack)<2 Then ReDim forum_pack(3)
%>
<input type=text size=21 name="forum_pack(1)" value="<%=forum_pack(1)%>|||<%=forum_pack(2)%>">
</td>
<input type="hidden" id="forum_pack2" value="<b>开启通知系统用户名与密码:</b><br><li>如要开启通知系统，请您先到动网官方论坛注册一个用户名并在动网官方通知系统里取得密码，并填写于此栏即可开启。">
<td class="td2"><a href=# onclick="helpscript(forum_pack2);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
</table><br />
<table border="0" cellspacing="1" cellpadding="3" align="center" width="100%">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting3"></a><b>论坛基本信息</b>[<a href="#top">顶部</a>]</th>
</tr>
<tr> 
<td width="50%" class="td1"> <U>论坛名称</U></td>
<td width="50%" class="td1">  
<input title="值不能为空 $!" name="Forum_info(0)" size="35" value="<%=Dvbbs.Forum_info(0)%>">
</td>
</tr>
<tr> 
<td width="50%" class="td2"> <U>论坛的访问地址</U></td>
<td width="50%" class="td2">  
<input title="值不能为空 $!" name="Forum_info(1)" size="35" value="<%=Dvbbs.Forum_info(1)%>">
</td>
</tr>
<tr> 
<td width="50%" class="td2"> <U>论坛的创建日期（格式：YYYY-M-D）</U></td>
<td width="50%" class="td2">  
<input type="text" name="forum_setting(74)" size="35" value="<%=Dvbbs.forum_setting(74)%>">
</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>论坛首页文件名</U></td>
<td width="50%" class="td1">  
<input title="值不能为空 $!" name="Forum_info(11)" size="35" value="<%=Dvbbs.Forum_info(11)%>">
</td>
</tr>
<tr> 
<td width="50%" class="td2"> <U>网站主页名称</U></td>
<td width="50%" class="td2">  
<input title="值不能为空 $!" name="Forum_info(2)" size="35" value="<%=Dvbbs.Forum_info(2)%>">
</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>网站主页访问地址</U></td>
<td width="50%" class="td1">  
<input title="值不能为空 $!" name="Forum_info(3)" size="35" value="<%=Dvbbs.Forum_info(3)%>">
</td>
</tr>
<tr> 
<td width="50%" class="td2"> <U>论坛管理员Email</U></td>
<td width="50%" class="td2">  
<input  name="Forum_info(5)" size="35" value="<%=Dvbbs.Forum_info(5)%>">
</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>联系我们的链接（不填写为Mailto管理员）</U></td>
<td width="50%" class="td1">  
<input type="text" name="Forum_info(7)" size="35" value="<%=Dvbbs.Forum_info(7)%>">
</td>
</tr>
<tr> 
<td width="50%" class="td2"> <U>论坛首页Logo图片地址</U><br />显示在论坛顶部左上角，可用相对路径或者绝对路径</td>
<td width="50%" class="td2">  
<input title="值不能为空 $!" type="text" name="Forum_info(6)" size="35" value="<%=Dvbbs.Forum_info(6)%>">
</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>论坛版权信息</U></td>
<td width="50%" class="td1" valign=top>  
<textarea name="Copyright" cols="50" rows="5" id=TdCopyright><%=Dvbbs.Forum_Copyright%></textarea>
<a href="javascript:admin_Size(-5,'TdCopyright')"><img src="skins/images/minus.gif" unselectable="on" border='0'></a> <a href="javascript:admin_Size(5,'TdCopyright')"><img src="skins/images/plus.gif" unselectable="on" border='0'></a>
</td>
</tr>
</table><br />
<table border="0" cellspacing="1" cellpadding="3" align="center" width="100%">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting21"></a><b>论坛系统数据设置</b>[<a href="#top">顶部</a>]--(以下信息不建议用户修改)</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>论坛会员总数</U></td>
<td width="50%" class="td1">  
<input title="请输入整数 $!cint" type="text" name="Forum_UserNum" size="25" value="<%=Dvbbs.CacheData(10,0)%>">
</td>
</tr>
<tr> 
<td width="50%" class="td2"> <U>论坛主题总数</U></td>
<td width="50%" class="td2">  
<input title="请输入整数 $!cint" type="text" name="Forum_TopicNum" size="25" value="<%=Dvbbs.CacheData(7,0)%>">
</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>论坛帖子总数</U></td>
<td width="50%" class="td1">  
<input title="请输入整数 $!cint" type="text" name="Forum_PostNum" size="25" value="<%=Dvbbs.CacheData(8,0)%>">
</td>
</tr>
<tr> 
<td width="50%" class="td2"> <U>论坛最高日发贴</U></td>
<td width="50%" class="td2">  
<input title="请输入整数 $!cint" type="text" name="Forum_MaxPostNum" size="25" value="<%=Dvbbs.CacheData(12,0)%>">
</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>论坛最高日发贴发生时间</U></td>
<td width="50%" class="td1">  
<input type="text" name="Forum_MaxPostDate" size="25" value="<%=Dvbbs.CacheData(13,0)%>">(格式：YYYY-M-D H:M:S)
</td>
</tr>
<tr> 
<td width="50%" class="td2"> <U>历史最高同时在线纪录人数</U></td>
<td width="50%" class="td2">  
<input title="请输入整数 $!cint" type="text" name="Forum_Maxonline" size="25" value="<%=Dvbbs.Maxonline%>">
</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>历史最高同时在线纪录发生时间</U></td>
<td width="50%" class="td1">  
<input type="text" name="Forum_MaxonlineDate" size="25" value="<%=Dvbbs.CacheData(6,0)%>">(格式：YYYY-M-D H:M:S)
</td>
</tr>
</table><br />

<table border="0" cellspacing="1" cellpadding="3" align="center" width="100%">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting6"></a><b>悄悄话选项</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>新短消息弹出窗口</U></td>
<td width="50%" class="td1">  
<input type=radio name="forum_setting(10)" value=0 <%if Dvbbs.forum_setting(10)="0" then%>checked<%end if%> class="radio">否&nbsp;
<input type=radio name="forum_setting(10)" value=1 <%if Dvbbs.forum_setting(10)="1" then%>checked<%end if%> class="radio">是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>发论坛短消息是否采用验证码</U><br />开启此项可以防止恶意短消息</td>
<td width="50%" class="td1">  
<input type=radio name="forum_setting(80)" value=0 <%if Dvbbs.forum_setting(80)="0" Then%>checked<%end if%> class="radio">否&nbsp;
<input type=radio name="forum_setting(80)" value=1 <%if Dvbbs.forum_setting(80)="1" Then%>checked<%end if%> class="radio">是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>多久检查一次是否有群发短信</U><br /></td>
<td width="50%" class="td1">  
<input type=text name="forum_setting(115)" size=8 value='<%if cint(Dvbbs.forum_setting(115))<1 then%>20<%else%><%=Dvbbs.forum_setting(115)%><%end if%>'> 分钟（最小为1，建议设置成20以上的整数）
</td>
</tr>
</table><br />

<table border="0" cellspacing="1" cellpadding="3" align="center" width="100%">
<tr> 
<th height=25 colspan=3 align=left id=tabletitlelink><a name="setting7"></a><b>论坛首页选项</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr>
<td width="50%" class="td1">
<U>首页显示论坛深度</U>
<input type="hidden" id="forum_depth" value="<b>首页显示论坛深度帮助:</b><br><li>0代表一级，1代表2级，以此类推；<li>设置过大的论坛深度将影响论坛整体性能，请根据自己论坛情况做设置，建议设置为1。">
</td>
<td width="43%" class="td1"> 
<input title="请输入整数 $!cint" type="text" size=10 name="forum_setting(5)" value="<%=Dvbbs.forum_setting(5)%>"> 级
</td>
<td class="td1"><a href=# onclick="helpscript(forum_depth);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class="td2"> <U>是否显示过生日会员</U>
<input type="hidden" id="forum_userbirthday" value="<b>首页显示过生日会员帮助:</b><br><li>凡当天有会员过生日则显示于论坛首页；<li>开启本功能较消耗资源。">
</td>
<td class="td2">  
<input type=radio name="forum_setting(29)" value=0 <%if Dvbbs.forum_setting(29)="0" then%>checked<%end if%> class="radio">否&nbsp;
<input type=radio name="forum_setting(29)" value=1 <%if Dvbbs.forum_setting(29)="1" then%>checked<%end if%> class="radio">是&nbsp;
</td>
<td class="td2"><a href=# onclick="helpscript(forum_userbirthday);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td width="50%" class="td1"><U>首页四格显示</U></td>
<td width="50%" class="td1">  
<input type=radio name="forum_setting(113)" value=0 <%if Dvbbs.forum_setting(113)="0" then%>checked<%end if%> class="radio">是&nbsp;
<input type=radio name="forum_setting(113)" value=1 <%if Dvbbs.forum_setting(113)="1" then%>checked<%end if%> class="radio">否&nbsp;
</td>
<td class="td2"><a href=# class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td width="50%" class="td1"><U>首页右侧信息显示</U></td>
<td width="50%" class="td1">  
<input type=radio name="forum_setting(114)" value=0 <%if Dvbbs.forum_setting(114)="0" then%>checked<%end if%> class="radio">是&nbsp;
<input type=radio name="forum_setting(114)" value=1 <%if Dvbbs.forum_setting(114)="1" then%>checked<%end if%> class="radio">否&nbsp;
</td>
<td class="td2"><a href=# class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
</table><br />

<table border="0" cellspacing="1" cellpadding="3" align="center" width="100%">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting8"></a><b>用户与注册选项</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>是否允许新用户注册</U><br />关闭后论坛将不能注册</td>
<td width="50%" class="td1">  
<input type=radio name="forum_setting(37)" value=0 <%if Dvbbs.forum_setting(37)="0" then%>checked<%end if%> class="radio">否&nbsp;
<input type=radio name="forum_setting(37)" value=1 <%if Dvbbs.forum_setting(37)="1" then%>checked<%end if%> class="radio">是&nbsp;
</td>
</tr>

<tr> 
<td width="50%" class="td2"> <U>注册是否采用验证码</U><br />开启此项可以防止恶意注册</td>
<td width="50%" class="td2">  
<input type=radio name="forum_setting(78)" value=0 <%if Dvbbs.forum_setting(78)="0" Then%>checked<%end if%> class="radio">否&nbsp;
<input type=radio name="forum_setting(78)" value=1 <%if Dvbbs.forum_setting(78)="1" Then%>checked<%end if%> class="radio">是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>注册是否采用问题验证</U><br />开启此项可以防止恶意注册<br />注意问题答案不要太BT哦，还想不想人注册？</td>
<td width="50%" class="td1">  
<input type=radio name="forum_setting(107)" value=0 <%if Dvbbs.forum_setting(107)="0" Then%>checked<%end if%> class="radio">否&nbsp;
<input type=radio name="forum_setting(107)" value=1 <%if Dvbbs.forum_setting(107)="1" Then%>checked<%end if%> class="radio">是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class="td2"> <U>注册验证问题：</U><br />可以设置多个验证问题，防止恶意注册<br />每个问题使用!(英文感叹号)分隔.<br /><b><font color="red">如：1+2=? ! 3*3=? ! 爱的英文单词是____ ?</font></b></td>
<td width="50%" class="td2"><textarea name="forum_setting(105)" rows="5" cols="60"><%=Dvbbs.forum_setting(105)%></textarea></td>
</tr>
<tr> 
<td width="50%" class="td2"> <U>注册验证答案：</U><br />设置回答上述问题的答案，防止恶意注册<br />每个答案使用!(英文感叹号)分隔，和上面的问题顺序对应! <br /><b><font color="red">如：3!9!love</font></b></td>
<td width="50%" class="td2"><textarea name="forum_setting(106)" rows="5" cols="60"><%=Dvbbs.forum_setting(106)%></textarea></td>
</tr>

<tr> 
<td width="50%" class="td1"> <U>登录是否采用验证码</U><br />开启此项可以防止恶意登录猜解密码</td>
<td width="50%" class="td1">  
<input type=radio name="forum_setting(79)" value=0 <%if Dvbbs.forum_setting(79)="0" Then%>checked<%end if%> class="radio">否&nbsp;
<input type=radio name="forum_setting(79)" value=1 <%if Dvbbs.forum_setting(79)="1" Then%>checked<%end if%> class="radio">是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class="td2"> <U>会员取回密码是否采用验证码</U><br />开启此项可以防止恶意登录猜解密码</td>
<td width="50%" class="td2">  
<input type=radio name="forum_setting(81)" value=0 <%if Dvbbs.forum_setting(81)="0" Then%>checked<%end if%> class="radio">否&nbsp;
<input type=radio name="forum_setting(81)" value=1 <%if Dvbbs.forum_setting(81)="1" Then%>checked<%end if%> class="radio">是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>会员取回密码次数限制</U><br />0则表示无限制，若取回问答错误超过此限制，则停止至24小时后才能再次使用取回密码功能。</td>
<td width="50%" class="td1">  
<input title="请输入整数 $!cint" type="text" name="forum_setting(84)" size="3" value="<%=Dvbbs.forum_setting(84)%>">
</td>
</tr>
<tr> 
<td width="50%" class="td2"> <U>最短用户名长度</U><br />填写数字，不能小于1大于50</td>
<td width="50%" class="td2">  
<input title="请输入整数 $!cint" type="text" name="forum_setting(40)" size="3" value="<%=Dvbbs.forum_setting(40)%>">
</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>最长用户名长度</U><br />填写数字，不能小于1大于50</td>
<td width="50%" class="td1">  
<input title="请输入整数 $!cint" type="text" name="forum_setting(41)" size="3" value="<%=Dvbbs.forum_setting(41)%>">
</td>
</tr>
<tr> 
<td width="50%" class="td2"> <U>同一IP注册间隔时间</U><br />如不想限制可填写0</td>
<td width="50%" class="td2">  
<input title="请输入整数 $!cint" type="text" name="forum_setting(22)" size="3" value="<%=Dvbbs.forum_setting(22)%>">&nbsp;秒
</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>Email通知密码</U><br />确认您的站点支持发送mail，所包含密码为系统随机生成</td>
<td width="50%" class="td1">  
<input type=radio name="forum_setting(23)" value=0 <%if Dvbbs.forum_setting(23)="0" then%>checked<%end if%> class="radio">关闭&nbsp;
<input type=radio name="forum_setting(23)" value=1 <%if Dvbbs.forum_setting(23)="1" then%>checked<%end if%> class="radio">打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class="td2"> <U>一个Email只能注册一个帐号</U></td>
<td width="50%" class="td2">  
<input type=radio name="forum_setting(24)" value=0 <%if Dvbbs.forum_setting(24)="0" then%>checked<%end if%> class="radio">关闭&nbsp;
<input type=radio name="forum_setting(24)" value=1 <%if Dvbbs.forum_setting(24)="1" then%>checked<%end if%> class="radio">打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>注册需要管理员认证</U></td>
<td width="50%" class="td1">  
<input type=radio name="forum_setting(25)" value=0 <%if Dvbbs.forum_setting(25)="0" then%>checked<%end if%> class="radio">关闭&nbsp;
<input type=radio name="forum_setting(25)" value=1 <%if Dvbbs.forum_setting(25)="1" then%>checked<%end if%> class="radio">打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class="td2"> <U>发送注册信息邮件</U><br />请确认您打开了邮件功能</td>
<td width="50%" class="td2">  
<input type=radio name="forum_setting(47)" value=0 <%if Dvbbs.forum_setting(47)="0" then%>checked<%end if%> class="radio">关闭&nbsp;
<input type=radio name="forum_setting(47)" value=1 <%if Dvbbs.forum_setting(47)="1" then%>checked<%end if%> class="radio">打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>开启短信欢迎新注册用户</U></td>
<td width="50%" class="td1">  
<input type=radio name="forum_setting(46)" value=0 <%if Dvbbs.forum_setting(46)="0" then%>checked<%end if%> class="radio">关闭&nbsp;
<input type=radio name="forum_setting(46)" value=1 <%if Dvbbs.forum_setting(46)="1" then%>checked<%end if%> class="radio">打开&nbsp;
</td>
</tr>

</table><br />
<table border="0" cellspacing="1" cellpadding="3" align="center" width="100%">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting10"></a><b>系统设置</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>论坛所在时区</U></td>
<td width="50%" class="td1">  
<input title="值不能为空 $!" type="text" name="Forum_info(9)" size="35" value="<%=Dvbbs.Forum_info(9)%>">
</td>
</tr>
<tr> 
<td width="50%" class="td2"> <U>服务器时差</U></td>
<td width="50%" class="td2">
<select name="forum_setting(0)">
<%for i=-23 to 23%>
<option value="<%=i%>" <%if i=CInt(Dvbbs.forum_setting(0)) then%>selected<%end if%>><%=i%></option>
<%next%>
</select>
</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>脚本超时时间</U><br />默认为300，一般不做更改</td>
<td width="50%" class="td1">  
<input title="请输入整数 $!cint" type="text" name="forum_setting(1)" size="3" value="<%=Dvbbs.forum_setting(1)%>">&nbsp;秒
</td>
</tr>
<tr> 
<td width="50%" class="td2"> <U>是否显示页面执行时间</U></td>
<td width="50%" class="td2">  
<input type=radio name="forum_setting(30)" value=0 <%If Dvbbs.forum_setting(30)="0" then%>checked<%end if%> class="radio">否&nbsp;
<input type=radio name="forum_setting(30)" value=1 <%if Dvbbs.forum_setting(30)="1" then%>checked<%end if%> class="radio">是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class="td1"><U>禁止的邮件地址</U><br />在下面指定的邮件地址将被禁止注册，每个邮件地址用“|”符号分隔<br />本功能支持模糊搜索，如设置了eway禁止，将禁止eway@aspsky.net或者eway@dvbbs.net类似这样的注册</td>
<td width="50%" class="td1"> 
<input title="值不能为空 $!" type="text" name="forum_setting(52)" size="50" value="<%=Dvbbs.forum_setting(52)%>">
</td>
</tr>
<tr> 
<td width="50%" class="td2"><U>论坛脚本过滤扩展设置</U><br />此设置为开启HTML解释的时候对脚本代码的识别设置，<br>您可以根据需要添加自定的过滤<br>格式是：过滤字| 如：abc|efg| 这样就添加了abc和efg的过滤</td>
<td width="50%" class="td2"> 
<Input title="值不能为空 $!" type="text" name="forum_setting(77)" size="50" value="<%=Dvbbs.forum_setting(77)%>"><br> 没有添加可以填0,如果添加了最后一个字符必须是"|"
</td>
</tr>
</table><br />
<table border="0" cellspacing="1" cellpadding="3" align="center" width="100%">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting12"></a><b>在线和用户来源</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>在线显示用户IP</U><br />关闭后如果所属用户组、论坛权限、用户权限中设置了用户可浏览则可见</td>
<td width="50%" class="td1">  
<input type=radio name="forum_setting(28)" value=0 <%if Dvbbs.forum_setting(28)="0" then%>checked<%end if%> class="radio">保密&nbsp;
<input type=radio name="forum_setting(28)" value=1 <%if Dvbbs.forum_setting(28)="1" then%>checked<%end if%> class="radio">公开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class="td2"> <U>在线显示用户来源</U><br />关闭后如果所属用户组、论坛权限、用户权限中设置了用户可浏览则可见<br />开启本功能较消耗资源</td>
<td width="50%" class="td2">  
<input type=radio name="forum_setting(36)" value=0 <%if Dvbbs.forum_setting(36)="0" then%>checked<%end if%> class="radio">保密&nbsp;
<input type=radio name="forum_setting(36)" value=1 <%if Dvbbs.forum_setting(36)="1" then%>checked<%end if%> class="radio">公开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>在线资料列表显示用户当前位置</U></td>
<td width="50%" class="td1">  
<input type=radio name="forum_setting(33)" value=0 <%if Dvbbs.forum_setting(33)="0" then%>checked<%end if%> class="radio">否&nbsp;
<input type=radio name="forum_setting(33)" value=1 <%if Dvbbs.forum_setting(33)="1" then%>checked<%end if%> class="radio">是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class="td2"> <U>在线资料列表显示用户登录和活动时间</U></td>
<td width="50%" class="td2">  
<input type=radio name="forum_setting(34)" value=0 <%if Dvbbs.forum_setting(34)="0" then%>checked<%end if%> class="radio">否&nbsp;
<input type=radio name="forum_setting(34)" value=1 <%if Dvbbs.forum_setting(34)="1" then%>checked<%end if%> class="radio">是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class="td1"> <U>在线资料列表显示用户.net绫讳技杩欐牱鐨勬敞鍐