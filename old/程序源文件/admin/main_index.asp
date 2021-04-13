<!--#include file="../Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
Head()
CheckAdmin(",")
Page_Main()
Footer()

Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Dvbbs.iCreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

Function CheckObj(objid)
	If Not IsObjInstalled(objid) Then
		CheckObj = "<font color="""&dvbbs.mainsetting(1)&""">×</font>"
	Else
		CheckObj = "√"
	End If
End Function

Sub Page_Main()
	Dim theInstalledObjects(20)
    theInstalledObjects(0) = "MSWC.AdRotator"
    theInstalledObjects(1) = "MSWC.BrowserType"
    theInstalledObjects(2) = "MSWC.NextLink"
    theInstalledObjects(3) = "MSWC.Tools"
    theInstalledObjects(4) = "MSWC.Status"
    theInstalledObjects(5) = "MSWC.Counters"
    theInstalledObjects(6) = "IISSample.ContentRotator"
    theInstalledObjects(7) = "IISSample.PageCounter"
    theInstalledObjects(8) = "MSWC.PermissionChecker"
    theInstalledObjects(9) = "Scripting.FileSystemObject"
    theInstalledObjects(10) = "adodb.connection"
    
    theInstalledObjects(11) = "SoftArtisans.FileUp"
    theInstalledObjects(12) = "SoftArtisans.FileManager"
    theInstalledObjects(13) = "JMail.SMTPMail"	'Jamil 4.2
    theInstalledObjects(14) = "CDONTS.NewMail"
    theInstalledObjects(15) = "Persits.MailSender"
    theInstalledObjects(16) = "LyfUpload.UploadFile"
    theInstalledObjects(17) = "Persits.Upload.1"
	theInstalledObjects(18) = "JMail.Message"	'Jamil 4.3
	theInstalledObjects(19) = "Persits.Upload"
	theInstalledObjects(20) = "SoftArtisans.FileUp"
	
	Dim Rs
	Dim Isaudituser
	Set Rs=Dvbbs.Execute("select count(*) from [dv_user] where usergroupid=5")
	Isaudituser=rs(0)
	if isnull(isaudituser) then isaudituser=0

	Dim BoardListNum
	set rs=dvbbs.execute("select count(*) from dv_board")
	BoardListNum=rs(0)
	If isnull(BoardListNum) then BoardListNum=0

%>
<table cellpadding="3" cellspacing="1" border="0" width="100%" align=center>
<tr><td colspan=2 height=25 class="td_title">论坛信息统计</td></tr>
<tr><td height=23 colspan=2>

系统信息：论坛帖子数 <B><%=Dvbbs.CacheData(8,0)%></B> 主题数 <B><%=Dvbbs.CacheData(7,0)%></B> 用户数 <B><%=Dvbbs.CacheData(10,0)%></B> 待审核用户数 <B><%=Isaudituser%></B> 版面总数 <B><%=BoardListNum%></B>

</td></tr>
<tr><td  class="forumRowHighlight" height=23 colspan=2>
本论坛由动网先锋（Dvbbs.Net）授权给 <%=Dvbbs.Forum_info(0)%> 使用，当前使用版本为 动网论坛
<%
If IsSqlDatabase=1 Then
	Response.Write "SQL数据库"
Else
	Response.Write "Access数据库"
End If
Response.Write " Dvbbs " & Dvbbs.Forum_Version
%>
</td></tr>
<tr>
<td width="50%"  height=23>服务器类型：<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
<td width="50%" class="forumRow">脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
</tr>
<tr>
<td width="40%" height=23>
FSO文本读写：<b><%=CheckObj(theInstalledObjects(9))%></b>

</td>
<td width="60%" class="forumRow">
Jmail4.3邮箱组件支持：
<b><%=CheckObj(theInstalledObjects(18))%></b>
&nbsp;&nbsp;<a href="data.asp?action=SpaceSize">>>查看更详细服务器信息检测</a>
</td>
</tr>
<tr><td height=23 colspan=2>
<a href="challenge.asp"><font color=red>网络支付相关设置</font></a>：主要用于论坛点券充值等相关导致论坛收益的服务，当前支付帐号为<%=Dvbbs.Forum_ChanSetting(4)%>
</td></tr>
<tr><td height=23 colspan=2>
<a href="data.asp"><font color=red>数据定期备份</font></a>：请注意做好定期数据备份，数据的定期备份可最大限度的保障您论坛数据的安全
</td></tr>
</table><br />


<%
Dim Forum_Pack
Set Rs=Dvbbs.Execute("Select Forum_Pack From Dv_Setup")
Forum_Pack = Rs(0)
Forum_Pack=Split(Forum_Pack,"|||")
If UBound(Forum_Pack)<2 Then ReDim Forum_Pack(3)
If Forum_Pack(0) = "1" Then
%>
<table cellpadding="3" cellspacing="1" border="0" align=center width="100%">
<tr><th colspan="2"><a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=143" target="SiteAdmin">论坛管理交流</a> | <a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=134" target="SiteAdmin">论坛新手区</a> | <a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=8" target="SiteAdmin">论坛技术区</a> | <a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=13" target="SiteAdmin">论坛插件区</a> | <a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=102" target="SiteAdmin">论坛风格区</a></td><tr>
<tr><td height=23 valign=top>
<iframe src="http://bbs.dvbbs.net/union_post.asp" name="SiteAdmin" height=180 width="100%" MARGINWIDTH=0 MARGINHEIGHT=0 HSPACE=0 VSPACE=0 FRAMEBORDER=0 SCROLLING=no></iframe>
</tr>
</table>
<%
End If
Rs.Close
Set Rs=Nothing
%>

<br />
<table cellpadding="3" cellspacing="1" border="0" align=center>
<tr><td class="td_title" colspan=2 height=25>论坛管理小贴士</td></tr>
<tr><td height=23 align="center" width="15%">
<img src="skins/images/user_tag.gif">
</td><td height=23 width="*"><B>用户组权限</B><br />动网论坛将注册用户分成不同的用户组，每个用户组可以拥有不同的论坛操作权限，并且在动网论坛7.0版本之后，用户等级结合到了用户组中，假如用户等级没有自定义权限，那么这个等级的权限就使用他所属的用户组权限，反之则拥有这个等级自己的权限。<font color=red>每个等级或者用户组所设定的权限都是是针对整个论坛的</font></td></tr>
<tr><td height=23 align="center">
<img src="skins/images/tag.gif">
</td><td height=23 width="*"><B>分版面权限</B><br />
每个用户组或有自定义权限设置的等级，都可以设置其在论坛中各个版面拥有不同的权限，比如说您可以设置注册用户或者新手上路在版面A不能发贴可以浏览等等权限设置，极大的扩充了论坛权限的设置，<font color=blue>从理论上来说可以分出很多个不同功能类型的论坛</font>
</td></tr>
<tr><td height=23 align="center">
<img src="skins/images/user_sys.gif">
</td><td height=23 width="*"><B>用户权限设定</B><br />
每个用户都可以设置其在论坛中各个版面拥有不同的权限或者特殊的权限，比如说您可以设置用户A在版面A中拥有所有管理权限。<U>对于上述三种权限需要注意的是其优先顺序为：用户权限设置(<font color=gray>自定义</font>)<font color=blue> <B>></B> </font>分版面权限设定(<font color=gray>自定义</font>)<font color=blue> <B>></B> </font>用户组权限设定(<font color=gray>默认</font>)</U>
</td></tr>
<tr><td height=23 align="center">
<img src="skins/images/skins.gif">
</td><td height=23 width="*"><B>对风格模板的管理</B><br />
其中包含对论坛所有模板的管理，模板中论坛的基本CSS设置，论坛主风格的更改，论坛分页面风格的更改，图片的设置，语言包的设置，新建模板页面和模板，模板中新建不同的语言、图片、风格等模板元素等等功能，并且拥有模板的导入导出功能，从真正意义上实现了论坛风格的在线编辑和切换
</td></tr>
<tr><td height=23 valign=top  align="right">
<B>一句话贴士</B>
</td><td height=23 width="*"><B>一句话贴士</B><br />
① 对于不同功能模块的页面，要仔细看页面中的说明，以免误操作
<BR>
② 用户组及其扩展的权限设置，对论坛的各种设置有极大扩充性，要充分明白其优先和有效顺序
<BR>
③ 添加论坛大分类的时候，别忘了回头看看该版面高级设置是否正确
<BR>
④ 有问题请到动网论坛官方站点提问，有很多热心的朋友会帮忙，<a href="help.asp">查看更多贴士请点击</a>
</td></tr>
</table><br />

<table width="100%" border="0" cellspacing="1" cellpadding="0">
    <form name="form1" method="post" action="">
  <tr>
    <td colspan="2" valign="middle" class="td_title">论坛管理帮助</td>
  </tr>
  <tr>
    <td align="right" width="15%">产品开发：</td>
    <td><a href="#01">海口动网先锋网络科技有限公司</a>  中国国家版权局著作权登记号2004SR00001 </td>
  </tr>
  <tr>
    <td align="right">产品负责：</td>
    <td>网站事业部 动网论坛项目组  企业典型案例 </td>
  </tr>
  <tr>
    <td align="right">联系方式：</td>
    <td>网站事业部：0898-68557467<br />
主机事业部：0898-68592224 68592294<br />
传　　　真：0898-68556467<br />
<a href="http://www.dvbbs.net/services_contect.asp" target="_blank">点击查看详细联系方法</a> </td>
  </tr>
  <tr>
    <td align="right">插件开发：</td>
    <td>
      动网论坛插件组织（Dvbbs Plus Organization） </td>
  </tr>
    </form>
</table>
<%
End Sub
%>