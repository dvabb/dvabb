<!--#include file="../Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
CheckAdmin(",")
%>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312" />
<TITLE>myspace</TITLE>
<style type="text/css">
body { margin:0px; background:transparent; overflow:hidden; background:url("skins/images/leftbg.gif"); }
.left_color { text-align:right; }
.left_color a { color: #083772; text-decoration: none; font-size:12px; display:block !important; display:inline; width:175px !important; width:180px; text-align:right; background:url("skins/images/menubg.gif") right no-repeat; height:23px; line-height:23px; padding-right:10px; margin-bottom:2px;}
.left_color a:hover { color: #7B2E00;  background:url("skins/images/menubg_hover.gif") right no-repeat; }
img { float:none; vertical-align:middle; }
#on { background:#fff url("skins/images/menubg_on.gif") right no-repeat; color:#f20; font-weight:bold; }
hr { width:90%; text-align:left; size:0; height:0px; border-top:1px solid #46A0C8;}
</style>
<script type="text/javascript">
<!--
	function disp(n){
		for (var i=0;i<10;i++)
		{
			if (!document.getElementById("left"+i)) return;			
			document.getElementById("left"+i).style.display="none";
		}
		document.getElementById("left"+n).style.display="";
	}
//-->
</script>
</head>
<BODY>


<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top" style="padding-top:10px;" class="left_color" id="menubar">
			<div id="left0" style="display:"> 
				<a href="setting.asp" target="frmright">论坛基本设置</a>
				<a href="board.asp?action=add" target="frmright">添加论坛版面</a>
				<a href="board.asp" target="frmright">论坛版面管理</a>
				<a href="user.asp" target="frmright">用户资料管理</a>
				<a href="group.asp" target="frmright">用户组(等级)管理</a>
				<a href="template.asp" target="frmright">风格模板管理</a>
				<a href="admin.asp" target="frmright">管理员管理</a>
				<a href="data.asp?action=BackupData" target="frmright">论坛数据备份</a>
	
				<!--<a href="plus_adsali.asp" target="frmright">站长营销(阿里妈妈广告联盟)</a>-->
				<a href="forumads.asp" target="frmright">论坛广告管理</a>
				<a href="ReloadForumCache.asp" target="frmright">更新论坛缓存</a>
				<a href="log.asp" target="frmright">论坛系统日志</a>
				<a href="Badlanguage.asp" target="frmright">论坛关键词设置</a>
				<a href="comparefileonlie.asp" target="frmright">论坛文件校检</a>
			</div>

			<div id="left1" style="display:none"> 
				<a href=setting.asp target="frmright">论坛基本设置</a>
				<a href=forumads.asp target="frmright">论坛广告设置</a>
				<!--<a href="plus_adsali.asp" target="frmright"><font color="red">站长营销(阿里妈妈广告联盟)</font></a>-->
				<a href="../announcements.asp?boardid=0&action=AddAnn" target="_blank">论坛公告管理</a>
				<a href=link.asp?action=add target="frmright">友情论坛添加</a>
				<a href=link.asp target="frmright">友情论坛管理</a>
				<a href="ForumPay.asp" target="frmright">论坛交易管理</a>
				<a href="ForumNewsSetting.asp" target="frmright">论坛首页调用</a>
				<a href="badword.asp?reaction=badword" target="frmright">脏话过滤设置</a>
				<a href="badword.asp?reaction=splitreg" target="frmright">注册过滤字符</a>
				<a href="lockip.asp?action=add" target="frmright">IP来访限定添加</a>
				<a href="lockip.asp" target="frmright">IP来访限定管理</a>
				<a href="Badlanguage.asp" target="frmright">论坛关键词设置</a>
				<a href="comparefileonlie.asp" target="frmright">论坛文件校检</a>
			</div>

			<div id="left2" style="display:none"> 
				<a href="board.asp?action=add" target="frmright" alt="">版面(分类)添加</a>
				<a href="board.asp" target="frmright" alt="">版面(分类)管理</a>
				<a href="board.asp?action=permission" target="frmright" alt="">分版面用户权限设置</a>
				<a href="boardunite.asp" target="frmright" alt="">合并版面数据</a>
				<a href="update.asp" target="frmright" alt="">重计论坛数据和修复</a>
			</div>

			<div id="left3" style="display:none">
				<a href="user.asp" target="frmright"  alt="">用户资料(权限)管理</a>
				<a href="group.asp" target="frmright" alt="">用户组(等级)管理</a>
				<a href="wealth.asp" target="frmright" alt="">用户积分设置</a>
				<a href="message.asp" target="frmright" alt="">用户短信管理</a>
				<a href="update.asp?action=updateuser" target="frmright" alt="">重计用户各项数据</a>
				<a href="SendEmail.asp" target="frmright" alt="">用户邮件群发管理</a>
				<a href="admin.asp?action=add" target="frmright" alt="">管理员添加</a>
				<a href="admin.asp" target="frmright" alt="">管理员管理</a>
				<a href="user.asp?action=audituser" target="frmright" alt="">批量设置审核</a>
				
			</div>

			<div id="left4" style="display:none"> 
				<a href="template.asp" target="frmright" alt="">风格界面模板总管理</a>
				<a href="Template_RegAndLogout.asp" target="frmright" alt="">模板注册与注销</a>
				<a href="Label.asp" target="frmright" alt="">自定义标签管理</a>
			</div>

			<div id="left5" style="display:none"> 
				<a href="alldel.asp" target="frmright" alt="">批量删除</a>
				<a href="alldel.asp?action=moveinfo" target="frmright" alt="">批量移动</a>
				<a href="../recycle.asp" target="frmright" alt="">回收站管理</a>
				<a href="postdata.asp?action=Nowused" target="frmright" alt="">当前帖子数据表管理 </a>
				<a href="postdata.asp" target="frmright" alt="">数据表间帖子转换 </a>
			</div>

			<div id="left6" style="display:none"> 
				<a href="data.asp?action=CompressData" target="frmright" alt="">压缩数据库</a>
				<a href="data.asp?action=BackupData" target="frmright" alt="">备份数据库</a>
				<a href="data.asp?action=RestoreData" target="frmright" alt="">恢复数据库</a>
				<a href="address.asp?action=add" target="frmright" alt="">论坛IP库添加</a>
				<a href="address.asp" target="frmright" alt="">论坛IP库管理 </a>
			</div>

			<div id="left7" style="display:none"> 
				<a href="upUserface.asp" target="frmright" alt="">上传头像管理</a>
				<a href="uploadlist.asp" target="frmright" alt="">上传文件管理</a>
				<a href="bbsface.asp?Stype=3" target="frmright" alt="">注册头像管理</a>
				<a href="bbsface.asp?Stype=2" target="frmright" alt="">发贴心情管理</a>
				<a href="bbsface.asp?Stype=1" target="frmright" alt="">发贴表情管理</a>
			</div>

			<div id="left8" style="display:none"> 
				<a href="plus.asp" target="frmright" alt="">论坛菜单管理</a>				
				
				<a href="../bokeadmin.asp" target="frmright" alt="">论坛博客管理</a>
				
				<a href="plus_Tools_Info.asp?action=Setting" target="frmright" alt="">道具中心设置</a>
				<a href="plus_Tools_Info.asp?action=List" target="frmright" alt="">道具资料设置</a>
				<a href="plus_Tools_User.asp" target="frmright" alt="">用户道具管理</a>
				<a href="plus_Tools_User.asp?action=paylist" target="frmright" alt="">交易信息管理</a>
				<a href="MoneyLog.asp" target="frmright" alt="">道具中心日志</a>
				<a href="plus_Tools_Magicface.asp" target="frmright" alt="">魔法表情设置</a>
				<a href="plus_cnzz_wss.asp" target='frmright'>流量统计</a>
				<a href="plus_ccvideo.asp" target="frmright" alt="">CC视频插件</a>
				<a href="plus_qcomic.asp" target='frmright'>组图参数设置</a>				
			</div>

			<div id="left9" style="display:none"> 
				<a href="data.asp?action=SpaceSize" target="frmright" alt="">系统信息检测</a>
				<a href="log.asp" target="frmright" alt="">论坛系统日志</a>
				<a href="help.asp" target="frmright" alt="">论坛帮助管理</a>
				<a href="ReloadForumCache.asp" target="frmright" alt="">更新论坛缓存</a>
			</div>
	</td>
 </tr>
</table>
</body>
</html>