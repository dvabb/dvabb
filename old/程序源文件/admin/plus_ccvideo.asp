<!--#include File="../Conn.Asp"-->
<!-- #include File="Inc/Const.Asp" -->
<!-- #include File="../Inc/Md5.Asp" -->
<!-- #include File="../Inc/Chkinput.Asp" -->
<%
Dim Admin_flag,Forum_api,Ccvideo_api,Action,Ccvideo,Boardlist,Rs,Sql,ccvideoid,ccvideotype,Xmldoc,Ccid,Cclist,Iscclist,Ccvideobtn,Xmldom,Adslist
Admin_flag=",2,"
Checkadmin(Admin_flag)
ccvideotype="Dvbbs"

Chkforum_api()
Head()
If Request("T")="1" Then
	Cc_save()
Else
	Page_main()
End If
Footer()
Set Ccvideo = Nothing
Set Forum_api = Nothing
Set Ccvideo_api = Nothing
Set Adslist = Nothing

Sub Chkforum_api()
	Set Rs = Dvbbs.Execute("Select Top 1 Forum_apis From Dv_setup")
	Xmldoc = Rs(0)
	Rs.Close
	Set Rs = Nothing
	If Isnull(Xmldoc) Or Xmldoc = "" Then
		Creat_forum_api("Y")
	Else
		Set Forum_api = Server.Createobject("Msxml2.Freethreadeddomdocument"& Msxmlversion)
		Forum_api.Loadxml(Xmldoc)
		Set Ccvideo_api = Forum_api.Documentelement.Selectsinglenode("ccvideo")
		Testapi()
	End If
End Sub
Sub Testapi()
	On Error Resume Next
	Dim testcc
	testcc=Ccvideo_api.Getattribute("ccvideoid")
	If Err Then
		Creat_forum_api("")
	End If
End Sub
Sub Creat_forum_api(Str)
	If Str="Y" Then
		Set Forum_api = Server.Createobject("Msxml2.Freethreadeddomdocument"& Msxmlversion)
		Forum_api.Loadxml("<Forum_api/>")
	End If
	Set Ccvideo_api = Forum_api.Documentelement.Appendchild(Forum_api.Createnode(1,"ccvideo",""))
	Ccvideo_api.Setattribute "ccvideoid","97510"
	Ccvideo_api.Setattribute "ccvideobtn","Plugin_1"
	Ccvideo_api.Setattribute "ccvideotype",ccvideotype
	Ccvideo_api.Setattribute "boardlist",",0,"
	Update_forum_api()
End Sub
Sub Update_forum_api()
	Dvbbs.Execute("Update Dv_setup Set Forum_apis='"&Dvbbs.Checkstr(Forum_api.Xml)&"'")
End Sub
Sub Cc_save()
	Ccid=Dvbbs.Checkstr(Request("Ccid"))
	Ccvideobtn=Dvbbs.Checkstr(Request("Ccvideobtn"))
	Boardlist=Replace(Dvbbs.Checkstr(Request("Boardlist"))," ","")
	Ccvideo_api.Setattribute  "ccvideoid",Ccid
	Ccvideo_api.Setattribute  "ccvideobtn",Ccvideobtn
	Ccvideo_api.Setattribute  "ccvideotype",ccvideotype
	Ccvideo_api.Setattribute  "boardlist",","&Boardlist&","
	Update_forum_api()
	Dv_suc("操作成功！")
End Sub 
Sub Page_main()
ccvideoid=Ccvideo_api.Getattribute("ccvideoid")
Boardlist=Ccvideo_api.Getattribute("boardlist")
Ccvideobtn=Ccvideo_api.Getattribute("ccvideobtn")
%>
<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">
	<tr>
		<th colspan="2" style="text-align: center;">Cc视频插件说明</th>
	</tr>
	<tr>
		<td width="20%" class="td1" align="center">
		<button style="width: 80; height: 50; border: 1px outset;" class="button">
		注意事项</button></td>
		<td width="80%" class="td2">
		<li>开启CC视频功能,您需要先<a href="http://union.bokecc.com/signup.bo" target="_blank"><font color="red">注册</font></a>一个CC视频联盟帐号</li>
		<li>开户此功能后,用户可以上传视频,所上传的视频会保存在CC服务器</li>
		</td>
	</tr>
</table>
<br />
<table border="0" cellspacing="1" cellpadding="3" align="center" width="100%">
	<tr>
		<th colspan="3" style="text-align: center;">Cc视频插件设置</th>
	</tr>
	<form method="post" action="">
		<input type="hidden" name="t" value="1" />
		<tr>
			<td align="right" width="25%">您的Cc联盟ID：</td>
			<td colspan="2">
			<input type="text" name="CCID" size="30" value="<%=ccvideoid%>" id="ccvideouserid" />&nbsp;&nbsp;
			<font class="font1">请填写您的CC视频联盟数字ID,如果没有,请点击<a href="http://union.bokecc.com/signup.bo" target="_blank"><font color="red">这里</font></a>注册</font></td>
		</tr>
		<tr>
			<td align="right" width="25%">请选择按钮样式：</td>
			<td width="6%">
			<select id="ccvideobtn" onchange="showbtnpre()" name="ccvideobtn">
<script language="javascript">
<!--
var plgnamelist=[
	["暗黑夜空","plugin"],
	["黑金时代","plugin_2"],
	["幽深海洋","plugin_3"],
	["海天交接","plugin_4"],
	["青青苹果","plugin_5"],
	["粉红浪漫","plugin_6"],
	["银白月光","plugin_7"],
	["纯洁无邪","plugin_8"],
	["灰色天空","plugin_9"],
	["含羞脉脉","plugin_10"],
	["雾遮青水","plugin_11"],
	["桔子欲黄","plugin_12"],
	["绿草如茵","plugin_13"],
	["淡黄发光","plugin_14"],
	["金秋十月","plugin_15"],
	["蓝田生玉","plugin_16"]
];
for (var i in plgnamelist){
	document.writeln('<option value="'+plgnamelist[i][1]+'"');
	if ('<%=Ccvideobtn%>'==plgnamelist[i][1]) document.writeln(' selected');
	document.writeln('>'+plgnamelist[i][0]);
	document.writeln('</option>');
}
//-->
</script>
			</select> </td>
			<td id="ccvideobtnpre" width="67%">
			<object height="22" width="86">
				<param name="wmode" value="transparent" />
				<param name="allowScriptAccess" value="always" />
				<param name="movie" value="http://union.bokecc.com/flash/<%=Ccvideobtn%>.swf?userID=97510&amp;type=<%=ccvideotype%>" />
				<embed src="http://union.bokecc.com/flash/<%=Ccvideobtn%>.swf?userID=97510&amp;type=<%=ccvideotype%>" type="application/x-shockwave-flash" width="86" height="22" allowfullscreen="true" />
			</object>
			</td>
		</tr>
		<script language="javascript">
    function showbtnpre(){
    	btn=document.getElementById("ccvideobtn").value;
    	userid=document.getElementById("ccvideouserid").value
    	btnurl="http://union.bokecc.com/flash/"+btn+".swf?userID="+userid+"&type=<%=ccvideotype%>";
    	cchtml="<param name='wmode' value='transparent' /><param name='allowScriptAccess' value='always' /><param name='movie' value='"+btnurl+"' /><embed src='"+btnurl+"' type='application/x-shockwave-flash' width='86' height='22' allowFullscreen=true ></embed></object>";
    	document.getElementById("ccvideobtnpre").innerHTML=cchtml;
    }
   </script>
		<tr>
			<td align="right" width="25%">选择启用板块：<br />
			请按 Ctrl 或者 Shift 键多选<br />
			板块不能继承</td>
			<td colspan="2">
			<select name="boardlist" size="20" style="width: 270px" multiple>
			<option value="0" style="color: #FF0000; background-color: #FFFFCC;" <%if Instr(boardlist,",0,")>0 Then Response.Write " selected"%>="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
			停止使用该插件请选中此项</option>
			<%
		Dim ii
		set rs=Dvbbs.Execute("select boardid,boardtype,depth from dv_board order by rootid,orders")
		do while not rs.eof
			Response.Write "<option "
			if Instr(boardlist,","&rs(0)&",")>0 then
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
		%></select> </td>
		</tr>
		<tr>
			<td class="td2" colspan="3" align="center">
			<input type="submit" name="submit" value="确认提交" /> </td>
		</tr>
	</form>
</table>
<%
End sub
%>