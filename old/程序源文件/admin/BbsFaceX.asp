<!--#include file="../conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
Head()
Dim admin_flag
admin_flag=",38,"
CheckAdmin(admin_flag)
If "save"=request("action") Then
	Dim sConfig,aTitle,aFrom,aTo,aWidth,aHeight,aNWidth,aNHeight,i,iDel
	aTitle=Split(request("title"),",")
	aFrom=Split(request("from"),",")
	aTo=Split(request("to"),",")
	aWidth=Split(request("width"),",")
	aHeight=Split(request("height"),",")
	aNWidth=Split(request("nwidth"),",")
	aNHeight=Split(request("nheight"),",")
	sConfig=""
	iDel=0
	For i=0 To UBound(aTitle)
		If "1"=Trim(request("isdel_"&i)) Then 
			iDel = iDel + 1
		ElseIf ""<>Trim(aTitle(i)) Then 
			If ""<>sConfig Then sConfig=sConfig&","
			sConfig=sConfig&"{t:"""&Trim(aTitle(i))&""",b:"&Dvbbs.CheckNumeric(Trim(aFrom(i)))&",e:"&Dvbbs.CheckNumeric(Trim(aTo(i)))&",w:"&Dvbbs.CheckNumeric(Trim(aWidth(i)))&",h:"&Dvbbs.CheckNumeric(Trim(aHeight(i)))&",nw:"&Dvbbs.CheckNumeric(Trim(aNWidth(i)))&",nh:"&Dvbbs.CheckNumeric(Trim(aNHeight(i)))&",p:'../images/emot/'}"&VBNewline
		End If 
	Next 
	If iDel<UBound(aTitle) Then 
		sConfig="var global_emot_config=["&sConfig&"];"&VBNewline
		If "1"=request("isdel_"&request("default_set")) Then 
			sConfig=sConfig&("var global_emot_default=0;")
		Else 
			sConfig=sConfig&("var global_emot_default="&CInt(request("default_set"))&";")
		End If 
		On Error Resume Next 
		DvStream.charset="gb2312"
		DvStream.Mode = 3
		DvStream.open()
		DvStream.WriteText(sConfig)
		DvStream.SaveToFile Server.MapPath("../images/emot/Config.js"),2
		DvStream.close()
		If Err Then
			Err.clear
			%>
			<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
			<tr> 
			<td height="23"><b><font color=red>更新配置文件失败！</font></b>可能是您的images/emot/目录没有写入和修改权限。您可以开启权限后再保存，或者复制下面的内容粘贴到images/emot/config.js，替换原来的所有内容。</td>
			</tr>
			<tr> 
			<td><textarea style="width:500px;height:200px;" onfocus="this.select()"><%=sConfig%></textarea></td>
			</tr>
			</table>
			<%
		End If 
	Else
		%>
		<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
		<tr> 
		<td height="23"><b><font color=red>不能全部删除，请至少保留一套表情。</font></b></td>
		</tr>
		</table>
		<%
	End If 
End If 
SetForm
Sub SetForm()
%>
<script language="javascript" src="../images/emot/config.js?rnd=<%=Now()%>"></script>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
<tr> 
<td height="23"><B>说明</B>：<br>①、图片统一存放于论坛Images/emot/目录。文件名为emX.gif，其中X表示两位或两位以上数字。不足两位前面补0。<br>②、此处调置为编辑器栏插入发贴表情时用<br />③、此处管理只对配置文件操作，不涉及删除图片操作。<br />④、添加非官方套图建议选择高序号段，比如从10000开始。以避免与官方套图序号冲突。
</td>
</tr>
</table>
<form name="form1" action="?action=save" method="post" style="margin:0px;padding:0px;">
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th colspan="2">套图设置（可以修改已有设置，也可以添加新套图）</th>
</tr>
<script language="javascript">
<!--
var a=global_emot_config,d=document;
for (var i=0; i<a.length; ++i){
	d.writeln('<tr><td colspan="2" height="5"></td></tr><tr> <td width="20%" align="right">标题：</td><td><input type="text" name="title" value="'+a[i]['t']+'" size="30" '+(global_emot_default==i?'style="font-weight:bold"':'')+' />&nbsp;<input type="checkbox" name="isdel_'+i+'" value="1" style="border:none" />删除&nbsp;<input type="radio" name="default_set" value="'+i+'" style="border:none;" '+(global_emot_default==i?'checked="checked"':'')+' />设为默认</td></tr>');
	d.writeln('<tr> <td width="20%" align="right">开始序号：</td><td><input type="text" name="from" value="'+a[i]['b']+'" size="10" /></td></tr>');
	d.writeln('<tr> <td width="20%" align="right">结束序号：</td><td><input type="text" name="to" value="'+a[i]['e']+'" size="10" /></td></tr>');
	d.writeln('<tr> <td width="20%" align="right">宽度：</td><td><input type="text" name="width" value="'+a[i]['w']+'" size="10" /></td></tr>');
	d.writeln('<tr> <td width="20%" align="right">高度：</td><td><input type="text" name="height" value="'+a[i]['h']+'" size="10" /></td></tr>');
	d.writeln('<tr> <td width="20%" align="right">一行显示个数：</td><td><input type="text" name="nwidth" value="'+a[i]['nw']+'" size="10" /></td></tr>');
	d.writeln('<tr> <td width="20%" align="right">显示多少行：</td><td><input type="text" name="nheight" value="'+a[i]['nh']+'" size="10" /></td></tr>');
}
//-->
</script>
<tr><td colspan="2" height="5"></td></tr>
<tr> 
<td width="20%" align="right">标题：</td>
<td><input type="text" name="title" value="" size="30" />&nbsp;*填写添加新套图，每项必填。不能含有英文逗号。</td>
</tr>
<tr> 
<td width="20%" align="right">开始序号：</td>
<td><input type="text" name="from" value="" size="10" /></td>
</tr>
<tr> 
<td width="20%" align="right">结束序号：</td>
<td><input type="text" name="to" value="" size="10" /></td>
</tr>
<tr> 
<td width="20%" align="right">宽度：</td>
<td><input type="text" name="width" value="" size="10" /></td>
</tr>
<tr> 
<td width="20%" align="right">高度：</td>
<td><input type="text" name="height" value="" size="10" /></td>
</tr>
<tr> 
<td width="20%" align="right">一行显示个数：</td>
<td><input type="text" name="nwidth" value="" size="10" /></td>
</tr>
<tr> 
<td width="20%" align="right">显示多少行：</td>
<td><input type="text" name="nheight" value="" size="10" /></td>
</tr>
<tr><td colspan="2" height="5"></td></tr>
<tr> 
<td width="20%" align="right"></td>
<td><input type="submit" name="sub1" value=" 提交保存 " /></td>
</tr>
</table>
</form>
<%
End Sub 
%>

