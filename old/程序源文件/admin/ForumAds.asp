<!--#include file =../conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="../inc/dv_clsother.asp" -->
<%
Call Head()
Dim admin_flag
admin_flag=",2,"
CheckAdmin(admin_flag)
If request("action")="save" Then 
	Call saveconst()
Else
	Call consted()
End If
If founderr then call dvbbs_error()
footer()

Sub consted()
dim sel
%>
<table width="100%" border="0" cellspacing="0" cellpadding="3" align="center">
<tr> 
<th colspan="2" style="text-align:center;"><b>论坛广告设置</b>（如为设置分论坛，就是分论坛首页广告，下属页面为帖子显示页面）</th>
</tr>
<tr> 
<td width="100%" class="td2" colspan=2><B>说明</B>：<BR>1、复选框中选择的为当前的使用设置模板，点击可查看该模板设置，点击别的模板直接查看该模板并修改设置。您可以将您下面的设置保存在多个论坛版面中<BR>2、您也可以将下面设定的信息保存并应用到具体的分论坛版面设置中，可多选<BR>3、如果您想在一个版面引用别的版面的配置，只要点击该版面名称，保存的时候选择要保存到的版面名称名称即可。
<hr size=1 width="100%" color=blue>
</td>
</tr>
<FORM METHOD=POST ACTION="">
<tr> 
<td width="100%" class="td2" colspan=2>
查看分版面广告设置，请选择左边下拉框相应版面&nbsp;&nbsp;
<select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
<option value="">查看分版面广告请选择</option>
<%
Dim ii,rs
set rs=Dvbbs.Execute("select boardid,boardtype,depth from dv_board order by rootid,orders")
do while not rs.eof
Response.Write "<option "
if rs(0)=dvbbs.boardid then
Response.Write " selected"
end if
Response.Write " value=""forumads.asp?boardid="&rs(0)&""">"
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
</FORM>
</table><BR>
<form method="POST" action="forumads.asp?action=save" name="advform">
<table width="100%" border="0" cellspacing="0" cellpadding="3" align="center">
<tr> 
<td width="100%" class="td2" colspan=2>
<input type=checkbox class=checkbox name="getskinid" value="1" <%if request("getskinid")="1" or request("boardid")="" then Response.Write "checked"%>><a href="forumads.asp?getskinid=1">论坛默认广告</a><BR> 点击此处返回论坛默认广告设置，默认广告设置包含所有<FONT COLOR="blue">除</FONT>包含具体版面内容（如帖子列表、帖子显示、版面精华、版面发贴等）<FONT COLOR="blue">以外</FONT>的页面。<hr size=1 width="90%" color=blue>
</td>
</tr>
<tr> 
<td width="200px" class="td1" valign=top>
版面广告保存选项<BR>
请按 CTRL 键多选<BR>
<select name="getboard" size="28" style="width:200px" multiple>
<%
set rs=Dvbbs.Execute("select boardid,boardtype,depth from dv_board order by rootid,orders")
do while not rs.eof
Response.Write "<option "
if rs(0)=dvbbs.boardid then
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
<td class="td1" valign=top>
<table>
<tr>
<td width="200" class="td1"><B>首页顶部广告代码</B></td>
<td width="*" class="td1"> 
<textarea name="Forum_ads_0" cols="50" rows="3"><%=server.htmlencode(Dvbbs.Forum_ads(0))%></textarea>
</td>
</tr>
<tr> 
<td width="200" class="td1"><B>首页尾部广告代码</B></td>
<td width="*" class="td1"> 
<textarea name="Forum_ads_1" cols="50" rows="3"><%=server.htmlencode(Dvbbs.Forum_ads(1))%></textarea>
</td>
</tr>
<tr> 
<td width="200" class="td1"><B>开启首页浮动广告</B></td>
<td width="*" class="td1"> 
<input type=radio class="radio" name="Forum_ads_2" value=0 <%if Dvbbs.Forum_ads(2)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio class="radio" name="Forum_ads_2" value=1 <%if Dvbbs.Forum_ads(2)="1" then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="200" class="td1"><B>论坛首页浮动广告图片地址</B></td>
<td width="*" class="td1"> 
<input type="text" name="Forum_ads_3" size="35" value="<%=Dvbbs.Forum_ads(3)%>">
</td>
</tr>
<tr> 
<td width="200" class="td1"><B>论坛首页浮动广告连接地址</B></td>
<td width="*" class="td1"> 
<input type="text" name="Forum_ads_4" size="35" value="<%=Dvbbs.Forum_ads(4)%>">
</td>
</tr>
<tr> 
<td width="200" class="td1"><B>论坛首页浮动广告图片宽度</B></td>
<td width="*" class="td1"> 
<input type="text" name="Forum_ads_5" size="3" value="<%=Dvbbs.Forum_ads(5)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td width="200" class="td1"><B>论坛首页浮动广告图片高度</B></td>
<td width="*" class="td1"> 
<input type="text" name="Forum_ads_6" size="3" value="<%=Dvbbs.Forum_ads(6)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td width="200" class="td1"><B>开启首页右下固定广告</B></td>
<td width="*" class="td1"> 
<input type=radio class="radio" name="Forum_ads_13" value=0 <%if Dvbbs.Forum_ads(13)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio class="radio" name="Forum_ads_13" value=1 <%if Dvbbs.Forum_ads(13)="1" then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="200" class="td1"><B>论坛首页右下固定广告图片地址</B></td>
<td width="*" class="td1"> 
<input type="text" name="Forum_ads_8" size="35" value="<%=Dvbbs.Forum_ads(8)%>">
</td>
</tr>
<tr> 
<td width="200" class="td1"><B>论坛首页右下固定广告连接地址</B></td>
<td width="*" class="td1"> 
<input type="text" name="Forum_ads_9" size="35" value="<%=Dvbbs.Forum_ads(9)%>">
</td>
</tr>
<tr> 
<td width="200" class="td1"><B>论坛首页右下固定广告图片宽度</B></td>
<td width="*" class="td1"> 
<input type="text" name="Forum_ads_10" size="3" value="<%=Dvbbs.Forum_ads(10)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td width="200" class="td1"><B>论坛首页右下固定广告图片高度</B></td>
<td width="*" class="td1"> 
<input type="text" name="Forum_ads_11" size="3" value="<%=Dvbbs.Forum_ads(11)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td width="200" class="td1"><B>是否开启帖间随机广告</B></td>
<td width="*" class="td1"> 
<input type=radio class="radio" name="Forum_ads_7" value=0 <%if Dvbbs.Forum_ads(7)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio class="radio" name="Forum_ads_7" value=1 <%if Dvbbs.Forum_ads(7)="1" then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="*" class="td1" valign="top" colspan=2><B>论坛帖间随机广告代码</B> <br>支持HTML语法和JS代码，每条随机用<font color="red">"#####"</font>(5个#号)分开。</td>
</tr>
<tr>
<%
Dim Ads_14
If UBound(Dvbbs.Forum_ads)>13 Then
	Ads_14=Dvbbs.Forum_ads(14)
End If
%>
<td width="*" class="td1" colspan=2> 
<textarea name="Forum_ads_14" style="width:100%" rows="10"><%=Ads_14%></textarea>
</td>
</tr>
<tr> 
<td width="200" class="td1"><B>是否开启页面文字广告位</B></td>
<td width="*" class="td1"> 
<input type=radio class="radio" name="Forum_ads_12" value=0 <%if Dvbbs.Forum_ads(12)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio class="radio" name="Forum_ads_12" value=1 <%if Dvbbs.Forum_ads(12)="1" then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr>
<%
Dim Ads_15
If UBound(Dvbbs.Forum_ads)>14 Then
	Ads_15=Dvbbs.Forum_ads(15)
End If
%>
<td width="200" class="td1"><B>页面文字广告位设置(版面)</B><BR>请确认已打开了页面文字广告位功能<BR></td>
<td width="*" class="td1"> 
<input type=radio class="radio" name="Forum_ads_15" value=0 <%if Ads_15="0" then%>checked<%end if%>>帖子列表&nbsp;
<input type=radio class="radio" name="Forum_ads_15" value=1 <%if Ads_15="1" then%>checked<%end if%>>帖子内容&nbsp;
<input type=radio class="radio" name="Forum_ads_15" value=2 <%if Ads_15="2" then%>checked<%end if%>>两者都显示&nbsp;
<input type=radio class="radio" name="Forum_ads_15" value=3 <%if Ads_15="3" then%>checked<%end if%>>两者都不显示&nbsp;
</td>
</tr>

<%
Dim Ads_17
If UBound(Dvbbs.Forum_ads)>16 Then
	Ads_17=Dvbbs.Forum_ads(17)
Else
	Ads_17 = 0
End If
%>
<tr>
<td width="200" class="td1"><B>文字广告每行广告个数</B></td>
<td width="*" class="td1"> 
<input type="text" name="Forum_ads_17" size="3" value="<%=Ads_17%>">&nbsp;个
</td>
</tr>
<tr> 
<td width="*" class="td1" valign="top" colspan=2>支持HTML语法和JS代码，每条用<font color="red">"#####"</font>(5个#号)分开。</td>
</tr>
<tr>
<td width="*" class="td1" colspan=2>
<%
Dim Ads_16
If UBound(Dvbbs.Forum_ads)>15 Then
	Ads_16=Dvbbs.Forum_ads(16)
End If
%>
<textarea name="Forum_ads_16" style="width:100%" rows="10"><%=Ads_16%></textarea>
</td>
</tr>



<%
Dim Ads_18
If UBound(Dvbbs.Forum_ads)>17 Then
	Ads_18=Dvbbs.Forum_ads(18)
Else
	Ads_18 = 0
End If
%>
<tr>
<td width="200" class="td1"><B>帖子顶楼顶部广告位</B></td>
<td width="*" class="td1">
<input type=radio class="radio" name="Forum_ads_18" value="0"/>关闭
<input type=radio class="radio" name="Forum_ads_18" value="1"/>开启
</td>
</tr>
<tr> 
<td width="*" class="td1" valign="top" colspan=2>支持HTML语法和JS代码，每条用<font color="red">"#####"</font>(5个#号)分开。</td>
</tr>
<tr>
<td width="*" class="td1" colspan=2>
<%
Dim Ads_19
If UBound(Dvbbs.Forum_ads)>18 Then
	Ads_19=Dvbbs.Forum_ads(19)
End If
%>
<textarea name="Forum_ads_19" style="width:100%" rows="10"><%=Ads_19%></textarea>
</td>
</tr>

<%
Dim Ads_20
If UBound(Dvbbs.Forum_ads)>19 Then
	Ads_20=Dvbbs.Forum_ads(20)
Else
	Ads_20 = 0
End If
%>
<tr>
<td width="200" class="td1"><B>帖子顶楼底部广告位</B></td>
<td width="*" class="td1">
<input type=radio class="radio" name="Forum_ads_20" value="0"/>关闭
<input type=radio class="radio" name="Forum_ads_20" value="1"/>开启
</td>
</tr>
<tr> 
<td width="*" class="td1" valign="top" colspan=2>支持HTML语法和JS代码，每条用<font color="red">"#####"</font>(5个#号)分开。</td>
</tr>
<tr>
<td width="*" class="td1" colspan=2>
<%
Dim Ads_21
If UBound(Dvbbs.Forum_ads)>20 Then
	Ads_21=Dvbbs.Forum_ads(21)
End If
%>
<textarea name="Forum_ads_21" style="width:100%" rows="10"><%=Ads_21%></textarea>
</td>
</tr>

<%
Dim Ads_22
If UBound(Dvbbs.Forum_ads)>21 Then
	Ads_22=Dvbbs.Forum_ads(22)
Else
	Ads_22 = 0
End If
%>
<tr>
<td width="200" class="td1"><B>帖子顶楼左右广告位</B></td>
<td width="*" class="td1">
<input type=radio class="radio" name="Forum_ads_22" value="0"/>关闭
<input type=radio class="radio" name="Forum_ads_22" value="1"/>左边
<input type=radio class="radio" name="Forum_ads_22" value="2"/>右边
</td>
</tr>
<tr> 
<td width="*" class="td1" valign="top" colspan=2>支持HTML语法和JS代码，每条用<font color="red">"#####"</font>(5个#号)分开。</td>
</tr>
<tr>
<td width="*" class="td1" colspan=2>
<%
Dim Ads_23
If UBound(Dvbbs.Forum_ads)>22 Then
	Ads_23=Dvbbs.Forum_ads(23)
End If
%>
<textarea name="Forum_ads_23" style="width:100%" rows="10"><%=Ads_23%></textarea>
</td>
</tr>

<tr> 
<td width="200" class="td1">&nbsp;</td>
<td width="*" class="td1"> 
<div align="center"> 
<input type="submit" name="Submit" value="提 交" class="button">
</div>
</td>
</tr>
</table>
</td>
</tr>
</table>
</form>
<script language="JavaScript">
<!--
chkradio(document.advform.Forum_ads_18,<%=Ads_18%>);
chkradio(document.advform.Forum_ads_20,<%=Ads_20%>);
chkradio(document.advform.Forum_ads_22,<%=Ads_22%>);
//-->
</script>
<%
end sub

Sub SaveConst()
	Dim iSetting,i,Sql,did
	For i = 0 To 30
		If Trim(Request.Form("Forum_ads_"&i))="" Then
			If i = 1 Or i = 0 Then
				iSetting = ""
			ElseIf i = 17 Then
				iSetting = 1
			Else
				iSetting = 0
			End If
		Else
			iSetting=Replace(Trim(Request.Form("Forum_ads_"&i)),"$","")
		End If

		If i = 0 Then
			Dvbbs.Forum_ads = iSetting
		Else
			Dvbbs.Forum_ads = Dvbbs.Forum_ads & "$" & iSetting
		End If
	Next
	For i = 1 To Request("getboard").Count
		If isNumeric(Request("getboard")(i)) Then
			If did = "" Then
				did = Request("getboard")(i)
			Else
				did = did & "," & Request("getboard")(i)
			End If
		End If
	Next
	If Request("getskinid")="1" Then
		Sql = "Update Dv_Setup Set Forum_ads='"&Replace(Dvbbs.Forum_ads,"'","''")&"'"
		Dvbbs.Execute(sql)
	End If
	If Request("getboard")<>"" Then
		Sql = "Update Dv_Board Set Board_Ads='"&Replace(Dvbbs.Forum_ads,"'","''")&"' Where BoardID In ("&did&")"
		Dvbbs.Execute(Sql)
	End If
	RestoreBoardCache()
	Dvbbs.loadSetup()
	Dv_suc("广告设置成功！")
End Sub
Sub RestoreBoardCache()
	Dim Board,node
	Dvbbs. LoadBoardList()
	For Each node in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style/@id")
		Application.Contents.Remove(Dvbbs.CacheName & "_showtextads_"&node.text)
		For Each board in Application(Dvbbs.CacheName&"_boardlist").documentElement.selectNodes("board/@boardid")
			Dvbbs.LoadBoardData board.text
			Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_"& board.text &"_"&node.text)
			Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_"& board.text &"_"&node.text&"_-time")
		Next
		Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_0_"& node.text)
		Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_0_"& node.text&"_-time")
	Next
End Sub
%>