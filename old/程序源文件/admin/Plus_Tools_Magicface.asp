<!--#include file =../conn.asp-->
<!--#include file="inc/const.asp"-->
<!--#include file="../inc/dv_clsother.asp"-->
<!--#include file="../Dv_plus/Tools/plus_MagicFace_const.asp"-->
<%
Head()
Dim Admin_flag
Admin_flag=",43,"
CheckAdmin(admin_flag)
Select Case Request("Action")
	Case "Addnew" : AddNew()
	Case "EditMagic" : EditMagic()
	Case Else
		Main_head()
		MagicFaceList()
End Select
If founderr then dvbbs_error()
Footer()

Sub Main_head()

%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr><th style="text-align:center;">魔法表情（头像）设置和管理</th></tr>
<tr><td class="td2" style="line-height: 130%"><B>魔法表情（头像）设置和管理</B>：<BR>
	1、魔法表情（头像）默认的图片和Flash效果图路径分别是：<B>Dv_Plus/Tools/magicface/gif/</B>和<B>Dv_Plus/Tools/magicface/swf/</B>，在添加或管理图片和flash效果的时候最好将相关文件上传到此位置。
	<BR>
	2、您可以分别设定使用每个魔法表情所需要的金币和点券数，所设置的金币和点券数为用户购买需要。可设置使用每个魔法表情（头像）需要的帖子、金钱、积分、魅力、威望数限制，这些只是限制达到此标准的才能使用，并不扣除相应设置的数值。</td></tr>
</table>
<br>
<%

End Sub

Sub MagicFaceList()
Dim Rs,Sql,iMagicFaceType,i,ii,stype
Dim Page,MaxRows,Endpage,CountNum,PageSearch,SqlString
Endpage = 0
MaxRows = 50
Page = Request("Page")
If IsNumeric(Page) = 0 or Page="" Then Page=1
Page = Clng(Page)
stype = Request("stype")
If IsNumeric(stype) = 0 or stype="" Then stype=-1
stype = Clng(stype)
Response.Write "<script language=""JavaScript"" src=""../inc/Pagination.js""></script>"
iMagicFaceType = Split(MagicFaceType,"|")
PageSearch = "stype="&stype
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr>
<td class="td1" colspan=15 style="line-height: 130%">
<li>图片和Flash请参照上述说明放在默认目录，<font color=red>图片用数字序号填写，图片和flash文件只是在显示时后缀不同，在此为统一名称</font>，点击图片可预览效果
<li>修改魔法表情（头像）类别请打开Dv_Plus/Tools/plus_tools_const.asp文件修改其中MagicFaceType参数
<li><font color=red>金币1和点券1是购买魔法表情的价格，金币2和点券2是购买魔法头像的价格</font>
<li><font color=blue>添加魔法表情（头像）请预先准备三个文件并传到相应目录，两个gif图片分别是小和大图片，小图片用于用户购买选择处显示，大图片用于用户购买魔法头像后在帖子及其资料中显示，一个swf文件是魔法效果</font><BR>
<B>快速查看分类</B>：<a href="?">全部</a> | 
<%
For i = 0 To Ubound(iMagicFaceType)
	If i <> Ubound(iMagicFaceType) Then
		Response.Write "<a href=""?stype="&i&""">"&iMagicFaceType(i)&"</a> | "
	Else
		Response.Write "<a href=""?stype="&i&""">"&iMagicFaceType(i)&"</a>"
	End If
Next
%>
</td>
</tr>
<tr>
<th>ID</th>
<th>预览</th>
<th>说明</th>
<th>类别</th>
<th>图片</th>
<th>金币1</th>
<th>点券1</th>
<th>金币2</th>
<th>点券2</th>
<th>帖子</th>
<th>金钱</th>
<th>积分</th>
<th>魅力</th>
<th>威望</th>
<th>操作</th>
</tr>
<FORM METHOD=POST ACTION="?Action=Addnew">
<tr align=center>
<td class="td2"><font color=red>新</font></td>
<td class="td2">&nbsp;</td>
<td class="td2"><input type=text size=13 name="ntitle"></td>
<td class="td2">
<Select Name="ntype" size=1>
<%
For i = 0 To Ubound(iMagicFaceType)
	Response.Write "<option value="""&i&""">"&iMagicFaceType(i)&"</option>"
Next
%>
</Select>
</td>
<td class="td2"><input type=text size=6 name="ngif" value="0"></td>
<td class="td2"><input type=text size=3 name="nmoney" value="1000"></td>
<td class="td2"><input type=text size=3 name="nticket" value="100"></td>
<td class="td2"><input type=text size=3 name="ntmoney" value="100"></td>
<td class="td2"><input type=text size=3 name="ntticket" value="10"></td>
<td class="td2"><input type=text size=3 name="ntopic" value="10"></td>
<td class="td2"><input type=text size=3 name="nwealth" value="100"></td>
<td class="td2"><input type=text size=3 name="nuserep" value="20"></td>
<td class="td2"><input type=text size=3 name="nusercp" value="10"></td>
<td class="td2"><input type=text size=3 name="npower" value="0"></td>
<td class="td2"><input type=submit class="button" name=submit value="添加"></td>
</tr>
</FORM>
<%
'[Dv_Plus_Tools_MagicFace]
'ID,Title,MagicFace_s,MagicFace_l,iMoney,iTicket,MagicSetting
Dim MagicSetting
If stype = -1 Then
	Sql="Select ID,Title,MagicFace_s,MagicFace_s As MagicFace_l,MagicType,iMoney,iTicket,MagicSetting,tMoney,tTicket From Dv_Plus_Tools_MagicFace Order By ID Desc"
Else
	Sql="Select ID,Title,MagicFace_s,MagicFace_s As MagicFace_l,MagicType,iMoney,iTicket,MagicSetting,tMoney,tTicket From Dv_Plus_Tools_MagicFace Where MagicType = "&stype&" Order By ID Desc"
End If
Set Rs = Dvbbs.iCreateObject ("adodb.recordset")
If Cint(Dvbbs.Forum_Setting(92))=1 Then
	If Not IsObject(Plus_Conn) Then Plus_ConnectionDatabase
	Rs.Open Sql,Plus_Conn,1,1
Else
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open Sql,conn,1,1
End If
If Not (Rs.Eof And Rs.Bof) Then
	CountNum = Rs.RecordCount
	If CountNum Mod MaxRows=0 Then
		Endpage = CountNum \ MaxRows
	Else
		Endpage = CountNum \ MaxRows+1
	End If
	Rs.MoveFirst
	If Page > Endpage Then Page = Endpage
	If Page < 1 Then Page = 1
	If Page >1 Then 				
		Rs.Move (Page-1) * MaxRows
	End if
	SQL=Rs.GetRows(MaxRows)
	For i=0 To Ubound(SQL,2)
		MagicSetting = Split(SQL(7,i),"|")
%>
		<FORM METHOD=POST ACTION="?Action=EditMagic">
		<tr align=center>
		<td class="td2"><%=SQL(0,i)%></td>
		<td class="td2"><a href="../Dv_plus/Tools/magicface/swf/<%=SQL(3,i)%>.swf" target=_blank><img src="../Dv_plus/Tools/magicface/gif/<%=SQL(2,i)%>.gif" border=0></a></td>
		<td class="td2"><input type=text size=13 name="ntitle_<%=SQL(0,i)%>" value="<%=SQL(1,i)%>"></td>
		<td class="td2">
		<Select Name="ntype_<%=SQL(0,i)%>" size=1>
		<%
		For ii = 0 To Ubound(iMagicFaceType)
			Response.Write "<option value="""&ii&""""
			If ii = SQL(4,i) Then Response.Write " Selected "
			Response.Write ">"&iMagicFaceType(ii)&"</option>"
		Next
		%>
		</Select>
		</td>
		<td class="td2"><input type=text size=6 name="ngif_<%=SQL(0,i)%>" value="<%=SQL(2,i)%>"></td>
		<td class="td2"><input type=text size=3 name="nmoney_<%=SQL(0,i)%>" value="<%=SQL(5,i)%>"></td>
		<td class="td2"><input type=text size=3 name="nticket_<%=SQL(0,i)%>" value="<%=SQL(6,i)%>"></td>
		<td class="td2"><input type=text size=3 name="ntmoney_<%=SQL(0,i)%>" value="<%=SQL(8,i)%>"></td>
		<td class="td2"><input type=text size=3 name="ntticket_<%=SQL(0,i)%>" value="<%=SQL(9,i)%>"></td>
		<td class="td2"><input type=text size=3 name="ntopic_<%=SQL(0,i)%>" value="<%=MagicSetting(0)%>"></td>
		<td class="td2"><input type=text size=3 name="nwealth_<%=SQL(0,i)%>" value="<%=MagicSetting(1)%>"></td>
		<td class="td2"><input type=text size=3 name="nuserep_<%=SQL(0,i)%>" value="<%=MagicSetting(2)%>"></td>
		<td class="td2"><input type=text size=3 name="nusercp_<%=SQL(0,i)%>" value="<%=MagicSetting(3)%>"></td>
		<td class="td2"><input type=text size=3 name="npower_<%=SQL(0,i)%>" value="<%=MagicSetting(4)%>"></td>
		<td class="td2"><input type=checkbox class="checkbox" name="ID" value="<%=SQL(0,i)%>"></td>
		</tr>
<%
	Next
End If
Rs.Close
Set Rs=Nothing
%>
<tr>
<td class="td1" colspan=15 align=right height="30">
请选中指定的魔法表情进行修改或删除操作&nbsp;&nbsp;全选<input name=chkall type=checkbox class="checkbox" value=on	onclick="CheckAll(this.form)">&nbsp;&nbsp;<input type=submit class="button" name=submit value="修改">&nbsp;<input type=submit class="button" name=submit value="删除">
</td>
</tr>
</FORM>
</table>
<%
PageSearch=Replace(Replace(PageSearch,"\","\\"),"""","\""")
Response.Write "<SCRIPT>PageList("&Page&",3,"&MaxRows&","&CountNum&","""&PageSearch&""",1);</SCRIPT>"

End Sub

Sub Addnew()
	Dim ntitle,ntype,ngif,nswf,nmoney,nticket,ntmoney,ntticket,ntopic,nwealth,nuserep,nusercp,npower
	If Request("ntitle")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入魔法表情说明。"
		founderr=True
	End If
	ntitle = Dvbbs.CheckStr(Request("ntitle"))
	If Request("ntype")="" Or Not IsNumeric(Request("ntype")) Then
		Errmsg=ErrMsg + "<BR><li>请选择魔法表情类型。"
		founderr=True
	End If
	ntype = Request("ntype")
	If Request("ngif")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入魔法表情小图片。"
		founderr=True
	End If
	ngif = Dvbbs.CheckStr(Request("ngif"))
	If Request("nmoney")="" Or Not IsNumeric(Request("nmoney")) Then
		Errmsg=ErrMsg + "<BR><li>请输入魔法表情需要的金币数。"
		founderr=True
	End If
	nmoney = Request("nmoney")
	If Request("nticket")="" Or Not IsNumeric(Request("nticket")) Then
		Errmsg=ErrMsg + "<BR><li>请输入魔法表情需要的点券数。"
		founderr=True
	End If
	nticket = Request("nticket")
	If Request("ntmoney")="" Or Not IsNumeric(Request("ntmoney")) Then
		Errmsg=ErrMsg + "<BR><li>请输入魔法表情需要的金币数。"
		founderr=True
	End If
	ntmoney = Request("ntmoney")
	If Request("ntticket")="" Or Not IsNumeric(Request("ntticket")) Then
		Errmsg=ErrMsg + "<BR><li>请输入魔法表情需要的点券数。"
		founderr=True
	End If
	ntticket = Request("ntticket")
	If Request("ntopic")="" Or Not IsNumeric(Request("ntopic")) Then
		Errmsg=ErrMsg + "<BR><li>请输入魔法表情需要的帖子数。"
		founderr=True
	End If
	ntopic = Request("ntopic")
	If Request("nwealth")="" Or Not IsNumeric(Request("nwealth")) Then
		Errmsg=ErrMsg + "<BR><li>请输入魔法表情需要的金钱数。"
		founderr=True
	End If
	nwealth = Request("nwealth")
	If Request("nuserep")="" Or Not IsNumeric(Request("nuserep")) Then
		Errmsg=ErrMsg + "<BR><li>请输入魔法表情需要的积分数。"
		founderr=True
	End If
	nuserep = Request("nuserep")
	If Request("nusercp")="" Or Not IsNumeric(Request("nusercp")) Then
		Errmsg=ErrMsg + "<BR><li>请输入魔法表情需要的魅力数。"
		founderr=True
	End If
	nusercp = Request("nusercp")
	If Request("npower")="" Or Not IsNumeric(Request("npower")) Then
		Errmsg=ErrMsg + "<BR><li>请输入魔法表情需要的威望数。"
		founderr=True
	End If
	npower = Request("npower")
	npower = Request("ntopic") & "|" & Request("nwealth") & "|" & Request("nuserep") & "|" & Request("nusercp") & "|" & Request("npower")
	If Founderr Then Exit Sub
	Dvbbs.Plus_Execute("Insert Into Dv_Plus_Tools_MagicFace (Title,MagicFace_s,MagicType,iMoney,iTicket,MagicSetting,tMoney,tTicket) Values ('"&ntitle&"',"&ngif&","&ntype&","&nmoney&","&nticket&",'"&npower&"',"&ntmoney&","&ntticket&")")
	Dv_suc("添加魔法表情成功！")
End Sub

Sub EditMagic()
	Dim ID,FixID,i
	Dim ntype,nmoney,nticket,ntmoney,ntticket,ntopic,nwealth,nuserep,nusercp,npower,ngif
	ID = Replace(Request("ID"),"'","")
	ID = Replace(ID,";","")
	ID = Replace(ID,"--","")
	ID = Replace(ID," ","")
	FixID = Replace(ID,",","")
	FixID = Left(FixID,300)
	If ID = "" Or Not IsNumeric(FixID) Then
		Errmsg=ErrMsg + "<BR><li>请选中指定的魔法表情进行修改更新或删除操作。"
		founderr=True
	End If
	For I=1 To Request.Form("ID").Count
		ID = Replace(Request.Form("ID")(I),"'","")
		ID = CLng(ID)
		ntype = Request.Form("ntype_"&ID)
		If Not IsNumeric(ntype) Then ntype = 0
		nmoney = Request.Form("nmoney_"&ID)
		If Not IsNumeric(nmoney) Then nmoney = 0
		nticket = Request.Form("nticket_"&ID)
		If Not IsNumeric(nticket) Then nticket = 0
		ntmoney = Request.Form("ntmoney_"&ID)
		If Not IsNumeric(ntmoney) Then ntmoney = 0
		ntticket = Request.Form("ntticket_"&ID)
		If Not IsNumeric(ntticket) Then ntticket = 0
		ntopic = Request.Form("ntopic_"&ID)
		If Not IsNumeric(ntopic) Then ntopic = 0
		nwealth = Request.Form("nwealth_"&ID)
		If Not IsNumeric(nwealth) Then nwealth = 0
		nuserep = Request.Form("nuserep_"&ID)
		If Not IsNumeric(nuserep) Then nuserep = 0
		nusercp = Request.Form("nusercp_"&ID)
		If Not IsNumeric(nusercp) Then nusercp = 0
		npower = Request.Form("npower_"&ID)
		If Not IsNumeric(npower) Then npower = 0
		npower = ntopic & "|" & nwealth & "|" & nuserep & "|" & nusercp & "|" & npower
		ngif = Request.Form("ngif_"&ID)
		If Not IsNumeric(ngif) Then ngif = 0

		If Request("submit")="修改" Then
			Dvbbs.Plus_Execute("Update Dv_Plus_Tools_MagicFace Set Title='"&Dvbbs.CheckStr(Request.Form("ntitle_"&ID))&"',MagicFace_s="&ngif&",MagicType="&ntype&",iMoney="&nmoney&",iTicket="&nticket&",tMoney="&ntmoney&",tTicket="&ntticket&",MagicSetting='"&npower&"' Where ID = " & ID)
		Else
			Dvbbs.Plus_Execute("Delete From Dv_Plus_Tools_MagicFace Where ID = " & ID)
		End If
	Next
	Dv_suc("批量修改魔法表情成功！")
End Sub
%>