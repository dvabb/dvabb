<!--#include file =../conn.asp-->
<!-- #include file="inc/const.asp" -->
<%	
Head()
Dim admin_flag,rs_c
admin_flag=",1,"
CheckAdmin(admin_flag)
If request("action")="save" Then
	Call savebadlanguage()
Else
	Call main()
end if
Footer()
Sub savebadlanguage()
    Dim content1
	content1=Replace(dvbbs.checkstr(Getabout("content")),Chr(10),"|")
	dvbbs.execute("update Dv_Badlanguage set content='"&content1&"'")
	Dv_suc("设置论坛脏话关键词成功!")
End Sub
Sub main()
%>
<table border="0" cellspacing="1" cellpadding="3"  align="center" width="100%">
<form method="POST" action="?action=save" name="theform">
<tr> 
<th width="100%" colspan="3" style="text-align:center;">论坛含有关键词不允许发布功能
</th></tr>
<tr> 
<td class="td2"><U>请输入关键词</U><br />以回车分隔</td>
<td class="td2"> 
<%
Dim rs,content
Set rs=dvbbs.execute("select top 1 content from Dv_Badlanguage")
If Not (rs.bof And Rs.eof) Then
    content=Replace(rs(0),"|",Chr(10)):Rs.close:Set Rs=nothing
Else
    content="":Rs.close:Set Rs=nothing
End If
%>
<TEXTAREA NAME="content" ROWS="30" COLS="100"><%=content%></TEXTAREA>
<td class="td2"><a href=# onclick="helpscript(forum_open);return false;" class="helplink"><img src="skins/images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class="td2"> &nbsp;</td>
<td class="td2">
<input type="submit" name="Submit" value="提 交" class="button">
</td>
<td class="td2">&nbsp;</td>
</tr>
</form></table>
<%
End Sub
Function Getabout(str)
    Dim ii
    For ii = 1 To Request.Form(str).Count
    Getabout = Getabout & Replace(Request.Form(str)(ii),"'","")
    Next
  End Function
%>