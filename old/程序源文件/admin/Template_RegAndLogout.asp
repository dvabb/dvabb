<!--#include file="../conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
Dim admin_flag,CssList,StyleConn
Head()
admin_flag=",23,"
CheckAdmin(admin_flag)
Select Case Request("action")
	Case "registerTemplate"
		RegisterTemplate()
	Case "logOutTemplate"
		LogOutTemplate()
	Case Else
		Main()
End Select
Footer()

Sub Main()
%>
<table border="0" cellspacing="1" cellpadding="5" align="center" width="100%">
	<tr>
		<th colspan="3" style="text-align:center;" id="TableTitleLink">模板注册和注销管理</th>
	</tr>
	<tr>
		<td class="forumHeaderBackgroundAlternate">
		注册模板说明：<br/>
		① 将需要注册的模板的目录上传到论坛根目录下的<font color="#FF0000">Resource</font>目录中（模板目录名称不能使用汉字）；<br/>
		② 填写下面“注册模板”的相关信息提交即可。<br/>
		③ 注意：<font color="#FF0000">模板目录只写模板的文件夹名称（如：Template_1），不需要Resource/前缀！</font><br/>
		<div id="formRegTemp">
		<form name="formRegTemplate" method="post" action="?action=registerTemplate" onsubmit="return checkform(this);" target="hiddenframe">
			模板名称：<input type="text" name="TemplateName" size="15"/>
			模板目录：<input type="text" name="TemplateFolder" size="15"/>
			<input type="submit" name="register" value="注册">
			<br/>注：①模板名称如果为空，则将以模板目录作为模板名称；
			<br/>&nbsp;&nbsp;&nbsp;&nbsp;②模板目录不能为空。
			<div id="ErrorInfo" style="display:none;"></div>
		</form>
		</div>
		</td>
	</tr>
	<tr>
		<td class="forumHeaderBackgroundAlternate">
		注销模板说明：<br/>
		①点击对应模板的“注销”按钮可以注销模板；<br/>
		②将已注销的模板的目录从论坛根目录下的<font color="#FF0000">Resource</font>目录中移除。<br/>
		<form name="formLogOutTemplate" method="post" action="?action=logOutTemplate" onsubmit="return logOutTemplate(0,'','');">
		<input type="hidden" name="id" value=""/>
		<input type="hidden" name="folder" value=""/>
		<table width="40%">
			<tr>
				<td>模板名称</td>
				<td>模板目录</td>
				<td>操作</td>
			</tr>
		<%
		Dim Rs,SQL
		Set Rs=Dvbbs.Execute("Select * From Dv_Templates")
		Do While Not Rs.Eof
			Response.write "<tr><td>"&Rs(1)&"</td><td>"&Rs(2)&"</td><td><input type=""button"" name=""LogOut"" value=""注销"" onclick=""logOutTemplate("&Rs(0)&",'"&Rs(1)&"','"&Rs(2)&"')""/></td></tr>"&chr(10)
			Rs.MoveNext
		Loop
		%>
		</table>
		</form>
		</td>
	</tr>
</table>
<iframe style="border:0px;width:0px;height:0px;" src="" name="hiddenframe" id="hiddenframe"></iframe>
<script languange="javascript">
	function checkform(theForm){
		if (''==theForm.TemplateFolder.value){
			document.getElementById('ErrorInfo').innerHTML='<img src="../skins/default/images/note_error.gif"/><font color="red">模板目录不能为空。</font>';
			document.getElementById('ErrorInfo').style.display='block';
			theForm.TemplateFolder.focus();
			return false;
		}
		if ('none'!=document.getElementById('ErrorInfo').style.display)
			document.getElementById('ErrorInfo').style.display='none';
		return true;
	}
	function logOutTemplate(id,templateName,templateFolder){
		if (0 != id){
			if (confirm('确定要注销模板“'+templateName+'”吗？')){
				document.formLogOutTemplate.id.value=id;
				document.formLogOutTemplate.folder.value=templateFolder;
				document.formLogOutTemplate.submit();
			}		
		}
		return false;
	}
</script>
<%
End Sub

Sub RegisterTemplate()
	Dim TemplateName,TemplateFolder
	Dim Rs
	TemplateName = Dvbbs.CheckStr(Request.Form("TemplateName"))
	TemplateFolder = Dvbbs.CheckStr(Request.Form("TemplateFolder"))
	If TemplateFolder="" Then
		Response.write "<script language=""javascript"">"
		Response.write "parent.document.getElementById('ErrorInfo').innerHTML='<img src=""../skins/default/images/note_error.gif""/><font color=""red"">模板目录不能为空。</font>';"
		Response.write "parent.document.getElementById('ErrorInfo').style.display='block';"
		Response.write "</script>"
		Response.end
	End If
	If TemplateName="" Then TemplateName=TemplateFolder
	Set Rs=Dvbbs.Execute("Select * From Dv_Templates Where Folder='"&TemplateFolder&"'")
	If Not Rs.Eof Then
		Response.write "<script language=""javascript"">"
		Response.write "parent.document.getElementById('ErrorInfo').innerHTML='<img src=""../skins/default/images/note_error.gif""/><font color=""red"">模板目录“"&TemplateFolder&"”已经注册过，不能重复注册。</font>';"
		Response.write "parent.document.getElementById('ErrorInfo').style.display='block';"
		Response.write "</script>"
		Response.end
	End If
	Rs.Close:Set Rs=Nothing
	Dvbbs.Execute("Insert Into Dv_Templates(Type,Folder) Values('"&TemplateName&"','"&TemplateFolder&"')")
	Response.write "<script language=""javascript"">"
	Response.write "parent.document.getElementById('formRegTemp').innerHTML='<img src=""../skins/default/images/note_ok.gif""/>"&TemplateName&" 模板注册成功。<a href=""Template_RegAndLogout.asp"">返回</a>继续注册模板';"
	Response.write "</script>"
	Dvbbs.Loadstyle()
End Sub

Sub LogOutTemplate()
	Dim ID,TemplateFolder,rs
	ID = Dvbbs.CheckNumeric(Request.Form("id"))
	TemplateFolder = Request.Form("folder")
	Dim count1
	Set Rs=Dvbbs.Execute("Select count(*) From Dv_Templates")
	count1=Rs(0)
	Rs.close
	Set Rs=Nothing
	If(count1>1)Then
		Dvbbs.Execute("Delete From Dv_Templates Where ID="&ID)
		Dvbbs.Loadstyle()
		Response.Write "<script language='javascript'>"
		Response.Write "alert('模板注销成功，请登录ftp把该模板对应的文件夹"&TemplateFolder&"删除。');"
		Response.Write "self.location='Template_RegAndLogout.asp';"
		Response.Write "</script>"
	Else
		Response.Write "<script language='javascript'>"
		Response.Write "alert('唯一的一份模板是不可以被注销的');"
		Response.Write "self.location='Template_RegAndLogout.asp';"
		Response.Write "</script>"
	End if
	Dvbbs.Loadstyle()
End Sub
%>