<!--#include file="../conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
Head()
Dim admin_flag
admin_flag=",22,"
CheckAdmin(admin_flag)

Dim action,SkinID,StyleID
StyleID=Dvbbs.CheckNumeric(request("StyleID"))
action=Request("action")

If Application(Dvbbs.CacheName &"_style").documentElement.selectSingleNode("style[@id='"& StyleID &"']") Is Nothing Then
	If Not Application(Dvbbs.CacheName &"_style").documentElement.selectSingleNode("style/@id") Is Nothing Then
		StyleID=Application(Dvbbs.CacheName &"_style").documentElement.selectSingleNode("style/@id").text 
	Else
		Response.Write "模板数据无法提取,请检查或重新导入"
		Response.End
	End If
End If

Dim StyleFolder,FilePath,typeList
StyleFolder = Application(Dvbbs.CacheName &"_style").documentElement.selectSingleNode("style[@id='"& StyleID &"']/@folder").text
FilePath = "../Resource/"& StyleFolder &"/"
typeList = Split("0,strings,pic,html",",")
'Response.Write FilePath

Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""3"" align=""center"" width=""100%"">"
Response.Write "<tr>"
Response.Write "<th width=""100%"" style=""text-align:center;"" colspan=""2"">论坛模板管理"
Response.Write "</th>"
Response.Write "</tr>"
Response.Write "<tr>"
Response.Write "<td class=""td2"" colspan=""2"">"
Response.Write "<p><b>注意</b>：<br />①在这里，您可以新建和修改模板，可以编辑论坛语言包和风格，可以新建模板页面，操作时请按照相关页面提示完整填写表单信息。<br />②论坛当前正在使用的默认模板不能删除<br />③如果修改分模板页面名称或删除分模板页面请在关闭论坛之后操作，否则可能会影响论坛访问。"
Response.Write "</td>"
Response.Write "</tr>"
Response.Write "<tr>"
Response.Write "<td class=""td2"" width=""20%"" height=""25"" align=""left"">"
Response.Write "<b>论坛模板操作选项</b></td>"
Response.Write "<td class=""td2"" width=""80%""><a href=""template.asp"">模板管理首页</a>"
Response.Write " | <a href=""http://bbs.dvbbs.net/loadtemplates.asp"" target = ""_blank"" title="""">获取官方模板数据</a></td>"
Response.Write "</tr>"
Response.Write "</table>"

Select Case action
	Case "edit"
		Call Edit() 
	Case "manage"
		If Request("mostyle")="编 辑" Then
			Main()
		ElseIf Request("mostyle") = "删 除" Then
			'DelStyle()
		End If
	Case "saveedit"
		Call Saveedit()
	Case "rename"
		rename()
	Case "editmain"
		editmain()
	Case "savemain"
		Savemain() 
	Case Else
		Main()
End Select
%>

<%
Rem 首页面 模板页列表 2007-10-9
Sub Main()
	Response.Write "<p></p>"
	Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""3"" align=""center"" width=""100%"">"
	Response.Write "<tr>"
	Response.Write "<th width=""100%"" style=""text-align:center;"" colspan=""2"">当前论坛模板管理"
	Response.Write "</th>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<form method=post action=""?action=manage"">"
	Response.Write "<td class=""td2"" height=30 align=left>"
	Response.Write "请选择相关模板： "
	'利用系统缓存数据取得所有模板名称和ID
	Dim Templateslist,rs,i
	Response.Write "<select name=""StyleID"" size=""1"">"
	For Each Templateslist in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style")
		Response.Write "<option value="""& Templateslist.selectSingleNode("@id").text &""""
		If CLng(Templateslist.selectSingleNode("@id").text) = CLng(StyleID) Then 
			Response.Write " selected"
		End If 
		Response.Write ">"&Templateslist.selectSingleNode("@type").text &"</option>"
	Next 
	Response.Write "</select>"
	Response.Write "&nbsp;&nbsp;"
	Response.Write "<input type=submit class=""button"" value=""编 辑"" name=""mostyle"">&nbsp;&nbsp;&nbsp;"
	'Response.Write "<input type=submit class=""button"" value=""删 除"" name=""mostyle"">"
	Response.Write "<br /><br /><b>说明：</b>删除操作将删除该模板所有数据，慎用。"
	Response.Write "</td>"
	Response.Write "</form>"
	Response.Write "<form method=post action=""?action=rename"">"
	Response.Write "<td class=""td2"" align=left>"
	Response.Write "<select name=""StyleID"" size=1>"
	For Each Templateslist in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style")
		Response.Write "<option value="""& Templateslist.selectSingleNode("@id").text &""""
		If CLng(Templateslist.selectSingleNode("@id").text) = CLng(StyleID) Then 
			Response.Write " selected"
		End If 
		Response.Write ">"&Templateslist.selectSingleNode("@type").text &"</option>"
	Next 
	Response.Write "</select>"
	Response.Write "&nbsp;&nbsp;"
	Response.Write "改名为：<input type=text size=20 name=""StyleName"" value="""
	Response.Write """>&nbsp;&nbsp;"
	Response.Write "<input type=submit class=""button"" name=submit value=""修改"">"
	Response.Write "</td>"
	Response.Write "</form>"

	Response.Write "</tr>"

	Response.Write "<tr>"
	Response.Write "<th style=""text-align:center;"" colspan=2>"
	Response.Write Application(Dvbbs.CacheName &"_style").documentElement.selectSingleNode("style[@id='"& StyleID &"']/@type").text &"－－模板资源管理</th></tr><tr><td height=25 class=""bodytitle"" colspan=2>"
	Response.Write "通常来说，分页面模板就是论坛中每个页面的风格模板，括号中是字段名，字段的命名规则为：Page_页面名（不要后缀）"
	Response.Write "</td>"
	Response.Write "</tr>"

	Dim pageList,pageName
	pageList = "Main_Style,index,dispbbs,showerr,login,online,usermanager,fmanage,boardstat,admin,paper_even_toplist,query,show,dispuser,help_permission,postjob,post,boardhelp,indivgroup"
	pageName = Split(pageList,",")
	pageName(0) = "pub"
	For i=0 To UBound(pageName)
		Response.Write "<tr onmouseover=""this.style.backgroundColor='#B3CFEE';this.style.color='red'"" onmouseout=""this.style.backgroundColor='';this.style.color=''"">"
		Response.Write "<td height=""25"">"
		Response.Write "<li>"
		If i=0 Then
			Response.Write "分页面模板(Main_Style) &nbsp;&nbsp;</li></td><td height=""25"" align=""left"">"
		Else
			Response.Write "分页面模板(page_"&pageName(i)&") &nbsp;&nbsp;</li></td><td height=""25"" align=""left"">"
		End If
		Response.Write "编辑该模块："
		Response.Write "<a href=""?action=edit&stype=1&page="&pageName(i)&"&StyleID="& StyleID &""">语言包</a>"
		Response.Write " | <a href=""?action=edit&stype=2&page="&pageName(i)&"&StyleID="& StyleID &""">图片</a>"
		Response.Write " | <a href=""?action=edit&stype=3&page="&pageName(i)&"&StyleID="& StyleID &""">界面风格</a>"
		If i=0 Then
			Response.Write " | <a href=""?action=editmain&stype=2&StyleID="&StyleID&""">基本设置</a>"
		End if
		Response.Write "</td>"
		Response.Write "</tr>"
	Next

	Response.Write "</table><p></p>"
End Sub

Rem 修改模板名称 2007-10-8 By Dv.唧唧
Sub rename()
	Dim stylename
	stylename=Dvbbs.checkStr(Request("stylename"))
	If Trim(stylename)=""  Then 
		Errmsg=ErrMsg + "<br /><li>修改名称请输入新的模板名称。"
		Dvbbs_error()
	End If
	Dvbbs.Execute("update [Dv_Templates] set Type='"&StyleName&"' where id="&StyleID&"")
	Dv_suc("模板名修改成功!")
	Dvbbs.loadSetup()
	Dvbbs.Loadstyle()
End Sub

Sub editmain() '基本设置部分
	Dim stype,NowEditinfo
	Dim mystr
	stype=Request("stype")	
	Select Case stype
		Case "2"
			NowEditinfo="基本设置"
			mystr="mainsetting"
		Case Else
			Errmsg=ErrMsg + "<br /><li>您提交了错误的参数."
			Dvbbs_error()	
	End Select
	Dim TemplateStr
	Response.Write "<form action=""?action=savemain&stype="&stype&"&StyleID="&StyleID&""" method=post>"
	Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">"
	Response.Write "<tr>"
	Response.Write "<th width=""100%"" style=""text-align:center;"" colspan=2>"
	Response.Write NowEditinfo
	Response.Write "("&mystr&")设置</th></tr>"
	Select Case stype
		Case "2"		
			TemplateStr = Dvbbs.ReadTextFile(FilePath &"pub_html0.htm")
			TemplateStr = split(TemplateStr,"||")
			If ubound(TemplateStr) < 13 Then
				TemplateStr = Dvbbs.ReadTextFile(FilePath &"pub_html0.htm")&"||skins/Default/"
				TemplateStr = split(TemplateStr,"||")
			End if
			Response.Write "<tr><td class=""td2"" height=40 align=""center"" colspan=2>"
			Response.Write "<table cellspacing=""1"" cellpadding=""0"" border=""0"" align=""left"" width=""100%"">"
			Response.Write "<tr>"
			Response.Write "<td width=""300"" align=""right"" colspan=""1"" class=""td2"">表格宽度：</td>"
			Response.Write "<td align=""left"" class=""td2"" colspan=""3"" >"
			Response.Write "<input type=""text"" size=""5"" name=""TemplateStr"" value="""&TemplateStr(0)&""">&nbsp;(实际像素:如<b>780px</b> 一定要写单位(px),或者百分比 ,如<b>98%</b>)&nbsp;"&mystr&"(0)"
			Response.Write "</td>"
			Response.Write "</tr>"
			Dim j,vtitle
			vtitle="aa|警告提醒语句的颜色：|显示帖子的时候，相关帖子，转发帖子，回复等的颜色：|首页连接颜色：|一般用户名称字体颜色：|一般用户名称上的光晕颜色：|版主名称字体颜色：|版主名称上的光晕颜色：|管理员名称字体颜色：|管理员名称上的光晕颜色：|贵宾名称字体颜色：|贵宾名称上的光晕颜色：|表格边框颜色：|风格图片路径："
			vtitle=split(vtitle,"|")
			For j=1 to UBound(vtitle)
				'If j=4 or j=6 or j=8 or j=10 Then
					'Response.Write "<input type=""hidden"" size=""10"" value="""&TemplateStr(j)&""" name=""TemplateStr"">"
				'Else
					Response.Write "<tr>"
					Response.Write "<td colspan=""4"" style=""background:#6595D6;height:3px; dislpay:none"">"
					Response.Write "</tr>"
					Response.Write "<tr>"
					Response.Write "<td height=""25"" width=""300"" align=""right"">"&vtitle(j)&"</td>"
					Response.Write "<td width=""20"" style=""background:"&TemplateStr(j)&";""></td>"
					Response.Write "<td width=""180"" align=""left"">"
					Response.Write "<input type=""text"" size=""10"" value="""&TemplateStr(j)&""" name=""TemplateStr"">&nbsp;"&mystr&"("&j&")"
					Response.Write "</td>"
					Response.Write "</tr>"
				'End If
			Next
			Response.Write "</table>"		
			Response.Write "</td></tr>"		
		Case Else
				
	End Select
	Response.Write "<tr><td class=""td2"" height=""25"" align=""center"">"
	Response.Write "<input type=""reset"" class=""button"" name=""Submit"" value=""重 填"">"
	Response.Write "</td>"
	Response.Write "<td class=""td2"" height=""25"" align=""center"">"
	Response.Write "<input type=""submit"" class=""button"" name=""B1"" value=""修 改"">"
	Response.Write "</table>"
	Response.Write "</form>"
End Sub

Sub savemain() ' 保存基本设置
	Dim stype,NowEditinfo,TemplateStr,tempstr,Main_Style,FileName
	stype=Request("stype")
	TemplateStr=""
	Select Case stype
		Case "2"
			NowEditinfo="基本设置"
			For Each TempStr in Request.form("TemplateStr")
				If TempStr<>"" Then
					TemplateStr=TemplateStr&TempStr&"||"
				Else
					TemplateStr=TemplateStr&Chr(1)&"||"
				End If
			Next
			TemplateStr=Left(TemplateStr,Len(TemplateStr)-2)
			
			FileName = FilePath &"pub_html0.htm"
		Case Else
			Errmsg=ErrMsg + "<br /><li>您提交了错误的参数."
			Dvbbs_error()	
	End Select
	TemplateStr=Dvbbs.checkStr(TemplateStr)
	'Response.Write TemplateStr
	Dvbbs.writeToFile FileName,TemplateStr
	Dvbbs.Loadstyle()
	Dv_suc("主模板"&NowEditinfo&"修改成功!")
	If stype=2 Then
			'createsccfile()
	End If
End Sub

Sub Edit()
	Dim Page,mystr,rs,i
	Dim FileName,TempStr,TemplateStr,stype
	Dim TempStyleHelp,StyleHelpValue
	stype=Dvbbs.checkStr(request("stype"))
	page=Dvbbs.checkStr(request("page"))

	FileName = page &"_"& typeList(stype)

	If Not IsNumeric(stype) Then 
		Errmsg=ErrMsg + "<br /><li>错误的样式参数"
		Dvbbs_error()
	End If
	
	Response.Write "<form name=""template"" action=""?action=saveedit&page="&page&"&stype="&stype&"&StyleID="&StyleID&""" method=post>"
	Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""3"" align=""center"" width=""100%"">"
	Response.Write "<tr>"
	Response.Write "<th width=""100%"" style=""text-align:center;"" colspan=3>"
	'Response.Write Rs(1)
	Response.Write "分页面模板("
	Response.Write page
	Response.Write ")"
	Response.Write "<input Type=""hidden"" name=""dvbbs"" value=""OK!"">"
	Select Case stype
		Case 1
			Response.Write "语言包"
			mystr="template.Strings"
			If page="main_style" Then mystr="Dvbbs.lanstr"
		Case 2
			Response.Write "图片资源(当前默认路径{$PicUrl}为："&Dvbbs.Forum_PicUrl&")"
			mystr="template.pic"
			If page="main_style" Then mystr="Dvbbs.mainpic"
		Case 3
			Response.Write "界面风格"
			mystr="template.html"
			If page="main_style" Then mystr="Dvbbs.mainhtml"
	End Select
	
	Response.Write "管理</th></tr>"

	'If TemplateStr(Ubound(TemplateStr))="" Then TemplateStr(Ubound(TemplateStr))="del"
	'Response.Write CountFilesNumber(FilePath,FileName)
	For i=0 To CountFilesNumber(FilePath,FileName)-1
		TemplateStr = Dvbbs.ReadTextFile(FilePath & FileName &i&".htm")

		Response.Write "<tr><td class=""td2"" width=20% height=40 align=left>"
		Response.Write mystr&"("&i&")"
		Response.Write "<br /><a href=""javascript:;"" onclick=""rundvscript(t"&i&",'page="&page&"&index="&i&"&stype="&stype&"');"" title=""点这里获取这部分模板的官方数据"">获取官方数据</a>"
		Response.Write "</td>"		
		Response.Write "<td class=""td2"" width=80% height=25 align=left>"
		Select Case stype
			Case 1
				If LenB(TemplateStr)>70 Then
				Response.Write "<textarea name=""TemplateStr"" id=""t"&i&"""  cols=""100"" rows=""3"">"
				Response.Write server.htmlencode(TemplateStr)
				Response.Write "</textarea>"
				Else
				Response.Write "<input Type=""text"" name=""TemplateStr"" id=""t"&i&""" value="""
				Response.Write server.htmlencode(TemplateStr)
				Response.Write """ size=50>"
				End If
				Response.Write "<INPUT TYPE=""hidden"" NAME=""ReadME"" id=""r"&i&""" value="""&StyleHelpValue&""">"
				Response.Write "<a href=# onclick=""helpscript(r"&i&");return false;"" class=""helplink""><img src=""skins/images/help.gif"" border=0 title=""点击查阅管理帮助！""></a>"
			Case 2
				Response.Write "<input Type=""text"" name=""TemplateStr"" id=""t"&i&""" value="""
				Response.Write server.htmlencode(TemplateStr)
				Response.Write """ size=20> "
				If server.htmlencode(TemplateStr)<>"" And (Instr(server.htmlencode(TemplateStr),".gif") or Instr(server.htmlencode(TemplateStr),".jpg")) Then
					If InStr(TemplateStr,"{$PicUrl}")>0 Then
						Response.Write "<img src="&server.htmlencode(Replace(TemplateStr,"{$PicUrl}",MyDbPath&Dvbbs.Forum_PicUrl))&"  border=0>"
					Else
						Response.Write "<img src="&server.htmlencode(MyDbPath & TemplateStr)&"  border=0>"
					End If
				End if
			Case 3
				If page="pub"  And i=0 Then 
					Response.Write "<input type=hidden name=""TemplateStr"" value="""
					Response.Write server.htmlencode(TemplateStr)
					Response.Write """>"
					Response.Write "此字段属于基本设置，  <a href=""?action=editmain&stype=2&StyleID="&StyleID&""">点这里修改基本设置</a>"
					Response.Write "</td><td class=""td2"">"
					Response.Write "<a href=# onclick=""helpscript(r"&i&");return false;"" class=""helplink""><img src=""skins/images/help.gif"" border=0 title=""点击查阅管理帮助！""></a>"
				Else
					
					Response.Write "<textarea name=""TemplateStr"" id=""t"&i&""" cols=""100"" rows=""5"">"
					Response.Write server.htmlencode(TemplateStr)
					Response.Write "</textarea>"
					Response.Write "</td><td class=""td2""><a href=""javascript:admin_Size(-5,'t"&i&"')""><img src=""skins/images/minus.gif"" unselectable=""on"" border='0'></a> <a href=""javascript:admin_Size(5,'t"&i&"')""><img src=""skins/images/plus.gif"" unselectable=""on"" border='0'></a>"
					Response.Write "<img src=skins/images/viewpic.gif onclick=runscript(t"&i&")>"
					Response.Write "<a href=# onclick=""helpscript(r"&i&");return false;"" class=""helplink""><img src=""skins/images/help.gif"" border=0 title=""点击查阅管理帮助！""></a> "		
				End If
				Response.Write "<INPUT TYPE=""hidden"" NAME=""ReadME"" id=""r"&i&""" value="""&StyleHelpValue&""">"
			End Select
			
		Response.Write "</td></tr>"
	Next
	Response.Write "<tr><td class=""td2"" height=""25"" align=""center"" colspan=""3"">&nbsp;"
	Response.Write "</td></tr>"
	Response.Write "<tr><td class=""td2"" height=""25"" align=""center"">"
	Response.Write "<input type=""reset"" class=""button"" name=""Submit"" value=""重 填"">"
	Response.Write "</td>"
	Response.Write "<td class=""td2"" height=""25"" colspan=2 align=""center"">"
	Response.Write "<input type=""submit"" class=""button"" name=""B1"" value=""修 改"">"
	Response.Write "</td></tr>"
	Response.Write "<tr>"
	Response.Write "<td colspan=3 Class=""td2"">"
	Response.Write "<br /><li>重要提示，模板中含XSLT代码的，修改必须严格按照XML语法标准。"
	Response.Write "<br /><li>模板编辑规则：如果想清除该字段，请在对应的文本框中输入""del""，那么模板数据的序号就会前移。"
	Response.Write "<br /><li>如果不想改变模板数据的序号，仅把该项目的数据清空，则只需要把内容清空。"
	Response.Write "</td></tr>"
	Response.Write "</table><p></p>"
	Response.Write "</form>"
End Sub

Sub SaveEdit()
	If Request("dvbbs")<>"OK!" Then
		Errmsg=ErrMsg + "<br /><li>您提交了非法数据"
		Dvbbs_error()
		Exit Sub
	End If
	Dim Page,i
	Dim TempStr,TemplateStr,stype
	Dim TempStyleHelp,StyleHelpValue
	stype=Dvbbs.checkStr(request("stype"))
	page=Dvbbs.checkStr(request("page"))
	If Not IsNumeric(stype) Then 
		Errmsg=ErrMsg + "<br /><li>错误的样式参数"
		Dvbbs_error()
	End If
	'模板查错,更新缓存.	
	If stype="3" Then
		Select Case Request("page")
			Case "page_dispbbs"
				TemplateStr=Request.form("TemplateStr")(1)
				Set TempStr=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
				If Not TempStr.Loadxml(TemplateStr) Then
					Errmsg=ErrMsg + "论坛首页模板template.html(0)未能通过XML校验,请重新编辑修改,确保无误."
					Set TempStr=Nothing
					Dvbbs_error()
					Exit Sub
				End If
			Case "page_index"
				TemplateStr=Request.form("TemplateStr")(1)
				Set TempStr=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
				If Not TempStr.Loadxml(TemplateStr) Then
					Errmsg=ErrMsg + "论坛首页模板template.html(0)未能通过XML校验,请重新编辑修改,确保无误."
					Set TempStr=Nothing
					Dvbbs_error()
					Exit Sub
				End If
				TemplateStr=Request.form("TemplateStr")(2)
				Set TempStr=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				If Not TempStr.Loadxml(TemplateStr)  Then
					Errmsg=ErrMsg + "论坛首页模板template.html(1)未能通过XML校验,请重新编辑修改,确保无误."
					Set TempStr=Nothing
					Dvbbs_error()
					Exit Sub
				End If
				TemplateStr=Request.form("TemplateStr")(4)
				Set TempStr=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				If Not TempStr.Loadxml(TemplateStr)  Then
					Errmsg=ErrMsg + "论坛首页模板template.html(3)未能通过XML校验,请重新编辑修改,确保无误."
					Set TempStr=Nothing
					Dvbbs_error()
					Exit Sub
				End If
			Case "page_query"
				TemplateStr=Request.form("TemplateStr")(1)
				Set TempStr=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
				If Not TempStr.Loadxml(TemplateStr) Then
					Errmsg=ErrMsg + "论坛首页模板template.html(0)未能通过XML校验,请重新编辑修改,确保无误."
					Set TempStr=Nothing
					Dvbbs_error()
					Exit Sub
				End If
			Case "main_style"
				TemplateStr=Request.form("TemplateStr")(23)
				Set TempStr=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
				If Not TempStr.Loadxml(TemplateStr) Then
					Errmsg=ErrMsg + "论坛首页模板Dvbbs.mainhtml(22)未能通过XML校验,请重新编辑修改,确保无误."
					Set TempStr=Nothing
					Dvbbs_error()
					Exit Sub
				End If
		End Select
	End If

	TemplateStr=""
	i=0
	Dim FileName
	FileName = page &"_"& typeList(stype)
	REM 获取表单字段内容 和 写入模板文件 By Dv.唧唧  2007-10-9
	For Each TempStr in Request.form("TemplateStr")
		Dvbbs.writeToFile FilePath & FileName &i&".htm",Replace(TempStr,Chr(13)&Chr(10)&Chr(13)&Chr(10),Chr(13)&Chr(10))
		Dvbbs.Name="tpl" & LCase(Replace(Replace(FilePath,"../",""),"/","\")&FileName&i)
		Dvbbs.RemoveCache
		i=i+1
	Next
	'Response.end
	If stype="3" Then
		Select Case Request("page")
			Case "page_dispbbs"
					Application.Lock
					Application.Contents.Remove(Dvbbs.CacheName & "_dispbbsemplate_"& Request("StyleID"))
					Application.unLock
			Case "page_index"
				Application.Lock
				Application.Contents.Remove(Dvbbs.CacheName & "_listtemplate_"& Request("StyleID"))
				Application.Contents.Remove(Dvbbs.CacheName & "_indextemplate_"& Request("StyleID"))
				Application.Contents.Remove(Dvbbs.CacheName & "_shownews_"&Request("StyleID"))
				Application.unLock
			Case "page_query"
					Application.Lock
					Application.Contents.Remove(Dvbbs.CacheName & "_querytemplate_"& Request("StyleID"))
					Application.unLock
			Case "main_style"
				RestoreBoardCache()
			Case Else
		End Select
	End If

	Select Case stype
		Case 1
			Dv_suc(page&"语言包修改成功!")
		Case 2
			Dv_suc(page&"图片资源修改成功!")
		Case 3
			Dv_suc(page&"界面风格修改成功!")		
	End Select
	'更新缓存。此处是在模板数据变化的时候需要更新的代码。如有漏掉，可以在这添加。
	Dvbbs.Loadstyle()
End Sub

'Response.Write CountFilesNumber("../Resource/Style_1/","pub_html")
Rem 返回指定文件夹下 相同前缀的文件数量  By Dv.唧唧  2007-10-9
Function CountFilesNumber(ByVal path,ByVal folderspec)
    Dim objfso,f,fc,i
    Set objfso=Dvbbs.iCreateObject("Scripting.FileSystemObject")
	path = Server.MapPath(path)
	Set f = objfso.getfolder(path) 

	For Each fc In f.Files
		fc = LCase(fc)
		folderspec = LCase(folderspec)
		path = LCase(path)

		fc=Replace(fc,path,"")		
		If InStr(fc,folderspec)=2 And Trim(Right(fc,4))=".htm" Then
			i=i+1
			'Response.Write fc & "----" & folderspec &"<br />"
		End If
	Next 

    CountFilesNumber = i
	Set fc = Nothing
    Set f = Nothing
    Set objfso=nothing
End Function
%>

<%footer()Rem End Html%>