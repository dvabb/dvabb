<%
Class Pager
Public int_totalRecord
Private LAM_PageCount,LAM_Conn,LAM_Rs,LAM_SQL,LAM_PageSize,Str_errors,int_curpage,str_URL,int_totalPage,LAM_sURL,LAM_Style,int_style,LAM_CountRs
Private str_TableName,LAM_TableName,str_KeyName,LAM_KeyName,str_PageWhere,LAM_PageWhere,LAM_OrderType,int_OrderType,str_Tablezd,LAM_Tablezd

Public Property Let TableName(str_TableName)
	LAM_TableName=Trim(str_TableName)
End Property
Public Property Let Tablezd(str_Tablezd)
	LAM_Tablezd=str_Tablezd
End Property

Public Property Let KeyName(str_KeyName)
	LAM_KeyName=Trim(str_KeyName)
End Property

Public Property Let PageSize(int_PageSize)
	LAM_PageSize=CLng(int_PageSize)
End Property

Public Property Let OrderType(int_OrderType)
	LAM_OrderType=CLng(int_OrderType)
End Property

Public Property Let PageWhere(str_PageWhere)
	LAM_PageWhere=Trim(Replace(str_PageWhere,"'","''"))
End Property

Public Property Get GetRs()
	Set LAM_CountRs = LAM_Conn.Execute("Dv_GetRecordCount '"&LAM_TableName&"','"&LAM_PageWhere&"'")

	if not LAM_CountRs.eof then
		int_totalRecord = Int(LAM_CountRs.Fields(0))
	else
		int_totalRecord = 0
		int_totalPage = 1
		
	end if
	LAM_CountRs.Close
	set LAM_CountRs=nothing

	if int_totalRecord<>0 then
		int_totalPage = Int(int_totalRecord/LAM_PageSize)
		if int_totalRecord mod LAM_PageSize <>0 then
			int_totalPage = int_totalPage +1
		end if
	else
		int_totalPage = 1
	end if
	if int_curpage>int_totalPage then
		int_curpage = int_totalPage
	end if
	LAM_SQL = "Dv_GetRecordFromPage '"&LAM_TableName&"','"&LAM_KeyName&"',"&LAM_PageSize&","&int_curpage&","&LAM_OrderType&",'"&LAM_PageWhere&"','"&LAM_Tablezd&"'"
	Set LAM_Rs=Server.createobject("adodb.recordset")
	LAM_Rs.Open LAM_SQL,LAM_Conn,1,1
	Set GetRs=LAM_RS
End Property


Public Property Let GetConn(obj_Conn)
	Set LAM_Conn=obj_Conn
End Property


Public Property Let GetStyle(int_style)
	LAM_Style=int_style
End Property 


Private Sub Class_Initialize
	If dvbbs.checknumeric(request("star"))=0 Then
		If dvbbs.checknumeric(request("page"))=0 Then
		int_curpage=1
		Else
			If Dvbbs.ScriptName="dispbbs.asp" Then 
					Int_curpage=1
			Else 
					Int_curpage=dvbbs.checknumeric(request("page"))	
			End If 

		End If
	Else
	    Int_curpage=dvbbs.checknumeric(request("star"))
	End If  
End Sub


Public Function ShowPage()
	Dim str_tmp
	LAM_sURL = GetUrl()
	int_TotalPage=int_totalPage 

	Dim strHtml,prevPage,nextPage,startPage,i
if LAM_Style=1 then
'模式1 (10页缩略,首页,前页,后页,尾页)	
	prevPage = Int_curpage - 1
	nextPage = Int_curpage + 1
	
	strHtml = "<form method=post style=""margin:0px"" onsubmit=""window.location.href='"& LAM_sURL &"'+document.getElementById('page').value;return false;"">"
	if int_totalrecord>0 then
		if (prevPage < 1) then
			strHtml = strHtml& "<span title=""第一页"" style=""margin: 0px 0px 0px 1px;color: #999999;"">首页</span>&nbsp;"
			strHtml = strHtml& "<span title=""上一页"" style=""margin: 0px 0px 0px 1px;color: #999999;"">上页</span>&nbsp;"
		else
			strHtml = strHtml& "<span title=""第一页""><a href="""& LAM_sURL &"1"" style=""margin: 0px 0px 0px 1px;"">首页</a></span>&nbsp;"
			strHtml = strHtml& "<span title=""上一页""><a href="""& LAM_sURL &prevPage&""" style=""margin: 0px 0px 0px 1px;"">上页</a></span>&nbsp;"
		end if
		if (Int_curpage mod 10 =0) then
			startPage = Int_curpage - 9
		else 
			startPage = Int_curpage - Int_curpage mod 10 + 1
		end if
		if (startPage > 10) then
			strHtml = strHtml& "<span title=""上十页"" style=""margin: 0px 0px 0px 1px;""><a href="""& LAM_sURL &startPage-1&""">上十页</a></span>&nbsp;"
		end if
		for i = startPage to startPage + 9
			if (i > int_totalpage) then
				exit for
			end if
			if (i =Int_curpage) then
				strHtml = strHtml& "<span title=""第" & i & "页"" style=""color: #999999;margin: 0px 0px 0px 1px;background:background:#cccccc;width:16px;text-align:center;height:16px;border:1px solid #888888;padding:0px 3px"">" & i & "</span>&nbsp;"
			else 
				strHtml = strHtml& "<span title=""第" & i & "页""style=""margin: 0px 0px 0px 1px;background:#cccccc;width:16px;text-align:center;height:16px;border:1px solid #888888;padding:0px 3px""><a href="""& LAM_sURL &i&""">" & i & "</a></span>&nbsp;"
			end if
		next
		if (int_totalpage>1) then
		strHtml = strHtml& "<input name=""page"" value="""&Int_curpage&""" type=""text"" style=""border: 1px solid #cccccc;height=18px;width:25px;text-align:right;background-color: #fff;vertical-align : middle ;"" onkeypress=""if (event.keyCode == 8 || (event.keyCode >= 48 && event.keyCode <= 57) || event.keyCode == 13) return true;else return false;"" onfocus=""this.select();""/>"
		end if
		if (int_totalpage >= startPage + 10) then
			strHtml = strHtml& "&nbsp;<span title=""下十页"" style=""margin: 0px 0px 0px 1px;""><a href="""& LAM_sURL &startPage+10&""">下十页</a></span>"
		end if
		if (nextPage > int_totalpage) then
			strHtml = strHtml& "&nbsp;<span title=""下一页"" style=""margin: 0px 0px 0px 1px;color: #999999;padding:0px 3px"">下一页</span>&nbsp;"
			strHtml = strHtml& "<span title=""尾页"" style=""margin: 0px 0px 0px 1px;color: #999999;padding:0px 3px"">尾页</span>"
		else
			strHtml = strHtml& "&nbsp;<span title=""下页""><a href="""& LAM_sURL &nextPage&""" style=""margin: 0px 0px 0px 1px;padding:0px 3px"">下一页</a></span>&nbsp;"
			strHtml = strHtml& "<span title=""尾页""><a href="""& LAM_sURL &int_totalpage&""" style=""margin: 0px 0px 0px 1px;padding:0px 3px"">尾页</a></span>"
		end if
	end if
	strHtml = strHtml& "&nbsp;&nbsp;<span style=""font-weight: normal;padding: 0px;text-decoration: none;margin: 0px ;"">"&int_curpage&"/"&int_totalpage&"页 共"&int_totalrecord&"条 "&LAM_PageSize&"条/页&nbsp;&nbsp;</span></form>"
	
end if
if LAM_Style=2 then
'模式1 (10页缩略,首页,前页,后页,尾页)	
	prevPage = Int_curpage - 1
	nextPage = Int_curpage + 1
	
	strHtml = "<table  height=""20""  border=""1"" cellpadding=""0"" cellspacing=""0"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""  class=""Pager"" style=""BORDER-COLLAPSE: collapse;font-weight: normal;padding: 0px;text-decoration: none;margin: 0px ;"" bgcolor=""#e4e4e4""><form method=post onsubmit=""window.location.href='"& LAM_sURL &"'+document.getElementById('page').value;return false;""><tr><td bgcolor=""#FFFFFF"">页次:"&int_curpage&"/"&int_totalpage&"页 共"&int_totalrecord&"条记录 "&LAM_PageSize&"条/页</td>"
	if int_totalrecord>0 then
		if (prevPage < 1) then
			strHtml = strHtml& "<td title=""首页"" width=20 align=middle style=""font-family: Webdings;margin: 0px 0px 0px 1px;color: #999999;"">9</td>"
			strHtml = strHtml& "<td title=""上页"" width=20 align=middle style=""font-family: Webdings;margin: 0px 0px 0px 1px;color: #999999;"">7</td>"
		else
			strHtml = strHtml& "<td title=""首页"" width=20 align=middle ><a href="""& LAM_sURL &"1"" style=""font-family: Webdings;margin: 0px 0px 0px 1px;"">9</a></td>"
			strHtml = strHtml& "<td title=""上页"" width=20 align=middle ><a href="""& LAM_sURL &prevPage&""" style=""font-family: Webdings;margin: 0px 0px 0px 1px;"">7</a></td>"
		end if
		if (Int_curpage mod 10 =0) then
			startPage = Int_curpage - 9
		else 
			startPage = Int_curpage - Int_curpage mod 10 + 1
		end if
		if (startPage > 10) then
			strHtml = strHtml& "<td title=""上十页"" width=20 align=middle  style=""margin: 0px 0px 0px 1px;""><a href="""& LAM_sURL &startPage-1&""">...</a></td>"
		end if
		for i = startPage to startPage + 9
			if (i > int_totalpage) then
				exit for
			end if
			if (i =Int_curpage) then
				strHtml = strHtml& "<td title=""第" & i & "页""  bgcolor=""#eaf0f8"" width=20 align=middle  style=""color: #999999;margin: 0px 0px 0px 1px;""><b>" & i & "</b></td>"
			else 
				strHtml = strHtml& "<td title=""第" & i & "页"" width=20 align=middle style=""margin: 0px 0px 0px 1px;""><a href="""& LAM_sURL &i&""">" & i & "</a></td>"
			end if
		next
		if (int_totalpage>1) then
		strHtml = strHtml& "<td width=20 align=middle ><input name=""page"" value="""&Int_curpage&""" type=""text"" style=""border: 1px solid #cccccc;height=18px;width:25px;text-align:right;background-color: #fff;vertical-align : middle ;"" onkeypress=""if (event.keyCode == 8 || (event.keyCode >= 48 && event.keyCode <= 57) || event.keyCode == 13) return true;else return false;"" onfocus=""this.select();""/></td>"
		end if
		if (int_totalpage >= startPage + 10) then
			strHtml = strHtml& "<td title=""下十页"" width=20 align=middle  style=""margin: 0px 0px 0px 1px;""><a href="""& LAM_sURL &startPage+10&""">...</a></td>"
		end if
		if (nextPage > int_totalpage) then
			strHtml = strHtml& "<td title=""下页"" width=20 align=middle  style=""font-family: Webdings;margin: 0px 0px 0px 1px;color: #999999;"">8</td>"
			strHtml = strHtml& "<td title=""尾页"" width=20 align=middle  style=""font-family: Webdings;margin: 0px 0px 0px 1px;color: #999999;"">:</td>"
		else
			strHtml = strHtml& "<td title=""下页"" width=20 align=middle ><a href="""& LAM_sURL &nextPage&""" style=""font-family: Webdings;margin: 0px 0px 0px 1px;"">8</a></td>"
			strHtml = strHtml& "<td title=""尾页"" width=20 align=middle ><a href="""& LAM_sURL &int_totalpage&""" style=""font-family: Webdings;margin: 0px 0px 0px 1px;"">:</a></td>"
		end if
	end if
	strHtml = strHtml& "</tr></form></table>"

end if
if LAM_Style=3 then
'模式3 (10页缩略,首页,前页,后页,尾页) input样式
	prevPage = Int_curpage - 1
	nextPage = Int_curpage + 1
	
	strHtml = "<table  border=""0"" cellpadding=""0"" cellspacing=""0"" style=""font:12px;""><form method=post onsubmit=""window.location.href='"& LAM_sURL &"'+document.getElementById('page').value;return false;""><tr><td>页次:"&int_curpage&"/"&int_totalpage&"页 共"&int_totalrecord&"条记录 "&LAM_PageSize&"条/页</td>"
	if int_totalrecord>0 then
		if (prevPage < 1) then
			strHtml = strHtml& "<td><input type=button  value=""|<<"" title=""第一页"" disabled></td>"
			strHtml = strHtml& "<td><input type=button  value=""<<"" title=""上一页"" disabled></td>"
		else
			strHtml = strHtml& "<td><input type=button  value=""|<<"" title=""第一页"" onclick=""window.location.href='"& LAM_sURL &"1';"" >"
			strHtml = strHtml& "<td><input type=button  value=""<<"" title=""上一页"" onclick=""window.location.href='"& LAM_sURL &prevPage&"';"" >"
		end if
		if (Int_curpage mod 10 =0) then
			startPage = Int_curpage - 9
		else 
			startPage = Int_curpage - Int_curpage mod 10 + 1
		end if
		if (startPage > 10) then
			strHtml = strHtml& "<td><input type=button  value=""..."" title=""上十页"" onclick=""window.location.href='"& LAM_sURL &startPage-1&"';"" >"
		end if
		for i = startPage to startPage + 9
			if (i > int_totalpage) then
				exit for
			end if
			if (i =Int_curpage) then
				strHtml = strHtml& "<td><input type=button  value=""" & i & """ title=""第" & i & "页"" disabled></td>"
			else 
				strHtml = strHtml& "<td><input type=button  value=""" & i & """ title=""第" & i & "页"" onclick=""window.location.href='"& LAM_sURL &i&"';"" >"
			end if
		next
		if (int_totalpage>1) then
		strHtml = strHtml& "<td><input name=""page"" title=""请输入要跳转的页码,然后按回车即可.""  value="""&Int_curpage&""" type=""text"" style=""border: 1px solid #cccccc;height=18px;width:25px;text-align:right;background-color: #fff;vertical-align : middle ;"" onkeypress=""javascript:if (event.keyCode == 8 || (event.keyCode >= 48 && event.keyCode <= 57) || event.keyCode == 13) return true;else return false;"" onfocus=""this.select();""/></td>"
		end if
		if (int_totalpage >= startPage + 10) then
			strHtml = strHtml& "<td><input type=button  value=""..."" title=""下十页"" onclick=""window.location.href='"& LAM_sURL &startPage+10&"';"" >"
		end if
		if (nextPage > int_totalpage) then
			strHtml = strHtml& "<td><input type=button  value="">>"" title=""下一页"" disabled></td>"
			strHtml = strHtml& "<td><input type=button  value="">>|"" title=""最后页"" disabled></td>"
		else
			strHtml = strHtml& "<td><input type=button  value="">>"" title=""下一页"" onclick=""window.location.href='"& LAM_sURL &nextPage&"';"" >"
			strHtml = strHtml& "<td><input type=button  value="">>|"" title=""最后页"" onclick=""window.location.href='"& LAM_sURL &int_totalpage&"';"" >"
		end if
	end if
	strHtml = strHtml& "</tr></form></table>"
end If
ShowPage=strHtml
End Function

Private Function GetURL()
	Dim url_host,url_string
	url_string=""
	url_host=request.ServerVariables("script_name")
	If Request.QueryString<>"" Then
		Dim Get_Query
		For Each Get_Query In Request.QueryString
			if LCase(Get_Query)<>"page" then
				if url_string="" then
					url_string=Get_Query&"="&Request.QueryString(Get_Query)
				else
					url_string=url_string&"&"&Get_Query&"="&Request.QueryString(Get_Query)
				end if
			end if
		Next
	End If
	If Request.Form<>"" Then
		Dim Post_Query
		For Each Post_Query In Request.Form
			if LCase(Post_Query)<>"page" then
				if url_string="" then
					url_string=Post_Query&"="&Request.Form(Post_Query)
				else
					url_string=url_string&"&"&Post_Query&"="&Request.Form(Post_Query)
				end if
			end if
		Next
	End If
	if 	url_string="" then
		GetURL=url_host&"?page="
	else
		GetURL=url_host&"?"&url_string&"&page="
	end if
End Function

Private Sub Class_Terminate
	'LAM_RS.close:Set LAM_RS=nothing
End Sub

End class
%>