<!-- #include file =../conn.asp-->
<!-- #include file="inc/const.asp" -->
<!--#include file="../inc/dv_clsother.asp"-->
<%
Head()
Dim Admin_Flag
Dim NewsConfigFile
Dim XmlDoc,Node
Dim NewsName,NewsType,Updatetime,Skin_Head,Skin_Main,Skin_Footer,NewsSql
Admin_flag=",10,"
CheckAdmin(admin_flag)
NewsConfigFile = MyDbPath & "Dv_ForumNews/Dv_NewsSetting.config"
NewsConfigFile = Server.MapPath(NewsConfigFile)

Main()
If FoundErr Then Call Dvbbs_Error()
Footer()

Sub Main()
%>
<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">
<tr><th colspan="2" height="23">论坛首页调用管理</th></tr>
<tr>
<td width="20%" class="td1" align="center">
<button Style="width:80;height:50;border: 1px outset;" class="button">注意事项</button>
</td>
<td width="80%" class="td2">
	①添加调用后，在列表中点击相应的预览可以看到效果，将调用代码复制到你的首页就可以了。
	<br>②如果你的首页是和论坛程序分开，在填写调用模板时建议用上绝对地址路径。
	<br>③若需要设置外部调用限制和设置临时文件名，修改Dv_News.asp文件，文件里附有说明。
	<br>④建议根据不同的调用设定更新时间间隔，如不是经常更新的版块调用可以设置长一些时间间隔，这样可以有效地减低消耗。
</td>
</tr>
<tr><td colspan="2" class="td2">
<a href="?Act=AddSetting">添加首页调用</a> | <a href="?Act=NewsList">首页调用列表</a> | <a href="<%=MyDbPath%>Dv_News_Demo.asp" target="_blank">查看所有调用演示</a>
</td></tr>
</table>
<%
	Select Case Request("Act")
		Case "NewsList": Call NewsList()
		Case "AddSetting" , "EditNewsInfo","CopyNewsInfo" : Call AddSetting()
		Case "SaveSetting" , "SaveEditSetting","SaveCopySetting" : Call SaveSetting()
		Case "DelNewsInfo" : Call DelNewsInfo()
		Case Else
		Call NewsList()
	End Select
End Sub

'删除记录
Sub DelNewsInfo()
	Dim DelNodes,DelChildNodes
	Set XmlDoc = Dvbbs.CreateXmlDoc("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	If Not XmlDoc.load(NewsConfigFile) Then
		ErrMsg = "调用列表中为空，请填写调用后再执行本操作!"
		Dvbbs_Error()
		Exit Sub
	End If
	'Response.Write Request.Form("DelNodes").count
	For Each DelNodes in Request.Form("DelNodes")
		Set DelChildNodes = XmlDoc.DocumentElement.selectSingleNode("NewsCode[@AddTime='"&DelNodes&"']")
		If Not (DelChildNodes is nothing) Then
			XmlDoc.DocumentElement.RemoveChild(DelChildNodes)
		End If
	Next
	Call SaveXml()
	Dv_suc("所选的记录已删除!")
End Sub

Sub SaveSetting()
	NewsName	= Replace(Request.Form("NewsName"),"""","")
	NewsType	= Replace(Request.Form("NewsType"),"""","")
	Updatetime	= Dvbbs.CheckNumeric(Request.Form("Updatetime"))
	Skin_Head	= Request.Form("Skin_Head")
	Skin_Main	= Request.Form("Skin_Main")
	Skin_Footer	= Request.Form("Skin_Footer")

	If NewsName="" Then
		Errmsg=ErrMsg + "<li>请填写调用标识！</li>"
	Else
		NewsName = Lcase(NewsName)
	End If
	If NewsType < "1" Then
		Errmsg=ErrMsg + "<li>选取调用类型！</li>"
	End If
	If Skin_Main = "" Then
		Errmsg=ErrMsg + "<li>模板_主体循环标记部分不能为空！</li>"
	End If
	If Errmsg<>"" Then Dvbbs_Error() : Exit Sub
	Call LoadXml()

	If FoundNewsName(NewsName) and Request("Act") <> "SaveEditSetting" Then
		Errmsg=ErrMsg + "<li>调用标识已存在，不能重复添加！</li>"
		Dvbbs_Error()
		Exit Sub
	End If
	Select Case NewsType
		Case "1"		'帖子调用
			Call NewsType_1()
		Case "2"		'信息调用
			Call NewsType_2()
		Case "3"		'版块调用
			Call NewsType_3()
		Case "4"		'会员调用
			Call NewsType_4()
		Case "5"		'公告调用
			Call NewsType_5()
		Case "6"		'展区调用
			Call NewsType_6()
		Case "7"		'圈子调用
			Call NewsType_7()
		Case "8"		'登录框调用
			Call NewsType_8()
		Case Else
			Errmsg=ErrMsg + "<li>请正确选取调用类型！</li>"
			Dvbbs_Error()
	End Select
	Call CreateXmlLog()
	Call SaveXml()
	Dv_suc("调用设置成功!")
End Sub

Sub LoadXml()
	Set XmlDoc = Dvbbs.CreateXmlDoc("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	If Not XmlDoc.load(NewsConfigFile) Then
		XmlDoc.loadxml "<?xml version=""1.0"" encoding=""gb2312""?><NewscodeInfo/>"
	End If
End Sub

'检查是否存在相同的标识
Function FoundNewsName(NewsName)
	Dim Test
	Set Test = XmlDoc.DocumentElement.selectSingleNode("NewsCode[@NewsName="""&NewsName&"""]")
	FoundNewsName = not (Test is nothing)
End Function

Sub SaveXml()
	XmlDoc.save NewsConfigFile
	Set XmlDoc = Nothing
End Sub

'公共记录
Sub CreateXmlLog()
	Dim attributes,createCDATASection,ChildNode
	Dim FormName,NoAttFormName
	Dim Addtime
	AddTime = Now()
	If Request("Act") = "SaveEditSetting" and Request.Form("AddTime")<>"" Then
		Set Node = XmlDoc.DocumentElement.selectSingleNode("NewsCode[@AddTime='"&Request.Form("AddTime")&"']")
		If Not (Node is nothing) Then
			AddTime = Node.getAttribute("AddTime")
			XmlDoc.DocumentElement.RemoveChild(Node)
		End If
	End If
	'创建节点
	Set Node=XmlDoc.createNode(1,"NewsCode","")
	NoAttFormName = ",Skin_Head,Skin_Main,Skin_Footer,Act,AddTime,Board_Input0,Board_Input1,Board_Input2,Board_Input3,Board_Input4,"
	For Each FormName In Request.Form
		If Instr(NoAttFormName,","&FormName&",")=0 Then
			Set attributes=XmlDoc.createAttribute(FormName)
			If FormName="NewsName" Then
				attributes.text = Lcase(Replace(Request.Form(FormName),"""",""))
			Else
				attributes.text = Replace(Request.Form(FormName),"""","")
			End If
			node.attributes.setNamedItem(attributes)
		End If
	Next
	Set attributes=XmlDoc.createAttribute("MasterName")
	attributes.text = Dvbbs.Membername
	node.attributes.setNamedItem(attributes)
	Set attributes=XmlDoc.createAttribute("MasterUserID")
	attributes.text = Dvbbs.UserID
	node.attributes.setNamedItem(attributes)
	Set attributes=XmlDoc.createAttribute("MasterIP")
	attributes.text = Dvbbs.UserTrueIP
	node.attributes.setNamedItem(attributes)
	Set attributes=XmlDoc.createAttribute("AddTime")
	attributes.text = AddTime
	node.attributes.setNamedItem(attributes)
	Set attributes=XmlDoc.createAttribute("LastTime")
	attributes.text = DateAdd("s", -Updatetime,now())
	node.attributes.setNamedItem(attributes)
	Set ChildNode = XmlDoc.createNode(1,"Search","")
	Set createCDATASection=XmlDoc.createCDATASection(replace(NewsSql,"]]>",""))
	ChildNode.appendChild(createCDATASection)
	node.appendChild(ChildNode)
	Set ChildNode = XmlDoc.createNode(1,"Skin_Head","")
	Set createCDATASection=XmlDoc.createCDATASection(replace(Skin_Head,"]]>","]]&gt;"))
	ChildNode.appendChild(createCDATASection)
	node.appendChild(ChildNode)
	Set ChildNode = XmlDoc.createNode(1,"Skin_Main","")
	Set createCDATASection=XmlDoc.createCDATASection(replace(Skin_Main,"]]>","]]&gt;"))
	ChildNode.appendChild(createCDATASection)
	node.appendChild(ChildNode)
	Set ChildNode = XmlDoc.createNode(1,"Skin_Footer","")
	Set createCDATASection=XmlDoc.createCDATASection(replace(Skin_Footer,"]]>","]]&gt;"))
	ChildNode.appendChild(createCDATASection)
	node.appendChild(ChildNode)
	''特殊版面增加
	If NewsType = "3" Then
		Set ChildNode = XmlDoc.createNode(1,"Board_Input0","")
		Set createCDATASection=XmlDoc.createCDATASection(Replace(Request.Form("Board_Input0"),"]]>","]]&gt;"))
		ChildNode.appendChild(createCDATASection)
		node.appendChild(ChildNode)
		Set ChildNode = XmlDoc.createNode(1,"Board_Input1","")
		Set createCDATASection=XmlDoc.createCDATASection(Replace(Request.Form("Board_Input1"),"]]>","]]&gt;"))
		ChildNode.appendChild(createCDATASection)
		node.appendChild(ChildNode)
		Set ChildNode = XmlDoc.createNode(1,"Board_Input2","")
		Set createCDATASection=XmlDoc.createCDATASection(Replace(Request.Form("Board_Input2"),"]]>","]]&gt;"))
		ChildNode.appendChild(createCDATASection)
		node.appendChild(ChildNode)
		Set ChildNode = XmlDoc.createNode(1,"Board_Input3","")
		Set createCDATASection=XmlDoc.createCDATASection(Replace(Request.Form("Board_Input3"),"]]>","]]&gt;"))
		ChildNode.appendChild(createCDATASection)
		node.appendChild(ChildNode)
		Set ChildNode = XmlDoc.createNode(1,"Board_Input4","")
		Set createCDATASection=XmlDoc.createCDATASection(Replace(Request.Form("Board_Input4"),"]]>","]]&gt;"))
		ChildNode.appendChild(createCDATASection)
		node.appendChild(ChildNode)
	ElseIf NewsType = "6" Then
		Set ChildNode = XmlDoc.createNode(1,"Board_Input0","")
		Set createCDATASection=XmlDoc.createCDATASection(Replace(Request.Form("Board_Input0"),"]]>","]]&gt;"))
		ChildNode.appendChild(createCDATASection)
		node.appendChild(ChildNode)
	End If
	XmlDoc.documentElement.appendChild(node)
End Sub


'帖子调用
Sub NewsType_1()
	Dim News_Total,Topiclen,Orders,TopicType,Boardid,BoardLimit,BoardType,UserIDList,Sdate
	News_Total = Dvbbs.CheckNumeric(Request.Form("Total"))
	Topiclen = Dvbbs.CheckNumeric(Request.Form("Topiclen"))
	Orders = Request.Form("Orders")
	Sdate = Dvbbs.CheckNumeric(Request.Form("Sdate"))
	TopicType = Request.Form("TopicType")
	Boardid = Dvbbs.CheckNumeric(Request.Form("Boardid"))
	BoardLimit = Dvbbs.CheckNumeric(Request.Form("BoardLimit"))
	BoardType = Request.Form("BoardType")
	UserIDList = Request.Form("UserIDList")
	If News_Total = 0 Then News_Total = 10
	Dim OrderBy,Searchstr,SearchBoard,Tempstr
	NewsSql = "SELECT TOP "& News_Total
	If Orders = "3" Then
		'修正按热帖排序显示精华出错 2007-7-6 Dv.Yz
		If TopicType = 1 Then
			OrderBy = " T.Hits Desc, "
		ELSE
			OrderBy = " Hits Desc, "
		End If
	ElseIf orders = "1" or orders = "2" Then
		If TopicType = 1 Then
			OrderBy = " B.Dateandtime Desc, "
		Else
			OrderBy = " Dateandtime Desc, "
		End If
	End If
	'指定版面
	If Boardid>0 Then
		If TopicType = 1 Then
			SearchBoard = " AND B.Boardid = " & Boardid
		Else
			SearchBoard = " AND Boardid = " & Boardid
		End If
		If BoardType > 0 Then
			Tempstr = GetChildBoardID(Boardid)
			If BoardType = 2 Then
				Tempstr = Boardid & "," &Tempstr
			End If
			If Tempstr<>"" Then
				Tempstr = Left(Tempstr,InStrRev(Tempstr, ",")-1)
				If TopicType = 1 Then
					SearchBoard = " AND B.Boardid IN (" & Tempstr &") "
				Else
					SearchBoard = " AND Boardid IN (" & Tempstr &") "
				End If
			End If
		End If
	Else
		Tempstr = Cstr(Boardid)
	End If
	If Boardid>0 Then
		If TopicType = 1 Then
			SearchBoard = " AND B.Boardid = " & Boardid
		Else
			SearchBoard = " AND Boardid = " & Boardid
		End If
		If BoardType > 0 Then
			Tempstr = GetChildBoardID(Boardid)
			If BoardType = 2 Then
				Tempstr = Boardid & "," &Tempstr
			End If
			If Tempstr<>"" Then
				Tempstr = Left(Tempstr,InStrRev(Tempstr, ",")-1)
				If TopicType = 1 Then
					SearchBoard = " AND B.Boardid IN (" & Tempstr &") "
				Else
					SearchBoard = " AND Boardid IN (" & Tempstr &") "
				End If
			End If
		End If
	Else
		Tempstr = Cstr(Boardid)
	End If
	'限制不显示特列版面
	If BoardLimit="1" and Tempstr<>"" Then
		Tempstr = GetBoardid(Tempstr)
		If Not Boardid = 0 Then
			If TopicType = 1 Then
				SearchBoard = " AND B.Boardid IN (" & Tempstr &") "
			Else
				SearchBoard = " AND Boardid IN (" & Tempstr &") "
			End If
		Else
			If Not Tempstr = "" Then
				If TopicType = 1 Then
					SearchBoard = " AND B.Boardid NOT IN (" & Tempstr &") "
				Else
					SearchBoard = " AND Boardid NOT IN (" & Tempstr &") "
				End If
			End If
		End If
	End If
	If Not SearchBoard = "" Then
		Searchstr = SearchBoard
	End If
	If UserIDList<>"" Then
		If Instr(UserIDList,",") Then
			If IsNumeric(Replace(UserIDList,",","")) Then
				If TopicType = 1 Then
					Searchstr = Searchstr & " AND B.PostUserID IN ("&UserIDList&")"
				Else
					Searchstr = Searchstr & " AND PostUserID IN ("&UserIDList&")"
				End If
			End If
		Else
			UserIDList = Dvbbs.CheckNumeric(UserIDList)
			If UserIDList > 0 Then
				If TopicType = 1 Then
					Searchstr = Searchstr & " AND B.PostUserID = " & UserIDList
				Else
					Searchstr = Searchstr & " AND PostUserID = " & UserIDList
				End If
			End If
		End If
	End If
	If Sdate>0 Then
		If IsSqlDataBase=1 Then
			If TopicType = 1 Then
				Searchstr = Searchstr & " AND Datediff(day, B.DateAndTime, " & SqlNowString & ") < " & Sdate
			Else
				Searchstr = Searchstr & " AND Datediff(day, DateAndTime, " & SqlNowString & ") < " & Sdate
			End If
		Else
			If TopicType = 1 Then
				Searchstr = Searchstr & " AND Datediff('d', B.DateAndTime, " & SqlNowString & ") < " & Sdate
			Else
				Searchstr = Searchstr & " AND Datediff('d', DateAndTime, " & SqlNowString & ") < " & Sdate
			End If
		End If
	End If
	If TopicType = 1 Then		'显示精华主题
		If Searchstr<>"" Then
			Searchstr = " WHERE " & Mid(Searchstr, InStr(Searchstr, "AND")+3)
		End If
		NewsSql = NewsSql & " B.PostUserName, B.Title, B.Rootid, B.Boardid, B.Dateandtime, B.Announceid, B.Id, B.Expression From Dv_BestTopic B INNER JOIN Dv_Topic T ON B.RootID = T.TopicID " & Searchstr & " ORDER BY " & OrderBy & " B.Id Desc"
	ElseIf TopicType=2 Then		'显示主题和回复
		NewsSql = NewsSql & " UserName,Topic,Rootid,Boardid,Dateandtime,Announceid,Body,Expression From "&Dvbbs.NowUseBBS&" Where not (Boardid in (444,777)) "& Searchstr &" ORDER BY "& OrderBy &" AnnounceID Desc"
	Else		'显示主题
		If Orders = 2 Then OrderBy = " Lastposttime Desc, "
		NewsSql = NewsSql & " PostUserName,Title,Topicid,Boardid,Dateandtime,Topicid,Hits,Expression,LastPost From [Dv_topic] Where not (Boardid in (444,777)) "& Searchstr & " ORDER BY "& OrderBy &" Topicid Desc"
	End If
End Sub

'信息调用
Sub NewsType_2()
End Sub

'版块调用
Sub NewsType_3()
End Sub

'会员调用
Sub NewsType_4()
	Dim News_Total,Orders
	News_Total = Dvbbs.CheckNumeric(Request.Form("Total"))
	Orders = Request.Form("Orders")
	Dim OrderBy
	If News_Total = 0 Then News_Total = 10
	NewsSql = "SELECT TOP "& News_Total &" UserID,UserName,UserTopic,UserPost,UserIsBest,UserWealth,UserCP,UserEP,UserDel,UserSex,JoinDate,UserLogins From [Dv_user] "
	Select Case Request.Form("UserOrders")
	Case "0"
		'OrderBy = " JoinDate desc, "
		OrderBy = ""
	Case "1"
		OrderBy = " UserPost desc, "
	Case "2"
		OrderBy = " UserTopic desc, "
	Case "3"
		OrderBy = " UserIsBest desc, "
	Case "4"
		OrderBy = " UserWealth desc, "
	Case "5"
		OrderBy = " UserEP desc, "
	Case "6"
		OrderBy = " UserCP desc, "
	Case "7"
		OrderBy = " UserDel desc, "
	Case "8"
		OrderBy = " UserLogins desc, "
	End Select
	NewsSql = NewsSql & " ORDER BY " & OrderBy & " UserID desc "
End Sub

'公告调用
Sub NewsType_5()
	Dim News_Total,Boardid
	News_Total = Dvbbs.CheckNumeric(Request.Form("Total"))
	Boardid = Dvbbs.CheckNumeric(Request.Form("Boardid"))
	If News_Total = 0 Then News_Total = 10
	NewsSql = "SELECT TOP "& News_Total &" ID,Boardid,Title,UserName,AddTime FROM [Dv_bbsnews] "
	If Boardid > 0 Then
		NewsSql = NewsSql & " WHERE Boardid="& Boardid
	End If
	NewsSql = NewsSql & " ORDER BY ID DESC"
End Sub

'展区调用
Sub NewsType_6()
	Dim News_Total,Boardid,FileOrders,BoardLock,FileType,BoardLimit
	Dim Searchstr,OrderBy
	News_Total = Dvbbs.CheckNumeric(Request.Form("Total"))
	Boardid = Dvbbs.CheckNumeric(Request.Form("Boardid"))
	FileOrders = Request.Form("FileOrders")
	BoardLock = Dvbbs.CheckNumeric(Request.Form("BoardLock"))
	FileType = Request.Form("FileType")
	BoardLimit = Dvbbs.CheckNumeric(Request.Form("BoardLimit"))
	If News_Total = 0 Then News_Total = 8
	If FileType<>"all" Then
		FileType = Dvbbs.CheckNumeric(FileType)
		Searchstr = " AND F_Type = "&FileType
	End If

	'指定版面
	Dim SearchBoard
	Dim Rs,Tempstr
	If Boardid > 0 Then
		Select Case BoardLock
		Case 1
			SearchBoard = " AND F_BoardID <> " & Boardid
			Tempstr = "0"
		Case 3,4
			Tempstr = GetChildBoardID(Boardid)
			If BoardLock = 4 Then
				Tempstr = Boardid & "," &Tempstr
			End If
			If TempStr<>"" Then
				Tempstr = Left(Tempstr,InStrRev(Tempstr, ",")-1)
				SearchBoard = " AND F_BoardID in (" & Tempstr &") "
			End If
		Case Else
			SearchBoard = " AND F_BoardID = " & Boardid
		End Select
	Else
		Tempstr = Cstr(Boardid)
	End If

	'限制不显示特列版面
	If BoardLimit="1" and Tempstr<>"" Then
		Tempstr = GetBoardid(Tempstr)
		If Boardid<>0 Then
			If BoardLock = 1 Then
				SearchBoard = " AND F_BoardID in (" & Boardid &","& Tempstr &") "
			Else
				SearchBoard = " AND F_BoardID in (" & Tempstr &") "
			End If
		Else
			If Tempstr<>"" Then
				SearchBoard = " AND F_BoardID not in (" & Tempstr &") "
			End If
		End If
	End If
	Select Case FileOrders
	Case 1
		OrderBy = " F_ViewNum DESC, "
	Case 2
		OrderBy = " F_DownNum DESC, "
	Case 3
		OrderBy = " F_FileSize DESC, "
	Case Else
		OrderBy = ""
	End Select
	Searchstr = Searchstr & SearchBoard
	NewsSql = "SELECT TOP "& News_Total &" F_ID,F_AnnounceID,F_BoardID,F_Username,F_Filename,F_Readme,F_Type,F_FileType,F_AddTime,F_Viewname,F_ViewNum,F_DownNum,F_FileSize FROM [DV_Upfile] WHERE F_Flag<>4 "
	NewsSql = NewsSql & Searchstr & " ORDER BY "& OrderBy &" F_ID DESC"
End Sub

Rem 圈子调用，已去掉，小易
Sub NewsType_7()
	
End Sub

Sub NewsType_8()
End Sub


'BoardidVal<>0 取出调用的版面ID，当BoardidVal=0 取出不被调用的版面ID
Function GetBoardid(BoardidVal)
	Dim TempData,Nodelist,Nodes
	If BoardidVal<>"0" Then
		BoardidVal = "," & BoardidVal & ","
	End If

	Set Nodelist = Application(Dvbbs.CacheName&"_boardlist").cloneNode(True).documentElement.getElementsByTagName("board")
	For Each Nodes in Nodelist
		If BoardidVal<>"0" Then
			If Instr(BoardidVal,","&Nodes.attributes.getNamedItem("boardid").text&",") and Nodes.attributes.getNamedItem("hidden").text="0" and Nodes.attributes.getNamedItem("checkout").text="0" Then
				TempData = TempData & Nodes.attributes.getNamedItem("boardid").text &","
			End If
		Else
			If Nodes.attributes.getNamedItem("hidden").text="1" or Nodes.attributes.getNamedItem("checkout").text="1" Then
				TempData = TempData & Nodes.attributes.getNamedItem("boardid").text &","
			End If
		End If
	Next
	If TempData<>"" Then
		GetBoardid = Left(TempData,InStrRev(TempData, ",")-1)
	End If
End Function

'获取下属版块ID
Private Function GetChildBoardID(BoardIDVal)
		Dim TempData,Nodelist,Node
		Set Nodelist = Application(Dvbbs.CacheName&"_boardlist").cloneNode(True).documentElement.getElementsByTagName("board")
		For Each Node in Nodelist
			If Instr(","&Node.attributes.getNamedItem("parentstr").text&",",","&BoardIDVal&",")>0 Then
				TempData = TempData & Node.attributes.getNamedItem("boardid").text &","
			End If
		Next
		GetChildBoardID = TempData
End Function

Sub AddSetting()
	Dim ChildNode,attributes,Action
	Call LoadXml()
	If Request("Act") = "EditNewsInfo" Then
		Set Node = XmlDoc.DocumentElement.selectSingleNode("NewsCode[@AddTime='"&Request("DelNodes")&"']")
		If (Node is nothing) Then
			ErrMsg = "<li>所选取的调用已不存在!</li>"
			Dvbbs_Error()
			Exit Sub
		End If
		Action = "SaveEditSetting"
	ElseIf Request("Act") = "CopyNewsInfo" Then
		Set Node = XmlDoc.DocumentElement.selectSingleNode("NewsCode[@AddTime='"&Request("DelNodes")&"']")
		If (Node is nothing) Then
			ErrMsg = "<li>所选取的调用已不存在!</li>"
			Dvbbs_Error()
			Exit Sub
		End If
		Action = "SaveCopySetting"
	Else 
		Set Node=XmlDoc.createNode(1,"NewsCode","")
		Set ChildNode = XmlDoc.createNode(1,"Skin_Head","")
		node.appendChild(ChildNode)
		Set ChildNode = XmlDoc.createNode(1,"Skin_Main","")
		node.appendChild(ChildNode)
		Set ChildNode = XmlDoc.createNode(1,"Skin_Footer","")
		node.appendChild(ChildNode)
		Action = "SaveSetting"
	End If
	'当不是编辑版面调用时创建临时节点
	If NewsType <> "3" or NewsType <> "6" Then
		Set ChildNode = XmlDoc.createNode(1,"Board_Input0","")
		node.appendChild(ChildNode)
		Set ChildNode = XmlDoc.createNode(1,"Board_Input1","")
		node.appendChild(ChildNode)
		Set ChildNode = XmlDoc.createNode(1,"Board_Input2","")
		node.appendChild(ChildNode)
		Set ChildNode = XmlDoc.createNode(1,"Board_Input3","")
		node.appendChild(ChildNode)
		Set ChildNode = XmlDoc.createNode(1,"Board_Input4","")
		node.appendChild(ChildNode)
	End If
	Set XmlDoc = Nothing
	Dim Boardid
	Boardid = "0"
	If Node.getAttribute("Boardid") <> "" Then
		Boardid = Node.getAttribute("Boardid")
	End If
%>
<br>
<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">
<form METHOD=POST ACTION="?Act=<%=Action%>" name="TheForm">
<tr><th colspan="2" height="23">首页调用管理</th></tr>
<tr>
<td width="30%" class="td2" align="right">
调用标识名称：
</td>
<%
If Request("Act") = "CopyNewsInfo" Then 
%>
<td width="70%" class="td1">
<INPUT TYPE="text" NAME="NewsName" size="20" Maxlength="10" onkeyup="OutputNewsCode(this.value);" value="<%=Node.getAttribute("NewsName")&"_copy"%>">(请使用英文或数字设定调用名称,并且是唯一标识.不能超出10个字符)
</td>
<%
Else 
%>
<td width="70%" class="td1">
<INPUT TYPE="text" NAME="NewsName" size="20" Maxlength="10" onkeyup="OutputNewsCode(this.value);" value="<%=Node.getAttribute("NewsName")%>">(请使用英文或数字设定调用名称,并且是唯一标识.不能超出10个字符)
</td>
<%
End If 
%>
</tr>
<tr>
<td width="15%" class="td2" align="right">
调用代码：
</td>
<%
If Request("Act") = "CopyNewsInfo" Then 
%>
<td width="85%" class="td1">
<INPUT TYPE="text" NAME="Newscode" size="70" disabled value="<script src=&quot;Dv_News.asp?GetName=<%=Node.getAttribute("NewsName")&"_copy"%>&quot;></script>">
</td>
<%
Else 
%>
<td width="85%" class="td1">
<INPUT TYPE="text" NAME="Newscode" size="70" disabled value="<script src=&quot;Dv_News.asp?GetName=<%=Node.getAttribute("NewsName")%>&quot;></script>">
</td>
<%
End If 
%>
</tr>
<tr>
<td class="td2" align="right">
调用说明：
</td>
<td class="td1">
<INPUT TYPE="text" NAME="Intro" size="30" Maxlength="30" value="<%=Node.getAttribute("Intro")%>">(提示说明,以作管理区分.不能超出30个字符)
</td>
</tr>
<tr>
<td class="td2" align="right">
调用类型：
</td>
<td class="td1">
	<SELECT NAME="NewsType" ID="NewsType" onchange="NewsTypeSel(this.selectedIndex)">
	<option value="0">选取调用类型</option>
	<option value="1">帖子调用</option>
	<option value="2">信息调用</option>
	<option value="3">版块调用</option>
	<option value="4">会员调用</option>
	<option value="5">公告调用</option>
	<option value="6">展区调用</option>
	<option value="8">登录框调用</option>
	</SELECT>
</td>
</tr>
<tr>
<td class="td2" align="right">
数据更新间隔：
</td>
<td class="td1"><INPUT TYPE="text" NAME="Updatetime" value="<%=Node.getAttribute("Updatetime")%>">(单位：秒)</td>
</tr>
<tr>
<td class="td2" align="right">
时间显示格式：
</td>
<td class="td1">
<SELECT NAME="FormatTime" ID="FormatTime">
	<option value="0" SELECTED>YYYY-M-D H:M:S(长格式)</option>
	<option value="1">YYYY年M月D</option>
	<option value="2">YYYY-M-D</option>
	<option value="3">H:M:S</option>
	<option value="4">hh:mm</option>
</SELECT>
(按服务器时间区域格式显示。)
</td>
</tr>

<tr>
<td class="td2" align="right" valign="top">调用设置：</td>
<td class="td2">
<div id="News"></div>
</td>
</tr>
<!-- 调用模板设置 -->
<tr><th colspan="2" height="23">调用模板设置(请用HTML语法填写)</th></tr>
<tr>
<td class="td2" align="right" valign="top">模板_开始标记部分
</td>
<td class="td2">
	<textarea name="Skin_Head" ID="Skin_Head" style="width:100%;" rows="3"><%=Server.Htmlencode(Node.selectSingleNode("Skin_Head").text&"")%></textarea>
	<br><a href="javascript:admin_Size(-3,'Skin_Head')"><img src="skins/images/minus.gif" unselectable="on" border='0'></a> <a href="javascript:admin_Size(3,'Skin_Head')"><img src="skins/images/plus.gif" unselectable="on" border='0'></a>
</td>
</tr>
<tr>
<td class="td2" align="right" valign="top">
模板_主体循环标记部分
<fieldset title="模板变量">
<legend>&nbsp;模板变量说明&nbsp;</legend>
<div id="skin_info" align="left"></div>
</fieldset>
</td>
<td class="td2" valign="top">
	<div id="DisInput"></div>
	<textarea name="Skin_Main" ID="Skin_Main" style="width:100%;" rows="10"><%=Server.Htmlencode(Node.selectSingleNode("Skin_Main").text&"")%></textarea>
	<br><a href="javascript:admin_Size(-3,'Skin_Main')"><img src="skins/images/minus.gif" unselectable="on" border='0'></a> <a href="javascript:admin_Size(3,'Skin_Main')"><img src="skins/images/plus.gif" unselectable="on" border='0'></a>
</td>
</tr>
<tr>
<td class="td2" align="right" valign="top">模板_结束标记部分
</td>
<td class="td2">
	<textarea name="Skin_Footer" ID="Skin_Footer" style="width:100%;" rows="3"><%=Server.Htmlencode(Node.selectSingleNode("Skin_Footer").text&"")%></textarea>
	<br><a href="javascript:admin_Size(-3,'Skin_Footer')"><img src="skins/images/minus.gif" unselectable="on" border='0'></a> <a href="javascript:admin_Size(3,'Skin_Footer')"><img src="skins/images/plus.gif" unselectable="on" border='0'></a>
</td>
</tr>
<!-- 调用模板设置 -->
<tr>
<td class="td2" align="right">&nbsp;
</td>
<td class="td2" align="center">
<INPUT TYPE="submit" class="button" value="提交">&nbsp;&nbsp;&nbsp;<INPUT TYPE="reset" class="button" value="重填">
<INPUT TYPE="hidden" name="AddTime" value="<%=Node.getAttribute("AddTime")%>">
</td>
</tr>
</form>
</table>
<!-- 设置信息部分 -->
<div id="News_1" style="display:none">
<!-- 帖子调用 -->
<table border="0" cellpadding="3" cellspacing="1" width="100%">
<tr>
<td class="td1">
显示记录数：<INPUT TYPE="text" NAME="Total" size="3" value="<%=Node.getAttribute("Total")%>">
</td><td class="td1">
标题长度：<INPUT TYPE="text" NAME="Topiclen" size="4" value="<%=Node.getAttribute("Topiclen")%>">
</td>
<td class="td1">
帖子排序：<SELECT NAME="Orders" ID="Orders">
	<option value="0" SELECTED>默认最新排序(推荐使用)</option>
	<option value="1">按照时间(按最新主题时间)</option>
	<option value="2">按照时间(按最新回复时间)</option>
	<option value="3">按照点击(最热帖)</option>
	</SELECT>
</td>
</tr>
<tr><td class="td1" colspan="3">
天数的限制：<INPUT TYPE="text" NAME="Sdate" value="<%=Node.getAttribute("Sdate")%>" size="3">(查询多少天内帖子，1为当天。若为空则日期不限，建议为空。)
</td></tr>
<tr><td class="td1" colspan="3">
显示的类型：<SELECT NAME="TopicType" ID="TopicType">
	<option value="0" SELECTED>显示主题</option>
	<option value="1">显示精华主题</option>
	<option value="2">显示主题和回复</option>
	</SELECT>
	(不推荐数据量大的用户使用调用主题和回复。)
</td>
</tr>
<tr><td class="td1" colspan="3">
调用的版面：<SELECT id="Boardid0" NAME="Boardid"></SELECT>
<BR>
版面&nbsp;&nbsp;设置：<SELECT NAME="BoardType" ID="BoardType">
	<option value="0" SELECTED>只显示该版面的数据</option>
	<option value="1">显示该版面的下级所有版面的数据</option>
	<option value="2">显示该版面和下级所有版面的数据</option>
	</SELECT>
<BR>版面的限制：<SELECT NAME="BoardLimit" ID="BoardLimit">
	<option value="0" SELECTED>显示所有数据</option>
	<option value="1">不显示特殊版面数据</option>
	</SELECT>（特殊版面指隐藏版面和认证版面）
</td>
</tr>
<tr><td class="td1" colspan="3">
单独用户ID：<INPUT TYPE="text" NAME="UserIDList" value="<%=Node.getAttribute("UserIDList")%>">(请填写用户会员ID,用英文逗号分隔)
</td>
</tr>
</table>
<SCRIPT LANGUAGE="JavaScript">
<!--
BoardJumpListSelect('<%=Boardid%>',"Boardid0","选取所有版面","",0);
//-->
</SCRIPT>
</div>
<div id="News_2" style="display:none">
<!-- 信息调用 -->
<table border="0" cellpadding="3" cellspacing="1" width="100%">
<tr>
<td></td>
</tr>
</table>
</div>
<div id="News_3" style="display:none">
<!-- 版块调用 -->
<table border="0" cellpadding="3" cellspacing="1" width="100%">
<tr>
<td class="td1">
显示模式：<SELECT NAME="Orders" ID="Orders">
	<option value="0"<%If Node.getAttribute("Orders") = 0 Then Response.Write " SELECTED"%>>树型结构</option>
	<option value="1"<%If Node.getAttribute("Orders") = 1 Then Response.Write " SELECTED"%>>地图结构</option>
	</SELECT>
</td>
<td class="td1">
<input type="text" name="BoardTab" value="<%=Node.getAttribute("BoardTab")%>" size="2">(地图结构时，限制每行显示数量)
</td>
</tr>
<tr>
<td class="td1" colspan="2">
限制调用版块的层数：<input type="text" name="Depth" size="2" value="<%=Node.getAttribute("Depth")%>"><BR>(如0,表示只调用第一级分类
;为空则表示调用所有，当地图结构模式时，层数超过2无效;)
</td>
</tr>
<tr>
<td class="td1">
调用的版面：<SELECT id="Boardid1" NAME="Boardid"></SELECT>
</td>
<td class="td1">
<input type="radio" class="radio" name="Stats" value="0">显示所有版块
<input type="radio" class="radio" name="Stats" value="1" checked>不显示隐藏版块
</td>
</tr>
</table>
<SCRIPT LANGUAGE="JavaScript">
<!--
BoardJumpListSelect('<%=Boardid%>',"Boardid1","选取所有版面","",0);
//-->
</SCRIPT>
</div>
<div id="News_4" style="display:none">
<!-- 会员调用 -->
<table border="0" cellpadding="3" cellspacing="1" width="100%">
<tr>
<td class="td1">
显示记录数：<INPUT TYPE="text" NAME="Total" size="3" value="<%=Node.getAttribute("Total")%>">
</td>
<td class="td1">
会员排序：<SELECT NAME="UserOrders" ID="UserOrders">
	<option value="0" SELECTED>按注册时间</option>
	<option value="1">按用户文章</option>
	<option value="2">按用户主题</option>
	<option value="3">按用户精华</option>
	<option value="4">按用户金钱</option>
	<option value="5">按用户积分</option>
	<option value="6">按用户魅力</option>
	<option value="7">按用户被删帖数</option>
	<option value="8">按用户登陆次数</option>
	</SELECT>
</td>
</tr>
</table>
</div>
<div id="News_5" style="display:none">
<!-- 公告调用 -->
<table border="0" cellpadding="3" cellspacing="1" width="100%">
<tr>
<td class="td1">
显示记录数：<INPUT TYPE="text" NAME="Total" value="<%=Node.getAttribute("Total")%>" size="3">
</td><td class="td1">
标题长度：<INPUT TYPE="text" NAME="Topiclen" value="<%=Node.getAttribute("Topiclen")%>" size="4">
</td>
</tr>
<tr>
<td class="td1" colspan="2">
调用的版面：<SELECT id="Boardid2" NAME="Boardid"></SELECT>
</td>
</tr>
</table>
<SCRIPT LANGUAGE="JavaScript">
<!--
BoardJumpListSelect('<%=Boardid%>',"Boardid2","选取所有版面","",0);
//-->
</SCRIPT>
</div>
<div id="News_6" style="display:none">
<!-- 展区调用 -->
<table border="0" cellpadding="3" cellspacing="1" width="100%">
<tr>
<td>
显示记录数：<INPUT TYPE="text" NAME="Total" value="<%=Node.getAttribute("Total")%>" size="3">
&nbsp;&nbsp;&nbsp;&nbsp;
每行显示个数：<INPUT TYPE="text" NAME="Tab" value="<%=Node.getAttribute("Tab")%>" size="3">
&nbsp;&nbsp;&nbsp;&nbsp;标题长度：<INPUT TYPE="text" NAME="Topiclen" value="<%=Node.getAttribute("Topiclen")%>" size="4">
<br>
调用的版面：<SELECT id="Boardid3" NAME="Boardid"></SELECT>
版面限制设置：
	<SELECT NAME="BoardLock" ID="BoardLock">
	<option value="0">不限制</option>
	<option value="1">该版面不被调用</option>
	<option value="2">只调用该版面</option>
	<option value="3">该版面的下级版面</option>
	<option value="4">该版及下级所有版面</option>
	</SELECT>
<BR>版面的限制：<SELECT NAME="BoardLimit" ID="BoardLimit">
	<option value="0" SELECTED>显示所有数据</option>
	<option value="1">不显示特殊版面数据</option>
	</SELECT>（特殊版面指隐藏版面和认证版面）
<br>
调用文件类型 ： <SELECT NAME="FileType" ID="FileType">
	<option value="all" SELECTED>所有文件</option>
	<option value="0">文件集</option>
	<option value="1">图片集</option>
	<option value="2">FLASH集</option>
	<option value="3">音乐集</option>
	<option value="4">电影集</option>
	</SELECT>
<br>
显示排序：<SELECT NAME="FileOrders" ID="FileOrders">
	<option value="0" SELECTED>默认</option>
	<option value="1">按浏览次数</option>
	<option value="2">按下载次数</option>
	<option value="3">按文件大小</option>
	</SELECT>
</td>
</tr>
</table>
<SCRIPT LANGUAGE="JavaScript">
<!--
BoardJumpListSelect('<%=Boardid%>',"Boardid3","选取所有版面","",0);
//BoardJumpListSelect(<%=Boardid%>,"Boardid3","选取所有版面","",0);
//-->
</SCRIPT>
</div>
<div id="News_7" style="display:none">

</div>
<div id="News_8" style="display:none">
<!-- 信息调用 -->
<table border="0" cellpadding="3" cellspacing="1" width="100%">
<tr>
<td></td>
</tr>
</table>
</div>
<!-- 变量说明 -->
<div id="skininfo_0" style="display:none"></div>
<div id="skininfo_1" style="display:none">
	<ol>
	
	<li>标题：{$Topic}</li>
	<li>作者：{$UserName}</li>
	<li>发表时间：{$PostTime}</li>
	<li>回复者：{$ReplyName}</li>
	<li>回复时间：{$ReplyTime}</li>
	<li>版块名称：{$BoardName}</li>
	<li>版块说明：{$BoardInfo}</li>
	<li>心情图标：{$Face}</li>
	<li>帖子ID：{$ID}</li>
	<li>帖子FileType" ID="FileType">
	<option value="all" SELECTED>鎵€鏈夋枃浠