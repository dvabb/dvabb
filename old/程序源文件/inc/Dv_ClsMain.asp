<object runat="server" id="DvStream" progid="ADODB.Stream"></object>
<%
'=========================================================
' File: Dv_ClsMain.asp
' Version:8.3.0
' Date: 20010-2-10
' Script Written by dvbbs.net
'=========================================================
' Copyright (C) 2003,2004 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: eway@aspsky.net
'=========================================================
'是否商业版，非官方SQL版本请在此设置为0以及在Conn中设置论坛为SQL数据库，否则显示不正常
Const IsBuss=1
Const Dvbbs_Server_Url = "http://server.dvbbs.net/"
Const Dvbbs_PayTo_Url = "http://pay.dvbbs.net/"
Const fversion="8.3.0"
Dim IP_MAX
Const guestxml="<?xml version=""1.0"" encoding=""gb2312""?><xml><userinfo statuserid=""0"" userid=""0"" username=""客人"" userclass=""客人"" usergroupid=""7"" cometime="""" boardid=""0"" activetime="""" statusstr=""""/></xml>"
Class Cls_Forum
	Rem Const
	Public BoardID,SqlQueryNum,Forum_Info,Forum_Setting,Forum_user,Forum_Copyright,Forum_ads,Forum_ChanSetting,Forum_UploadSetting
	Public Forum_sn,Forum_Version,Stats,StyleName,ErrCodes,NowUseBBS,Cookiepath,ScriptFolder,BoardInfoData,UserSession
	Public MainSetting,sysmenu,UserToday,BoardJumpList,BoardList,CacheData,Maxonline
	Public VipGroupUser,Vipuser,Boardmaster,Superboardmaster,Master,FoundIsChallenge,FoundUser
	Public ScriptName,MemberName,MemberWord,MemberClass,UserHidden,UserID,UserTrueIP,UserPermission
	Public sendmsgnum,sendmsgid,sendmsguser,Page_Admin
	Public BadWords,rBadWord,Forum_emot,Forum_PostFace,Forum_UserFace,SkinID,Forum_PicUrl
	Private Forum_CSS,Main_Sid,Nowstats,CssID
	Public Reloadtime,CacheName,UserGroupID,Lastlogin,GroupSetting,FoundUserPer,UserGroupParent,UserGroupParentID,LastMsg
	Private LocalCacheName,IsTopTable,ShowErrType
	Public Board_Setting,LastPost,Board_user,BoardType,Board_Data,Sid,Boardreadme,BoardRootID,BoardParentID
	Private Is_Isapi_Rewrite,iArchiverUrl
	Public ModHtmlLinked,ArchiverUrl,ArchiverType
	Public Browser,version ,platform,IsSearch,Cls_IsSearch
	Public IsUserPermissionOnly,IsUserPermissionAll,ShowSQL,actforip,DvRegExp,DvRegExp1
	Public GroupName,ScriptPath,Forum_apis,TyClsGroup,TyReadOnly,TyClsGroupM	'fish
	Rem Const
	Function iCreateObject(str)
		'iis5创建对象方法Server.CreateObject(ObjectName);
		'iis6创建对象方法CreateObject(ObjectName);
		'默认为iis6，如果在iis5中使用，需要改为Server.CreateObject(str);
		Set iCreateObject=CreateObject(str)
	End Function

	Function CreateXmlDoc(str)
		Set CreateXmlDoc = iCreateObject(str)
		CreateXmlDoc.async=false
	End Function

	Public Function ReadTextFile(fileName)
		On Error Resume Next
			'Response.Write Server.MapPath(ScriptPath&fileName)
			DvStream.charset="gb2312"
			DvStream.Mode = 3
			DvStream.open()
			DvStream.LoadFromFile(Server.MapPath(ScriptPath&fileName))
			ReadTextFile=DvStream.ReadText
			DvStream.close()
		If Err Then
			err.Clear
			PageEnd()
			Response.Clear
			Response.Write ScriptPath&fileName & "文件不存在！请检查,或者恢复官方模板数据！"
			Response.End
		End If
	End Function

	Function writeToFile(fileName,Text)
		DvStream.charset="gb2312"
		DvStream.Mode = 3
		DvStream.open()
		DvStream.WriteText(Text)
		DvStream.SaveToFile Server.MapPath(ScriptPath&fileName),2
		DvStream.close()
	End Function

	Private Sub Class_Initialize()		
	End Sub

	public Sub PageInit()
		ScriptPath="./"
		Forum_sn="DvForum 8.3"'如果一个虚拟目录或站点开多个论坛，则每个要错开，不能定义同一个名称
		Forum_sn=Forum_sn & "_" & Request.servervariables("SERVER_NAME")
		CacheName="DvCache 8.3"'如果一个虚拟目录或站点开多个论坛，则每个要错开，不能定义同一个名称
		IsUserPermissionOnly = 0
		IsUserPermissionAll = 0
		ShowErrType = 0 '错误信息显示模式
		SqlQueryNum = 0
		Reloadtime=600
		IsTopTable = 0
		VipGroupUser = False:IsSearch=False:Cls_IsSearch=False
		Vipuser = False:Boardmaster = False
		Superboardmaster = False:Master = False:FoundIsChallenge = False:FoundUser = False
		BoardID = Request("BoardID")
		If IsNumeric(BoardID) = 0 or BoardID = "" Then BoardID = 0
		BoardID = Clng(BoardID)
		MemberName = checkStr(Trim(Request.Cookies(Forum_sn)("username")))
		MemberWord = checkStr(Trim(Request.Cookies(Forum_sn)("password")))
		UserHidden = Trim(Request.Cookies(Forum_sn)("userhidden"))
		UserID = Trim(Request.Cookies(Forum_sn)("UserID"))
		If IsNumeric(UserHidden) = 0 or Userhidden = "" Then UserHidden = 2
		If IsNumeric(UserID) = 0 Or UserID="" Then UserID=0
		UserID = Clng(UserID)
		UserTrueIP = getIP()
		IP_MAX=0
		Dim Tmpstr
		Tmpstr = Request.ServerVariables("PATH_INFO")
		Tmpstr = Split(Tmpstr,"/")
		ScriptName = Lcase(Tmpstr(UBound(Tmpstr)))
		ScriptFolder = Lcase(Tmpstr(UBound(Tmpstr)-1)) & "/"
		MemberClass = checkStr(Request.Cookies(Forum_sn)("userclass"))
		Page_Admin=False
		If InStr(ScriptName,"showerr")>0 Or InStr(ScriptName,"login")>0 Or InStr(ScriptName,"admin_")>0 Or InStr(ScriptName,"ajax")>0 Then Page_Admin=True
		sendmsgnum=0:sendmsgid=0:sendmsguser=""
		'模拟HTML部分开始
		Is_Isapi_Rewrite = 0
		If Is_Isapi_Rewrite = 0 Then ModHtmlLinked = "?"
		ArchiverType = 0

		If InStr(ScriptName,"indexhtml.asp") > 0 Then
			iArchiverUrl = Lcase(Request.ServerVariables("QUERY_STRING"))
			If iArchiverUrl <> "" Then
				ArchiverUrl = iArchiverUrl
				iArchiverUrl = Split(iArchiverUrl,"_")
				If iArchiverUrl(0) = "list" And Ubound(iArchiverUrl) = 5 Then
					If IsNumeric(iArchiverUrl(1)) Then
						ArchiverType = 1
						BoardID = Clng(iArchiverUrl(1))
					End If
				End If
			End If
		End If
	End Sub
	'isapi_write
	Public Function ArchiveHtml(Textstr)
		Str=Textstr
		If isUrlreWrite = 1 Then
			Dim Str,re,Matches,Match
			Set re=new RegExp
			re.IgnoreCase =True
			re.Global=True
			re.Pattern = "<a(.[^>]*)index\.asp\?boardid=(\d+)(&|&amp;)topicmode=(\d+)?(&|&amp;)list_type=([\d,]+)?(&|&amp;)page=(\d+)?"
			str = re.Replace(str,"<a$1index_$2_$4_$6_$8.html")
			re.Pattern = "<a(.[^>]*)index\.asp\?boardid=(\d+)(&|&amp;)action=(.[^&]*)?(&|&amp;)topicmode=(\d+)?(&|&amp;)list_type=([\d,]+)?(&|&amp;)page=(\d+)?"
			str = re.Replace(str,"<a$1index_$2_$4_$6_$8_$10.html")
			re.Pattern = "<a(.[^>]*)index\.asp\?boardid=(\d+)(&|&amp;)page=(\d+)?(&|&amp;)action=(.[^<>""\'\s]*)?"
			str = re.Replace(str,"<a$1index_$2_$4_$6.html")
			re.Pattern = "<a(.[^>]*)index\.asp\?boardid=(\d+)(&|&amp;)topicmode=(\d+)?"
			str = re.Replace(str,"<a$1index_$2_$4.html")
			re.Pattern = "<a(.[^>]*)index\.asp\?boardid=(\d+)(&|&amp;)page=(\d+)?"
			str = re.Replace(str,"<a$1index_$2_$4_.html")
			re.Pattern = "<a(.[^>]*)index\.asp\?boardid=(\d+)"
			str = re.Replace(str,"<a$1index_$2.html")
			re.Pattern = "<a(.[^>|_]*)index\.asp"
			str = re.Replace(str,"<a$1index.html")
			re.Pattern = "<a(.[^>]*)dispbbs\.asp\?boardid=(\d+)(&|&amp;)replyid=(\d+)?(&|&amp;)id=(\d+)?(&|&amp;)skin=(\d+)?(&|&amp;)page=(\d+)?(&|&amp;)star=(\d+)?"
			str = re.Replace(str,"<a$1dispbbs_$2_$4_$6_skin$8_$10_$12.html")
			re.Pattern = "<a(.[^>]*)dispbbs\.asp\?boardid=(\d+)(&|&amp;)replyid=(\d+)?(&|&amp;)id=(\d+)?(&|&amp;)skin=(\d+)?(&|&amp;)star=(\d+)?"
			str = re.Replace(str,"<a$1dispbbs_$2_$4_$6_skin$8_$10.html")
			re.Pattern = "<a(.[^>]*)dispbbs\.asp\?boardid=(\d+)(&|&amp;)replyid=(\d+)?(&|&amp;)id=(\d+)?(&|&amp;)skin=(\d+)?"
			str = re.Replace(str,"<a$1dispbbs_$2_$4_$6_skin$8.html")
			re.Pattern = "<a(.[^>]*)dispbbs\.asp\?boardid=(\d+)(&|&amp;)id=(\d+)?(&|&amp;)authorid=(\d+)?(&|&amp;)page=(\d+)?(&|&amp;)(star|move)=([\w\d]+)?"
			str = re.Replace(str,"<a$1dispbbs_$2_$4_$6_$8_$11.html")
			re.Pattern = "<a(.[^>]*)dispbbs\.asp\?boardid=(\d+)(&|&amp;)id=(\d+)?(&|&amp;)page=(\d+)?(&|&amp;)(star|move)=([\w\d]+)?"
			str = re.Replace(str,"<a$1dispbbs_$2_$4_$6_$9.html")
			re.Pattern = "<a(.[^>]*)dispbbs\.asp\?boardid=(\d+)(&|&amp;)id=(\d+)?(&|&amp;)page=(\d+)?"
			str = re.Replace(str,"<a$1dispbbs_$2_$4_$6.html")
			re.Pattern = "<a(.[^>]*)dispbbs\.asp\?boardid=(\d+)(&|&amp;)id=(\d+)?"
			str = re.Replace(str,"<a$1dispbbs_$2_$4.html")
			re.Pattern = "<a(.[^>]*)dv_rss\.asp\?s=(.[^<|>|""|\'|\s]*)"
			str = re.Replace(str,"<a$1dv_rss_$2.html")
			re.Pattern = "<a(.[^>]*)dv_rss\.asp"
			str = re.Replace(str,"<a$1dv_rss.html")
			Set Re=Nothing
		End If
		ArchiveHtml = Str
	End Function

	Private Function getIP() 
		Dim strIPAddr 
		If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" OR InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then 
			strIPAddr = Request.ServerVariables("REMOTE_ADDR") 
		ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then 
			strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1) 
			actforip=Request.ServerVariables("REMOTE_ADDR")
		ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then 
			strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1)
			actforip=Request.ServerVariables("REMOTE_ADDR")
		Else 
			strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
			actforip=Request.ServerVariables("REMOTE_ADDR")
		End If 
		getIP = CheckStr(Trim(Mid(strIPAddr, 1, 30)))
	End Function 

	Private Sub class_terminate()		
	End Sub

	Public Sub PageEnd()
		If EnabledSession Then
			If Not UserSession Is Nothing  Then Session(CacheName & "UserID")= UserSession.xml
		End If
		Set UserSession=Nothing 
		If IsObject(Conn) Then Conn.Close : Set Conn = Nothing
		If IsObject(Plus_Conn) Then Plus_Conn.Close : Set Plus_Conn = Nothing
		Set DvRegExp= Nothing
		Set DvRegExp1= Nothing
		CacheData=Null
		Forum_Setting = Null
		Forum_UploadSetting = Null
		Forum_user =Null
		Forum_ChanSetting =Null
		BadWords = Null
		rBadWord = Null
		Forum_ads=Null
	End Sub

	Public Sub Sendmessanger(touserid,senduser,messangertext)
		Dim Node
		If Not IsObject( Application(Dvbbs.CacheName&"_messanger")) Then
			Set  Application(Dvbbs.CacheName&"_messanger")=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			 Application(Dvbbs.CacheName&"_messanger").appendChild( Application(Dvbbs.CacheName&"_messanger").createElement("xml"))
		End If
		For Each Node in Application(Dvbbs.CacheName&"_messanger").documentElement.SelectNodes("messanger")
			If datediff("s",Node.selectSingleNode("@sendtime").text,Now()) > 72000 Then
				Application(Dvbbs.CacheName&"_messanger").documentElement.removeChild(Node)
			End If
		Next
		Set Node=Application(Dvbbs.CacheName&"_messanger").documentElement.appendChild(Application(Dvbbs.CacheName&"_messanger").createNode(1,"messanger",""))
		Node.attributes.setNamedItem(Application(Dvbbs.CacheName&"_messanger").createNode(2,"sendtime","")).text=Now()
		Node.attributes.setNamedItem(Application(Dvbbs.CacheName&"_messanger").createNode(2,"touserid","")).text=touserid
		Node.attributes.setNamedItem(Application(Dvbbs.CacheName&"_messanger").createNode(2,"senduser","")).text=senduser
		Node.text=messangertext
	End Sub

	Public Property Let Name(ByVal vNewValue)
		LocalCacheName = LCase(vNewValue)
	End Property

	Public Property Let Value(ByVal vNewValue)
		If LocalCacheName<>"" Then
			Application.Lock
			Application(CacheName & "_" & LocalCacheName &"_-time")=Now()
			Application(CacheName & "_" & LocalCacheName) = vNewValue
			Application.unLock
		End If
	End Property

	Public Property Get Value()
		If LocalCacheName<>"" Then 	
				Value=Application(CacheName & "_" & LocalCacheName)
		End If
	End Property

	Public Function ObjIsEmpty()
		'Response.Write DateDiff("s",CDate(Application(CacheName & "_" & LocalCacheName &"_-time")),Now())&"秒"
		ObjIsEmpty=False
		If  IsDate(Application(CacheName & "_" & LocalCacheName &"_-time")) Then
			If DateDiff("s",CDate(Application(CacheName & "_" & LocalCacheName &"_-time")),Now()) > (60*Reloadtime) Then ObjIsEmpty=True
		Else
			ObjIsEmpty=True
		End If
		If ObjIsEmpty Then RemoveCache()
	End Function

	Public Sub RemoveCache()
		Application.Lock
		Application.Contents.Remove(CacheName & "_" & LocalCacheName)
		Application.Contents.Remove(CacheName & "_" & LocalCacheName &"_-time")
		Application.unLock
	End Sub
	'取得基本设置数据
	Public Sub loadSetup()
		Dim Rs,locklist,ip,ip1,XMLDom,Node,i
		Name="setup"
		Set Rs = Dvbbs.Execute("Select id, Forum_Setting, Forum_ads, Forum_Badwords, Forum_rBadword, Forum_Maxonline, Forum_MaxonlineDate, Forum_TopicNum, Forum_PostNum, Forum_TodayNum, Forum_UserNum, Forum_YesTerdayNum, Forum_MaxPostNum, Forum_MaxPostDate, Forum_lastUser, Forum_LastPost, Forum_BirthUser, Forum_Sid, Forum_Version, Forum_NowUseBBS, Forum_IsInstall, Forum_challengePassWord, Forum_Ad, Forum_ChanName, Forum_ChanSetting, Forum_LockIP, Forum_Cookiespath, Forum_Boards, Forum_alltopnum, Forum_pack, Forum_Cid, Forum_AvaSiteID, Forum_AvaSign, Forum_AdminFolder, Forum_BoardXML, Forum_Css, Forum_apis From [Dv_Setup]")
		Value = Rs.GetRows(1)
		CacheData=value
		Set XMLDom=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			XMLDom.appendChild(XMLDom.createElement("xml"))
			locklist=Trim(CacheData(25,0))
			locklist=Split(locklist,"|")
			For Each Ip in locklist
				Ip1=Split(Ip,".")
				Set Node=XMLDom.documentElement.appendChild(XMLDom.createNode(1,"lockip",""))
				For i=0 to UBound(ip1)
					Node.attributes.setNamedItem(XMLDom.createNode(2,"number"& (i+1),"")).text=ip1(i)
				Next
			Next
			Application.Lock
			Set Application(CacheName & "_forum_lockip")=XMLDom
			Application.UnLock
		Set XMLDom=Nothing
		If Not isobject(Application(CacheName & "_getbrowser")) Then
			Dim stylesheet
			Set stylesheet=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			stylesheet.load Server.MapPath(MyDbPath &"inc\GetBrowser.xslt")
			Application.Lock
			Set Application(CacheName & "_getbrowser")=Dvbbs.iCreateObject("msxml2.XSLTemplate" & MsxmlVersion)
			Application(CacheName & "_getbrowser").stylesheet=stylesheet
			Application.unLock
		End If
		Application.Lock
		Set Application(CacheName & "_accesstopic")=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		Application(CacheName & "_accesstopic").Loadxml Replace(Replace(CacheData(34,0),Chr(10),""),Chr(13),"")
		Application.unLock
	End Sub

	Public Sub LoadBbsBoard()		
	End Sub

	Public Sub LoadBoardList()
		Dim TempXmlDoc,TempMasterDoc,ChildNode
		Dim Rs,boardmaster,master,node,Board_setting
		Set Rs=Execute("select boardid,boardtype,ParentID,depth,rootid,Child,indeximg,parentstr,cid as checkout,cid as hidden,cid as nopost,cid as checklock,cid as mode,cid as simplenesscount,readme,boardmaster From Dv_board Order by rootid,Orders")
		Set TempXmlDoc = RecordsetToxml(rs,"board","BoardList")
		Rs.Close
		Set TempMasterDoc = Dvbbs.CreateXmlDoc("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		TempMasterDoc.documentElement = TempMasterDoc.createElement("masterlist")
		Set Rs=Execute("select boardmaster,boardid,Board_setting From Dv_board Order by Orders")
		Do While Not Rs.EOF
			Set Node = TempMasterDoc.documentElement.appendChild(TempMasterDoc.createNode(1,"boardmaster",""))
			Node.setAttribute "boardid",Rs(1)
			boardmaster=split(Rs("boardmaster")&"","|")
			For Each Master In boardmaster
				Node.appendChild(TempMasterDoc.createNode(1,"master","")).text=Master
			Next
			Board_setting=Split(Rs("Board_setting"),",")
			TempXmlDoc.documentElement.selectSingleNode("board[@boardid='"& Rs("Boardid")&"']/@checkout").text=Board_setting(2)
			TempXmlDoc.documentElement.selectSingleNode("board[@boardid='"& Rs("Boardid")&"']/@hidden").text=Board_setting(1)
			TempXmlDoc.documentElement.selectSingleNode("board[@boardid='"& Rs("Boardid")&"']/@nopost").text=Board_setting(43)
			TempXmlDoc.documentElement.selectSingleNode("board[@boardid='"& Rs("Boardid")&"']/@checklock").text=Board_setting(0)
			TempXmlDoc.documentElement.selectSingleNode("board[@boardid='"& Rs("Boardid")&"']/@mode").text=Board_setting(39)
			TempXmlDoc.documentElement.selectSingleNode("board[@boardid='"& Rs("Boardid")&"']/@simplenesscount").text=Board_setting(41)
		Rs.MoveNext
		Loop
		Rs.Close
		Set Rs= Nothing
		Application.Lock
		Set Application(CacheName&"_boardmaster") = TempMasterDoc
		Set Application(CacheName&"_boardlist") = TempXmlDoc
		Application(CacheName&"_boardmaster_xml") = TempMasterDoc.xml
		Application(CacheName&"_boardlist_xml") = TempXmlDoc.xml
		Application.unLock
	End Sub

	Public Sub LoadPlusMenu()
		Name = "ForumPlusMenu"
		Dim Rs,XMLDom,Node,plus_setting,stylesheet,XMLStyle,proc
		Set Rs=Execute("Select id,plus_type,plus_name,mainpage,plus_copyright,plus_setting,isshowmenu as width,isshowmenu as height From Dv_Plus Where  Isuse=1 Order By ID")
		Set XMLDom=RecordsetToxml(rs,"plus","")
		Rs.close()
		Set Rs=Nothing
		For Each Node In XMLDom.documentElement.selectNodes("plus")
			plus_setting=Split(Split(node.selectSingleNode("@plus_setting").text,"|||")(0),"|")
			node.selectSingleNode("@plus_setting").text=plus_setting(0)
			node.selectSingleNode("@width").text=plus_setting(1)
			node.selectSingleNode("@height").text=plus_setting(2)
		Next
		Set XMLStyle=Dvbbs.iCreateObject("msxml2.XSLTemplate" & MsxmlVersion)

		Set stylesheet=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		
		stylesheet.load Server.MapPath(MyDbPath &"inc\Templates\plusmenu.xslt")
		XMLStyle.stylesheet=stylesheet
		Set proc=XMLStyle.createProcessor()
		proc.input = XMLDom
  		proc.transform()
  		value=proc.output
	End Sub

	Public Sub LoadBoardData(bid)
		Dim Rs
		Application.Lock
		Set Rs=Execute("select boardid,boarduser,board_ads,board_user,isgroupsetting,rootid,board_setting,sid,cid,Rules From Dv_board Where Boardid="&bid)
		Set Application(CacheName &"_boarddata_" & bid)=RecordsetToxml(rs,"boarddata","")
		Rs.Close
		Set Rs= Nothing
		Application.unLock
	End Sub

	Public Sub LoadBoardinformation(bid)'加载动态板面信息数据
		Dim Rs,lastpost,i
		Application.Lock
		Set Rs=Execute("select boardid,boardtopstr,postnum,topicnum,todaynum,lastpost as lastpost_0 From Dv_board Where Boardid="&bid)
		Set Application(CacheName &"_information_" & bid)=RecordsetToxml(rs,"information","")
		lastpost=Split(Application(CacheName &"_information_" & bid).documentElement.selectSingleNode("information/@lastpost_0").text,"$")
		For i=0 to UBound(lastpost)
			Application(CacheName &"_information_" & bid).documentElement.firstChild.setAttribute "lastpost_"& i,lastpost(i)
			If i = 7 Then Exit For
		Next
		Rs.Close
		Set Rs= Nothing
		Application.unLock
	End Sub

	Public Sub LoadAllBoardinformation()'加载所有板面信息数据
		Dim Rs,lastpost,i
		Dim TempXmlDom,Node,TempNode,TempXmlDom1
		Set Rs=Execute("select boardid,boardtopstr,postnum,topicnum,todaynum,lastpost as lastpost_0 From Dv_board Order by Orders")
		Set TempXmlDom = RecordsetToxml(rs,"information","")
		Rs.Close
		Set Rs = Nothing
		For Each Node In TempXmlDom.documentElement.selectNodes("information")
			lastpost=Split(Node.getAttribute("lastpost_0"),"$")
			For i=0 to UBound(lastpost)
				Node.setAttribute "lastpost_"& i,lastpost(i)
				If i = 7 Then Exit For
			Next
			Set TempXmlDom1=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			Set TempNode = TempXmlDom1.appendChild(TempXmlDom1.createNode(1,"xml",""))
			TempNode.appendChild(Node)
			Application.Lock
			Set Application(CacheName &"_information_" & Node.getAttribute("boardid")) = TempXmlDom1
			Application.UnLock
		Next
		If IsObject(TempXmlDom1) Then Set TempXmlDom1 = Nothing
	End Sub

	Public Sub LoadGroupSetting()
		Dim Rs
		Set Rs=Dvbbs.Execute("Select GroupSetting,UserGroupID,ParentGID,IsSetting,UserTitle,TyClsGroup,TyClsGroupM From Dv_UserGroups")	'fish，新增用户组开关
		Set Application(CacheName &"_groupsetting")=RecordsetToxml(rs,"usergroup","")
		Set Rs=Dvbbs.Execute("Select UserGroupID,usertitle,titlepic,orders From Dv_UserGroups order by orders")
		Set Application(CacheName &"_grouppic")=RecordsetToxml(rs,"usergroup","grouppic")
		Rs.close()
		Set Rs=Nothing
	End Sub

	Public Sub Loadstyle()
		Dim Rs
		Application.Lock
		Set Rs=Dvbbs.Execute("Select *  From Dv_Templates")
		Set Application(CacheName &"_style")=RecordsetToxml(rs,"style","")
		Rs.close()
		Set Rs=Nothing
		Application.UnLock
		LoadStyleMenu()
	End Sub

	Public Sub LoadStyleMenu()'生成风格选单数据
		Name="style_list"
		Dim HTMLstr
		HTMLStr="<a href=""cookies.asp?action=stylemod&amp;boardid=$boardid"" >恢复默认设置</a>"
		Dim Node
		For Each Node in Application(CacheName &"_style").documentElement.selectNodes("style")
			HTMLstr=(HTMLstr&"<br /><a href=""cookies.asp?action=stylemod&amp;skinid="& node.selectSingleNode("@id").text & "&amp;boardid=$boardid"">"& node.selectSingleNode("@type").text& "</a>")
		Next
  	value=HTMLstr
	End Sub

	Public Sub UpdateForum_Info(act)'act=0 不处理缓存,act=1 处理缓存
		If value <> "1900-1-1" Then 
			value="1900-1-1"
			Dim Rs,LastPostInfo,TempStr,i,Board
			Dim Forum_YesterdayNum,Forum_TodayNum,Forum_LastPost,Forum_MaxPostNum,Forum_MaxPostDate
			Set Rs=Execute("Select Top 1 Forum_YesterdayNum,Forum_TodayNum,Forum_LastPost,Forum_MaxPostNum From Dv_Setup")
			Forum_YesterdayNum=Rs(0)
			Forum_TodayNum=Rs(1)
			Forum_LastPost=Rs(2)
			Forum_MaxPostNum=Rs(3)
			Rs.close()
			Set Rs=Nothing
			LastPostInfo = Split(Forum_LastPost,"$")
			If Not IsDate(LastPostInfo(2)) Then LastPostInfo(2)=Now()	
			If DateDiff("d",CDate(LastPostInfo(2)),Now())<>0 Then'最后发帖时间不是今天，	
				TempStr=LastPostInfo(0)&"$"&LastPostInfo(1)&"$"&Now()&"$"&LastPostInfo(3)&"$"&LastPostInfo(4)&"$"&LastPostInfo(5)&"$"&LastPostInfo(6)&"$"&LastPostInfo(7)
				Execute("Update Dv_Setup Set Forum_YesterdayNum="&Forum_TodayNum&",Forum_LastPost='"&TempStr&"',Forum_TodayNum=0")
				Execute("update Dv_board Set TodayNum=0")
				If act=1 Then
					If not IsObject(Application(CacheName&"_boardlist")) Then LoadBoardList()
					For Each board in Application(CacheName&"_boardlist").documentElement.selectNodes("board/@boardid")
						LoadBoardinformation board.text
					Next
				End If
			End If
			If Forum_TodayNum >Forum_MaxPostNum Then
				Execute("Update Dv_Setup Set Forum_MaxPostNum="&Forum_TodayNum&",Forum_MaxPostDate="&SqlNowString)
			End If
			If act=1 Then loadSetup()
			Dim xmlhttp
			If IsSqlDataBase =0 Then
				On Error Resume Next
				Set xmlhttp = Dvbbs.CreateXmlDoc("msxml2.ServerXMLHTTP")
				xmlhttp.setTimeouts 65000, 65000, 65000, 65000
		  	xmlhttp.Open "POST",Get_ScriptNameUrl& "Loadservoces.asp",false
		  	xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		  	xmlhttp.send()
		  	Set xmlhttp = Nothing
			End If
		End If
		Name="Date"
		value=Date()
	End Sub

	Public Sub GetForum_Setting()
		Name="Date"
		If ObjIsEmpty() Then
			UpdateForum_Info(0)
		ElseIf  Cstr(value) <> Cstr(Date()) Then 
			UpdateForum_Info(1)
		End If
		Name="setup"
		If ObjIsEmpty Then loadSetup()
		If Not IsObject(Application(CacheName&"_boardlist")) Then
				LoadBoardList()
		End If
		If Not IsObject(Application(CacheName &"_style")) Then
				Loadstyle()
		End If
		Name="setup"
		CacheData=value
		Forum_apis=CacheData(36,0)
		Dim Setting,OpenTime,ischeck:Setting= Split(CacheData(1,0),"|||"):Forum_Info =  Split (Setting(0),",")
		Forum_Setting = Split (Setting(1),","):Forum_UploadSetting = Split(Forum_Setting(7),"|")
		Forum_user = Setting(2):Forum_user = Split (Forum_user,","):Forum_Copyright = Setting(3)
		Forum_ChanSetting = Split(CacheData(24,0),",")
		Forum_Version = fversion 'CacheData(18,0)
		BadWords = Split(CacheData(3,0),"|")
		Set DvRegExp=new RegExp
		DvRegExp.IgnoreCase =True
		DvRegExp.Global=true
		Set DvRegExp1=new RegExp
		DvRegExp1.IgnoreCase =True
		DvRegExp1.Global=false
		if(CacheData(3,0)<>"") Then
			DvRegExp.Pattern="(" & CacheData(3,0) &")"
			DvRegExp1.Pattern="(" & CacheData(3,0) &")"
		End if
		rBadWord = Split(CacheData(4,0),"|"):	Main_Sid=CacheData(17,0):Maxonline = CacheData(5,0):NowUseBBS = CacheData(19,0):Cookiepath = CacheData(26,0)
		If ScriptFolder = Lcase(CacheData(33,0)) Then Page_Admin = True
		Rem 禁止代理服务器访问开始,如需要允许访问，请屏蔽此段代码。
		If Forum_Setting(100)="1" Then
			If actforip <> "" Then
				Session(CacheName & "UserID")=empty
				Set Dvbbs=Nothing
				Response.Status = "302 Object Moved" 
				Response.End 
			End If
			If UBound(Forum_Setting)> 101 Then
				IP_MAX=CLng(Forum_Setting(101))
			Else
				IP_MAX=0
			End If
		End If
		Rem 禁止代理服务器访问结束
		Rem Hantg 2007-12-05
		If UBound(Forum_Setting)<107 Then
			Redim Preserve Forum_Setting(106)
		End If
		If Forum_Setting(21)="1" And Not Page_Admin Then Set Dvbbs=Nothing:Response.redirect "showerr.asp?action=stop"	
		If BoardID <>0 Then
			If Application(CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"&BoardID&"']") Is Nothing Then
				Set Dvbbs=Nothing
				Response.Write "错误的版面参数"
  				Response.End
			End If
		End If
		If BoardID > 0 Then
			If Not IsObject(Application(CacheName &"_boarddata_" & Boardid)) Then LoadBoardData boardid
			If Not IsObject (Application(CacheName &"_information_" & boardid)) Then LoadBoardinformation BoardID
			Dim Nodelist,node
			Forum_ads = Split(Application(CacheName &"_boarddata_" & Boardid).documentElement.selectSingleNode("boarddata/@board_ads").text,"$")
			
			Forum_user = Split(Application(CacheName &"_boarddata_" & Boardid).documentElement.selectSingleNode("boarddata/@board_user").text,",")
			board_Setting = Split(Application(CacheName &"_boarddata_" & Boardid).documentElement.selectSingleNode("boarddata/@board_setting").text,",")
			BoardType = Application(CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"&BoardID&"']/@boardtype").text
			BoardRootID = Application(CacheName &"_boarddata_" & Boardid).documentElement.selectSingleNode("boarddata/@rootid").text
			BoardParentID=CLng(Application(CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"&BoardID&"']/@parentid").text)	
			Sid = Application(CacheName &"_boarddata_" & Boardid).documentElement.selectSingleNode("boarddata/@sid").text
			Boardreadme=Application(CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"&BoardID&"']/@readme").text
			If Len(Board_Setting(22))< 24 Then Board_Setting(22)="1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1"
			OpenTime=Split(Board_Setting(22),"|")
			setting=Board_Setting(21)
			ischeck=Clng(Board_Setting(18))
			If Board_Setting(50)<>"0" And Board_Setting(50)<>"" Then Set Dvbbs=Nothing:Response.Redirect Board_Setting(50)
		Else
			Forum_ads = Split(CacheData(2,0),"$")
			If Len(Forum_Setting(70))< 24 Then Forum_Setting(70)="1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1"
			OpenTime=Split(Forum_Setting(70),"|")
			setting=Forum_Setting(69)
			ischeck=Forum_Setting(26)
			If Not IsNumeric(ischeck) Then ischeck=0
			ischeck=CLng(ischeck)		
		End If
		'定时开放判断
		If Not Page_Admin And Cint(setting)=1 Then
			If OpenTime(Hour(Now))="1" Then Response.redirect "showerr.asp?action=stop&boardid="&Dvbbs.BoardID&""
		End If
		'在线人数限制
		If ischeck > 0 And Not Page_Admin Then
			If MyBoardOnline.Forum_Online > ischeck And BoardID=0 Then
				If Not IsONline(Membername,1) Then Set Dvbbs=Nothing:Response.Redirect "showerr.asp?action=limitedonline&lnum="&ischeck
			End If
			If BoardID > 0 Then
				If (Not IsONline(Membername,1)) And MyBoardOnline.Board_Online > ischeck Then Set Dvbbs=Nothing:Response.Redirect "showerr.asp?action=limitedonline&lnum="&ischeck
			End If
		End If
		Dim CookiesSid
		CookiesSid = Request.Cookies("skin")("skinid_"&BoardID)
		
		If  CookiesSid = "" Then
			If BoardID = 0 Then 
				SkinID = Main_Sid
			Else
				SkinID = Sid
			End If
		Else
			SkinID = CookiesSid
		End If
		Setting=empty
	End Sub

	Public Function IsReadonly()
		IsReadonly=False
		Dim TimeSetting
		If Forum_Setting(69)="2" Then
			TimeSetting=split(Forum_Setting(70),"|")
			If TimeSetting(Hour(Now))="1" Then
				IsReadonly=True
				Exit Function
			End If
		End If
		If BoardID>0 Then 
			If Board_Setting(21)="2" Then
				TimeSetting=split(Board_Setting(22),"|")
				If TimeSetting(Hour(Now))="1" Then IsReadonly=True
			End If
		End If 
	End Function

	Public Function IsONline(UserName,action)
		IsONline=False
		If Trim(UserName)="" Then Exit Function
		If IsObject(Session(CacheName & "UserID")) And action=1 Then
				IsONline=True:Exit Function 
		End If
		Dim Rs
		Set Rs =Execute("Select UserID From Dv_Online Where Username='"&UserName&"'")
		If Not Rs.EOF  Then IsONline=True
		Rs.close()
		Set rs=Nothing  
	End Function  

	Public Sub LoadTemplates(Page_Fields)
		Dim Style_Pic,Main_Style,TempStyle,cssfilepath
		Forum_PicUrl="skins/Default/"
		Template.TPLname=Page_Fields
		Dim Node
		Set Node=Application(CacheName &"_style").documentElement.selectSingleNode("style[@id='"& Skinid &"']")
		If(Node is Nothing) Then
			Set Node=Application(CacheName &"_style").documentElement.selectSingleNode("style")
			If (Node is Nothing) Then
				Response.Write "没有注册可用风格模板，请到后台设置"
				Response.End
			Else
				Skinid=node.selectSingleNode("@id").text
			End if
		End If
		Template.ChildFolder=CheckStr(node.getAttribute("folder"))
		If Template.ChildFolder="" Then Template.ChildFolder="template_1"

		if (lcase(Page_Fields)="index" or lcase(Page_Fields)="dispbbs" or lcase(Page_Fields)="showerr") Then 
			Template.Cache=True '
		End if
		If Not (Instr(ScriptName,"index")>0 Or Page_Admin) Then
			Dim FolderPath
			If MyDbPath = "../../../" Then
				FolderPath = ""
			Else
				FolderPath = MyDbPath
			End If
			Style_Pic = Dvbbs.ReadTextFile(FolderPath & Template.Folder &"\"& Template.ChildFolder &"\pub_FaceEmot.htm")
			'Style_Pic = Dvbbs.ReadTextFile(Template.Folder &"\"& Template.ChildFolder &"\pub_FaceEmot.htm")
			Style_Pic = Split(Style_Pic,"@@@")
			Forum_UserFace = Style_Pic(0)
			Forum_PostFace = Style_Pic(1)
			Forum_Emot = Style_Pic(2)
		End If
		MainSetting = Split(MainHtml(0),"||")
		if ubound(MainSetting) >12 Then Forum_PicUrl=MainSetting(13)
	End Sub

	Public Function MainPic(Index)
		Dim Tplname,Cache
		Tplname=Template.TPLname
		Cache=Template.Cache
		Template.TPLname="pub"
		Template.Cache=true
		MainPic=Template.Pic(Index)
		Template.TPLname=Tplname
		Template.Cache=Cache
	End Function

	Public Function LanStr(Index)
		Dim Tplname,Cache
		Tplname=Template.TPLname
		Cache=Template.Cache
		Template.TPLname="pub"
		Template.Cache=true
		LanStr=Template.Strings(Index)
		Template.TPLname=Tplname
		Template.Cache=Cache
	End Function

	Public Function MainHtml(Index)
		Dim Tplname,Cache
		Tplname=Template.TPLname
		Cache=Template.Cache
		Template.TPLname="pub"
		Template.Cache=True
		MainHtml=Template.HTML(Index)
		Template.TPLname=Tplname
		Template.Cache=Cache
	End Function

	Rem 判断发言是否来自外部
	Public Function ChkPost()
		Dim server_v1,server_v2
		Chkpost=False 
		server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
		server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
		If Mid(server_v1,8,len(server_v2))=server_v2 Then Chkpost=True 
	End Function

	Public Sub ReloadSetupCache(MyValue,N)'更新总设置表部分缓存数组，入口：更新内容、数组位置
		CacheData(N,0) = MyValue
		Name="setup"
		value=CacheData
	End Sub

	Public Sub NeedUpdateList(username,act)'更新用户资料缓存(缓存用户名,是否需要添加)[0=不添加,只作清理,1=需要添加]
		Dim Tmpstr,TmpUsername
		Name="NeedToUpdate"
		If ObjIsEmpty() Then Value=""
		Tmpstr=Value
		TmpUsername=","&username&","
		Tmpstr=Replace(Tmpstr,TmpUsername,",")
		Tmpstr=Replace(Tmpstr,",,",",")
		If act=1 Then 
			If IsONline(username,0) Then
				If Tmpstr="" Then
					Tmpstr=TmpUsername
				Else
					Tmpstr=Tmpstr&TmpUsername
				End If
			End If
		End If
		Tmpstr=Replace(Tmpstr,",,",",")
		Value=Tmpstr
	End Sub

	Public Sub LetGuestSession()'写入客人session
		Dim StatUserID,UserSessionID
		StatUserID = checkStr(Trim(Request.Cookies(Forum_sn)("StatUserID")))
		If IsNumeric(StatUserID) = 0 or StatUserID = "" Then
			StatUserID = Replace(UserTrueIP,".","")
			UserSessionID = Replace(Startime,".","")
			If IsNumeric(StatUserID) = 0 or StatUserID = "" Then StatUserID = 0
			StatUserID = Ccur(StatUserID) + Ccur(UserSessionID)
		End If
		StatUserID = Ccur(StatUserID)
		Response.Cookies(Forum_sn).Expires=DateAdd("s",3600,Now())
		Response.Cookies(Forum_sn).path=cookiepath
		Response.Cookies(Forum_sn)("StatUserID") = StatUserID
		Set UserSession=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		UserSession.Loadxml guestxml
		UserSession.documentElement.se