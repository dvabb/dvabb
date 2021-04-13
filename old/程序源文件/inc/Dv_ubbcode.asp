<%
'最后修改2007.4.24
'最后修改2008.3.6 增加文件下载信息显示
Dim ServerHttp :ServerHttp = Dvbbs.Get_ScriptNameUrl
Dim UserPointInfo(4)
Dim FileInfo
FileInfo = 1	'是否显示下载文件的相关信息，不显示请设置为0 by 雨·漫步 2008.3.6 鸣谢：唧唧.NET hxyman
If CInt(Request.Form("ajaxPost")) Then
	FileInfo = 0
End If
Const NOScript=1        Rem 是否开启脚本过滤功能
Const NOWrongHTML=0     Rem 是否开启过滤错误HTML标记功能
Const MaxLoopcount=100	Rem UBB代码勘套循环的最多次数，避免死循环加入此变量
Const Issupport=1		Rem 部分服务器vbscript可能不支持SubMatches集合，请设置 Issupport=0
Const Maxsize=4			Rem 签名最大字体值
Const can_Post_Style="1,2,3"	Rem can_Post_Style是不限制style使用的用户组别列表，你可以根据自己需要修改
Dim Mtinfo
Mtinfo="<fieldset style=""border : 1px dotted #ccc;text-align : left;line-height:22px;text-indent:10px""><legend><b>媒体文件信息</b></legend><div>文件来源：$4</div>"&_
"<div>您可以点击控件上的播放按钮在线播放。注意，播放此媒体文件存在一些风险。</div>"&_
"<div>附加说明：动网论坛系统禁止了该文件的自动播放功能。</div>"&_
"<div>由于该用户没有发表自动播放多媒体文件的权限或者该版面被设置成不支持多媒体播放。</div></fieldset>"
Const DV_UBB_TITLE=" title=""dvubb"" "
Const UBB_TITLE="dvubb"

%>
<script language=vbscript runat=server>
Dim Ubblists
'[/img]编号:1.[/upload]编号:2.[/dir]编号:3.[/qt]编号:4.[/mp]编号:5.
'[/rm]编号:6.[/sound]编号:7.[/flash]编号:8.[/money]编号:9.[/point]编号:10.
'[/usercp]编号:11.[/power]编号:12.[/post]编号:13.[/replyview]编号:14.[/usemoney]编号:15.
'[/url]编号:16.[/email]编号:17.http编号:18.https编号:19.ftp编号:20.rtsp编号:21.
'mms编号:22.[/html]编号:23.[/code]编号:24.[/color]编号:25.[/face]编号:26.[/align]编号:27.
'[/quote]编号:28.[/fly]编号:29.[/move]编号:30.[/shadow]编号:31.[/glow]编号:32.[/size]编号:33.
'[/i]编号:34.[/b]编号:35.[/u]编号:36.[em编号:37.www.编号:38.[/payto]编号:40.[/username]编号:41.[/center]编号:42.

Class Dvbbs_UbbCode
	Public Re,reed,isgetreed,Board_Setting,WapPushUrl,xml,isxhtml,pageReload
	Public UpFileInfoScript,UpFileCount'UpFileInfoScript用来保存显示文件相关信息的脚本，UpFileCount用来保存整个页面中文件的数量 2008.3.6
	Public ismanager1
	Public Property Let PostType(ByVal vNewvalue)
		If PostType=2 Then
			Board_Setting=Split("1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1",",")
			Board_Setting(6)=1
			Board_Setting(5)=0:Board_Setting(7)=1
			Board_Setting(8)=1:Board_Setting(9)=1
			Board_Setting(10)=0:Board_Setting(11)=0
			Board_Setting(12)=0:Board_Setting(13)=0
			Board_Setting(14)=0:Board_Setting(15)=0
			Board_Setting(23)=0:Board_Setting(44)=0
		Else
			If Dvbbs.BoardID >0 Then
				Board_Setting=Dvbbs.Board_Setting
			Else
				Board_Setting=Split("1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1",",")
				Board_Setting(6)=1
				Rem 如果需要短信 支持 URL 自动识别  修改 Board_Setting(5)=1
				Board_Setting(5)=0:Board_Setting(7)=1
				Board_Setting(8)=1:Board_Setting(9)=1
				Board_Setting(10)=0:Board_Setting(11)=0
				Board_Setting(12)=0:Board_Setting(13)=0
				Board_Setting(14)=0:Board_Setting(15)=0
				Board_Setting(23)=0:Board_Setting(44)=0
			End If
		End If
	End Property
	Private Sub Class_Initialize()
		Set re=new RegExp
		re.IgnoreCase =true
		re.Global=true
		Set xml=Dvbbs.iCreateObject("msxml2.DOMDocument"& MsxmlVersion)
		If Dvbbs.UserID=0 Then
			UserPointInfo(0)=0:UserPointInfo(1)=0:UserPointInfo(2)=0:UserPointInfo(3)=0:UserPointInfo(4)=0
		Else
			UserPointInfo(0)=CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text)
			UserPointInfo(1)=CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text)
			UserPointInfo(2)=CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text)
			UserPointInfo(3)=CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpower").text)
			UserPointInfo(4)=CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpost").text)
		End If
	End Sub
	Private Sub class_terminate()
		Set xml=Nothing
		Set Re=Nothing
	End Sub
	Function istext(Str)
			Dim text,text1
			text=Str
			text1=Str
		 If text1=Dvbbs.Replacehtml(text) Then
		 	istext=True
		 End If
	End Function
	Function TextFormat(Str)
		Dim tmp,i
		Str=replace(Str,Chr(13)& Chr(10),Chr(13))
		Str=replace(Str,Chr(10),Chr(13))
		TMP=Split(Str,Chr(13))
		Str=""
		For i=0 to UBound(tmp)
			If i=UBound(tmp) Then
				Str=Str & tmp(i)
			Else
				Str=Str & tmp(i) &"<br />"
			End If
		Next
		TextFormat=Str
	End Function
	Rem 处理老DHTML贴子
	Public Function Dv_UbbCode_DHTML(s,PostUserGroup,PostType,sType)
		Dim matches,match,CodeStr
        If InStr(Ubblists,",39,")>0 And (InStr(Ubblists,",table,")>0 Or InStr(Ubblists,",td,")>0 Or InStr(Ubblists,",th,")>0 Or InStr(Ubblists,",tr,")>0 ) And NOWrongHTML = 1 Then
                s = server.htmlencode(s)
                s="<form name=""scode"&replyid_a&""" method=""post"" action=""""><table class=""tableborder2"" cellspacing=""1"" cellpadding=""3"" width=""100%"" align=""center"" border=""0""><tr><th height=""22"">以下内容含错误标记</th></tr><tr><td class=""tablebody1"" align=""middle"" width=""98%""><textarea id=""CodeText"" style=""BORDER-RIGHT: 1px dotted; BORDER-TOP: 1px dotted; OVERFLOW-Y: visible; OVERFLOW: visible; BORDER-LEFT: 1px dotted; WIDth: 98%; COLOR: #000000; BORDER-BOTTOM: 1px dotted"" rows=""20"" cols=""120"">"&s&"</textarea></td></tr><tr><td class=""tablebody2"" align=""middle"" width=""98%""></td></tr></table></form>"
                Dv_UbbCode_DHTML=s
                Exit Function
        Else
            If Board_Setting(5)="0" Then
                    re.Pattern ="<(\/?(i|b|p))>"
                    s=re.Replace(s,Chr(1)&"$1"&Chr(2))
                    re.Pattern="(>)("&vbNewLine&"){1,2}(<)"
                    s=re.Replace(s,"$1$3")
                    re.Pattern="(<div class=""quote"">)((.|\n)*?)(<\/div>)"
                    Do While re.Test(s)
                        s=re.Replace(s,"[quote]$2[/quote]")
                    Loop
                    re.Pattern = "(<\/tr>)"
                    s = re.Replace(s,"[br]")
                    re.Pattern = "(<br/>)"
                    s = re.Replace(s,"[br]")
                    re.Pattern = "(<br>)"
                    s = re.Replace(s,"[br]")
                    re.Pattern = "<(\/?s(ub|up|trike))>"
                    s = re.Replace(s,"[$1]")
                    re.Pattern = "(<)(\/?font[^>]*)(>)"
                    s = re.Replace(s,CHR(1)&"$2"&CHR(2))
                    re.Pattern="<([^<>]*?)>"
                    Do while re.Test(s)
                        s=re.Replace(s,"")
                    Loop
                    re.Pattern = "(\x01)(\/?font[^\x02]*)(\x02)"
                    s = re.Replace(s,"<$2>")
                    re.Pattern = "\[(\/?s(ub|up|trike))\]"
                    s = re.Replace(s,"<$1>")
                    re.Pattern="(\[quote\])((.|\n)*?)(\[\/quote\])"
                    Do While re.Test(s)
                        s=re.Replace(s,"<div class=""quote"">$2</div>")
                    Loop
                    re.Pattern="\x01(\/?(i|b|p))\x02"
                    s=re.Replace(s,"<$1>")
                    re.Pattern = "(\[br\])"
                    s = re.Replace(s,"<br/>")
                End If
                re.Pattern="<((asp|\!|%))"
                s=re.Replace(s,"&lt;$1")
        End If
        Dv_UbbCode_DHTML=s
	End Function
	'论坛内容部分UBBCODE，入口：内容、用户组ID、模式(1=帖子/2=公告、短信等)、模式2(0=新版/1=老版)
	Public Function Dv_UbbCode(s,PostUserGroup,PostType,sType)
		Dim mt,i,tmp
		If FileInfo Then
			Rem 防止标签名字被占用,如果要强制替换请去掉IF,2008.3.6.
			re.Pattern = "\<scr"&"ipt[\s\S]*\<\/scri"&"pt\>"
			s = re.Replace(s,"")
			s = Replace(s,"UpFileSize","UpFileSize.")
			s = Replace(s,"LoadTime","LoadTime.")
		End If

		'If Not xml.loadxml("<div>" & replace(s,"&","&amp;") &"</div>") Then
            'If NOScript = 1 Then
                'If Dv_FilterJS(s) Then
                    're.Pattern = "(&nbsp;)"
                    's = re.Replace(s,Chr(9))
                    're.Pattern = "(<br/>)"
                    's = re.Replace(s,vbNewLine)
                    're.Pattern = "(<br>)"
                    's = re.Replace(s,vbNewLine)
                    're.Pattern = "(<p>)"
                    's = re.Replace(s,"")
                    're.Pattern = "(<\/p>)"
                    's = re.Replace(s,vbNewLine)
                    's=server.htmlencode(s)
                    's="<form name=""scode"&replyid_a&""" method=""post"" action=""""><table class=""tableborder2"" cellspacing=""1"" cellpadding=""3"" width=""100%"" align=""center"" border=""0""><tr><th height=""22"">以下内容含脚本,或可能导致页面不正常的代码</th></tr><tr><td class=""tablebody1"" align=""middle"" width=""98%""><textarea id=""CodeText"" style=""BORDER-RIGHT: 1px dotted; BORDER-TOP: 1px dotted; OVERFLOW-Y: visible; OVERFLOW: visible; BORDER-LEFT: 1px dotted; WIDth: 98%; COLOR: #000000; BORDER-BOTTOM: 1px dotted"" rows=""20"" cols=""120"">"&s&"</textarea></td></tr><tr><td class=""tablebody2"" align=""middle"" width=""98%""><b>说明：</b>上面显示的是代码内容。您可以先检查过代码没问题，或修改之后再运行.</td></tr><tr><td class=""tablebody1"" align=""middle"" width=""98%""><input type=""button"" name=""run"" value=""运行代码"" onclick=""Dvbbs_ViewCode("&replyid_a&");""></td></tr></table></form>"
                    'Dv_UbbCode=s
                    'Exit Function
                'End If
            'End If
		'End If
		mt=canusemt(PostUserGroup)
		re.Pattern = "(\[br\])"
		s = re.Replace(s,"<br />")
		'Ubb转换
		'[img]图片标签
		If InStr(Ubblists,",1,")>0 Or sType=1 Then
				s=Dv_UbbCode_iS2(s,"img",_
				"<a href=""$1"" target=""_blank"" ><img "& DV_UBB_TITLE &" src=""$1"" border=""0"" /></a>",_
				"<img  "& DV_UBB_TITLE &" src=""skins/default/filetype/gif.gif"" border=""0"" alt="" /><a  href=""$1"" target=""_blank"" >$1</a>",_
				PostUserGroup,Cint(Board_Setting(7)),_
				"")
		End If
		'upload code
		If InStr(Ubblists,",2,")>0 Or sType=1 Then
			s=Dv_UbbCode_U(s,PostUserGroup,Cint(Board_Setting(7)))
		End If
		'media code
		If InStr(Ubblists,",3,")>0 Or sType=1 Then
			s=Dv_UbbCode_iS2(s,"DIR",_
			"<object "& DV_UBB_TITLE &" classid=""clsid:166B1BCA-3F9C-11CF-8075-444553540000"" "&_
			"codebase=""http://download.macromedia.com/pub/shockwave/cabs/director/sw.cab#version=7,0,2,0"" "&_
			"width=""$1"" height=""$2""><param name=""src"" value=""$3"" /><embed "& DV_UBB_TITLE &" src=""$3"""&_
			" pluginspage=""http://www.macromedia.com/shockwave/download/"" width=""$1"" height=""$2""></embed></object>",_
			"<a href=""$3"" target=""_blank"">$3</a>",_
			PostUserGroup,Cint(Board_Setting(9) * mt),_
			"=*([0-9]*),*([0-9]*)")
		End If
		'qt
		If InStr(Ubblists,",4,")>0 Or sType=1 Then
			s=Dv_UbbCode_iS2(s,"QT",_
			"<embed "& DV_UBB_TITLE &" src=""$3"" width=""$1"" height=""$2"" autoplay=""true"" loop=""false"" controller=""true"" playeveryframe=""false"" cache=""false"" scale=""TOFIT"" bgcolor=""#000000"" kioskmode=""false"" targetcache=""false"" pluginspage=""http://www.apple.com/quicktime/"" />",_
			"<embed "& DV_UBB_TITLE &" src=""$3"" width=""$1"" height=""$2"" autoplay=""false"" loop=""false"" controller=""true"" playeveryframe=""false"" cache=""false"" scale=""TOFIT"" bgcolor=""#000000"" kioskmode=""false"" targetcache=""false"" pluginspage=""http://www.apple.com/quicktime/"" />"&_
			 replace(Mtinfo,"$4","$3"),_
			PostUserGroup,Cint(Board_Setting(9) * mt),_
			"=*([0-9]*),*([0-9]*)")
		End If
		'mp
		If InStr(Ubblists,",5,")>0 Or sType=1 Then
			s=Dv_UbbCode_iS2(s,"mp",_
			"<object "& DV_UBB_TITLE &" align=""middle"" classid=""CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95"" class=""object"" id=""MediaPlayer"" width=""$1"" height=""$2"" >"&_
			"<param name=""ShowStatusBar"" value=""-1"" /><param name=""Filename"" value=""$3"" />"&_
			"<embed "& DV_UBB_TITLE &" type=""application/x-oleobject"" "&_
			"codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701"" flename=""mp"" src=""$3"" width=""$1"" height=""$2""></embed></object>",_
			"<object "& DV_UBB_TITLE &" align=""middle"" classid=""CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95"" class=""object"" id=""MediaPlayer"" width=""$1"" height=""$2"" >"&_
			"<param name=""ShowStatusBar"" value=""-1"" /><param name=""Filename"" value=""$3"" /><param name=""AUTOSTART"" value=""false"" />"&_
			"<embed "& DV_UBB_TITLE &" type=""application/x-oleobject"" "&_
			"codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701"" flename=""mp"" src=""$3"" width=""$1"" height=""$2""></embed></object>"&_
			replace(Mtinfo,"$4","$3"),_
			PostUserGroup,Cint(Board_Setting(9) * mt),"=*([0-9]*),*([0-9]*)")
			'Dv7 MediaPlayer自定义播放模式；
			s=Dv_UbbCode_iS2(s,"mp",_
			"<object "& DV_UBB_TITLE &" align=""middle"" classid=""CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95"" class=""object"" id=""MediaPlayer"" width=""$1"" height=""$2"" >"&_
			"<param name=""AUTOSTART"" value=""$3"" /><param name=""ShowStatusBar"" value=""-1"" /><param name=""Filename"" value=""$4"" />"&_
			"<embed "& DV_UBB_TITLE &" type=""application/x-oleobject"" codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701"" flename=""mp"" src=""$4"" width=""$1"" height=""$2""></embed></object>",_
			"<object "& DV_UBB_TITLE &" align=""middle"" classid=""CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95"" class=""object"" id=""MediaPlayer"" width=""$1"" height=""$2"" >"&_
			"<param name=""AUTOSTART"" value=""false"" /><param name=""ShowStatusBar"" value=""-1"" /><param name=""Filename"" value=""$4"" />"&_
			"<embed "& DV_UBB_TITLE &" type=""application/x-oleobject"" codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701"" flename=""mp"" src=""$4"" width=""$1"" height=""$2""></embed></object>"&_
			 Mtinfo,PostUserGroup,Cint(Board_Setting(9) * mt),"=*([0-9]*),*([0-9]*),*([0|1|true|false]*)")
		End If
		'rm
		If InStr(Ubblists,",6,")>0 Or sType=1 Then
			s=Dv_UbbCode_iS2(s,"rm",_
			"<div><object "& DV_UBB_TITLE &" classid=""clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA"" class=""object"" id=""RAOCX"" width=""$1"" height=""$2"">"&_
			"<param name=""src"" value=""$3"" />"&_
			"<param name=""CONSOLE"" value=""Clip1"" />"&_
			"<param name=""CONtrOLS"" value=""imagewindow"" />"&_
			"<param name=""AUTOSTART"" value=""true"" /></object></div>"&_
			"<div><object "& DV_UBB_TITLE &" classid=""CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA"" height=""32"" id=""video2"" width=""$1"">"&_
			"<param name=""src"" value=""$3"" /><param name=""AUTOSTART"" value=""-1"" />"&_
			"<param name=""CONtrOLS"" value=""controlpanel"" />"&_
			"<param name=""CONSOLE"" value=""Clip1"" /></object></div>",_
			"<div><object "& DV_UBB_TITLE &" classid=""clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA"" class=""object"" id=""RAOCX"" width=""$1"" height=""$2"">"&_
			"<param name=""src"" value=""$3"" />"&_
			"<param name=""CONSOLE"" value=""Clip1"" />"&_
			"<param name=""CONtrOLS"" value=""imagewindow"" />"&_
			"<param name=""AUTOSTART"" value=""false"" /></object></div>"&_
			"<div><object "& DV_UBB_TITLE &" classid=""CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA"" height=""32"" id=""video2"" width=""$1"">"&_
			"<param name=""src"" value=""$3"" /><param name=""AUTOSTART"" value=""false"" />"&_
			"<param name=""CONtrOLS"" value=""controlpanel"" />"&_
			"<param name=""CONSOLE"" value=""Clip1"" /></object></div>"& replace(Mtinfo,"$4","$3"),_
			PostUserGroup,Cint(Board_Setting(9) * mt),"=*([0-9]*),*([0-9]*)")
			'Dv7 RealPlayer自定义播放模式；
			s=Dv_UbbCode_iS2(s,"rm",_
			"<div><object "& DV_UBB_TITLE &" classid=""clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA"" class=""object"" id=""RAOCX"" width=""$1"" height=""$2"">"&_
			"<param name=""src"" value=""$4"" /><param name=""CONSOLE"" value=""$4"" /><param name=""CONtrOLS"" value=""imagewindow"" />"&_
			"<param name=""AUTOSTART"" value=""$3"" /></object></div>"&_
			"<div><object "& DV_UBB_TITLE &" classid=""CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA"" height=""32"" id=""video"" width=""$1"">"&_
			"<param name=""src"" value=""$4"" />"&_
			"<param name=""AUTOSTART"" value=""$3"" />"&_
			"<param name=""CONtrOLS"" value=""controlpanel"" /><param name=""CONSOLE"" value=""$4"" /></object></div>",_
			"<div><object "& DV_UBB_TITLE &" classid=""clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA"" class=""object"" id=""RAOCX"" width=""$1"" height=""$2"">"&_
			"<param name=""src"" value=""$4"" /><param name=""CONSOLE"" value=""$4"" /><param name=""CONtrOLS"" value=""imagewindow"" />"&_
			"<param name=""AUTOSTART"" value=""false"" /></object></div>"&_
			"<div><object "& DV_UBB_TITLE &" classid=""CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA"" height=""32"" id=""video"" width=""$1"">"&_
			"<param name=""src"" value=""$4"" />"&_
			"<param name=""AUTOSTART"" value=""false"" />"&_
			"<param name=""CONtrOLS"" value=""controlpanel"" /><param name=""CONSOLE"" value=""$4"" /></object></div>"&_
			Mtinfo,PostUserGroup,Cint(Board_Setting(9) * mt),"=*([0-9]*),*([0-9]*),*([0|1|true|false]*)")
		End If
		'背景音乐
		If InStr(Ubblists,",7,")>0 Or sType=1 Then
			s=Dv_UbbCode_iS2(s,"sound",_
			"<a href=""$1"" target=""_blank""><img "& DV_UBB_TITLE &" src=""skins/default/filetype/mid.gif"" border=""0"" alt=""背景音乐"" /></a><bgsound src=""$1"" loop=""-1"" />",_
			"<a href=""$1"" target=""_blank"">$1</a>"& replace(Mtinfo,"$4","$1"),_
			PostUserGroup,Cint(Board_Setting(9) * mt),"")
		End If
		'flash code
		If InStr(Ubblists,",8,")>0 Or sType=1 Then
			s=Dv_UbbCode_iS2(s,"flash",_
			"<a href=""$1"" target=""_blank""><img "& DV_UBB_TITLE &" src=""skins/default/filetype/swf.gif"" border=""0"" alt=""点击开新窗口欣赏该FLASH动画!"" height=""16"" width=""16"" />[全屏欣赏]</a><br/>"&_
			"<object "& DV_UBB_TITLE &" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0"" classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000""  width=""500"" height=""400"">"&_
			"<param name=""movie"" value=""$1"" /><PARAM NAME=""AllowScriptAccess"" VALUE=""never""><param name=""quality"" value=""high"" />"&_
			"<embed "& DV_UBB_TITLE &" src=""$1"" quality=""high"" pluginspage=""http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"" type=""application/x-shockwave-flash"" width=""500"" height=""400"">$1</embed></object>",_
			"<img "& DV_UBB_TITLE &" src=""skins/default/filetype/swf.gif"" border=""0""> <a href=$1 target=""_blank"">$1</a>"& replace(Mtinfo,"$4","$1"),_
			PostUserGroup,Cint(Board_Setting(44)),"")

			s=Dv_UbbCode_iS2(s,"flash",_
			"<a href=""$3"" target=""_blank""><img "& DV_UBB_TITLE &" src=""skins/default/filetype/swf.gif"" border=""0"" alt=""点击开新窗口欣赏该FLASH动画!"" height=""16"" width=""16"" />[全屏欣赏]</a><br/>"&_
			"<object "& DV_UBB_TITLE &" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0"" classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000""  width=""$1"" height=""$2"">"&_
			"<param name=""movie"" value=""$3"" /><PARAM NAME=""AllowScriptAccess"" VALUE=""never""><param name=""quality"" value=""high"" />"&_
			"<embed "& DV_UBB_TITLE &" src=""$3"" quality=""high"" pluginspage=""http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"" type=""application/x-shockwave-flash"" width=""$1"" height=""$2"">$3</embed></object>",_
			"<a href=""$3"" target=""_blank"">$3</a>",PostUserGroup,Cint(Board_Setting(44)),"=*([0-9]*),*([0-9]*),*(?:true|false)*")
		End If

		'point view
		If InStr(Ubblists,",9,")>0 Or sType=1 Then
			s=Dv_UbbCode_Get(s,PostUserGroup,PostType,"money",_
			"<hr/><font color=""gray"">以下内容需要金钱数达到<b>$1</b>才可以浏览</font><br />$2<hr/>",_
			"<span class=""info""><font color="""&Dvbbs.Mainsetting(1)&""">以下内容需要金钱数达到<b>$1</b>才可以浏览</font></span>",_
			UserPointInfo(0),Cint(Board_Setting(10)))
		End If
		If InStr(Ubblists,",10,")>0 Or sType=1 Then
			s=Dv_UbbCode_Get(s,PostUserGroup,PostType,"point",_
			"<hr/><font color=""gray"">以下内容需要积分达到<b>$1</b>才可以浏览</font><br/>$2<hr/>",_
			"<span class=""info""><font color="""&Dvbbs.Mainsetting(1)&""">以下内容需要积分达到<b>$1</b>才可以浏览</font></span>",_
			UserPointInfo(1),Cint(Board_Setting(11)))
		End If
		If InStr(Ubblists,",11,")>0 Or sType=1 Then
			s=Dv_UbbCode_Get(s,PostUserGroup,PostType,_
			"UserCP","<hr/><font color=""gray"">以下内容需要魅力达到<b>$1</b>才可以浏览</font><br/>$2<hr/>",_
			"<span class=""info""><font color="""&Dvbbs.Mainsetting(1)&""">以下内容需要魅力达到<b>$1</b>才可以浏览</font></span>",_
			UserPointInfo(2),Cint(Board_Setting(12)))
		End If
		If InStr(Ubblists,",12,")>0 Or sType=1 Then
			s=Dv_UbbCode_Get(s,PostUserGroup,PostType,_
			"Power","<hr /><font color=""gray"">以下内容需要威望达到<b>$1</b>才可以浏览</font><br/>$2<hr/>",_
			"<span class=""info""><font color="""&Dvbbs.Mainsetting(1)&""">以下内容需要威望达到<b>$1</b>才可以浏览</font></span>",_
			UserPointInfo(3),Cint(Board_Setting(13)))
		End If
		If InStr(Ubblists,",13,")>0 Or sType=1 Then
			s=Dv_UbbCode_Get(s,PostUserGroup,PostType,"Post",_
			"<hr /><font color=""gray"">以下内容需要帖子数达到<b>$1</b>才可以浏览</font><br/>$2<hr />",_
			"<span class=""info""><font color="""&Dvbbs.Mainsetting(1)&""">以下内容需要帖子数达到<b>$1</b>才可以浏览</font></span>",_
			UserPointInfo(4),Cint(Board_Setting(14)))
		End If
		If InStr(Ubblists,",14,")>0 Or sType=1 Then
			s=UBB_REPLYVIEW(s,PostUserGroup,PostType)
		End If
		If InStr(Ubblists,",15,")>0 Or sType=1 Then
			s=UBB_USEMONEY(s,PostUserGroup,PostType)
		End If
		'url code
		If InStr(Ubblists,",16,")>0 Or sType=1 Then
			s=Dv_UbbCode_S1(s,"url","<a href=""$1"" target=""_blank"">$1</a>")
			s=Dv_UbbCode_UF(s,"url","<a href=""$1"" target=""_blank"">$2</a>","0")
		End If
		'email code
		If InStr(Ubblists,",17,")>0 Or sType=1 Then
			s=Dv_UbbCode_S1(s,"email","<img "& DV_UBB_TITLE &" align=""absmiddle"" src=""skins/default/email1.gif"" alt=""""/><a href=""mailto:$1"">$1</a>")
			s=Dv_UbbCode_UF(s,"email","<img "& DV_UBB_TITLE &" align=""absmiddle"" src=""skins/default/email1.gif"" alt=""""/><a href=""mailto:$1"" target=""_blank"">$2</a>","0")
		End If
		If InStr(Ubblists,",37,")>0 Or sType=1 Then
			If (Cint(Board_Setting(8)) = 1 Or PostUserGroup<4) And InStr(Lcase(s),"[em")>0 Then
				re.Pattern="\[em([0-9]+)\]"
				s=re.Replace(s,"<img "& DV_UBB_TITLE &" src="""&EmotPath&"em$1.gif"" border=""0"" align=""middle"" alt="""" />")
			End If
		End If
		If InStr(Ubblists,",23,")>0 Or sType=1 Then
			s=Dv_UbbCode_C(s,"html")
		End If
		If InStr(Ubblists,",24,")>0 Or sType=1 Then
			s=Dv_UbbCode_S1(s,"code","<div class=""htmlcode""><b>以下内容为程序代码:</b><br/>$1</div>")
		End If
		If InStr(Ubblists,",25,")>0 Or sType=1 Then
			s=Dv_UbbCode_UF(s,"color","<font color=""$1"">$2</font>","1")
		End If
		If InStr(Ubblists,",26,")>0 Or sType=1 Then
			s=Dv_UbbCode_UF(s,"face","<font face=""$1"">$2</font>","1")
		End If
		If InStr(Ubblists,",27,")>0 Or sType=1 Then
			s=Dv_UbbCode_Align(s)
		End If
		If InStr(Ubblists,",42,")>0 Or sType=1 Then
			s=Dv_UbbCode_S1(s,"center","<div align=""center"">$1</div>")
		End If
		If InStr(Ubblists,",28,")>0 Or sType=1 Then
			s=Dv_UbbCode_Q(s)
		End If
		If InStr(Ubblists,",29,")>0 Or sType=1 Then
			s=Dv_UbbCode_S1(s,"fly","<marquee width=""90%"" behavior=""alternate"" scrollamount=""3"">$1</marquee>")
		End If
		If InStr(Ubblists,",30,")>0 Or sType=1 Then
			s=Dv_UbbCode_S1(s,"move","<marquee scrollamount=""3"">$1</marquee>")
		End If
		If InStr(Ubblists,",31,")>0 Or sType=1 Then
			s=Dv_UbbCode_iS1(s,"shadow","<div style=""width:$1px;filter:shadow(color=$2, strength=$3)"">$4</div>")
		End If
		If InStr(Ubblists,",32,")>0 Or sType=1 Then
			s=Dv_UbbCode_iS1(s,"glow","<div style=""width:$1px;filter:glow(color=$2, strength=$3)"">$4</div>")
		End If
		If InStr(Ubblists,",33,")>0 Or sType=1 Then
			s=Dv_UbbCode_UF(s,"size","<font size=""$1"">$2</font>","1")
		End If
		If InStr(Ubblists,",34,")>0 Or sType=1 Then
			s=Dv_UbbCode_S1(s,"i","<i>$1</i>")
		End If
		If InStr(Ubblists,",35,")>0 Or sType=1 Then
			s=Dv_UbbCode_S1(s,"b","<b>$1</b>")
		End If
		If InStr(Ubblists,",36,")>0 Or sType=1 Then
			s=Dv_UbbCode_S1(s,"u","<u>$1</u>")
		End If
		If InStr(Ubblists,",41,")>0 Or sType=1 Then
			s= Dv_UbbCode_name(s)
		End If
		'如果没有更新过帖子数据，而定员帖失效的，请把下面的注释去掉，建议进行帖子数据更新，以提高性能 2005.10.10 By Winder.F
		'If InStr(Lcase(s),"[username")>0 Then s= Dv_UbbCode_name(s)

		If InStr(s,"payto:") = 0 Then
			s = Replace(s,"https://www.alipay.com/payt","https://www.alipay.com/payto:")
		End If
		If InStr(Ubblists,",40,")>0 Then
			s=Dv_Alipay_PayTo(s)
		End If

		If xml.loadxml("<div>" & replace(s,"&","&amp;") &"</div>") Then
			isxhtml=True
			'checkimg已经写入checkXHTML
			'增加管理员是否允许发iframe标签 by 牛头
			s=checkXHTML(mt,PostUserGroup,ismanager1)
		Else
			Rem 处理老DHTML贴子
			isxhtml=False
			s=Dv_UbbCode_DHTML(s,PostUserGroup,PostType,sType)
			s=bbimg(s)
		End If

		If FileInfo Then
			s = s & UpFileInfoScript   '把增加文件相关信息的JS增加到变量S中 2008.3.6
			UpFileInfoScript = ""
		End If
		If pageReload Then
			Rem  回复可见帖子 ajax刷新一下
			s = s & "<scr"&"ipt type=""text/javascript"">var reload=1;</scr"&"ipt>"
		End If

		Dv_UbbCode = s
	End Function
	Private Function checkXHTML(mt,PostUserGroup,ismanager)
		Dim node,newnode,nodetext,attributes1,attributes2
		Dim NodeName,Attribute,AttName
		Dim hasname,hasvalue
		Rem  新xhtml 格式处理
		Rem 检索有害标记实行过滤
		Dim Stylestr,style,style1,newstyle,style_a,style_b
		Dim XML1,titletext,thissrc,objcount

		For Each Node in xml.documentElement.getElementsByTagName("*")
			NodeName = LCase(Node.nodeName)
			If NodeName="link" _
			Or NodeName="meta" _
			Or NodeName="script"  _
			Or NodeName="layer"  _
			Or NodeName="xss"  _
			Or NodeName="base"  _
			Or NodeName="html"  _
			Or NodeName="xhtml"  _
			Or NodeName="xml"  _
			Then
				Set newnode=xml.createTextNode(node.xml)
				node.parentNode.replaceChild newnode,node
			ElseIf NodeName="iframe" Or NodeName="frameset" Then
			    If ismanager-1<>0 Then
				    Set newnode=xml.createTextNode(node.xml)
				    node.parentNode.replaceChild newnode,node
				End If
			End If
            'response.Write G_UserList(26, G_ItemList(10, G_Floor)-1)&"<hr />"
			If NodeName="a" Then
				Node.setAttribute "target","_blank"
			End If

			'去掉STYLE标记
			If NodeName="style" Then
				node.parentNode.removeChild(Node)
			End If
			If NodeName="embed" Then
				node.setAttribute "quality","high"
'				node.setAttribute "wmode","opaque"
			End If
			'所有的属性的检查过滤

			For Each Attribute in node.attributes
				AttName = LCase(Attribute.nodeName)

				If Left(AttName,2) = "on" Then
					node.removeAttribute AttName
				Else
					nodetext=replaceasc(Attribute.text)
					If InStr(nodetext,"script:")>0 or InStr(nodetext,"document.")>0 Or InStr(nodetext,"xss:") > 0 Or InStr(nodetext,"expression") > 0 Then
						node.removeAttribute AttName
					End If
				End If

				Select Case NodeName
					Case "object"
						If AttName = "data" Then
							node.removeAttribute AttName
						End If
					Case "param"
						If Cint(Board_Setting(9) * mt)=0 Then
							hasname=0
							hasvalue=0
							If AttName="name" and Attribute.text = "autostart" Then
								hasname=1
							ElseIf AttName = "value" Then
								If hasvalue=1 Then
									node.setAttribute AttName,"false"
								End If
							End If
						End If
					Case "embed"
						If Cint(Board_Setting(9) * mt)=0 Then
							If AttName="autoplay" Then
								node.setAttribute AttName,"false"
							ElseIf AttName = "title" Then
								If Attribute.text<>UBB_TITLE Then
									node.setAttribute "title",UBB_TITLE
								End If
							ElseIf AttName = "src" Then
								node.setAttribute "src",Attribute.text
							Else
								'node.removeAttribute AttName
							End If
						End If

				End Select
			Next
			'把对图片的处理移到这里，去除原来的checkimg函数 hxyman 2008-1-6
			If NodeName="img" Then
				Set titletext=node.attributes.getNamedItem("title")
				If titletext is nothing Then
					titletext=""
				Else
					titletext=titletext.text
				End If
				If titletext=UBB_TITLE Then
					Rem 是否开启滚轮改变图片大小的功能，如果不需要可以屏蔽
					Rem Node.attributes.setNamedItem(xml.createNode(2,"onmousewheel","")).text="return bbimg(this);"
					Node.attributes.setNamedItem(xml.createNode(2,"onload","")).text="imgresize(this);"
					Node.attributes.setNamedItem(xml.createNode(2,"alt","")).text="图片点击可在新窗口打开查看"
				Else
					Rem 是否开启滚轮改变图片大小的功能，如果不需要可以屏蔽
					Rem Node.attributes.setNamedItem(xml.createNode(2,"onmousewheel","")).text="return bbimg(this);"
					Node.attributes.setNamedItem(xml.createNode(2,"onload","")).text="imgresize(this);"
					Node.attributes.setNamedItem(xml.createNode(2,"style","")).text="cursor: pointer;"
					Node.attributes.setNamedItem(xml.createNode(2,"alt","")).text="图片点击可在新窗口打开查看"
					Node.attributes.setNamedItem(xml.createNode(2,"onclick","")).text="javascript:window.open(this.src);"
					If Not node.parentNode is Nothing Then
						If node.parentNode.nodename = "a" Then
								node.attributes.removeNamedItem("onclick")
						End If
					End If
				End If
			End If
		Next

		Dim i

		If instr(","& can_Post_Style &",",","& PostUserGroup &",") = 0 Then
			For Each Node in xml.documentElement.selectNodes("//@*")
				If LCase(Node.nodeName)="style" Then
					Stylestr=node.text
					Stylestr=split(Stylestr,";")
					newstyle=""
				 	For each style in Stylestr
				 		style1=split(style,":")
				 		If UBound(style1)>0 Then
				 			style_a=LCase(Trim(style1(0)))
					 		style_b=LCase(Trim(style1(1)))
					 		If UBound(style1)>1 Then
					 				For i =2 to UBound(style1)
					 				style_b=style_b& ":"& style1(i)
					 				Next
					 		End If
					 		'吃掉POSITION:,top,left几个属性
					 		If (style_a<>"top" and style_a<>"left" and style_a<>"bottom" and style_a<>"right" and style_a<>"" and style_a<> "position") Then
						 			'去掉过宽的属性
						 			If style_a="width" Then
						 				If InStr(style_b,"px")>0 Then
						 					style_b=replace(style_b,"px","")
						 					If IsNumeric(style_b) Then
						 						If CLng(style_b)>600 Then style_b=600
						 					End If
						 					style_b=style_b&"px"
						 				ElseIf InStr(style_b,"%")>0 Then
						 					style_b=replace(style_b,"%","")
						 					If IsNumeric(style_b) Then
						 						If CLng(style_b)>100 Then style_b=100
						 					End If
						 					style_b=style_b&"%"
						 			End If
					 				'去掉过大的字体
						 			If style_a = "font-size" Then
						 				If InStr(style_b,"px")>0 Then
						 					style_b=replace(style_b,"px","")
						 					If IsNumeric(style_b) Then
						 						If CLng(style_b)> 200 Then style_b=200
						 					End If
						 					style_b=style_b&"px"
						 				ElseIf InStr(style_b,"%")> 0 Then
						 					style_b=replace(style_b,"%","")
						 					If IsNumeric(style_b) Then
						 						If CLng(style_b)>100 Then style_b=100
						 					End If
						 					style_b=style_b&"%"
						 				End If
						 			End If
					 			End If
					 			newstyle=newstyle&style_a&":"&style_b&";"
					 		End If
				 		End If
					Next
					node.text=newstyle
				End If
			Next
		End If
		checkXHTML=replace(Mid(xml.documentElement.xml,6,Len (xml.documentElement.xml)-11),"&amp;","&")
	End Function
	Function checkimg(textstr)
		Dim node,titletext
		If xml.loadxml("<div>" & replace(textstr,"&","&ayle_b&"%"
						 			End If
					 				'鍘绘帀杩囧ぇ鐨勫瓧浣