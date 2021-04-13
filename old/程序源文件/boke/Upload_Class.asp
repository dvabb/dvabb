<!--#include FILE="Upload.inc"-->
<%
'-----------------------------------------------------------------------
'--- 上传处理类模块
'--- Copyright (c) 2004 Aspsky, Inc.
'--- Mail: Sunwin@artbbs.net   http://www.aspsky.net
'--- 2004-12-18
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'-- InceptFileType	: 设置上传类型属性 (以逗号分隔多个文件类型) String
'-- MaxSize			: 设置上传文件大小上限 (单位：kb) Long
'-- InceptMaxFile	: 设置一次上传文件最大个数 Long
'-- UploadPath		: 设置保存的目录相对路径 String
'-- UploadType		: 设置上传组件类型 （0=无组件上传类，1=Aspupload3.0 ,2=SA-FileUp 4.0 ,3=DvFile.Upload V1.0）
'-- SaveUpFile		: 执行上传
'-- GetBinary		: 设置上传是否返回文件数据流  Bloon值 : True/False
'-- ChkSessionName	: 设置SESSION名，防止重复提交，SESSION名与提交的表单名要一致。
'-- RName设置文件名	: 定义文件名前缀 (如默认生成的文件名为200412230402587123.jpg
'									设置：RName="PRE_",生成的文件名为：PRE_200412230402587123.jpg)
'-----------------------------------------------------------------------
'-- 设置图片组件属性
'-- PreviewType		: 设置组件(0=CreatePreviewImage组件，1=AspJpegV1.2 ,2=SoftArtisans ImgWriter V1.21)
'-- PreviewImageWidth	: 设置预览图片宽度
'-- PreviewImageHeight	: 设置预览图片高度
'-- DrawImageWidth	: 设置水印图片或文字区域宽度
'-- DrawImageHeight	: 设置水印图片或文字区域高度
'-- DrawGraph		: 设置水印图片或文字区域透明度
'-- DrawFontColor	: 设置水印文字颜色
'-- DrawFontFamily	: 设置水印文字字体格式
'-- DrawFontSize	: 设置水印文字字体大小
'-- DrawFontBold	: 设置水印文字是否粗体
'-- DrawInfo		: 设置水印文字信息或图片信息
'-- DrawType		: 设置加载水印模式：0=不加载水印 ，1=加载水印文字 ，2=加载水印图片
'-- DrawXYType		: 图片添加水印LOGO位置坐标："0" =左上，"1"=左下,"2"=居中,"3"=右上,"4"=右下
'-- DrawSizeType	: 生成预览图片大小规则："0"=固定缩小，"1"=等比例缩小
'-----------------------------------------------------------------------
'-- 获取上传信息
'-- ObjName			: 采用的组件名称
'-- Count			: 上传文件总数
'-- CountSize		: 上传总大小字节数
'-- ErrCodes		: 错误NUMBER (默认为0)
'-- Description		: 错误描述
'-----------------------------------------------------------------------
'-- CreateView Imagename,TempFilename,FileExt
'	创建预览图片过程: 原始文件的相对路径,生成预览文件相对路径,原文件后缀
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'-- 获取文件对象属性 : UploadFiles
'-- FormName		: 表单名称
'-- FileName		: 生成的文件名称
'-- sFileName		: 文件原始名称
'-- FilePath		: 保存文件的相对路径
'-- FileSize		: 文件大小
'-- FileContentType	: ContentType文件类型
'-- FileType		: 0=其它,1=图片,2=FLASH,3=音乐,4=电影,5=压缩,6=文档
'-- FileData		: 文件数据流 (若组件不支持直接获取，则返回Null)
'-- FileExt			: 文件后缀
'-- FileWidth		: 图片/Flash文件宽度	（其他文件默认=-1）
'-- FileHeight		: 图片/Flash文件高度	（其他文件默认=-1）
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'-- 获取表单对象属性 : UploadForms
'-- Count			: 表单数
'-- key				: 表单内容
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------

Class UpFile_Cls
	Private UploadObj,ImageObj
	Private FilePath,InceptFile,FileMaxSize,MaxFile,Upload_Type,FileInfo,IsBinary,SessionName
	Private Preview_Type,View_ImageWidth,View_ImageHeight,Draw_ImageWidth,Draw_ImageHeight,Draw_Graph
	Private Draw_FontColor,Draw_FontFamily,Draw_FontSize,Draw_FontBold,Draw_Info,Draw_Type,Draw_XYType,Draw_SizeType
	Private RName_Str,Transition_Color
	Public ErrCodes,ObjName,UploadFiles,UploadForms,Count,CountSize
	'-----------------------------------------------------------------------------------
	'初始化类
	'-----------------------------------------------------------------------------------
	Private Sub Class_Initialize
		SessionName = Empty
		IsBinary = False
		ErrCodes = 0
		Count = 0
		CountSize = 0
		FilePath = "./"
		InceptFile = ""
		FileMaxSize = -1
		MaxFile = 1
		Upload_Type = -1
		Preview_Type = 999
		ObjName = "未知组件"
		View_ImageWidth = 0
		View_ImageHeight = 0
		Draw_FontColor	= &H000000
		Draw_FontFamily	= "Arial"
		Draw_FontSize	= 10
		Draw_FontBold	= False
		Draw_Info		= "BBS.DVBBS.NET"
		Draw_Type		= -1
		Set UploadFiles = Dvbbs.iCreateObject ("Scripting.Dictionary")
		Set UploadForms = Dvbbs.iCreateObject ("Scripting.Dictionary")
		UploadFiles.CompareMode = 1
		UploadForms.CompareMode = 1
	End Sub

	'-----------------------------------------------------------------------------------
	'销毁类
	'-----------------------------------------------------------------------------------
	Private Sub Class_Terminate
		If IsObject(UploadObj) Then
			Set UploadObj = Nothing
		End If
		If IsObject(ImageObj) Then
			Set ImageObj = Nothing
		End If
		UploadFiles.RemoveAll
		UploadForms.RemoveAll
		Set UploadForms = Nothing
		Set UploadFiles = Nothing
	End Sub

	'-----------------------------------------------------------------------------------
	'设置上传是否返回文件数据流
	'-----------------------------------------------------------------------------------
	Public Property Let GetBinary(Byval Values)
		IsBinary = Values
	End Property

	'-----------------------------------------------------------------------------------
	'设置上传类型属性 (以逗号分隔多个文件类型)
	'-----------------------------------------------------------------------------------
	Public Property Let InceptFileType(Byval Values)
		InceptFile = Lcase(Values)
	End Property

	'-----------------------------------------------------------------------------------
	'设置上传类型属性 (以逗号分隔多个文件类型)
	'-----------------------------------------------------------------------------------
	Public Property Let ChkSessionName(Byval Values)
		SessionName = Values
	End Property

	'-----------------------------------------------------------------------------------
	'设置上传文件大小上限 (单位：kb)
	'-----------------------------------------------------------------------------------
	Public Property Let MaxSize(Byval Values)
		FileMaxSize = ChkNumeric(Values) * 1024
	End Property
	Public Property Get MaxSize
		MaxSize = FileMaxSize
	End Property

	'-----------------------------------------------------------------------------------
	'设置每次上传文件上限
	'-----------------------------------------------------------------------------------
	Public Property Let InceptMaxFile(Byval Values)
		MaxFile = ChkNumeric(Values)
	End Property

	'-----------------------------------------------------------------------------------
	'设置上传目录路径
	'-----------------------------------------------------------------------------------
	Public Property Let UploadPath(Byval Path)
		FilePath = Replace(Path,Chr(0),"")
		If Right(FilePath,1)<>"/" Then FilePath = FilePath & "/"
	End Property

	Public Property Get UploadPath
		UploadPath = FilePath
	End Property

	'-----------------------------------------------------------------------------------
	'获取错误信息
	'-----------------------------------------------------------------------------------
	Public Property Get Description
		Select Case ErrCodes
			Case 1 : Description = "不支持 " & ObjName & " 上传，服务器可能未安装该组件。"
			Case 2 : Description = "暂未选择上传组件！"
			Case 3 : Description = "请先选择你要上传的文件!"
			Case 4 : Description = "文件大小超过了限制 " & (FileMaxSize\1024) & "KB!"
			Case 5 : Description = "文件类型不正确!"
			Case 6 : Description = "已达到上传数的上限！"
			Case 7 : Description = "请不要重复提交！"
			Case Else
				Description = Empty
		End Select
	End Property

	'-----------------------------------------------------------------------------------
	'设置文件名前缀
	'-----------------------------------------------------------------------------------
	Public Property Let RName(Byval Values)
		RName_Str = Values
	End Property

	'-----------------------------------------------------------------------------------
	'设置上传组件属性
	'-----------------------------------------------------------------------------------
	Public Property Let UploadType(Byval Types)
		Upload_Type = Types
		If Upload_Type = "" or Not IsNumeric(Upload_Type) Then
			Upload_Type = -1
		End If
	End Property

	'-----------------------------------------------------------------------------------
	'设置上传图片组件属性
	'-----------------------------------------------------------------------------------
	Public Property Let PreviewType(Byval Types)
		Preview_Type = Types
		On Error Resume Next
		If Preview_Type = "" or Not IsNumeric(Preview_Type) Then
			Preview_Type = 999
		Else
			If PreviewType <> 999 Then
				Select Case Preview_Type
					Case 0
					'---------------------CreatePreviewImage---------------
						ObjName = "CreatePreviewImage组件"
						Set ImageObj = Dvbbs.iCreateObject("CreatePreviewImage.cGvbox")
					Case 1
					'---------------------AspJpegV1.2---------------
						ObjName = "AspJpegV1.2组件"
						Set ImageObj = Dvbbs.iCreateObject("Persits.Jpeg")
					Case 2
					'---------------------SoftArtisans ImgWriter V1.21---------------
						ObjName = "SoftArtisans ImgWriter V1.21组件"
						Set ImageObj = Dvbbs.iCreateObject("SoftArtisans.ImageGen")
					Case Else
						Preview_Type = 999
				End Select
				If Err.Number<>0 Then
					ErrCodes = 1
				End If
			End If
		End If
	End Property

	Public Property Get PreviewType
		PreviewType = Preview_Type
	End Property

	'-----------------------------------------------------------------------------------
	'设置预览图片宽度属性
	'-----------------------------------------------------------------------------------
	Public Property Let PreviewImageWidth(Byval Values)
		View_ImageWidth = ChkNumeric(Values)
	End Property

	'-----------------------------------------------------------------------------------
	'设置预览图片高度属性
	'-----------------------------------------------------------------------------------
	Public Property Let PreviewImageHeight(Byval Values)
		View_ImageHeight = ChkNumeric(Values)
	End Property

	'-----------------------------------------------------------------------------------
	'设置水印图片或文字区域宽度属性
	'-----------------------------------------------------------------------------------
	Public Property Let DrawImageWidth(Byval Values)
		Draw_ImageWidth = ChkNumeric(Values)
	End Property

	'-----------------------------------------------------------------------------------
	'设置水印图片或文字区域高度属性
	'-----------------------------------------------------------------------------------
	Public Property Let DrawImageHeight(Byval Values)
		Draw_ImageHeight = ChkNumeric(Values)
	End Property

	'-----------------------------------------------------------------------------------
	'设置水印图片或文字区域透明度属性
	'-----------------------------------------------------------------------------------
	Public Property Let DrawGraph(Byval Values)
		If IsNumeric(Values) Then
			Draw_Graph = Formatnumber(Values,2)
		Else
			Draw_Graph = 1
		End If
	End Property

	'-----------------------------------------------------------------------------------
	'设置水印图片透明度去除底色值
	'-----------------------------------------------------------------------------------
	Public Property Let TransitionColor(Byval Values)
		If Values<>"" or Values<>"0" Then
			Transition_Color = Replace(Values,"#","&h")
		End If
	End Property

	'-----------------------------------------------------------------------------------
	'设置水印文字颜色
	'-----------------------------------------------------------------------------------
	Public Property Let DrawFontColor(Byval Values)
		If Values<>"" or Values<>"0" Then
			Draw_FontColor = Replace(Values,"#","&h")
		End If
	End Property

	'-----------------------------------------------------------------------------------
	'设置水印文字字体格式
	'-----------------------------------------------------------------------------------
	Public Property Let DrawFontFamily(Byval Values)
		Draw_FontFamily = Values
	End Property

	'-----------------------------------------------------------------------------------
	'设置水印文字字体大小
	'-----------------------------------------------------------------------------------
	Public Property Let DrawFontSize(Byval Values)
		Draw_FontSize = Values
	End Property

	'-----------------------------------------------------------------------------------
	'设置水印文字是否粗体 Boolean
	'-----------------------------------------------------------------------------------
	Public Property Let DrawFontBold(Byval Values)
		Draw_FontBold = ChkBoolean(Values)
	End Property
	'-----------------------------------------------------------------------------------
	'设置水印文字信息或图片信息
	'-----------------------------------------------------------------------------------
	Public Property Let DrawInfo(Byval Values)
		Draw_Info = Values
	End Property

	'-----------------------------------------------------------------------------------
	'加载模式：0=不加载水印 ，1=加载水印文字 ，2=加载水印图片
	'-----------------------------------------------------------------------------------
	Public Property Let DrawType(Byval Values)
		Draw_Type = ChkNumeric(Values)
	End Property

	'-----------------------------------------------------------------------------------
	'图片添加水印LOGO位置坐标："0" =左上，"1"=左下,"2"=居中,"3"=右上,"4"=右下
	'-----------------------------------------------------------------------------------
	Public Property Let DrawXYType(Byval Values)
		 Draw_XYType = Values
	End Property

	'-----------------------------------------------------------------------------------
	'生成预览图片大小规则："0"=固定缩小，"1"=等比例缩小
	'-----------------------------------------------------------------------------------
	Public Property Let DrawSizeType(Byval Values)
		Draw_SizeType = Values
	End Property

	Private Function ChkNumeric(Byval Values)
		If Values<>"" and Isnumeric(Values) Then
			ChkNumeric = Int(Values)
		Else
			ChkNumeric = 0
		End If
	End Function

	Private Function ChkBoolean(Byval Values)
		If Typename(Values)="Boolean" or IsNumeric(Values) or Lcase(Values)="false" or Lcase(Values)="true" Then
			ChkBoolean = CBool(Values)
		Else
			ChkBoolean = False
		End If
	End Function

	'-----------------------------------------------------------------------------------
	'日期时间定义文件名
	'-----------------------------------------------------------------------------------
	Private Function FormatName(Byval FileExt)
		Dim RanNum,TempStr
		Randomize
		RanNum = Int(90000*rnd)+10000
		TempStr = Year(now) & Month(now) & Day(now) & Hour(now) & Minute(now) & Second(now) & RanNum & "." & FileExt
		If RName_Str<>"" Then
			TempStr = RName_Str & TempStr
		End If
		FormatName = TempStr
	End Function
	
	'-----------------------------------------------------------------------------------
	'格式后缀
	'-----------------------------------------------------------------------------------
	Private Function FixName(Byval UpFileExt)
		If IsEmpty(UpFileExt) Then Exit Function
		FixName = Lcase(UpFileExt)
		FixName = Replace(FixName,Chr(0),"")
		FixName = Replace(FixName,".","")
		FixName = Replace(FixName,"'","")
		FixName = Replace(FixName,"asp","")
		FixName = Replace(FixName,"asa","")
		FixName = Replace(FixName,"aspx","")
		FixName = Replace(FixName,"cer","")
		FixName = Replace(FixName,"cdx","")
		FixName = Replace(FixName,"htr","")
	End Function

	'-----------------------------------------------------------------------------------
	'判断文件类型是否合格
	'-----------------------------------------------------------------------------------
	Private Function CheckFileExt(FileExt)
		Dim Forumupload,i
		CheckFileExt=False
		If FileExt="" or IsEmpty(FileExt) Then
			CheckFileExt = False
			Exit Function
		End If
		If FileExt="asp" or FileExt="asa" or FileExt="aspx" Then
			CheckFileExt = False
			Exit Function
		End If
		Forumupload = Split(InceptFile,",")
		For i = 0 To ubound(Forumupload)
			If FileExt = Trim(Forumupload(i)) Then
				CheckFileExt = True
				Exit Function
			Else
				CheckFileExt = False
			End If
		Next
	End Function

	'-----------------------------------------------------------------------------------
	'判断文件类型:0=其它,1=图片,2=FLASH,3=音乐,4=电影
	'-----------------------------------------------------------------------------------
	Private Function CheckFiletype(Byval FileExt)
		FileExt = Lcase(Replace(FileExt,".",""))
		Select Case FileExt
				Case "gif", "jpg", "jpeg","png","bmp","tif","iff"
					CheckFiletype=1
				Case "swf", "swi"
					CheckFiletype=2
				Case "mid", "wav", "mp3","rmi","cda"
					CheckFiletype=3
				Case "avi", "mpg", "mpeg","ra","ram","wov","asf"
					CheckFiletype=4
				Case "rar", "zip", "tar", "cab", "exe"
					CheckFiletype=5
				Case "doc", "txt", "mdb", "ppt","xls","asp","aspx","php","jsp"
					CheckFiletype=6
				Case Else
					CheckFiletype=0
		End Select
	End Function

	'-----------------------------------------------------------------------------------
	'执行保存上传文件
	'-----------------------------------------------------------------------------------
	Public Sub SaveUpFile()
		On Error Resume Next
		Select Case (Upload_Type) 
			Case 0
				ObjName = "无组件"
				Set UploadObj = New UpFile_Class
				If Err.Number<>0 Then
					ErrCodes = 1
				Else
					SaveFile_0
				End If
			Case 1
				ObjName = "Aspupload3.0组件"
				Set UploadObj = Dvbbs.iCreateObject("Persits.Upload") 
				If Err.Number<>0 Then
					ErrCodes = 1
				Else
					SaveFile_1
				End If
			Case 2
				ObjName = "SA-FileUp 4.0组件"
				Set UploadObj = Dvbbs.iCreateObject("SoftArtisans.FileUp")
				If Err.Number<>0 Then
					ErrCodes = 1
				Else
					SaveFile_2
				End If
			Case 3
				ObjName = "DvFile.Upload V1.0组件"
				Set UploadObj = Dvbbs.iCreateObject("DvFile.Upload")
				If Err.Number<>0 Then
					ErrCodes = 1
				Else
					SaveFile_3
				End If
			Case Else
				ErrCodes = 2
		End Select
	End Sub

	''-----------------------------------------------------------------------------------
	' 上传处理过程
	''-----------------------------------------------------------------------------------
	''-----------------------------------------------------------------------------------
	''无组件上传
	''-----------------------------------------------------------------------------------
	Private Sub SaveFile_0()
		Dim FormName,Item,File
		Dim FileExt,FileName,sFileName,FileType,FileToBinary
		UploadObj.InceptFileType = InceptFile
		UploadObj.MaxSize = FileMaxSize
		UploadObj.GetDate ()	'取得上传数据
		FileToBinary = Null
		If Not IsEmpty(SessionName) Then
			If Session(SessionName) <> UploadObj.Form(SessionName) or Session(SessionName) = Empty Then
				ErrCodes = 7
				Exit Sub
			End If
		End If
		If UploadObj.Err > 0 then
			Select Case UploadObj.Err
				Case 1 : ErrCodes = 3
				Case 2 : ErrCodes = 4
				Case 3 : ErrCodes = 5
			End Select
			Exit Sub
		Else
			For Each FormName In UploadObj.File		''列出所有上传了的文件
				If Count>MaxFile Then
					ErrCodes = 6
					Exit Sub
				End If
				Set File = UploadObj.File(FormName)
				sFileName = File.FileName
				FileExt = FixName(File.FileExt)
				If CheckFileExt(FileExt) = False then
					ErrCodes = 5
					EXIT SUB
				End If
				FileName = FormatName(FileExt)
				FileType = CheckFiletype(FileExt)
				If IsBinary Then
					FileToBinary = File.FileData
				End If
				If File.FileSize>0 Then
					File.SaveToFile Server.Mappath(FilePath & FileName)
					AddData FormName , _ 
							FileName , _
							sFileName , _
							FilePath , _
							File.FileSize , _
							File.FileType , _
							FileType , _
							FileToBinary , _
							FileExt , _
							File.FileWidth , _
							File.FileHeight
					Count = Count + 1
					CountSize = CountSize + File.FileSize
				End If
				Set File=Nothing
			Next
			For Each Item in UploadObj.Form
				If UploadForms.Exists (Item) Then _
					UploadForms(Item) = UploadForms(Item) & ", " & UploadObj.Form(Item) _
				Else _
				UploadForms.Add Item , UploadObj.Form(Item)
			Next
			If Not IsEmpty(SessionName) Then Session(SessionName) = Empty
		End If
	End Sub
	''-----------------------------------------------------------------------------------
	''Aspupload3.0组件上传
	''-----------------------------------------------------------------------------------
	Private Sub SaveFile_1()
		Dim FileCount
		Dim FormName,Item,File
		Dim FileExt,FileName,sFileName,FileType,FileToBinary
		UploadObj.OverwriteFiles = False		'不能复盖
		UploadObj.IgnoreNoPost = True
		UploadObj.SetMaxSize FileMaxSize, True	'限制大小
		FileCount = UploadObj.Save
		FileToBinary = Null
		If Not IsEmpty(SessionName) Then
			If Session(SessionName) <> UploadObj.Form(SessionName) or Session(SessionName) = Empty Then
				ErrCodes = 7
				Exit Sub
			End If
		End If

		If Err.Number = 8 Then
				ErrCodes = 4
				EXIT SUB
		Else 
				If Err <> 0 Then
					ErrCodes = -1
					Response.Write "错误信息: " & Err.Description
					EXIT SUB
				End If
				If FileCount < 1 Then 
					ErrCodes = 3
					EXIT SUB
				End If
				For Each File In UploadObj.Files	'列出所有上传文件
					If Count>MaxFile Then
						ErrCodes = 6
						Exit Sub
					End If
					sFileName = File.Name
					FileExt = FixName(Replace(File.Ext,".",""))
					If CheckFileExt(FileExt) = False then
						ErrCodes = 5
						EXIT SUB
					End If
					FileName = FormatName(FileExt)
					FileType = CheckFiletype(FileExt)
					If IsBinary Then
						FileToBinary = File.Binary
					End If
					'File.Filename
					If File.Size>0 Then
						File.SaveAs Server.Mappath(FilePath & FileName)
						AddData File.Name , _ 
							FileName , _
							sFileName , _
							FilePath , _
							File.Size , _
							File.ContentType , _
							FileType , _
							FileToBinary , _
							FileExt , _
							File.ImageWidth , _
							File.ImageHeight
						Count = Count + 1
						CountSize = CountSize + File.Size
					End If
				Next
				For Each Item in UploadObj.Form
					If UploadForms.Exists (Item) Then _
						UploadForms(Item) = UploadForms(Item) & ", " & Item.Value _
					Else _
						UploadForms.Add Item.Name , Item.Value
				Next
				If Not IsEmpty(SessionName) Then Session(SessionName) = Empty
		End If
	End Sub
	''-----------------------------------------------------------------------------------
	''SA-FileUp 4.0组件上传FileUpSE V4.09
	''-----------------------------------------------------------------------------------
	Private Sub SaveFile_2()
		Dim FormName,Item,File,FormNames
		Dim FileExt,FileName,sFileName,FileType,FileToBinary
		Dim Filesize
		FileToBinary = Null
		If Not IsEmpty(SessionName) Then
			If Session(SessionName) <> UploadObj.Form(SessionName) or Session(SessionName) = Empty Then
				ErrCodes = 7
				Exit Sub
			End If
		End If
		For Each FormName In UploadObj.Form
			FormNames = ""
			If IsObject(UploadObj.Form(FormName)) Then
				If Not UploadObj.Form(FormName).IsEmpty Then
					UploadObj.Form(FormName).Maxbytes = FileMaxSize	'限制大小
					UploadObj.OverWriteFiles = False
					Filesize = UploadObj.Form(FormName).TotalBytes
					If Err.Number<>0 Then
						ErrCodes = -1
						Response.Write "错误信息: " & Err.Description
						EXIT SUB
					End If
					If Filesize>FileMaxSize then
						ErrCodes = 4
						Exit sub
					End If
					FileName	= UploadObj.Form(FormName).ShortFileName	 '原文件名
					sFileName	= FileName
					FileExt		= Mid(Filename, InStrRev(Filename, ".")+1)
					FileExt		= FixName(FileExt)
					If CheckFileExt(FileExt) = False then
						ErrCodes = 5
						EXIT SUB
					End If
					FileName = FormatName(FileExt)
					FileType = CheckFiletype(FileExt)
					'If IsBinary Then
						'FileToBinary = UploadContents (2)
					'End If
					'保存文件
					If Filesize>0 Then
						UploadObj.Form(FormName).SaveAs Server.MapPath(FilePath & FileName)
						AddData FormName , _ 
								FileName , _
								sFileName , _
								FilePath , _
								FileSize , _
								UploadObj.Form(FormName).ContentType , _
								FileType , _
								FileToBinary , _
								FileExt , _
								-1 , _
								-1
						Count = Count + 1
						CountSize = CountSize + Filesize
					End If
				Else
					ErrCodes = 3
					EXIT SUB
				End If
			Else
				If UploadObj.FormEx(FormName).Count > 1 Then
					For Each FormNames In UploadObj.FormEx(FormName)
						FormNames = FormNames & ", " & FormNames
					Next
					UploadForms.Add FormName , FormNames
				Else
					UploadForms.Add FormName , UploadObj.Form(FormName)
				End If
			End If
		Next
		If Not IsEmpty(SessionName) Then Session(SessionName) = Empty
	End Sub
	''-----------------------------------------------------------------------------------
	''DvFile.Upload V1.0组件上传
	''-----------------------------------------------------------------------------------
	Private Sub SaveFile_3()
		Dim FormName,Item,File
		Dim FileExt,FileName,sFileName,FileType,FileToBinary
		UploadObj.InceptFileType = InceptFile
		UploadObj.MaxSize = FileMaxSize
		UploadObj.Install
		FileToBinary = Null
		If Not IsEmpty(SessionName) Then
			If Session(SessionName) <> UploadObj.Form(SessionName) or Session(SessionName) = Empty Then
				ErrCodes = 7
				Exit Sub
			End If
		End If
		If UploadObj.Err > 0 then
			Select Case UploadObj.Err
				Case 1 : ErrCodes = 3
				Case 2 : ErrCodes = 4
				Case 3 : ErrCodes = 5
				Case 4 : ErrCodes = 5
				Case 5 : ErrCodes = -1
			End Select
			Exit Sub
		Else
			For Each FormName In UploadObj.File		''列出所有上传了的文件
				If Count>MaxFile Then
					ErrCodes = 6
					Exit Sub
				End If
				Set File = UploadObj.File(FormName)
				sFileName = File.FileName
				FileExt = FixName(File.FileExt)
				If CheckFileExt(FileExt) = False then
					ErrCodes = 5
					EXIT SUB
				End If
				FileName = FormatName(FileExt)
				FileType = CheckFiletype(FileExt)
				If IsBinary Then
					FileToBinary = File.FileData
				End If
				If File.FileSize>0 Then
					UploadObj.SaveToFile Server.mappath(FilePath & FileName),FormName
					AddData FormName , _ 
							FileName , _
							sFileName , _
							FilePath , _
							File.FileSize , _
							File.FileType , _
							FileType , _
							FileToBinary , _
							FileExt , _
							File.FileWidth , _
							File.FileHeight
					Count = Count + 1
					CountSize = CountSize + File.FileSize
				End If
				Set File=Nothing
			Next
			For Each Item in UploadObj.Form
				UploadForms.Add Item.Name , Item.Value
			Next
			If Not IsEmpty(SessionName) Then Session(SessionName) = Empty
		End If
	End Sub

	Private Sub AddData( Form_Name,File_Name,sFile_Name,File_Path,File_Size,File_ContentType,File_Type,File_Data,File_Ext,File_Width,File_Height )
		Set FileInfo = New FileInfo_Cls
			FileInfo.FormName = Form_Name
			FileInfo.FileName = File_Name
			FileInfo.sFileName = sFile_Name
			FileInfo.FilePath = File_Path
			FileInfo.FileSize = File_Size
			FileInfo.FileType = File_Type
			FileInfo.FileContentType = File_ContentType
			FileInfo.FileExt = File_Ext
			FileInfo.FileData = File_Data
			FileInfo.FileHeight = File_Height
			FileInfo.FileWidth = File_Width
			UploadFiles.Add Form_Name , FileInfo
		Set FileInfo = Nothing
	End Sub

	'创建预览图片:Call CreateView(原始文件的路径,预览文件名及路径,原文件后缀)
	Public Sub CreateView(Imagename,TempFilename,FileExt)
		If ErrCodes <>0 Then Exit Sub
		Select Case Preview_Type
			Case 0
				Image_Obj_0 Imagename,TempFilename,FileExt
			Case 1
				Image_Obj_1 Imagename,TempFilename,FileExt
			Case 2
				Image_Obj_2 Imagename,TempFilename,FileExt
			Case Else
				Preview_Type = 999
		End Select
	End Sub

	Sub Image_Obj_0(Imagename,TempFilename,FileExt)
			ImageObj.SetSavePreviewImagePath = Server.MapPath(TempFilename)			'预览图存放路径
			ImageObj.SetPreviewImageSize = SetPreviewImageSize						'预览图宽度
			ImageObj.SetImageFile = Trim(Server.MapPath(Imagename))					'Imagename原始文件的物理路径
			'创建预览图的文件
			If ImageObj.DoImageProcess = False Then
				ErrCodes = -1
				Response.Write "生成预览图错误: " & ImageObj.GetErrString
			End If
	End Sub

	'---------------------AspJpegV1.2---------------
	Sub Image_Obj_1(Imagename,TempFilename,FileExt)
			' 读取要处理的原文件
			Dim Draw_X,Draw_Y,Logobox
			Draw_X = 0
			Draw_Y = 0
			FileExt = Lcase(FileExt)
			ImageObj.Open Trim(Server.MapPath(Imagename))
			If ImageObj.OriginalWidth<View_ImageWidth or ImageObj.Originalheight<View_ImageHeight Then
				TempFilename = ""
				Exit Sub
			Else
				If FileExt<>"gif" and ImageObj.OriginalWidth > Draw_ImageWidth * 2 and Draw_Type >0 Then
					Draw_X = DrawImage_X(ImageObj.OriginalWidth,Draw_ImageWidth,2)
					Draw_Y = DrawImage_y(ImageObj.Originalheight,Draw_ImageHeight,2)
					If Draw_Type=2 Then
						Set Logobox = Dvbbs.iCreateObject("Persits.Jpeg")
						'*添加水印图片	添加时请关闭水印字体*
						'//读取添加的图片
						Logobox.Open Server.MapPath(Draw_Info)
						Logobox.Width = Draw_ImageWidth								'// 加入图片的原宽度
						Logobox.Height = Draw_ImageHeight							'// 加入图片的原高度
						ImageObj.DrawImage Draw_X, Draw_Y, Logobox, Draw_Graph,Transition_Color,90	'// 加入图片的位置价坐标（添加水印图片）
						'ImageObj.Sharpen 1, 130
						ImageObj.Save Server.MapPath(Imagename)
						Set Logobox=Nothing
					Else
						'//关于修改字体及文字颜色的
						ImageObj.Canvas.Font.Color		= Draw_FontColor	'// 文字的颜色
						ImageObj.Canvas.Font.Family		= Draw_FontFamily	'// 文字的字体
						ImageObj.Canvas.Font.Bold		= Draw_FontBold
						ImageObj.Canvas.Font.Size		= Draw_FontSize					'//字体大小
						' Draw frame: black, 2-pixel width
						ImageObj.Canvas.Print Draw_X, Draw_Y, Draw_Info	'// 加入文字的位置坐标
						ImageObj.Canvas.Pen.Color		= &H000000		'// 边框的颜色
						ImageObj.Canvas.Pen.Width		= 1				'// 边框的粗细
						ImageObj.Canvas.Brush.Solid	= False			'// 图片边框内是否填充颜色
						'ImageObj.Canvas.Bar 0, 0, ImageObj.Width, ImageObj.Height	'// 图片边框线的位置坐标
						ImageObj.Save Server.MapPath(Imagename)
					End If
				End If
				If ImageObj.Width > ImageObj.height Then
					ImageObj.Width = View_ImageWidth
					ImageObj.Height = ViewImage_Height(ImageObj.OriginalWidth,ImageObj.Originalheight,View_ImageWidth,View_ImageHeight)
				Else
					ImageObj.Width = ViewImage_Width(ImageObj.OriginalWidth,ImageObj.Originalheight,View_ImageWidth,View_ImageHeight)
					ImageObj.Height = View_ImageHeight
				End If
				ImageObj.Sharpen 1, 120
				ImageObj.Save Server.MapPath(TempFilename)		'// 生成预览文件
			End If
	End Sub

	'SoftArtisans ImgWriter V1.21
	Public Sub Image_Obj_2(Imagename,TempFilename,FileExt)
			'定义变量
			Dim Draw_X,Draw_Y
			FileExt = Lcase(FileExt)
			Draw_X = 0
			Draw_Y = 0
			' 读取要处理的原文件
			ImageObj.LoadImage Trim(Server.MapPath(Imagename))
			If ImageObj.ErrorDescription <> "" Then
				TempFilename = ""
				ErrCodes = -1
				Response.Write "生成预览图错误: " &ImageObj.ErrorDescription
				Exit Sub
			End If
			If ImageObj.Width<Cint(View_ImageWidth) or ImageObj.Height<Cint(View_ImageHeight) Then
				TempFilename=""
				Exit Sub
			Else
				IF FileExt<>"gif" and ImageObj.Width > Draw_ImageWidth * 2 and Draw_Type>0 Then
					Draw_X = DrawImage_X(ImageObj.Width,Draw_ImageWidth,2)
					Draw_Y = DrawImage_y(ImageObj.Height,Draw_ImageHeight,2)
					Dim saiTopMiddle
					Select Case Draw_XYType
						Case "0" '左上
							saiTopMiddle = 3
						Case "1" '左下
							saiTopMiddle = 5
						Caidth,ImageObj.Originalheight,View_ImageWidth,View_ImageHeight)
				Else
					ImageObj.Width = ViewImage_Width(ImageObj.OriginalWidth,ImageObj.Originalheight,View_ImageWidth,View_ImageHeight)
					ImageObj.Height = View_ImageHeight
				End If
				ImageObj.Sharpen 1, 120
				ImageObj.Save Server.MapPath(TempFilename)		'// 鐢熸垚棰勮