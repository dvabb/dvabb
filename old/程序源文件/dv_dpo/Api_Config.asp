<%
'=========================================================
' File: Api_Config.asp
' Version:8.3.0
' Date: 2006-3-28
' Script Written by dvbbs.net
'=========================================================
' Copyright (C) 2003,2006 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: eway@aspsky.net
'=========================================================


'================================================================================================
'多系统整合设置
'================================================================================================
'DvApi_Enable 是否打开系统整合（默认闭关: False ,打开：True ）
Const DvApi_Enable	= False
'DvApi_SysKey 设置系统密钥 (系统整合，必须保证与其它系统设置的密钥一致。)
Const DvApi_SysKey	= "API_TEST"
'DvApi_Urls :整合的其它程序的接口文件路径。多个程序接口之间用半角"|"分隔。
'例如：DvApi_Urls = "http://你的网站地址/博客安装目录/oblogresponse.asp|http://你的网站地址/动易安装目录/API/API_Response.asp"
Const DvApi_Urls	= "http://你的网站地址/博客安装目录/oblogresponse.asp|http://你的网站地址/动易安装目录/API/API_Response.asp" 
%>