<!--#include file="../conn.asp"-->
<!--#include file="inc/Const.asp"-->
<%
Dim iCacheName,iCache,mCacheName
MyDbPath = "../"
'获得论坛基本信息和检测用户登陆状态
'Dvbbs.GetForum_Setting
'Dvbbs.CheckUserLogin
'重新赋予用户是否可进入后台权限
'If Dvbbs.GroupSetting(70)="1" Then Dvbbs.Master = True
CheckAdmin(",")
Head()
Dim CacheName
CacheName=Dvbbs.CacheName
Call delallcache()

Function  GetallCache()
	Dim Cacheobj
	For Each Cacheobj in Application.Contents
	If CStr(Left(Cacheobj,Len(CacheName)+1))=CStr(CacheName&"_") Then	
		GetallCache=GetallCache&Cacheobj&","
	End If
	Next
End Function
Sub delallcache()
	Dim cachelist,i
	Cachelist=split(GetallCache(),",")
	If UBound(cachelist)>1 Then
		For i=0 to UBound(cachelist)-1
			DelCahe Cachelist(i)
			Response.Write "更新 <b>"&Replace(cachelist(i),CacheName&"_","")&"</b> 完成<br>"		
		Next
		Response.Write "更新了"
		Response.Write UBound(cachelist)-1
		Response.Write "个缓存对象<br>"	
	Else
		Response.Write "所有对象已经更新。"
	End If
End Sub 
Sub DelCahe(MyCaheName)
	Application.Lock
	Application.Contents.Remove(MyCaheName)
	Application.unLock
End Sub
%>