<%
Dim Dv_Tools
Set Dv_Tools=new Plus_Tools_Cls

Class Plus_Tools_Cls
	Public ToolsID,ToolsInfo,ToUserInfo,UserToolsInfo,ToolsSetting
	Private buyCount

	Private Sub Class_Initialize()
		buyCount = 1
		ToolsID = CheckNumeric(Request("ToolsID"))
		If DVbbs.Forum_Setting(90)=0 and IsEmpty(session("flag")) Then ShowErr(1)	'中心已关闭
	End Sub

	Public Sub ChkToolsLogin()
		If Dvbbs.UserID=0 Then Dvbbs.AddErrCode(6):Dvbbs.Showerr()	'判断用户是否在线。
		If ToolsID=0 Then ShowErr(3):Exit Sub
		GetToolsInfo	'提取道具设置信息
	End Sub

	'---------------------------------------------------
	'读取道具系统信息
	'---------------------------------------------------
	Private Sub GetToolsInfo()
		Dim Sql,Rs
		'ID=0 ,ToolsName=1 ,ToolsInfo=2 ,IsStar=3 ,SysStock=4 ,UserStock=5 ,UserMoney=6 ,UserPost=7 ,UserWealth=8 ,UserEp=9 ,UserCp=10 ,UserGroupID=11 ,boardID=12,UserTicket=13,buyType=14,ToolsImg=15,ToolsSetting=16
		Sql = "Select ID,ToolsName,ToolsInfo,IsStar,SysStock,UserStock,UserMoney,UserPost,UserWealth,UserEp,UserCp,UserGroupID,boardID,UserTicket,buyType,ToolsImg,ToolsSetting From [Dv_Plus_Tools_Info] Where ID="& ToolsID
		Set Rs = Dvbbs.Plus_Execute(Sql)
		If Rs.Eof Then
			ShowErr(3):Exit Sub
		Else
			Sql = Rs.GetString(,1, "§§§", "", "")
			Sql = Split(Sql,"§§§")
		End If
		Rs.Close : Set Rs = Nothing
		ToolsInfo = Sql
		ToolsSetting = Split(ToolsInfo(16),",")
	End Sub
	'---------------------------------------------------
	'读取用户道具信息
	'---------------------------------------------------
	Public Sub GetUserToolsInfo(G_USerID,G_ToolsID)
		Dim Sql,Rs
		G_USerID = CheckNumeric(G_USerID)
		G_ToolsID = CheckNumeric(G_ToolsID)
		If G_USerID = 0 or G_ToolsID = 0 Then ShowErr(3):Exit Sub
		'ID=0 ,UserID=1 ,UserName=2 ,ToolsID=3 ,ToolsName=4 ,ToolsCount=5 ,SaleCount=6 ,UpdateTime=7 ,SaleMoney=8 ,SaleTicket=9
		Sql = "Select ID,UserID,UserName,ToolsID,ToolsName,ToolsCount,SaleCount,UpdateTime,SaleMoney,SaleTicket From [Dv_Plus_Tools_buss] Where ToolsCount>0 and UserID="& G_USerID &" and ToolsID="& G_ToolsID
		Set Rs = Dvbbs.Plus_Execute(Sql)
		If Rs.Eof Then
			ShowErr(3):Exit Sub
		Else
			Sql = Rs.GetString(,1, "§§§", "", "")
			Sql = Split(Sql,"§§§")
		End If
		Rs.Close : Set Rs = Nothing
		UserToolsInfo = Sql
	End Sub
	'---------------------------------------------------
	'读取目标用户信息
	'---------------------------------------------------
	Public Sub GetToUserInfo(ToUserID)
		Dim Sql,Rs
'		ToUserID = ToUserID
		If ToUserID = 0 Then ShowErr(11):Exit Sub
		'UserID=0,UserName=1,LockUser=2,UserPost=3,UserTopic=4,UserMoney=5,UserTicket=6,userWealth=7,userEP=8,userCP=9,UserPower=10,UserGroupID=11
		Sql = "Select UserID,UserName,LockUser,UserPost,UserTopic,UserMoney,UserTicket,userWealth,userEP,userCP,UserPower,UserGroupID From [Dv_User] Where UserID="& ToUserID
		Set Rs = Dvbbs.Execute(Sql)
		If Rs.Eof Then
			ShowErr(11):Exit Sub
		Else
			Sql = Rs.GetString(,1, "§§§", "", "")
			Sql = Split(Sql,"§§§")
		End If
		Rs.Close : Set Rs = Nothing
		ToUserInfo = Sql
	End Sub
	'---------------------------------------------------
	'检查用户使用道具权限
	'---------------------------------------------------
	Public Sub ChkUseTools()
		If Not IsArray(ToolsInfo) Then GetToolsInfo
		ChkUserGroup
		If Dvbbs.boardID>0 Then Chkboard
		If cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpost").text)<=cCur(ToolsInfo(7)) or cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text)<=cCur(ToolsInfo(8)) or cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text)<=cCur(ToolsInfo(9)) or cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text)<=cCur(ToolsInfo(10))Then ShowErr(12):Exit Sub
		Call GetUserToolsInfo(Dvbbs.UserID,ToolsID)
	End Sub

	'---------------------------------------------------
	'检查目标用户使用道具权限
	'---------------------------------------------------
	Public Sub ChkToUseTools(ToUserID)
		If Not IsArray(ToUserInfo) Then GetToUserInfo(ToUserID)
		If cCur(ToUserInfo(3))<=cCur(ToolsSetting(0)) or cCur(ToUserInfo(7))<=cCur(ToolsSetting(1)) or cCur(ToUserInfo(8))<=cCur(ToolsSetting(2)) or cCur(ToUserInfo(9))<=cCur(ToolsSetting(3)) Then ShowErr(13):Exit Sub
	End Sub

	'---------------------------------------------------
	'检查用户组限制使用道具权限
	'---------------------------------------------------
	Public Sub ChkUserGroup()
		If Not IsArray(ToolsInfo) Then GetToolsInfo
		If Cint(ToolsInfo(3)) = 0 Then ShowErr(6):Exit Sub
		If ToolsInfo(11) = "" or Instr(","& ToolsInfo(11) &",",","& Dvbbs.UserGroupID &",") = 0 Then ShowErr(4):Exit Sub
	End Sub
	'---------------------------------------------------
	'检查版块限制使用道具权限
	'---------------------------------------------------
	Public Sub Chkboard()
		If Not IsArray(ToolsInfo) Then GetToolsInfo
		If ToolsInfo(12) = "" or Instr(","& ToolsInfo(12) &",",","& Dvbbs.boardID &",") = 0 Then ShowErr(5):Exit Sub
	End Sub
	
	Public Property Let buySum(byVal Value)
		buyCount = Value
	End Property

	'---------------------------------------------------
	'检查用户购买道具权限： bType 数字型，为用户选取的购买类型
	'---------------------------------------------------
	Public Sub ChkbuyTools(byval bType)
		Dim CanbuyTools
		CanbuyTools = False
		If bType="" or Not Isnumeric(bType) Then
			bType = -1
		Else
			bType = Cint(bType)
		End If
		If Not IsArray(ToolsInfo) Then GetToolsInfo
		If Int(ToolsInfo(4)) = 0 OR buyCount>Int(ToolsInfo(4)) OR buyCount = 0 Then ShowErr(8):Exit Sub '库存不足
		Select Case Cint(ToolsInfo(14))
			Case 0 '只需金币
				If cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text)>=Int(ToolsInfo(6))*buyCount and bType=0 Then
					CanbuyTools = True
				End If
			Case 1 '只需点券
				If cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)>=Int(ToolsInfo(13))*buyCount and bType=1 Then
					CanbuyTools = True
				End If
			Case 2 '金币+点券
				If cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text)<Int(ToolsInfo(6))*buyCount Or cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)<Int(ToolsInfo(13))*buyCount Then
					CanbuyTools = False
				Else
					CanbuyTools = True
				End If
			Case 3 '金币或点券
				If bType=0 Then
					If cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text)>Int(ToolsInfo(6))*buyCount Then CanbuyTools = True
				ElseIf bType=1 Then
					If cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)>Int(ToolsInfo(13))*buyCount Then CanbuyTools = True
				Else
					CanbuyTools = False
				End If
			Case Else
				ShowErr(10):Exit Sub
		End Select
		If CanbuyTools = False Then ShowErr(7):Exit Sub
	End Sub
	'---------------------------------------------------
	'购买方式
	'---------------------------------------------------
	Public Property Get buyType(byval bType)
		Select Case Cint(bType)
			Case 0 : buyType = "只需金币"
			Case 1 : buyType = "只需点券"
			Case 2 : buyType = "金币+点券"
			Case 3 : buyType = "金币或点券"
			Case Else : buyType = "暂停购买"
		End Select
		buyType = "<font class=redfont>"&buyType&"</font>"
	End Property

	Public Sub ShowErr(byval Code)
		If Code<>"" Then Response.redirect "showerr.asp?ErrCodes="& ErrCodes(Code) &"&action=NoHeadErr"
	End Sub
	'---------------------------------------------------
	'错误信息
	'---------------------------------------------------
	Public Function ErrCodes(byval ErrNum)
		Select Case ErrNum
			Case 1 : ErrCodes = "<li>道具中心已经关闭！</li>"
			Case 2 : ErrCodes = "<li>道具交易中心已经关闭，不能进行道具交易！</li>"
			Case 3 : ErrCodes = "<li>该道具不存在或参数不正确！</li>"
			Case 4 : ErrCodes = "<li>您没有购买或使用该道具的权限！</li>"
			Case 5 : ErrCodes = "<li>本版块不能使用该道具！</li>"
			Case 6 : ErrCodes = "<li>该道具已被系统禁止使用！</li>"
			Case 7 : ErrCodes = "<li>您的金币或点券不足或选取的购买方式不正确，不能购买该道具！</li>"
			Case 8 : ErrCodes = "<li>该道具系统库存不足，暂停购买！</li>"
			Case 9 : ErrCodes = "<li>转让的数量已超过了您拥有的道具数据或没有填写正确的道具数量，出售中止！</li>"
			Case 10 : ErrCodes = "<li>暂停购买！</li>"
			Case 11 : ErrCodes = "<li>道具使用目标用户不存在或参数不正确！"
			Case 12 : ErrCodes = "<li>由于你的文章数或金钱值或积分值或魅力值不足，所以没有使用该道具的权限！</li>"
			Case 13 : ErrCodes = "<li>由于使用的目标用户的文章数或金钱值或积分值或魅力值不足，所以你不能使用该道具！</li>"
			Case 14 : ErrCodes = "<li>此操作不能在相同用户名之间进行！</li>"
			Case 15 : ErrCodes = "<li>后悔药只能用在自己发表的帖子上！</li>"
			Case 16 : ErrCodes = "<li>您设置的转让金币或点券数不正确！</li>"
			Case 17 : ErrCodes = "<li>您的金币或点券不足，不能转让！</li>"
			Case 18 : ErrCodes = "<li>该用户没有任何道具。</li>"
		End Select
	End Function

	Public Function CheckNumeric(byval CHECK_ID)
		If CHECK_ID<>"" and IsNumeric(CHECK_ID) Then _
			CHECK_ID = Int(CHECK_ID) _
		Else _
			CHECK_ID = 0
		CheckNumeric = CHECK_ID
	End Function

End Class

'--------------------------------------------------------------------------------
'用户信息
'--------------------------------------------------------------------------------
Sub UserInfo()
	Dim Sql,Rs,UserToolsCount
	Sql = "Select Sum(ToolsCount) From [Dv_Plus_Tools_buss] where UserID="& Dvbbs.UserID
	Set Rs = Dvbbs.Plus_Execute(Sql)
	UserToolsCount = Rs(0)
	If IsNull(UserToolsCount) Then UserToolsCount = 0
%>
<table border="0" cellpadding="3" cellspacing="1" align="center" class="tableborder1" style="Width:100%">
	<tr>
		<th>个人资料</th>
	</tr>
	<tr>
		<td align="center" class="tablebody1">
			<table border="0" cellpadding="3" cellspacing="1" align="center" style="Width:90%">
				<tr>
					<td class="tablebody2" style="text-align:left">金币：<b>
						<font color="<%=Dvbbs.mainsetting(1)%>">
							<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text%>
						</font></b> 个
					</td>
				</tr>
				<tr>
					<td class="tablebody1" style="text-align:left">点券：<b>
						<font color="<%=Dvbbs.mainsetting(1)%>">
							<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text%>
						</font></b> 张
					</td>
				</tr>
				<tr>
					<td class="tablebody2" style="text-align:left">道具：
						<a href="?action=UserTools_List"><b>
							<font color="<%=Dvbbs.mainsetting(1)%>"><%=UserToolsCount%></font></b></a> 个
					</td>
				</tr>
				<tr>
					<td class="tablebody1" style="text-align:left">金钱：<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text%>
					</td>
				</tr>
				<tr>
					<td class="tablebody2" style="text-align:left">文章：<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpost").text%>
					</td>
				</tr>
				<tr>
					<td class="tablebody1" style="text-align:left">积分：<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text%>
					</td>
				</tr>
				<tr>
					<td class="tablebody2" style="text-align:left">魅力：<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text%>
					</td>
				</tr>
				<tr>
					<td class="tablebody1" style="text-align:left">威望：<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpower").text%>
					</td>
				</tr>
				<tr><td class="tablebody2"></td></tr>
			</table>
		</td>
	</tr>
</table>
<%
End Sub

Sub Tools_Nav_Link()
%>
	<table border="0" width="<%=Dvbbs.mainsetting(0)%>" cellpadding="2" cellspacing="0" align="center">
		<tr>
		<th><a href="plus_Tools_Center.asp">系统交易中心</a></th>
		<th><a href="plus_Tools_Center.asp?action=UserbussTools_List" >用户交易中心</a></th>
		<th ><a href="?action=UserTools_List">我的道具箱</a></th>
		<th><a href="UserPay.asp">购买论坛点券</a></th>
		</tr>
	</table>
<%
End Sub
%>
