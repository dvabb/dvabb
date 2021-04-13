<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/chan_const.asp"-->
<!--#include file="inc/md5.asp"-->
<%
Dvbbs.stats="购买论坛点券"
Dvbbs.LoadTemplates("")
Dvbbs.nav()

Dvbbs.Head_var 0,0,"用户控制面板","usermanager.asp"

If Request("raction")="alipay_return" Then
	AliPay_Return()
	Dvbbs.Footer()
	Response.End
ElseIf Request("action")="alipay_return" Then
	AliPay_Return()
	Dvbbs.Footer()
	Response.End
End If

If Dvbbs.userid=0 Then Dvbbs.AddErrCode(6):Dvbbs.Showerr()
Dvbbs.TrueCheckUserLogin()
CenterMain()
Dvbbs.Showerr()
Dvbbs.Footer()
'Dvbbs.PageEnd()
Sub CenterMain()
%>
	<table border="0" width="<%=Dvbbs.mainsetting(0)%>" cellpadding=2 cellspacing=0 align=center>
		<tr>
		<td width="180" valign=top>
		<%UserInfo()%>
		</td>
		<td width="*" valign=top>
		<%
		Select Case Request.QueryString("action")
			Case "alipay"
				AliPay()
			Case "alipay_1"
				AliPay_1()
			Case "alipay_return"
				AliPay_Return()
			Case "UserCenter"
				UserCenter()
			Case "UserToolsLog_List"
				UserToolsLog_List()
			Case "PayList"
				PayList()
			Case Else
				SmsPayMain()
		End Select
		%>
		</td>
		</tr>
	</table>
<%
End Sub

Sub SmsPayMain()
	MainReadMe(0)

	If Dvbbs.Forum_ChanSetting(3)="0" Then
%>
	<tr><td height=23 class="tablebody2"><B>网络银行支付购买点券</B>：使用前请到 <a href="https://www.alipay.com/" target=_blank><font color=red>阿里巴巴.支付宝</font></a> 申请一个支付宝账号，支付过程不收取手续费</td>
	</tr>
	<FORM TARGET="_blank" METHOD=POST ACTION="?action=alipay">
	<tr><td height=23 class="tablebody1">
	请输入要支付的金额：
	<input type=text size=5 name="paymoney" value="1" onkeyup="ShowChange(this.value,this,'PAY_M',1)">
	获取<FONT ID="PAY_M" CLASS="REDFONT"><%=CCur(Dvbbs.Forum_ChanSetting(14))*1%></FONT>张论坛点券。
	（最低 1 元人民币 ）
	<input type=submit name=submit value="网上支付">
	</td>
	</tr>
	</FORM>
	<tr><td height=24 class="tablebody1">
	<B>您成功支付后有系统可能需要几分钟的时间等待支付结果，因此可能无法瞬间入账，支付成功后请刷新此页面并查看点券数是否正确。</B>
	</td>
	</tr>
	<tr><td height=24 class="tablebody1">
	<iframe src="<%=Dvbbs_Server_Url%>dvbbs/DvDefaultTextAd_1.asp" height=23 width="100%" MARGINWIDTH=0 MARGINHEIGHT=0 HSPACE=0 VSPACE=0 FRAMEBORDER=0 SCROLLING=no></iframe>
	</td>
	</tr>
	<%End If%>
	<tr><td height=23 class="tablebody2" style="line-height: 18px"><B>点券使用小贴士</B>：<BR>
	① 论坛点券可用于购买论坛中出售的各种趣味性道具<BR>
	② 论坛点券和金币可用于参与论坛中一些需要点券购买贴的浏览、当您的帖子有人正确回答后赠与回复用户等操作<BR>
	③ 各种论坛道具有其不同的功能，比如机遇卡可让目标用户（也可是您自己）随机出现一些机遇（如增减金钱获丢失道具等）<BR>
	④ 论坛点券可在论坛用户中相互转让，前提是目标用户必须符合论坛设置以及购买了道具转让器<BR>
	⑤ 系统中部分特殊的道具出于限制使用的目的，是需要用户同时拥有金币和点券才能购买的，有部分道具只有在特殊的情况下才会出现，这部分道具是用点券或金币都不能购买到的。</td>
	</tr>
</table>

<SCRIPT LANGUAGE="JavaScript">
<!--
var ProductMoney = <%=Dvbbs.Forum_ChanSetting(14)%>;
function getinfo(v){
	v=parseFloat(v);
	var pag=document.getElementById('pay');
	pag.innerHTML=ProductMoney*v;
}
function ShowChange(Ivalue,Iname,ShowID,Min){
	if(isNaN(Ivalue)){
		Iname.value = Min;
		alert('请填写正确的数值！');
	}
	else{
		Ivalue = parseFloat(Ivalue);
		Min = parseFloat(Min);
		if (Ivalue<Min){
			Iname.value = Min;
			document.getElementById(ShowID).innerHTML = Min;
			alert('填写数值低于限制！');
		}
		else{
			document.getElementById(ShowID).innerHTML = (Ivalue * ProductMoney).toFixed(1);
		}
	}
}
//-->
</SCRIPT>
<%
End Sub

Sub AliPay()
	Dim PayMoney
	PayMoney = Request("paymoney")
	If PayMoney = "" Or Not IsNumeric(PayMoney) Then
		Response.redirect "showerr.asp?ErrCodes=<li>错误，非法的付款参数。&action=OtherErr"
		Exit Sub
	End If
	If PayMoney < 1 Then
		Response.redirect "showerr.asp?ErrCodes=<li>错误，每笔订单金额最小为 <B>1</B> 元人民币。&action=iOtherErr"
		Exit Sub
	End If
	PayMoney = FormatNumber(PayMoney,2,True,False,False)

	'生成订单号:01+yyyyMMddhhmmss+六位随机数
	'生成日期字串
	Dim NowTimes,PayMonth,PayDay,PayHour,PayMin,PaySe,PayDayStr,RandomizeStr,num1
	Dim PayCode,PayCodeEnCode
	NowTimes = Now()
	PayMonth = Month(NowTimes)
	If Len(PayMonth)=1 Then PayMonth = "0" & PayMonth
	PayDay = Day(NowTimes)
	If Len(PayDay)=1 Then PayDay = "0" & PayDay
	PayHour = Hour(NowTimes)
	If Len(PayHour)=1 Then PayHour = "0" & PayHour
	PayMin = Minute(NowTimes)
	If Len(PayMin)=1 Then PayMin = "0" & PayMin
	PaySe = Second(NowTimes)
	If Len(PaySe)=1 Then PaySe = "0" & PaySe
	PayDayStr = Year(NowTimes) & PayMonth & PayDay & PayHour & PayMin & PaySe
	'生成随机字串
	Randomize
	Do While Len(RandomizeStr)<5
		num1 = CStr(Chr((57-48)*rnd+48))
		RandomizeStr = RandomizeStr & num1
	Loop
	'Response.Write RandomizeStr
	'Response.Write "<BR>"
	'Response.Write PayDayStr
	If Dvbbs.Forum_ChanSetting(5) <> "0" Then
		PayCode = "01" & Dvbbs.Forum_ChanSetting(5) & PayDayStr & RandomizeStr
	Else
		PayCode = PayDayStr & RandomizeStr & Left(MD5(Dvbbs.Forum_ChanSetting(4)&Dvbbs.Forum_ChanSetting(6),32),8)
	End If
	Dim EnCodeStr
	
	EnCodeStr="body=Forum points Certificates&notify_url="&Dvbbs_PayTo_Url&"newpay.asp?action=newpay&out_trade_no="&PayCode&"&partner=2088002048522272&payment_type=1&return_url="&Dvbbs_PayTo_Url&"newpay.asp?action=newpay&seller_email="&Lcase(Dvbbs.Forum_ChanSetting(4))&"&service=create_direct_pay_by_user&show_url="&Dvbbs.Get_ScriptNameUrl&"&subject=Forum points Certificates&total_fee="&PayMoney&Dvbbs.Forum_ChanSetting(6)
	EnCodeStr = MD5(EnCodeStr,32)

	'进入论坛订单库
	Dvbbs.Execute("InSert Into Dv_ChanOrders (O_type,O_Username,O_isApply,O_issuc,O_PayMoney,O_Paycode,O_AddTime) Values (1,'"&Dvbbs.MemberName&"',0,0,"&PayMoney&",'"&PayCode&"','"&NowTimes&"')")

	'提交到动网官方主服务器
	If Dvbbs.Forum_ChanSetting(5) <> "0" Then
%>
正在提交数据，如果您的论坛地址设置了URL转发，将不能正确传输信息，请稍后……
<form name="redir" action="<%=Dvbbs_Server_Url%>alipay_t1.aspx?action=pay" method="post">
<INPUT type=hidden name="username" value="<%=Dvbbs.MemberName%>">
<INPUT type=hidden name="paycode" value="<%=PayCode%>">
<INPUT type=hidden name="returnurl" value="<%=Dvbbs.Get_ScriptNameUrl%>UserPay.asp?action=alipay_return">
<INPUT type=hidden name="paymoney" value="<%=PayMoney%>">
</form>
<script LANGUAGE=javascript>
<!--
redir.submit();
//-->
</script>
<%
	Else
%>
正在提交数据，如果您的论坛地址设置了URL转发，将不能正确传输信息，请稍后……
<form name="redir" action="<%=Dvbbs_PayTo_Url%>newpay.asp?action=pay" method="post">
<INPUT type=hidden name="buyer" value="<%=Dvbbs.MemberName%>">
<INPUT type=hidden name="returnurl" value="<%=Dvbbs.Get_ScriptNameUrl%>">
<INPUT type=hidden name="out_trade_no" value="<%=PayCode%>">
<INPUT type=hidden name="seller_email" value="<%=Lcase(Dvbbs.Forum_ChanSetting(4))%>">
<INPUT type=hidden name="total_fee" value="<%=PayMoney%>">
<INPUT type=hidden name="sign" value="<%=EnCodeStr%>">
</form>
<script LANGUAGE=javascript>
<!--
redir.submit();
//-->
</script>
<%
	End If
End Sub

'在线支付返回结果处理，不登陆也可执行
Sub AliPay_Return()
	If Dvbbs.Forum_ChanSetting(5) <> "0" Then
		AliPay_Return_Old()
		Exit sub
	Else
		Dim Rs,Order_No,EnCodeStr,UserInMoney
		Order_No=Dvbbs.Checkstr(Request("out_trade_no"))
		Set Rs = Dvbbs.Execute("Select * From [Dv_ChanOrders] Where O_IsSuc=3 And O_PayCode='"&Order_No&"'")
		If not(Rs.Eof And Rs.Bof) Then
			AliPay_Return_Old()
			Exit sub
		End If
		Set Rs = Dvbbs.Execute("Select * From [Dv_ChanOrders] Where O_IsSuc=1 And O_PayCode='"&Order_No&"'")
		If not(Rs.Eof And Rs.Bof) Then
			AliPay_Return_Old()
			Exit sub
		End if
		Response.Clear
		Set Rs = Dvbbs.Execute("Select * From [Dv_ChanOrders] Where O_IsSuc=0 And O_PayCode='"&Order_No&"'")
		If Rs.Eof And Rs.Bof Then
			Response.Write "fail"
		Else
			Response.Write "success"
			Dvbbs.Execute("Update Dv_ChanOrders Set O_IsSuc=3 Where O_ID = " & Rs("O_ID"))
		End If
		Response.End
	End If
End Sub

Sub AliPay_Return_Old()		
	'得到和判断返回参数
	Dim PayCode,SignStr,Success,UserInMoney
	PayCode = Replace(Request("out_trade_no"),"'","")
	Success = Request("is_success")
	If PayCode = "" Or Success = "" Then
		Response.redirect "showerr.asp?ErrCodes=<li>错误，非法的订单参数。&action=OtherErr"
		Exit Sub
	End If
	If Success<>"T" Then
		Response.redirect "showerr.asp?ErrCodes=<li>订单支付失败，请详细检查您的支付信息，<a href=""UserPay.asp"">重新进入支付页面</a>。&action=iOtherErr"
		Exit Sub
	End If

	'验证订单信息
	Dim Rs
	Set Rs = Dvbbs.Execute("Select * From [Dv_ChanOrders] Where O_PayCode='"&PayCode&"'")
	If Rs.Eof And Rs.Bof Then
		Response.redirect "showerr.asp?ErrCodes=<li>错误，找不到该订单信息或该订单已支付成功。&action=OtherErr"
		Exit Sub
	Else
		If CInt(rs("O_issuc"))=3 Or CInt(rs("O_issuc"))=1 Then
			dim alipayNotifyURL,ResponseTxt,Retrieval
			alipayNotifyURL="http://notify.alipay.com/trade/notify_query.do?partner=2088002048522272&notify_id="&request("notify_id")
			Set Retrieval=Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
			Retrieval.setOption 2,13056 
			Retrieval.open "GET",alipayNotifyURL,False,"","" 
			Retrieval.send()
			ResponseTxt=Retrieval.ResponseText
			Set Retrieval=Nothing
			If ResponseTxt="false" Then Response.redirect "showerr.asp?ErrCodes=<li>错误，非法的订单参数。&action=OtherErr":Exit Sub
			'更新数据库资料
			UserInMoney = Rs("O_PayMoney")
			If CInt(rs("O_issuc"))=3 then
			'更新用户资料
			Dvbbs.Execute("Update Dv_User Set UserTicket = UserTicket + " & Dvbbs.Forum_ChanSetting(14) * UserInMoney & " Where UserName='"&Rs("O_UserName")&"'")
			If Dvbbs.UserID > 0 And Lcase(Dvbbs.MemberName)=Lcase(Rs("O_UserName")) Then
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text=CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text) + cCur(Dvbbs.Forum_ChanSetting(14) * UserInMoney)
			End If
			'更新订单状态
			Dvbbs.Execute("Update Dv_ChanOrders Set O_IsSuc=1 Where O_ID = " & Rs("O_ID"))
			End if
		Else
			Response.redirect "showerr.asp?ErrCodes=<li>错误，找不到该订单信息或该订单已支付成功。&action=OtherErr"
			Exit Sub
		End if
	End If
	Rs.Close
	Set Rs=Nothing
%>
<!--论坛操作成功信息-->
<br>
<table cellpadding=0 cellspacing=1 align=center class="tableborder1" style="width:75%">
<tr align=center>
<th width="100%">论坛成功信息
</td>
</tr>
<tr>
<td width="100%" class="tablebody1">
<b>操作成功：</b><br><br>
<li>成功，您本次兑换了 <B><font color=red><%=(Dvbbs.Forum_ChanSetting(14) * UserInMoney)%></font></B> 张论坛点券。
</td></tr>
<tr align=center><td width="100%" class="tablebody2">
<a href="usermanager.asp"> << 返回用户控制面板</a> &nbsp;&nbsp;||&nbsp;&nbsp; <a href="UserPay.asp?action=UserCenter"> 去把点券转换成论坛金币>></a> 
</td></tr>
</table><br>
<%
End sub
'--------------------------------------------------------------------------------
'用户信息
'--------------------------------------------------------------------------------
Sub UserInfo()
	Dim Sql,Rs,UserToolsCount
	'Sql = "Select Sum(ToolsCount) From [Dv_Plus_Tools_Buss] where UserID="& Dvbbs.UserID
	'Set Rs = Dvbbs.Plus_Execute(Sql)
	'UserToolsCount = Rs(0)
	'If IsNull(UserToolsCount) Then UserToolsCount = 0
%>
<table cellpadding="0" cellspacing="1" align="center" class="tableborder1" Style="Width:100%">
	<tr>
		<th height=23 >个人资料</th>
	</tr>
	<tr>
		<td align=center class="tablebody1">
			<table border="0" cellpadding="0" cellspacing="1" align="center" Style="Width:90%">
				<tr>
					<td class="tablebody2" style="text-align:left;">金币：
						<B>
							<font color="<%=Dvbbs.mainsetting(1)%>">
								<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text %>
							</font>
						</B> 个
					</td>
				</tr>
				<tr>
					<td class="tablebody1" style="text-align:left;">点券：<B>
						<font color="<%=Dvbbs.mainsetting(1)%>">
							<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text%>
						</font></B> 张
					</td>
				</tr>
				<tr>
					<td class="tablebody2" style="text-align:left;">金钱：<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text%></td>
				</tr>
				<tr>
					<td class="tablebody1" style="text-align:left;">文章：<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpost").text%></td>
				</tr>
				<tr>
					<td class="tablebody2" style="text-align:left;">积分：<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text%></td>
				</tr>
				<tr>
					<td class="tablebody1" style="text-align:left;">魅力：<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text%></td>
				</tr>
				<tr>
					<td class="tablebody2" style="text-align:left;">威望：<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpower").text%></td>
				</tr>
				<tr><td class="tablebody1"></td></tr>
			</table>
		</td>
	</tr>
</table>
<%
End Sub

'--------------------------------------------------------------------------------
'金币转换
'--------------------------------------------------------------------------------
Sub UserCenter()
	If Request("react") = "Savechange" Then
		If Not Dvbbs.ChkPost() Then Dvbbs.AddErrCode(16):Dvbbs.Showerr()
		Dim userWealth,userep,usercp,userticket,UpUserMoney
		Dim Sql,Rs
		userWealth = Dvbbs.CheckNumeric(Request.Form("userWealth"))
		userep = Dvbbs.CheckNumeric(Request.Form("userep"))
		usercp = Dvbbs.CheckNumeric(Request.Form("usercp"))
		userticket = Dvbbs.CheckNumeric(Request.Form("userticket"))
		UpUserMoney = 0
		If userWealth<0 or userep<0 or usercp<0 or userticket<0 Then Dvbbs.AddErrCode(35):Dvbbs.Showerr()

		Dim ErrMsg
		If userWealth>0 And CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text)<CCur(Dvbbs.Forum_setting(93)) Then ErrMsg="你的金钱不足。"
		If userep>0 And CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text)<CCur(Dvbbs.Forum_setting(94)) Then ErrMsg="你的积分不足。"
		If usercp>0 And CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text)<CCur(Dvbbs.Forum_setting(95)) Then ErrMsg="你的魅力不足。"
		If userticket And CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)<CCur(Dvbbs.Forum_setting(96)) Then ErrMsg="你的点券不足。"
		If Trim(ErrMsg)<>"" Then
			Response.redirect "showerr.asp?ErrCodes=<li>"&ErrMsg&"&action=OtherErr"
		End If

		If userWealth>=1 and userWealth<=CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text) and cCur(Dvbbs.Forum_setting(93))<>0 Then
			If Cint(userWealth / cCur(Dvbbs.Forum_setting(93))) > 0 Then
				UpUserMoney = UpUserMoney + Cint(userWealth / cCur(Dvbbs.Forum_setting(93)))
				userWealth = Cint(userWealth / cCur(Dvbbs.Forum_setting(93))) * cCur(Dvbbs.Forum_setting(93))
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text = cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text) - userWealth
			Else
				userWealth = 0
			End If
		Else
			userWealth = 0
		End If

		If userep>=1 and userep<=cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text) and cCur(Dvbbs.Forum_setting(94))<>0 Then
			If Cint(userep / cCur(Dvbbs.Forum_setting(94))) > 0 Then
				UpUserMoney = UpUserMoney + Cint(userep / cCur(Dvbbs.Forum_setting(94)))
				userep = Cint(userep / cCur(Dvbbs.Forum_setting(94))) * cCur(Dvbbs.Forum_setting(94))
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text = cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text) - userep
			Else
				userep = 0
			End If
		Else
			userep = 0
		End If
		If usercp>=1 and usercp<=cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text) and cCur(Dvbbs.Forum_setting(95))<>0 Then
			If Cint(usercp / cCur(Dvbbs.Forum_setting(95))) > 0 Then
				UpUserMoney = UpUserMoney + Cint(usercp / cCur(Dvbbs.Forum_setting(95)))
				usercp = Cint(usercp / cCur(Dvbbs.Forum_setting(95))) * cCur(Dvbbs.Forum_setting(95))
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text = cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text) - usercp
			Else
				usercp = 0
			End If
		Else
			usercp = 0
		End If
		If userticket>=1 and userticket<=cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text) and Dvbbs.Forum_setting(96) <> 0 Then
			Userticket = Clng(Userticket)
			If Cint(userticket / Dvbbs.Forum_setting(96)) > 0 Then
				UpUserMoney = UpUserMoney + Cint(userticket / Dvbbs.Forum_setting(96))
				userticket = Cint(userticket / Dvbbs.Forum_setting(96)) * Dvbbs.Forum_setting(96)
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text = cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text) - userticket
			Else
				userticket = 0
			End If
		Else
			userticket = 0
		End If
		If UpUserMoney < 1 Then 
			 Response.redirect "showerr.asp?ErrCodes=<li>请填写转换的数据或获得的金币数太少！&action=OtherErr"
		Else
			Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text  = cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text ) + UpUserMoney
			Sql = "Update Dv_user set userWealth = "&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text&",userEP="&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text&",userCP="&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text&",UserMoney="&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text &",UserTicket="&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text&" where UserID="&Dvbbs.UserID
			Dvbbs.Execute(Sql)
			Dim LogMsg
			LogMsg = "金币转换成功，获得总金币数为<b>"&UpUserMoney&"</b>,金钱减少<b>"&userWealth&"</b>,积分减少<b>"&userep&"</b>,魅力减少<b>"&usercp&"</b>,点券减少<b>"&userticket&"</b>。"
			'Call Dvbbs.ToolsLog(0,0,0,0,0,LogMsg,Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text &"|"&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)
			Dvbbs.Dvbbs_Suc(LogMsg)
		End If
	Else
%>
	<table border=0 cellpadding=3 cellspacing=1 class="tableborder1" align=center style="width:100%">
	<tr><th height=20 colspan="5">论坛金币转换</th></tr>
	<tr><td height=20 colspan="5" class="tablebody1"><li>允许用户将金钱、积分、魅力、点券转换成金币。</td></tr>
    <tr>
      <th width="30%" height="20">金币转换汇率</th>
      <th width="15%">转换项目</th>
	  <th width="20%">转换信息</th>
      <th width="15%">转换设置</th>
	  <th width="20%">转换所得金币</th>
    </tr>
	<form action="UserPay.asp?action=UserCenter&react=Savechange" method=post NAME=CenterForm>
    <tr>
      <td rowspan="5" class="tablebody1">
		<table border="0" cellpadding=0 cellspacing=1 align=center Style="Width:90%">
			<tr><td class="tablebody1">&nbsp;&nbsp;&nbsp;&nbsp;<a href="UserPay.asp"><font color=red>前往购买论坛点券</font></a></td></tr>
			<tr><td class="tablebody2">&nbsp;&nbsp;&nbsp;&nbsp;<b><font class=redfont>1</font> 金币 = <font class=redfont><%=Dvbbs.Forum_setting(93)%></font> 金钱</b></td></tr>
			<tr><td class="tablebody1">&nbsp;&nbsp;&nbsp;&nbsp;<b><font class=redfont>1</font> 金币 = <font class=redfont><%=Dvbbs.Forum_setting(94)%></font> 积分</b></td></tr>
			<tr><td class="tablebody2">&nbsp;&nbsp;&nbsp;&nbsp;<b><font class=redfont>1</font> 金币 = <font class=redfont><%=Dvbbs.Forum_setting(95)%></font> 魅力</b></td></tr>
			<tr><td class="tablebody1">&nbsp;&nbsp;&nbsp;&nbsp;<b><font class=redfont>1</font> 金币 = <font class=redfont><%=Dvbbs.Forum_setting(96)%></font> 点券</b></td></tr>
			<tr><td class="tablebody2"></td></tr>
		</table>
	  </td>
      <td class="tablebody2" align=center>拥有金钱值：</td>
      <td class="tablebody1"><font class=redfont><%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text%></font></td>
	  <td class="tablebody1"><INPUT TYPE="text" NAME="userWealth" value="0" onkeyup="ShowChange(this.value,this,'Show_Money',<%=Dvbbs.Forum_setting(93)%>,<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text%>)"></td>
	  <td class="tablebody1" ID=Show_Money>0</td>
    </tr>
    <tr>
      <td class="tablebody2" align=center>拥有积分值：</td>
      <td class="tablebody1"><font class=redfont><%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text%></font></td>
	  <td class="tablebody1"><INPUT TYPE="text" NAME="userep" value="0" onkeyup="ShowChange(this.value,this,'Show_EP',<%=Dvbbs.Forum_setting(94)%>,<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text%>)"></td>
	  <td class="tablebody1" ID=Show_EP>0</td>
    </tr>
    <tr>
      <td class="tablebody2" align=center>拥有魅力值：</td>
      <td class="tablebody1"><font class=redfont><%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text%></font></td>
	  <td class="tablebody1"><INPUT TYPE="text" NAME="usercp" value="0" onkeyup="ShowChange(this.value,this,'Show_CP',<%=Dvbbs.Forum_setting(95)%>,<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text%>)"></td>
	  <td class="tablebody1" ID=Show_CP>0</td>
    </tr>
    <tr>
      <td class="tablebody2" align=center>拥有点券值：</td>
      <td class="tablebody1"><font class=redfont><%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text%></font></td>
	  <td class="tablebody1"><INPUT TYPE="text" NAME="userticket" value="0" onkeyup="ShowChange(this.value,this,'Show_Ticket',<%=Dvbbs.Forum_setting(96)%>,<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text%>)"></td>
	  <td class="tablebody1" ID=Show_Ticket>0</td>
    </tr>
	<tr>
      <td class="tablebody2" align=center colspan="4">
	  <INPUT TYPE="submit" value="确定转换">&nbsp;&nbsp;<INPUT TYPE="reset" value="重新设置"></td>
    </tr>
	</form>
	</table>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	function ShowChange(Ivalue,Iname,ShowID,Sys,User){
		if(isNaN(Ivalue)){
			Iname.value = 0;
			alert('请填写正确的数值！');
		}
		else{
			Ivalue = parseFloat(Ivalue);
			Sys = parseFloat(Sys);
			User = parseFloat(User);
			if (Ivalue>User||Ivalue<0){
				Iname.value = 0;
				document.getElementById(ShowID).innerHTML = 0;
				alert('填写数值超过限制！');
			}
			else{
				document.getElementById(ShowID).innerHTML = (Ivalue / Sys).toFixed(1);
			}
		}
	}
	//-->
	</SCRIPT>
<%
	End If
End Sub

'用户订单列表
Sub PayList()
	Dim Success
	Success = Dvbbs.CheckNumeric(Request("Suc"))

	Dim Page,MaxRows,Endpage,CountNum,PageSearch,SqlString
	PageSearch = "action=PayList&Suc=" & Success
	Endpage = 0
	MaxRows = 20
	Page = Request("Page")
	If IsNumeric(Page) = 0 or Page="" Then Page=1
	Page = Clng(Page)
	Response.Write "<script language=""JavaScript"" src=""inc/Pagination.js""></script>"

	MainReadMe(1)
%>
		</td>
		</tr>
		<tr><td colspan=3><hr style="BORDER: #807d76 1px dotted;height:1px;">


<table border="0" cellpadding=3 cellspacing=1 align=center class="tableborder1" style="width:100%">
	<tr><td height=23 class="tablebody2" colspan=6 style="line-height: 18px">
<%
	Dim Rs,Sql
	Select Case Success
	Case 0
		Response.Write Dvbbs.MemberName & " 的所有论坛网络支付交易订单"
		Sql = "Select O_Type,O_PayCode,O_PayMoney,O_IsSuc,O_AddTime,O_ID From Dv_ChanOrders Where O_UserName = '"&Dvbbs.MemberName&"' Order By O_AddTime Desc"
	Case 1
		Response.Write Dvbbs.MemberName & " 的所有论坛网络支付交易成功订单"
		Sql = "Select O_Type,O_PayCode,O_PayMoney,O_IsSuc,O_AddTime,O_ID From Dv_ChanOrders Where O_IsSuc = 1 And O_UserName = '"&Dvbbs.MemberName&"' Order By O_AddTime Desc"
	Case 2
		Response.Write Dvbbs.MemberName & " 的所有论坛网络支付交易失败订单"
		Sql = "Select O_Type,O_PayCode,O_PayMoney,O_IsSuc,O_AddTime,O_ID From Dv_ChanOrders Where O_IsSuc = 0 And O_UserName = '"&Dvbbs.MemberName&"' Order By O_AddTime Desc"
	End Select
%>
	</td></tr>
	<tr>
	<th height=23 width="15%">订单类型</th>
	<th width="20%">订单号</th>
	<th width="15%">支付金额</th>
	<th width="15%">交易状态</th>
	<th width="15%">交易时间</th>
	<th width="20%">操作</th>
	</tr>
<%
	Dim i
	Set Rs = server.CreateObject ("adodb.recordset")
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open Sql,Conn,1,1
	If Rs.Eof And Rs.Bof Then
		Response.Write "<tr><td height=23 class=""tablebody1"" colspan=6>当前还没有订单。</td></tr>"
		Response.Write "</table>"
	Else
		CountNum = Rs.RecordCount
		If CountNum Mod MaxRows=0 Then
			Endpage = CountNum \ MaxRows
		Else
			Endpage = CountNum \ MaxRows+1
		End If
		Rs.MoveFirst
		If Page > Endpage Then Page = Endpage
		If Page < 1 Then Page = 1
		If Page >1 Then 				
			Rs.Move (Page-1) * MaxRows
		End if
		SQL=Rs.GetRows(MaxRows)
		'O_Type,O_PayCode,O_PayMoney,O_IsSuc,O_AddTime,O_ID
		For i=0 To Ubound(SQL,2)
%>
	<tr align=center>
	<td height=23 class="tablebody1">
<%
	Select Case SQL(0,i)
	Case 1
		Response.Write "网络支付"
	Case Else
		Response.Write "<font color=gray>未知</font>"
	End Select
%>
	</td>
	<td class="tablebody1"><%=SQL(1,i)%></td>
	<td class="tablebody1"><%=SQL(2,i)%></td>
	<td class="tablebody1">
<%
	Select Case SQL(3,i)
	Case 0
		Response.Write "<font color=gray>失败</font>"
	Case 1
		Response.Write "成功"
	Case Else
		Response.Write "<font color=gray>未知</font>"
	End Select
%>
	</td>
	<td class="tablebody1"><%=SQL(4,i)%></td>
	<td class="tablebody1">&nbsp;
	</td>
	</tr>
<%
		Next
	Response.Write "</table>"
	PageSearch=Replace(Replace(PageSearch,"\","\\"),"""","\""")
	Response.Write "<SCRIPT>PageList("&Page&",3,"&MaxRows&","&CountNum&","""&PageSearch&""",1);</SCRIPT>"
	End If
	Rs.Close
	Set Rs=Nothing

End Sub

'重新获得交易状态
Sub AliPay_1()
	Dim ID,Rs
	Dim PayMoney,PayCode
	ID = Request("ID")
	If ID = "" Or Not IsNumeric(ID) Then
		Response.redirect "showerr.asp?ErrCodes=<li>错误，非法的订单参数。&action=OtherErr"
		Exit Sub
	Else
		ID = cCur(ID)
	End If
	Set Rs = Dvbbs.Execute("Select * From Dv_ChanOrders Where O_ID = "&ID&" And O_UserName = '"&Dvbbs.MemberName&"'")
	If Rs.Eof And Rs.Bof Then
		Response.redirect "showerr.asp?ErrCodes=<li>错误，找不到相关的订单信息。&action=OtherErr"
		Exit Sub
	Else
		PayMoney = Rs("O_PayMoney")
		PayMoney = FormatNumber(PayMoney,2,True,False,False)
		PayCode = Rs("O_PayCode")
	End If
	Rs.Close
	Set Rs=Nothing
	'提交到动网官方主服务器
%>
正在提交数据，如果您的论坛地址设置了URL转发，将不能正确传输信息，请稍后……
<form name="redir" action="<%=Dvbbs_Server_Url%>alipay_t1.aspx?action=pay_1" method="post">
<INPUT type=hidden name="username" value="<%=Dvbbs.MemberName%>">
<INPUT type=hidden name="paycode" value="<%=PayCode%>">
<INPUT type=hidden name="returnurl" value="<%=Dvbbs.Get_ScriptNameUrl%>UserPay.asp?action=alipay_return">
<INPUT type=hidden name="paymoney" value="<%=PayMoney%>">
</form>
<script LANGUAGE=javascript>
<!--
redir.submit();
//-->
</script>
<%
End Sub

Sub UserToolsLog_List()

	Dim Rs,Sql,i,LogType
	Dim Page,MaxRows,Endpage,CountNum,PageSearch,SqlString
	LogType = "未知|使用|转让|充值|购买|奖励|VIP交易"
	LogType = Split(LogType,"|")
	PageSearch = "action=UserToolsLog_List"
	Endpage = 0
	MaxRows = 20
	Page = Request("Page")
	If IsNumeric(Page) = 0 or Page="" Then Page=1
	Page = Clng(Page)
	Response.Write "<script language=""JavaScript"" src=""inc/Pagination.js""></script>"

	If Request.QueryString("UserID")<>"" and IsNumeric(Request.QueryString("UserID")) Then _
	SqlString = "and UserID="&Dvbbs.CheckNumeric(Request.QueryString("UserID"))

	MainReadMe(1)
%>
		</td>
		</tr>
		<tr><td colspan=3><hr style="BORDER: #807d76 1px dotted;height:1px;">
<table border="0" cellpadding=3 cellspacing=1 align=center class="tableborder1" Style="Width:100%">
	<tr>
	<th height=23 width="15%">道具名称</th>
	<th width="10%">操作</th>
	<th width="*%">操作内容</th>
	<th width="5%">金币</th>
	<th width="5%">点券</th>
	<th width="5%">数量</th>
	<th width="13%">使用IP</th>
	<th width="12%">时间</th>
	</tr>
<%
	Dim ToolsNames
	Dvbbs.forum_setting(90)=0
	If Dvbbs.forum_setting(90)="1" Then
		Set Rs = Dvbbs.Plus_Execute("Select ID,ToolsName From Dv_Plus_Tools_Info Order By ID")
		If Not (Rs.Eof And Rs.Bof) Then
			Sql = Rs.GetRows(-1)
		End If
		Rs.Close
		Set ToolsNames = server.CreateObject ("adodb.recordset")
		For i=0 to Ubound(Sql,2)
			ToolsNames.add Sql(0,i),Sql(1,i)
		Next
		ToolsNames.add -88,"魔法表情或头像"		'添加道具名魔法表情或头像，ID为-88
	End If

	'T.ToolsName=0,L.CountNum=1,L.Log_Money=2,L.Log_Ticket=3,L.Log_IP=4,L.Log_Time=5,L.Log_Type=6,L.Conect=7
	Sql = "Select ToolsID,CountNum,Log_Money,Log_Ticket,Log_IP,Log_Time,Log_Type,Conect From Dv_MoneyLog Where AddUserID="&Dvbbs.UserID&" And Not BoardID=-1 Order By Log_Time Desc"
	'Response.Write Sql
	Set Rs = server.CreateObject ("adodb.recordset")
	If Cint(Dvbbs.Forum_Setting(92))=1 Then
		If Not IsObject(Plus_Conn) Then Plus_ConnectionDatabase
		Rs.Open Sql,Plus_Conn,1,1
	Else
		If Not IsObject(Conn) Then ConnectionDatabase
		Rs.Open Sql,conn,1,1
	End If

	If Not (Rs.Eof And Rs.Bof) Then
		CountNum = Rs.RecordCount
		If CountNum Mod MaxRows=0 Then
			Endpage = CountNum \ MaxRows
		Else
			Endpage = CountNum \ MaxRows+1
		End If
		Rs.MoveFirst
		If Page > Endpage Then Page = Endpage
		If Page < 1 Then Page = 1
		If Page >1 Then 				
			Rs.Move (Page-1) * MaxRows
		End if
		SQL=Rs.GetRows(MaxRows)
	Else
		Response.Write "<tr><td class=""Tablebody1"" colspan=""8"" align=center>道具还未添加！</td></tr></table>"
		Exit Sub
	End If
	Rs.close:Set Rs = Nothing
	
	'输出道具列表
	For i=0 To Ubound(SQL,2)
%>
	<tr>
	<td class="Tablebody1" align=center height=24>
<%
	If Dvbbs.forum_setting(90)="1" Then
		Response.Write ToolsNames(SQL(0,i))
	Else
		Response.Write "<font color=gray>未知</font>"
	End If
%>
	</td>
	<td class="Tablebody1" align=center><%=LogType(SQL(6,i))%></td>
	<td class="Tablebody1"><%=SQL(7,i)%></td>
	<td class="Tablebody1" align=center><%=SQL(2,i)%></td>
	<td class="Tablebody1" align=center><%=SQL(3,i)%></td>
	<td class="Tablebody1" align=center><%=SQL(1,i)%></td>
	<td class="Tablebody1" align=center><%=SQL(4,i)%></td>
	<td class="Tablebody1" align=center><%=SQL(5,i)%></td>
	</tr>
<%
	Next
	Set ToolsNames = Nothing
	Response.Write "</table>"
	PageSearch=Replace(Replace(PageSearch,"\","\\"),"""","\""")
	Response.Write "<SCRIPT>PageList("&Page&",3,"&MaxRows&","&CountNum&","""&PageSearch&""",1);</SCRIPT>"
End Sub

Sub MainReadMe(str)
%>
<table border="0" cellpadding=0 cellspacing=1 align=center class="tableborder1" Style="Width:100%">
	<tr>
	<th height=23>购买论坛点券</th></tr>
	<tr><td height=24 class="tablebody2" align=center><a href="?action=PayList">所有交易记录</a> | <a href="?action=PayList&Suc=1">已成功订单</a> | <a href="?action=PayList&Suc=2">未成功订单</a> | <a href="?action=UserToolsLog_List">金币或点券使用记录</a> | <a href="?action=UserCenter"><font color=red>兑换论坛金币</font></a> | <a href="UserPay.asp"><font color=red>购买论坛点券</font></a></td>
	</tr>
	<tr><td height=23 class="tablebody1" style="line-height: 18px"><B>说明</B>：<BR>
	① 通过网络支付可获<font color=red>奖励</font>相应的论坛点券<BR>
	② 每通过网络支付 <font color=red><B>1</B></font> 元可获奖励 <font color=red><B><%=Dvbbs.Forum_ChanSetting(14)%></B></font> 张论坛点券<BR>
	③ 论坛点券的作用：可购买论坛中各种趣味道具，享受更多有趣的论坛功能<BR>
	④ 点券的获取流程：根据下面提示选择网络支付后，通过网络)%></td>
	</tr>
<%
	Next
	Set ToolsNames = Nothing
	Response.Write "</table>"
	PageSearch=Replace(Replace(PageSearch,"\","\\"),"""","\""")
	Response.Write "<SCRIPT>PageList("&Page&",3,"&MaxRows&","&CountNum&","""&PageSearch&""",1);</SCRIPT>"
End Sub

Sub MainReadMe(str)
%>
<table border="0" cellpadding=0 cellspacing=1 align=center class="tableborder1" Style="Width:100%">
	<tr>
	<th height=23>璐