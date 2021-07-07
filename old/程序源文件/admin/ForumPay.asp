<!--#include file =../conn.asp-->
<!--#include file="inc/const.asp"-->
<!--#include file="../inc/dv_clsother.asp"-->
<%
Head()
Dim admin_flag
admin_flag=",9,"
CheckAdmin(admin_flag)
If Request("action")="readme" Then
	NetPay()
Else
	Main()
End If
If FoundErr Then Call Dvbbs_Error()
Footer()

Sub Main()
	Dim StartTime,EndTime,sType,KeyWord,IsSuc,MoneySize,PayMoney,SqlString
	StartTime = Request("StartTime")
	EndTime = Request("EndTime")
	sType = Request("sType")
	KeyWord = Replace(Request("keyword"),"'","''")
	MoneySize = Request("MoneySize")
	PayMoney = Request("PayMoney")
	IsSuc = Request("IsSuc")

	If IsSuc = "" Or Not IsNumeric(IsSuc) Then IsSuc = 0
	If IsSuc = 1 Then
		If SqlString = "" Then
			SqlString = " Where O_IsSuc = 1"
		Else
			SqlString = SqlString & " And O_IsSuc = 0"
		End If
	ElseIf IsSuc = 2 Then
	End If
	If StartTime <> "" And IsDate(StartTime) Then
		If SqlString = "" Then
			SqlString = " Where O_AddTime >= '"&StartTime&"'"
		Else
			SqlString = SqlString & " And O_AddTime >= '"&StartTime&"'"
		End If
	End If
	If EndTime <> "" And IsDate(EndTime) Then
		If SqlString = "" Then
			SqlString = " Where O_AddTime <= '"&EndTime&"'"
		Else
			SqlString = SqlString & " And O_AddTime <= '"&EndTime&"'"
		End If
	End If
	If sType = "" Or Not IsNumeric(sType) Then sType = 0
	If sType = 1 Then
		If SqlString = "" Then
			SqlString = " Where O_Type = 1"
		Else
			SqlString = SqlString & " And O_Type = 1"
		End If
	ElseIf sType = 2 Then
		If SqlString = "" Then
			SqlString = " Where O_Type = 2"
		Else
			SqlString = SqlString & " And O_Type = 2"
		End If
	End If
	If KeyWord <> "" Then
		If SqlString = "" Then
			SqlString = " Where (O_UserName Like '%"&keyword&"%' Or O_PayCode Like '%"&keyword&"%')"
		Else
			SqlString = SqlString & " And (O_UserName Like '%"&keyword&"%' Or O_PayCode Like '%"&keyword&"%')"
		End If
	End If
	If MoneySize = "" Or Not IsNumeric(MoneySize) Then MoneySize=0
	MoneySize = Cint(MoneySize)
	If PayMoney <> "" And IsNumeric(PayMoney) Then
		If MoneySize = 0 Then
			If SqlString = "" Then
				SqlString = " Where O_PayMoney > "&PayMoney&""
			Else
				SqlString = SqlString & " And O_PayMoney > "&PayMoney&""
			End If
		Else
			If SqlString = "" Then
				SqlString = " Where O_PayMoney < "&PayMoney&""
			Else
				SqlString = SqlString & " And O_PayMoney < "&PayMoney&""
			End If
		End If
	End If

	Dim Page,MaxRows,Endpage,CountNum,PageSearch
	PageSearch = "StartTime="&StartTime&"&EndTime="&EndTime&"&keyword="&KeyWord&"&sType="&sType&"&IsSuc="&IsSuc&"&MoneySize="&MoneySize&"&PayMoney="&PayMoney&""
	Endpage = 0
	MaxRows = 20
	Page = Request("Page")
	If IsNumeric(Page) = 0 or Page="" Then Page=1
	Page = Clng(Page)
	Response.Write "<script language=""JavaScript"" src=""../inc/Pagination.js""></script>"
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr><th style="text-align:center;">论坛交易信息管理</th></tr>
<tr><td class="td1" style="line-height : 18px ;">
<B>说明</B>：<br />
1、建议您开启网络支付或手机短信通道用于用户购买用户点券，<a href="Challenge.asp"><font color=red>前往网络支付或手机短信设置</font></a><br />
2、论坛中的交易贴、VIP用户、道具中心等交易的货币为论坛金币或点券<br />
3、金币可通过道具中心赠与用户或论坛版主奖励，点券可通过网络支付或手机短信通道购买，详见相关的设置<br />
4、此页面功能为查询网络支付或手机短信购买点券的详细情况，由于论坛存在数据破坏等未知因素，此处数据仅供参考，请管理员参考动网官方相应文档做好论坛安全工作
</td>
</tr>
<FORM METHOD=POST ACTION="ForumPay.asp">
<tr>
<td class="td1" style="line-height : 18px ;">
<B>说明</B>：开始和结束时间不填写为查询所有，下拉菜单中选项不选则查询所有信息，关键字可输入用户名、订单号
</td>
</tr>
<tr>
<td class="td1" style="line-height : 18px ;">
关键字：
<input size=15 name="keyword" type=text value="<%=keyword%>">
开始时间：
<input size=15 name="StartTime" type=text value="<%=StartTime%>">
结束时间：
<input size=15 name="EndTime" type=text value="<%=EndTime%>">
格式：yyyy-mm-dd
</td>
</tr>
<tr>
<td class="td1" style="line-height : 18px ;">
类　型：
<Select Size=1 Name="sType">
<Option value=0>所有</Option>
<Option value=1 <%If sType = 1 Then Response.Write "Selected"%>>网络支付</Option>
<Option value=2 <%If sType = 2 Then Response.Write "Selected"%>>手机短信</Option>
</Select>
交易状态：
<Select Size=1 Name="IsSuc">
<Option value=0>所有</Option>
<Option value=1 <%If IsSuc = 1 Then Response.Write "Selected"%>>成功</Option>
<Option value=2 <%If IsSuc = 2 Then Response.Write "Selected"%>>失败</Option>
</Select>
交易金额：
多于
<input type=radio class="radio" value="0" name="MoneySize" <%If MoneySize = 0 Then Response.Write "Checked"%>>
少于
<input type=radio class="radio" value="1" name="MoneySize" <%If MoneySize = 1 Then Response.Write "Checked"%>>
<input type=text size=10 name="PayMoney" value="<%=PayMoney%>">
<input name="submit" value="提交查询" type=Submit class="button">
</td>
</tr>
</FORM>
</table><br>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr>
<th width="150">用户名</th>
<th>订单号</th>
<th width="35">金额</th>
<th width="70">订单类型</th>
<th width="35">状态</th>
<th width="100">交易时间</th>
</tr>
<%
	'Response.Write SqlString
	Dim PageAmount,AllAmount,Rs,Sql,i
	PageAmount = 0
	AllAmount = 0
	Set Rs = Dvbbs.Execute("Select Sum(O_PayMoney) From Dv_ChanOrders "&SqlString&"")
	AllAmount = Rs(0)
	If IsNull(AllAmount) Then AllAmount = 0
	Set Rs = Dvbbs.iCreateObject ("adodb.recordset")
	Sql = "Select O_PayMoney,O_UserName,O_PayCode,O_Type,O_IsSuc,O_AddTime From Dv_ChanOrders "&SqlString&" Order By O_ID Desc"
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open Sql,conn,1,1
	If Rs.Eof And Rs.Bof Then
		Response.Write "<tr><td class=td1 height=23 colspan=6>未找到相关的交易信息。</td></tr>"
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
		For i=0 To Ubound(SQL,2)
			PageAmount = PageAmount + SQL(0,i)
%>
<tr align=center>
<td class="td1" height=23><a href="../dispuser.asp?name=<%=Server.HtmlEncode(SQL(1,i))%>" target=_blank><%=Server.HtmlEncode(SQL(1,i))%></a></td>
<td class="td1"><%=SQL(2,i)%></td>
<td width="35" class="td1"><%=SQL(0,i)%></td>
<td width="70" class="td1">
<%
Select Case SQL(3,i)
Case 1
	Response.Write "网络支付"
Case 2
	Response.Write "手机短信"
End Select
%>
</td>
<td width="35" class="td1">
<%
Select Case SQL(4,i)
Case 0
	Response.Write "<font color=gray>失败</font>"
Case 1
	Response.Write "成功"
End Select
%>
</td>
<td><%=FormatDateTime(SQL(5,i),2)%>&nbsp;<%=FormatDateTime(SQL(5,i),4)%></td>
</tr>
<%
		Next
	End If
	Rs.Close
	Set Rs=Nothing
	Response.Write "<tr><td class=td1 height=23 colspan=6><B>本页交易金额</B>：<B>"&PageAmount&"</B> 元人民币，<B>本次查询交易总金额</B>：<B>"&AllAmount&"</B> 元人民币</td></tr>"
	Response.Write "</table>"
	PageSearch=Replace(Replace(PageSearch,"\","\\"),"""","\""")
	If CountNum > 0 Then Response.Write "<SCRIPT>PageList("&Page&",3,"&MaxRows&","&CountNum&","""&PageSearch&""",1);</SCRIPT>"
End Sub

Sub NetPay()
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr><th style="text-align:center;">网络支付接口二次开发文档</th></tr>
<tr>
<td class="td1" style="line-height : 18px ;">
<B>重要说明</B>：<br />
1、所有采用此接口网站必须成为动网联盟网站成员，申请时将自动提交部分信息至官方服务器登记<br />
2、所有采用此网络支付接口功能的成员默认为接受动网的相关网络支付条款，动网保留所有的解释权利，请随时留意动网官方以获得最新通知<br />
3、开启此网络支付功能不需开通和服务费，但动网收取一定的手续费，单比交易默认为 <B>2%</B>，默认单比交易最低收费标准为 <B>1</B> 元人民币<br />
4、动网保留随时更改手续费用和根据不同网站采取不同收费策略的行为，更加详细的收费标准和收费策略请随时关注动网官方站点<br />
5、使用者独立承担因其网站信息违法、虚假、陈旧或不详实造成的投诉、退货、纠纷等的责任，更详细说明请参见动网网络支付服务条款
<p></p>
<B>网络支付接口和使用步骤</B>：
<p></p>
<a href="http://www.dvbbs.net/netpay/pay.rar"><font color=red>详细的网络支付接口说明和开发包下载</font></a>
<p></p>
<B>（一）成为动网联盟网站成员</B><br />
访问动网官方站点，申请成为动网网络支付联盟网站成员
<p></p>
<B>（二）提交支付请求页面</B><br />
1、生成订单号，订单号规则为：02 + 联盟网站ID + 时间(yyyymmddhhMMss) + 5位随机数，如第一项中的联盟用户信息所生成的订单可以为020000000012005031513355519597，其中000000001是联盟网站ID，20050315133555就是订单生成的时间（时间中除了年份外其它不足2位数的需补0，如里面的03表示3月），19597是随机数<br />
2、生成正式订单，正式订单包含内容为：订单号（参数为paycode），交易用户名（参数为username，可为空），返回地址（参数为returnurl），支付金额（参数为paymoney）<br />
3、生成的订单信息应记录一份在本地数据库，内容应包含订单号、交易用户、支付金额、订单时间、是否成功等信息<br />
4、将订单内容用post方式提交到：http://server.dvbbs.net/alipay_t1.aspx?action=pay
<p></p>
<B>（三）接收支付结果页面</B><br />
1、Get方式返回的参数：订单状态（参数为success），订单号（参数为paycode），验证字串（参数为sign）<br />
2、验证过程：<br />
　　A. success=1为成功订单，0为失败订单<br />
　　B. 根据订单号查询本地数据库，如订单存在，则进行字符验证过程，首先需要生成一份本地验证字串，规则为MD5进行32位加密以下字符：订单号:返回的订单状态:订单金额:本地验证字串（就是第一项中的加密字符key），注意其中每个项目的分割使用英文状态下的冒号，然后将加密后字符和返回的字符(sign)比较，相同则表示验证通过<br />
3、验证通过后进行相应的用户数据操作并将订单设置为成功状态
<p></p>
<B>（四）重新获得订单状态</B><br />
主要用于因网络故障引起的交易失败<br />
提取对应订单中的相关信息重新生成一份订单，注意相关内容必须沿用原订单信息，包括订单号，金额等信息<br />
将生成的订单信息以post方式提交到：http://server.dvbbs.net/alipay_t1.aspx?action=pay_1
</td>
</tr>
</table><P>
<%
End Sub
%>