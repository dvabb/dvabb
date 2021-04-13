<!--#include File="../Conn.Asp"-->
<!-- #include File="Inc/Const.Asp" -->
<!-- #include File="../Inc/Md5.Asp" -->
<!-- #include File="../Inc/Chkinput.Asp" -->
<%
Dim Admin_flag,Forum_api,CnzzWss_api,Action,Rs,Sql,CnzzWssID,CnzzWssPassword,CnzzWssIsOpen,Xmldoc,Xmldom
Admin_flag=",2,"
Checkadmin(Admin_flag)
Const DV_CNZZ_RSA_PUBLIC_KEY = "KslLiq8H"
CnzzWssID = ""
CnzzWssPassword = ""
CnzzWssIsOpen = 0

Call Head()
Call PageLoad()
Call Footer()
Set Forum_api = Nothing
Set CnzzWss_api = Nothing

Sub PageLoad()
	Action = request("action")
	Chkforum_api()
	Select Case Action
		Case "savereg"
			Dim arr
			arr=Split(request("wss"),",")
			CnzzWssID=arr(0)
			CnzzWssPassword=arr(1)
			CnzzWssIsOpen = 0
			Call CnzzWss_Save()
		Case "saveconfig"
			CnzzWssIsOpen = Dvbbs.CheckNumeric(request("CnzzWssIsOpen"))
			Call CnzzWss_Save()
			Call Update_CopyRight(Replace(request("CnzzWssCode"), "'", """"))
		Case "register"
			CnzzWssID = ""
			CnzzWssIsOpen = 0
			Call CnzzWss_Save()
			Call Update_CopyRight(Replace(request("CnzzWssCode"), "'", """"))
	End Select 
	If ""=CnzzWssID Then 
		Call TemplateApply()
	Else 
		Call TemplateDefault()
	End If 
End Sub 

Sub Chkforum_api()
	Set Rs = Dvbbs.Execute("Select Top 1 Forum_apis From Dv_setup")
	Xmldoc = Rs(0)
	Rs.Close
	Set Rs = Nothing
	If Isnull(Xmldoc) Or Xmldoc = "" Then
		Creat_forum_api("Y")
	Else
		Set Forum_api = Server.Createobject("Msxml2.Freethreadeddomdocument"& Msxmlversion)
		Forum_api.Loadxml(Xmldoc)
		Set CnzzWss_api = Forum_api.Documentelement.Selectsinglenode("CnzzWss")
		Testapi()
		CnzzWssID = CnzzWss_api.Getattribute("CnzzWssID")
		CnzzWssPassword = CnzzWss_api.Getattribute("CnzzWssPassword")
		CnzzWssIsOpen = CnzzWss_api.Getattribute("CnzzWssIsOpen")
	End If
End Sub

Sub Update_CopyRight(sCode)
	Dim strSetting,arr,i
	Set rs=server.CreateObject("ADODB.RecordSet")
	rs.Open "select forum_setting from dv_setup", Conn, 1,3
	If Not rs.eof Then 
		strSetting = rs(0)
		arr=split(strSetting,"|||")
		strSetting = arr(0) & "|||" & arr(1) & "|||" & arr(2)
		Dim re
		Set re=new RegExp
		re.IgnoreCase =true
		re.Global=True
		re.Pattern ="<!--cnzz_wss_code_begin-->(.+)<!--cnzz_wss_code_end-->"
		arr(3) = re.Replace(arr(3), "")
		Set re=Nothing 
		If 1=CnzzWssIsOpen Then arr(3) = arr(3) & "<!--cnzz_wss_code_begin-->" & sCode & "<!--cnzz_wss_code_end-->"
		For i=3 To UBound(arr)
			strSetting = strSetting & "|||" & arr(i)
		Next 
		rs(0)=strSetting
		rs.update()
	End If 
	rs.Close
	Set rs=Nothing 
End Sub 

Sub Testapi()
	On Error Resume Next
	Dim testcc
	testcc=CnzzWss_api.Getattribute("CnzzWssID")
	If Err Then
		Creat_forum_api("")
	End If
End Sub
Sub Creat_forum_api(Str)
	If Str="Y" Then
		Set Forum_api = Server.Createobject("Msxml2.Freethreadeddomdocument"& Msxmlversion)
		Forum_api.Loadxml("<Forum_api/>")
	End If
	Set CnzzWss_api = Forum_api.Documentelement.Appendchild(Forum_api.Createnode(1,"CnzzWss",""))
	CnzzWss_api.Setattribute "CnzzWssID",CnzzWssID
	CnzzWss_api.Setattribute "CnzzWssPassword",CnzzWssPassword
	CnzzWss_api.Setattribute "CnzzWssIsOpen",CnzzWssIsOpen
	Update_forum_api()
End Sub
Sub Update_forum_api()
	Dvbbs.Execute("Update Dv_setup Set Forum_apis='"&Dvbbs.Checkstr(Forum_api.Xml)&"'")
End Sub
Sub CnzzWss_Save()
	CnzzWss_api.Setattribute "CnzzWssID",CnzzWssID
	CnzzWss_api.Setattribute "CnzzWssPassword",CnzzWssPassword
	CnzzWss_api.Setattribute "CnzzWssIsOpen",CnzzWssIsOpen
	Update_forum_api()
End Sub 

Sub TemplateApply()
%>
<table width="100%"  border="0" cellspacing="0" cellpadding="0" class="tableborder" style="h">
  <tr>
    <td colspan="2" class="tdtit">
       正在开通论坛统计,请稍候......
    </td>
  </tr>
</table>
<script>
function   utf8to16(str)   {   
	var   out,   i,   len,   c;   
	var   char2,   char3;   

	out   =   "";   
	len   =   str.length;   
	i   =   0;   
	while(i   <   len)   {   
	c   =   str.charCodeAt(i++);   
	switch(c   >>   4)   
	{     
	case   0:   case   1:   case   2:   case   3:   case   4:   case   5:   case   6:   case   7:   
	//   0xxxxxxx   
	out   +=   str.charAt(i-1);   
	break;   
	case   12:   case   13:   
	//   110x   xxxx       10xx   xxxx   
	char2   =   str.charCodeAt(i++);   
	out   +=   String.fromCharCode(((c   &   0x1F)   <<   6)   |   (char2   &   0x3F));   
	break;   
	case   14:   
	//   1110   xxxx     10xx   xxxx     10xx   xxxx   
	char2   =   str.charCodeAt(i++);   
	char3   =   str.charCodeAt(i++);   
	out   +=   String.fromCharCode(((c   &   0x0F)   <<   12)   |   
	((char2   &   0x3F)   <<   6)   |   
	((char3   &   0x3F)   <<   0));   
	break;   
	}   
	}   
	return   out;   
}   

var base64DecodeChars = [
    -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
    -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
    -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, 62, -1, -1, -1, 63,
    52, 53, 54, 55, 56, 57, 58, 59, 60, 61, -1, -1, -1, -1, -1, -1,
    -1,  0,  1,  2,  3,  4,  5,  6,  7,  8,  9, 10, 11, 12, 13, 14,
    15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, -1, -1, -1, -1, -1,
    -1, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40,
    41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, -1, -1, -1, -1, -1
];

function base64decode(str) {
    var c1, c2, c3, c4;
    var i, j, len, out;

    len = str.length;
    i = j = 0;
    out = [];
    while (i < len) {
        /* c1 */
        do {
            c1 = base64DecodeChars[str.charCodeAt(i++) & 0xff];
        } while (i < len && c1 == -1);
        if (c1 == -1) break;

        /* c2 */
        do {
            c2 = base64DecodeChars[str.charCodeAt(i++) & 0xff];
        } while (i < len && c2 == -1);
        if (c2 == -1) break;

        out[j++] = String.fromCharCode((c1 << 2) | ((c2 & 0x30) >> 4));

        /* c3 */
        do {
            c3 = str.charCodeAt(i++) & 0xff;
            if (c3 == 61) return out.join('');
            c3 = base64DecodeChars[c3];
        } while (i < len && c3 == -1);
        if (c3 == -1) break;

        out[j++] = String.fromCharCode(((c2 & 0x0f) << 4) | ((c3 & 0x3c) >> 2));

        /* c4 */
        do {
            c4 = str.charCodeAt(i++) & 0xff;
            if (c4 == 61) return out.join('');
            c4 = base64DecodeChars[c4];
        } while (i < len && c4 == -1);
        if (c4 == -1) break;
        out[j++] = String.fromCharCode(((c3 & 0x03) << 6) | c4);
    }
    return out.join('');
}

function unserialize(ss) { 
    var p = 0, ht = [], hv = 1; r = null; 
    function unser_null() { 
        p++; 
        return null; 
    } 
    function unser_boolean() { 
        p++; 
        var b = (ss.charAt(p++) == '1'); 
        p++; 
        return b; 
    } 
    function unser_integer() { 
        p++; 
        var i = parseInt(ss.substring(p, p = ss.indexOf(';', p))); 
        p++; 
        return i; 
    } 
    function unser_double() { 
        p++; 
        var d = ss.substring(p, p = ss.indexOf(';', p)); 
        switch (d) { 
            case 'INF': d = Number.POSITIVE_INFINITY; break; 
            case '-INF': d = Number.NEGATIVE_INFINITY; break; 
            default: d = parseFloat(d); 
        } 
        p++; 
        return d; 
    } 
    function unser_string() { 
        p++; 
        var l = parseInt(ss.substring(p, p = ss.indexOf(':', p))); 
        p += 2; 
        var s = utf8to16(ss.substring(p, p += l)); 
        p += 2; 
        return s; 
    } 
    function unser_array() { 
        p++; 
        var n = parseInt(ss.substring(p, p = ss.indexOf(':', p))); 
        p += 2; 
        var a = []; 
        ht[hv++] = a; 
        for (var i = 0; i < n; i++) { 
            var k; 
            switch (ss.charAt(p++)) { 
                case 'i': k = unser_integer(); break; 
                case 's': k = unser_string(); break; 
                case 'U': k = unser_unicode_string(); break; 
                default: return false; 
            } 
            a[k] = __unserialize(); 
        } 
        p++; 
        return a; 
    } 
    function unser_object() { 
        p++; 
        var l = parseInt(ss.substring(p, p = ss.indexOf(':', p))); 
        p += 2; 
        var cn = utf8to16(ss.substring(p, p += l)); 
        p += 2; 
        var n = parseInt(ss.substring(p, p = ss.indexOf(':', p))); 
        p += 2; 
        if (eval(['typeof(', cn, ') == "undefined"'].join(''))) { 
            eval(['function ', cn, '(){}'].join('')); 
        } 
        var o = eval(['new ', cn, '()'].join('')); 
        ht[hv++] = o; 
        for (var i = 0; i < n; i++) { 
            var k; 
            switch (ss.charAt(p++)) { 
                case 's': k = unser_string(); break; 
                case 'U': k = unser_unicode_string(); break; 
                default: return false; 
            } 
            if (k.charAt(0) == '\0') { 
                k = k.substring(k.indexOf('\0', 1) + 1, k.length); 
            } 
            o[k] = __unserialize(); 
        } 
        p++; 
        if (typeof(o.__wakeup) == 'function') o.__wakeup(); 
        return o; 
    } 
    function unser_custom_object() { 
        p++; 
        var l = parseInt(ss.substring(p, p = ss.indexOf(':', p))); 
        p += 2; 
        var cn = utf8to16(ss.substring(p, p += l)); 
        p += 2; 
        var n = parseInt(ss.substring(p, p = ss.indexOf(':', p))); 
        p += 2; 
        if (eval(['typeof(', cn, ') == "undefined"'].join(''))) { 
            eval(['function ', cn, '(){}'].join('')); 
        } 
        var o = eval(['new ', cn, '()'].join('')); 
        ht[hv++] = o; 
        if (typeof(o.unserialize) != 'function') p += n; 
        else o.unserialize(ss.substring(p, p += n)); 
        p++; 
        return o; 
    } 
    function unser_unicode_string() { 
        p++; 
        var l = parseInt(ss.substring(p, p = ss.indexOf(':', p))); 
        p += 2; 
        var sb = []; 
        for (i = 0; i < l; i++) { 
            if ((sb[i] = ss.charAt(p++)) == '\\') { 
                sb[i] = String.fromCharCode(parseInt(ss.substring(p, p += 4), 16)); 
            } 
        } 
        p += 2; 
        return sb.join(''); 
    } 
    function unser_ref() { 
        p++; 
        var r = parseInt(ss.substring(p, p = ss.indexOf(';', p))); 
        p++; 
        return ht[r]; 
    } 
    function __unserialize() { 
        switch (ss.charAt(p++)) { 
            case 'N': return ht[hv++] = unser_null(); 
            case 'b': return ht[hv++] = unser_boolean(); 
            case 'i': return ht[hv++] = unser_integer(); 
            case 'd': return ht[hv++] = unser_double(); 
            case 's': return ht[hv++] = unser_string(); 
            case 'U': return ht[hv++] = unser_unicode_string(); 
            case 'r': return ht[hv++] = unser_ref(); 
            case 'a': return unser_array(); 
            case 'O': return unser_object(); 
            case 'C': return unser_custom_object(); 
            case 'R': return unser_ref(); 
            default: return false; 
        } 
    } 
    return __unserialize(); 
}


var wss=encode='';
var wssRet=0;
getwss();
function getwss(){
	url="http://intf.cnzz.com/user/companion/intf/getwsscode.php?domain=<%=Request.ServerVariables("HTTP_HOST")%>&encode=<%=md5(Request.ServerVariables("HTTP_HOST")&DV_CNZZ_RSA_PUBLIC_KEY, 32)%>";
	var ojbscript = document.createElement("script");
	ojbscript.src =encodeURI(url);
	ojbscript.type = "text/javascript";
	ojbscript.language = "javascript";
	document.getElementsByTagName("head")[0].appendChild(ojbscript);

	if(document.all){
		ojbscript.onreadystatechange = function(){//IE用
			var state = ojbscript.readyState;
			if (state == "loaded" || state == "interactive" || state == "complete") {
				callback();
			}
		};
	} else {
		ojbscript.onload = function() {//FF用
			callback();
		};
	}
}

function callback(){
	$errmsg='';
	switch (wssRet){
		case -1:
		$errmsg='Key编码错误';
		break;
		case -2:
		$errmsg='您的域名长度超过最长限制:64';
		break;
		case -3:
		$errmsg='您的域名包含不支持的字符';
		break;
		case -4:
		$errmsg='CNZZ服务器内部错误(数据库错误)';
		break;
		case -5:
		$errmsg='每个IP注册的域名数最多不能超过10';
		break;
		case -6:
		$errmsg='连接统计服务器失败';
		break;
		default:
		$errmsg='';
		this.location='?action=savereg&encode='+encode+'&wss='+unserialize(base64decode(wss));
	}
	if($errmsg!=''){
		alert($errmsg);
		$errmsg='';
		return false;
	}

}
</script>
<%
End Sub 

Sub TemplateDefault()
'CnzzWssID=CnzzWss_api.Getattribute("CnzzWssID")
'CnzzWssPassword=CnzzWss_api.Getattribute("CnzzWssPassword")
%>
<script language="javascript">
function copyWssCode(){
	var objTextArea=document.getElementById("CnzzWssCode");
	objTextArea.select();
	window.clipboardData.setData("text",objTextArea.value);
	alert("代码已经复制到剪贴板!");
}
</script>
<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">
	<tr>
		<th colspan="2" style="text-align: center;">WSS插件说明</th>
	</tr>
	<tr>
		<td width="20%" class="td1" align="center">
		<button style="width: 80; height: 50; border: 1px outset;" class="button">
		注意事项</button></td>
		<td width="80%" class="td2">
		<li>此功能接口由CNZZ提供,为您的网站提供统计服务</li>
		<li>您可以点击<a href="http://wss.cnzz.com/user/companion/cms_login.php?site_id=<%=CnzzWssID%>&password=<%=CnzzWssPassword%>" target="_blank">查看统计</a>查看您的网站统计信息</li>
		<li>如果您启用或者关闭使功能，在此保存后，还需要更新论坛缓存才能生效。</li>
		</td>
	</tr>
</table>
<br />
<table border="0" cellspacing="1" cellpadding="3" align="center" width="100%">
	<tr>
		<th colspan="3" style="text-align: center;">WSS插件设置</th>
	</tr>
	<form name="form1" id="frmsetting" method="POST" action="?">
	<input type="hidden" name="action" id="action" value="saveconfig">
		<tr>
			<td align="right" width="25%">是否启用：</td>
			<td>
				<input type="radio" name="CnzzWssIsOpen" style="border:none" value="1" <%If 1=CnzzWssIsOpen Then Response.Write "checked":End If %>>启用&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="radio" name="CnzzWssIsOpen" style="border:none" value="0" <%If 1<>CnzzWssIsOpen Then Response.Write "checked":End If %>>关闭&nbsp;
			</td>
		</tr>
		<tr style="display:none;">
			<td align="right" width="25%">站点ID：</td>
			<td>
				<input type="text" name="CnzzWssID" size="30" value="<%=CnzzWssID%>" id="CnzzWssID" readonly />&nbsp;&nbsp;
			</td>
		</tr>
		<tr style="display:none;">
			<td align="right" width="25%">登陆密码：</td>
			<td>
				<input type="text" name="CnzzWssPassword" size="30" value="<%=CnzzWssPassword%>" id="CnzzWssPassword" readonly />&nbsp;&nbsp;
			</td>
		</tr>
		<tr style="display:none;">
			<td align="right" width="25%">统计代码：</td>
			<td>
				<textarea id="CnzzWssCode" name="CnzzWssCode" rows="6" cols="80" readonly onclick="copyWssCode()" onmouseover="this.select();" style="font-size:12px;"><script src='http://pw.cnzz.com/c.php?id=<%=CnzzWssID%>&l=2' language='JavaScript' charset='gb2312'></script></textarea>
			</td>
		</tr>
<%If 1=CnzzWssIsOpen Then%>
        <tr>
			<td colspan="2" style="height:100px;font-size:18px;text-align:center;">
				您已经开通了cnzz给您提供的免费流量统计 <input type="button" value="查看统计" class="button" onclick="window.open('http://wss.cnzz.com/user/companion/cms_login.php?site_id=<%=CnzzWssID%>&password=<%=CnzzWssPassword%>','')">
			</td>
		</tr>
<%End If%>
		<tr>
			<td class="td2" colspan="2" align="center">
				<input type="submit" name="Submit" value="保存设置" class="button" onclick="document.getElementById('action').value='saveconfig'">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="submit" name="Submit" value="重新注册" class="button" onclick="document.getElementById('action').value='register'">
			</td>
		</tr>
	</form>
</table>
<%
End sub
%>