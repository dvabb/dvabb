<%
Function GroupPermission(GroupSetting)
	Dim reGroupSetting,Rs,UserHtml,UserHtmlA,UserHtmlB
	If GroupSetting="" Then
		Set Rs = Dvbbs.Execute("Select GroupSetting From Dv_UserGroups Where UserGroupID=4")
		reGroupSetting = Split(Rs(0),",")
	Else
		reGroupSetting = Split(GroupSetting,",")
	End If
	If reGroupSetting(58)="0" Then reGroupSetting(58)="§"
	UserHtml = Split(reGroupSetting(58),"§")
	If Ubound(UserHtml)=1 Then
		UserHtmlA=UserHtml(0)
		UserHtmlB=UserHtml(1)
	Else
		UserHtmlA=""
		UserHtmlB=""
	End If
%>

<tr> 
<th colspan="4"><a name="setting2"></a>＝＝浏览相关选项</th>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(58)"></td>
<td class=tablebody1>用户名在帖子内容中显示标记<br />HTML语法，左右标记代码将加于用户名前后两头</td>
<td class=tablebody1>左标记 <input name="GroupSetting(58)A" type=text size=30 value="<%=Server.HtmlEncode(UserHtmlA)%>"> <br>右标记 <input name="GroupSetting(58)B" type=text size=30 value="<%=Server.HtmlEncode(UserHtmlB)%>"></td>
<td class=tablebody1><input type="hidden" id="g1" value="<b>用户名在帖子内容中显示标记</b><br><li>HTML语法，左右标记代码将加于用户名前后两头<br><li>如您设置了前后分别为《b》和《/b》，则在帖子内容中该组用户或者相关等级用户名显示为<B>粗体</B>">
<a href=# onclick="helpscript(g1);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(57)"></td>
<td class=tablebody1>允许用户自选风格</td>
<td class=tablebody1>是<input name="GroupSetting(57)" type=radio class="radio" value="1" <%if reGroupSetting(57)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(57)" type=radio class="radio" value="0" <%if reGroupSetting(57)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g2" value="<b>允许用户自选风格</b><br><li>如果关闭了本选项，论坛中用户将不能自己选择浏览显示的风格（包括用户在个人信息中设定的风格）">
<a href=# onclick="helpscript(g2);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(0)"></td>
<td class=tablebody1>可以浏览论坛</td>
<td class=tablebody1>是<input name="GroupSetting(0)" type=radio class="radio" value="1" <%if reGroupSetting(0)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(0)" type=radio class="radio" value="0" <%if reGroupSetting(0)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g3" value="<b>用户名在帖子内容中显示标记</b><br><li>关闭此选项，相关组或等级用户将不能浏览论坛<br><li>使用技巧：您可以设定某个用户组不能使用本设置，而当其身份变化后的用户组可使用本设置，如设置客人不能使用本设置，这样将迫使他登录">
<a href=# onclick="helpscript(g3);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(1)"></td>
<td class=tablebody1>可以查看会员信息（包括其他会员的资料和会员列表）
</td>
<td  class=tablebody1>是<input name="GroupSetting(1)" type=radio class="radio" value="1" <%if reGroupSetting(1)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(1)" type=radio class="radio" value="0" <%if reGroupSetting(1)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g4" value="<b>可以查看会员信息</b><br><li>关闭此选项，相关组或等级用户将不能浏览论坛用户资料，包括会员资料和会员列表资料<br><li>使用技巧：您可以设定某个用户组不能使用本设置，而当其身份变化后的用户组可使用本设置，如设置客人不能使用本设置，这样将迫使他登录">
<a href=# onclick="helpscript(g4);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(2)"></td>
<td  class=tablebody1>可以查看其他人发布的主题
</td>
<td  class=tablebody1>是<input name="GroupSetting(2)" type=radio class="radio" value="1" <%if reGroupSetting(2)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(2)" type=radio class="radio" value="0" <%if reGroupSetting(2)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g5" value="<b>可以查看其他人发布的主题</b><br><li>关闭此选项，相关组或等级用户将不能浏览论坛中其他人发布的帖子<br><li>使用技巧：您可以设定某个用户组不能使用本设置，而当其身份变化后的用户组可使用本设置，如设置客人不能使用本设置，这样将迫使他登录">
<a href=# onclick="helpscript(g5);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(41)"></td>
<td  class=tablebody1>可以浏览精华帖子
</td>
<td  class=tablebody1>是<input name="GroupSetting(41)" type=radio class="radio" value="1" <%if reGroupSetting(41)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(41)" type=radio class="radio" value="0" <%if reGroupSetting(41)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g6" value="<b>可以浏览精华帖子</b><br><li>关闭此选项，相关组或等级用户将不能浏览论坛中的精华帖子<br><li>使用技巧：您可以设定某个用户组不能使用本设置，而当其身份变化后的用户组可使用本设置，如设置客人不能使用本设置，这样将迫使他登录">
<a href=# onclick="helpscript(g6);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<th colspan="4"><a name="setting3"></a>＝＝发帖权限</th>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(3)"></td>
<td  class=tablebody1>可以发布新主题</td>
<td  class=tablebody1>是<input name="GroupSetting(3)" type=radio class="radio" value="1" <%if reGroupSetting(3)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(3)" type=radio class="radio" value="0" <%if reGroupSetting(3)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g9" value="<b>可以发布新主题</b><br><li>打开此选项，相关组或等级用户将可以可以发布新主题。鉴于国家规定，论坛默认的未登录用户组将即使设置此选项也不能发贴">
<a href=# onclick="helpscript(g9);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(64)"></td>
<td class=tablebody1>在审核模式下可直接发贴而不需经过审核</td>
<td class=tablebody1>是<input name="GroupSetting(64)" type=radio class="radio" value="1" <%if reGroupSetting(64)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(64)" type=radio class="radio" value="0" <%if reGroupSetting(64)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g10" value="<b>在审核模式下可直接发贴而不需经过审核</b><br><li>打开此选项，相关组或等级用户将可以可以发布新主题或回复而不经审核<br><li>当论坛版面设置为审核状态时该选项有效">
<a href=# onclick="helpscript(g10);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(62)"></td>
<td class=tablebody1>一天最多发贴数目
</td>
<td  class=tablebody1><input name="GroupSetting(62)" type=text size=4 value="<%=reGroupSetting(62)%>"></td>
<td class=tablebody1><input type="hidden" id="g11" value="<b>一天最多发贴数目</b><br><li>填写0为不作限制，出于对付灌水或者使用软件发贴的用户，请在此设置合理的数字<br><li>使用技巧：您可以给不同用户组设置不同的数字">
<a href=# onclick="helpscript(g11);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(4)"></td>
<td  class=tablebody1>可以回复自己的主题
</td>
<td  class=tablebody1>是<input name="GroupSetting(4)" type=radio class="radio" value="1" <%if reGroupSetting(4)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(4)" type=radio class="radio" value="0" <%if reGroupSetting(4)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g12" value="<b>可以回复自己的主题</b><br><li>打开此选项，相关用户组或等级用户可以回复自己发布的主题">
<a href=# onclick="helpscript(g12);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(5)"></td>
<td  class=tablebody1>可以回复其他人的主题
</td>
<td  class=tablebody1>是<input name="GroupSetting(5)" type=radio class="radio" value="1" <%if reGroupSetting(5)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(5)" type=radio class="radio" value="0" <%if reGroupSetting(5)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g13" value="<b>可以回复其他人的主题</b><br><li>打开此选项，相关用户组或等级用户可以回复其他人的主题">
<a href=# onclick="helpscript(g13);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(8)"></td>
<td  class=tablebody1>可以发布新投票</td>
<td  class=tablebody1>是<input name="GroupSetting(8)" type=radio class="radio" value="1" <%if reGroupSetting(8)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(8)" type=radio class="radio" value="0" <%if reGroupSetting(8)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g21" value="<b>可以发布新投票</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以发布新投票">
<a href=# onclick="helpscript(g21);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(9)"></td>
<td  class=tablebody1>可以参与投票</td>
<td  class=tablebody1>是<input name="GroupSetting(9)" type=radio class="radio" value="1" <%if reGroupSetting(9)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(9)" type=radio class="radio" value="0" <%if reGroupSetting(9)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g22" value="<b>可以发布新投票</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以参与投票">
<a href=# onclick="helpscript(g22);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(68)"></td>
<td  class=tablebody1>投票可以使用HTML语法</td>
<td  class=tablebody1>是<input name="GroupSetting(68)" type=radio class="radio" value="1" <%if reGroupSetting(68)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(68)" type=radio class="radio" value="0" <%if reGroupSetting(68)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g_u_HTML" value="<b>投票可以使用HTML语法</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以在投票中使用HTML语法">
<a href=# onclick="helpscript(g_u_HTML);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(17)"></td>
<td  class=tablebody1>可以发布小字报</td>
<td  class=tablebody1>是<input name="GroupSetting(17)" type=radio class="radio" value="1"  <%if reGroupSetting(17)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(17)" type=radio class="radio" value="0" <%if reGroupSetting(17)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g23" value="<b>可以发布小字报</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以发布小字报">
<a href=# onclick="helpscript(g23);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(46)"></td>
<td  class=tablebody1>发布小字报所需金钱</td>
<td  class=tablebody1><input name="GroupSetting(46)" type=text value="<%=reGroupSetting(46)%>" size=4></td>
<td class=tablebody1><input type="hidden" id="g24" value="<b>发布小字报所需金钱</b><br><li>在这里您可以根据需要设置不同用户组或等级用户发布小字报所需金钱">
<a href=# onclick="helpscript(g24);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(51)"></td>
<td  class=tablebody1>可以发布特殊标题帖子（如标题加红、UBB语法等）</td>
<td  class=tablebody1>是<input name="GroupSetting(51)" type=radio class="radio" value="1"  <%if reGroupSetting(51)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(51)" type=radio class="radio" value="0" <%if reGroupSetting(51)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g25" value="<b>可以发布特殊标题帖子</b><br><li>在这里您可以根据需要设置不同用户组或等级用户可以发布特殊标题帖子，如标题加颜色、HTML语法、UBB语法等，您可针对个别用户组可使用此特殊功能">
<a href=# onclick="helpscript(g25);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>

<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(59)"></td>
<td  class=tablebody1>可以发布赠送金币贴，获赠金币贴，论坛交易帖</td>
<td  class=tablebody1>是<input name="GroupSetting(59)" type=radio class="radio" value="1"  <%if reGroupSetting(59)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(59)" type=radio class="radio" value="0" <%if reGroupSetting(59)="0" then%>checked<%end if%>></td>
<td class=tablebody1>　</td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(67)"></td>
<td  class=tablebody1>发表模式选择</td>
<td  class=tablebody1>
<select name="GroupSetting(67)" >
<option value="0"  <%if reGroupSetting(67)="0" then%>selected<%end if%>>关闭HTML编辑
<option value="1"  <%if reGroupSetting(67)="1" then%>selected<%end if%>>允许HTML编辑
<option value="2"  <%if reGroupSetting(67)="2" then%>selected<%end if%>>简单模式编辑
<option value="3"  <%if reGroupSetting(67)="3" then%>selected<%end if%>>全功能编辑
</select>
</td>
<td class=tablebody1><input type="hidden" id="g0" value="<b>发表模式选择</b><br><li>发表模式包括：Design编辑模式,Ubb简单模式，HTML可编辑模式；<li>关闭HTML编辑：当版块允许发表高级模式下，用户只保留Design编辑模式和Ubb简单模式；<li>允许HTML编辑：当版块允许发表高级模式下，用户拥有Design编辑模式和HTML可编辑模式；<li>简单模式编辑：当版块允许发表高级模式下，用户只保留Ubb简单模式；<li>全功能编辑：当版块在发表简单模式下，拥有所有发表模式；<li>为避免用户滥用HTML的各种语法，建议只对部分用户关闭HTML编辑；">
<a href=# onclick="helpscript(g0);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(65)"></td>
<td  class=tablebody1>可以发表论坛专题</td>
<td  class=tablebody1>是<input name="GroupSetting(65)" type=radio class="radio" value="1"  <%if reGroupSetting(65)="1" then%>checked<%end if%>>&nbsp;必选<input name="GroupSetting(65)" type=radio class="radio" value="2"  <%if reGroupSetting(65)="2" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(65)" type=radio class="radio" value="0" <%if reGroupSetting(65)="0" then%>checked<%end if%>></td>
<td class=tablebody1><a href=# class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(52)"></td>
<td  class=tablebody1>新注册用户多少分钟后才能发言</td>
<td  class=tablebody1><input name="GroupSetting(52)" type=text value="<%=reGroupSetting(52)%>" size=4> 分钟</td>
<td class=tablebody1><input type="hidden" id="g26" value="<b>新注册用户多少分钟后才能发言</b><br><li>在这里您可以根据需要设置不同用户组或等级用户新注册需要多少分钟后才能发言，建议合理设置此选项，以避免一些恶意用户乱注册散发非法帖子或广告帖子">
<a href=# onclick="helpscript(g26);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(69)"></td>
<td class=tablebody1>是否允许使用魔法表情</td>
<td class=tablebody1>
<input type=radio class="radio" name="GroupSetting(69)" value=1 <%if reGroupSetting(69)="1" then%>checked<%end if%>>是&nbsp;
<input type=radio class="radio" name="GroupSetting(69)" value=0 <%if reGroupSetting(69)="0" then%>checked<%end if%>>否&nbsp;
</td>
<td class=tablebody1>　</td>
</tr>
<tr> 
<th colspan="4"><a name="setting4"></a>＝＝<b>帖子/主题编辑权限</b></th>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(10)"></td>
<td  class=tablebody1>可以编辑自己的帖子
</td>
<td  class=tablebody1>是<input name="GroupSetting(10)" type=radio class="radio" value="1" <%if reGroupSetting(10)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(10)" type=radio class="radio" value="0" <%if reGroupSetting(10)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g27" value="<b>可以编辑自己的帖子</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以编辑自己的帖子">
<a href=# onclick="helpscript(g27);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(11)"></td>
<td  class=tablebody1>可以删除自己的帖子
</td>
<td  class=tablebody1>是<input name="GroupSetting(11)" type=radio class="radio" value="1" <%if reGroupSetting(11)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(11)" type=radio class="radio" value="0" <%if reGroupSetting(11)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g28" value="<b>可以删除自己的帖子</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以删除自己的帖子，请根据自己的需要合理设置此选项">
<a href=# onclick="helpscript(g28);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(12)"></td>
<td  class=tablebody1>可以移动自己的帖子到其他论坛
</td>
<td  class=tablebody1>是<input name="GroupSetting(12)" type=radio class="radio" value="1" <%if reGroupSetting(12)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(12)" type=radio class="radio" value="0" <%if reGroupSetting(12)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g29" value="<b>可以移动自己的帖子到其他论坛</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以移动自己的帖子到其他论坛，请根据自己的需要合理设置此选项">
<a href=# onclick="helpscript(g29);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(13)"></td>
<td  class=tablebody1>可以打开/关闭自己发布的主题
</td>
<td  class=tablebody1>是<input name="GroupSetting(13)" type=radio class="radio" value="1" <%if reGroupSetting(13)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(13)" type=radio class="radio" value="0" <%if reGroupSetting(13)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g30" value="<b>可以打开/关闭自己发布的主题</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以打开/关闭自己发布的主题，请根据自己的需要合理设置此选项">
<a href=# onclick="helpscript(g30);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<th colspan="4"><a name="setting5"></a>＝＝上传权限设置</th>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(7)"></td>
<td  class=tablebody1>可以上传附件
</td>
<td  class=tablebody1>是<input name="GroupSetting(7)" type=radio class="radio" value="1" <%if reGroupSetting(7)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(7)" type=radio class="radio" value="0" <%if reGroupSetting(7)="0" then%>checked<%end if%>>
&nbsp;发帖可以上传<input name="GroupSetting(7)" type=radio class="radio" value="2" <%if reGroupSetting(7)="2" then%>checked<%end if%>>&nbsp;回复可以上传<input name="GroupSetting(7)" type=radio class="radio" value="3" <%if reGroupSetting(7)="3" then%>checked<%end if%>>
</td>
<td class=tablebody1><input type="hidden" id="g16" value="<b>可以上传附件</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以上传附件，选择是则发贴和回贴都可以上传，否则不行。您也可以可以根据需要分别设置发贴或回帖是否可以上传">
<a href=# onclick="helpscript(g16);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(66)"></td>
<td  class=tablebody1>一次批量上传数量（设置为0，即不允许使用此功能；建议不要超过5个）
</td>
<td  class=tablebody1><input name="GroupSetting(66)" type=text size=4 value="<%=reGroupSetting(66)%>"></td>
<td class=tablebody1><input type="hidden" id="GroupSetting66" value="<b>一次批量上传数量</b><br><li>设置为0，即不允许使用此功能;<li>建议不要超过5个，因为上传操作将消耗大量服务器资源">
<a href=# onclick="helpscript(GroupSetting66);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(40)"></td>
<td  class=tablebody1>一次最多上传文件个数
</td>
<td  class=tablebody1><input name="GroupSetting(40)" type=text size=4 value="<%=reGroupSetting(40)%>"></td>
<td class=tablebody1><input type="hidden" id="g17" value="<b>一次最多上传文件个数</b><br><li>在这里您可以根据需要设置不同用户组或等级用户一次最多上传文件个数，建议不要设置过大，因为上传操作将消耗大量服务器资源">
<a href=# onclick="helpscript(g17);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(50)"></td>
<td  class=tablebody1>一天最多上传文件个数
</td>
<td  class=tablebody1><input name="GroupSetting(50)" type=text size=4 value="<%=reGroupSetting(50)%>"></td>
<td class=tablebody1><input type="hidden" id="g18" value="<b>一天最多上传文件个数</b><br><li>在这里您可以根据需要设置不同用户组或等级用户一天最多上传文件个数，建议不要设置过大，因为上传操作将消耗大量服务器资源">
<a href=# onclick="helpscript(g18);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(44)"></td>
<td  class=tablebody1>上传文件大小限制
</td>
<td  class=tablebody1><input name="GroupSetting(44)" type=text size=4 value="<%=reGroupSetting(44)%>"> KB</td>
<td class=tablebody1><input type="hidden" id="g19" value="<b>上传文件大小限制</b><br><li>在这里您可以根据需要设置不同用户组或等级用户上传文件大小，建议不要设置过大，因为上传操作将消耗大量服务器资源">
<a href=# onclick="helpscript(g19);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(61)"></td>
<td  class=tablebody1>可以下载附件</td>
<td  class=tablebody1>是<input name="GroupSetting(61)" type=radio class="radio" value="1" <%if reGroupSetting(61)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(61)" type=radio class="radio" value="0" <%if reGroupSetting(61)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g20" value="<b>可以下载附件</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以下载附件，比如可以设置未登录用户不许下载">
<a href=# onclick="helpscript(g20);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<th colspan="4"><a name="setting6"></a>＝＝管理权限</th>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(18)"></td>
<td  class=tablebody1>可以删除其它人帖子
</td>
<td  class=tablebody1>是<input name="GroupSetting(18)" type=radio class="radio" value="1" <%if reGroupSetting(18)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(18)" type=radio class="radio" value="0"  <%if reGroupSetting(18)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g38" value="<b>可以删除其它人帖子</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以删除其它人帖子，请根据自己的需要合理设置此选项，建议对版主及其以上用户组设置此权限">
<a href=# onclick="helpscript(g38);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(19)"></td>
<td  class=tablebody1>可以移动其它人帖子
</td>
<td  class=tablebody1>是<input name="GroupSetting(19)" type=radio class="radio" value="1" <%if reGroupSetting(19)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(19)" type=radio class="radio" value="0"  <%if reGroupSetting(19)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g39" value="<b>可以移动其它人帖子</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以移动其它人帖子，请根据自己的需要合理设置此选项，建议对版主及其以上用户组设置此权限">
<a href=# onclick="helpscript(g39);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(20)"></td>
<td  class=tablebody1>可以打开/关闭其它人帖子
</td>
<td  class=tablebody1>是<input name="GroupSetting(20)" type=radio class="radio" value="1" <%if reGroupSetting(20)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(20)" type=radio class="radio" value="0"  <%if reGroupSetting(20)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g40" value="<b>可以打开/关闭其它人帖子</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以打开/关闭其它人帖子，请根据自己的需要合理设置此选项，建议对版主及其以上用户组设置此权限">
<a href=# onclick="helpscript(g40);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(21)"></td>
<td  class=tablebody1>可以固顶/解除固顶帖子
</td>
<td  class=tablebody1>是<input name="GroupSetting(21)" type=radio class="radio" value="1" <%if reGroupSetting(21)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(21)" type=radio class="radio" value="0"  <%if reGroupSetting(21)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g41" value="<b>可以固顶/解除固顶帖子</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以固顶/解除固顶帖子，请根据自己的需要合理设置此选项，建议对版主及其以上用户组设置此权限">
<a href=# onclick="helpscript(g41);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(54)"></td>
<td  class=tablebody1>可以进行帖子区域固顶操作
</td>
<td  class=tablebody1>是<input name="GroupSetting(54)" type=radio class="radio" value="1" <%if reGroupSetting(54)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(54)" type=radio class="radio" value="0"  <%if reGroupSetting(54)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g42" value="<b>可以进行帖子区域固顶操作</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以进行帖子区域固顶操作，请根据自己的需要合理设置此选项，建议对超级版主及其以上用户组设置此权限">
<a href=# onclick="helpscript(g42);return false;" class="helplink"><img src="<%=MyDbPath%>images/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" class="checkbox" name="CheckGroupSetting(38)"></td>
<td  class=tablebody1>可以进行帖子总固顶操作
</td>
<td  class=tablebody1>是<input name="GroupSetting(38)" type=radio class="radio" value="1"  <%if reGroupSetting(38)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(38)" type=radio class="radio" value="0" <%if reGroupSetting(38)="0" then%>checked<%end if%>></td>
<td class=tablebody1><input type="hidden" id="g43" val class="checkbox" name="CheckGroupSetting(20)"></td>
<td  class=tablebody1>鍙