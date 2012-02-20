<!-- #include file="common/header.asp" -->
<p>通过自带工具生成的URL为:<% = m.cu(m.seg("news","read","3"),m.kv("page","3")) %></p>
<p><a href="<% = m.cu(m.seg("index","index","test"),m.kv("welcome","")) %>">点击这个试试</a></p>
<p>url的第1段为：<% = m.gs(0) %></p>
<p>url的第2段为：<% = m.gs(1) %></p>
<% if( vartype( m.gs(2) ) <> 11 ) then %>
<p>url的第3段为：<% = m.gs(2) %></p>
<% end if %>
<p>根目录是：<%=m.root()%></p>
<p>来个虚拟文件:<%=m.root("images/MagicASP-Logo.png")%></p>
<div>
	<form action="<% = m.cu(m.seg("news"),m.kv("page","3")) %>" method="get">
	<input name="test" type="text" value="test"><input type="submit">
	<% m.fgm(m.seg("index","index")) : '此处用于修复get方式提交数据 %>
	</form>
</div>

以下是此页的源代码
<pre style="font-size:12px;">

&lt;!-- #include file="common/header.asp" --&gt;
&lt;p&gt;通过自带工具生成的URL为:&lt;% = m.cu(m.seg("news","read","3"),m.kv("page","3")) %&gt;&lt;/p&gt;
&lt;p&gt;&lt;a href="&lt;% = m.cu(m.seg("index","index","test"),m.kv("welcome","")) %&gt;"&gt;点击这个试试&lt;/a&gt;&lt;/p&gt;
&lt;p&gt;url的第1段为：&lt;% = m.gs(0) %&gt;&lt;/p&gt;
&lt;p&gt;url的第2段为：&lt;% = m.gs(1) %&gt;&lt;/p&gt;
&lt;% if( vartype( m.gs(2) ) &lt;&gt; 11 ) then %&gt;
&lt;p&gt;url的第3段为：&lt;% = m.gs(2) %&gt;&lt;/p&gt;
&lt;% end if %&gt;
&lt;p&gt;根目录是：&lt;%=m.root()%&gt;&lt;/p&gt;
&lt;p&gt;来个虚拟文件:&lt;%=m.root("images/MagicASP-Logo.png")%&gt;&lt;/p&gt;
&lt;div&gt;
	&lt;form action="&lt;% = m.cu(m.seg("news"),m.kv("page","3")) %&gt;" method="get"&gt;
	&lt;input name="test" type="text" value="test"&gt;&lt;input type="submit"&gt;
	&lt;% m.fgm(m.seg("index","index")) : '此处用于修复get方式提交数据 %&gt;
	&lt;/form&gt;
&lt;/div&gt;
&lt;!-- #include file="common/footer.asp" --&gt;

</pre>
<!-- #include file="common/footer.asp" -->