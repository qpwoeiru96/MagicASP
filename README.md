MagicASP - 轻量级ASP整合框架
=====================

通过静态分析技术，使得ASP可以支持单一入口。**放弃乱七八糟的文件名吧，转而投向干净整洁的目录吧。**

比如 index.php?_=index/test 代表 模块为index action为test，自动拼接文件并执行。

在此基础上可以实现简单的 ACL。

高效函数
----
```
get_ 代替 Request.QueryString
post_ 代替 Request.Form
echo 代替 Response.Write 支持多参数 例如 <% echo "Magic" , "ASP" %> = <% = "MagicASP"%>
```

链接产生函数
----

```
<% = url.root("images/register.gif") %> // http://www.example.com/images/register.gif
<% = url.create("user/register/",kv("name","demo"))  %> // http://www.example.com/index.asp?_=user/register/&name=demo
//kv 代表 key/value 结构
```

目录无关include
---
将原有的 include 语法进行替换变成与web容器目录无关的语法 方便迁移项目
同时增加 require 语法以兼容老的语法
```
<!-- #include virtual="lib/const.asp" -->
<!-- #include virtual="lib/admin.class.asp" -->

<!-- #require virtual="lib/const.asp" -->
<!-- #require virtual="lib/admin.class.asp" -->
```
（*现在看看这个设计其实应该反过来的，直接修改原有的设计其实是不对的*）

自带日志记录
----
日志格式为：`RemoteIP` `DateTime` `HTTP Method` `Request` `HTTP Protocol Version` `HTTP Referer` `UserAgent`
站长么可以没事用用文本处理工具或者语言处理下日志文件，统计下访问量啥的。

PS: 这是本人于2011为一个门户网站项目开发的框架。估计是唯一一个不用eval实现了ASP的单一路口和动态加载（至少我没碰到过）。