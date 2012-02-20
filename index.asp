<%
	Response.Write "<h1>请根据实际环境到index.asp设置您的的MagicASP，谢谢合作！</h1>"
	Response.Write "<h1>设置完毕之后请删除本段代码！</h1>"
	Response.End
%>
<!-- #include file="lib/MagicASP.asp" -->
<script Language="JavaScript" RunAt="Server">
/*
 * MagicASP
 * @version: beta
 * @author: qpwoeiru96
 * @link: http://sou.la/blog
 * @date: 2012-02-18
 * @description: 
 *    将本文件放入根目录作为默认目录 新建文件夹 module action runtime logs module里面建立index.asp  action里面建立index文件夹并在其里面建立index.asp 即可
 *    然后配置一下的内容，即可启动。
 */
app = new MagicASP({
	codePage : 936 , //生成文件的代码页 默认是UTF-8的代码页 这里是936 为GBK UTF-8的是65001 
	directory : '' , //配置之一的目录 如果根目录请保留为空 后面请加 / 例如 MagicASP/
	debug : true , //产品中请设置为false 是否开启调试模式
	version : 'beta' , //版本设置用于不同版本的调用
	logPath :  'logs/' , //日志保存路径 后面请加 / 例如 logs/
	accessLog : true ,// 是否记录访问信息
	scriptFileName : '' ,//站点脚本文件 一般是index.asp  用于URL生成
	siteAddress : 'http://www.example.com/', //站点地址 用于URL生成
    redirectAddress : false //404出错跳转地址  用于出错时直接跳转到页面而不显示错误
});
app.init(['index']); //可执行 module 列表
app.run();
</script>