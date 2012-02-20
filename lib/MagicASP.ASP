<!-- #include file="MagicASP_Utility.ASP" -->
<script Language="JScript" RunAt="Server">
/*
 * MagicASP 一个ASP文件自动合并工具 用于统一URL. 
 * @version: beta
 * @author:qpwoeiru96
 * @link: http://sou.la/blog
 * @date: 2012-02-18 05:29:29
 * @version：2012-02-18 05:29:29
 */

MagicASP = function( config ) {
	
	var config = ( typeof config === 'undefined' ) ? {} : config ;
	this['config'] = _copy( config, {
		codePage : 65001 , //生成文件的代码页 默认是UTF-8的代码页 
		directory : '' , //配置之一的目录 如果根目录请保留为空 后面请加 / 例如 MagicASP/
		debug : true , //产品中请设置为false 是否开启调试模式
		version : 'alpha' , //版本设置用于不同版本的调用
		logPath :  'logs/' , //日志保存路径 后面请加 / 例如 logs/
		accessLog :  true ,// 是否记录访问信息
        scriptFileName : 'index.asp' ,
		siteAddress : 'http://www.example.com/' ,
		redirectAddress : false //404出错跳转地址 
	});	
	
	this['fso'] = new ActiveXObject("Scripting.FileSystemObject");
	this['str'] = $S.CreateObject("ADODB.Stream");
	this['rootPath'] = $S.Mappath('\\' + this['config']['directory'] ) + '\\';
	this['CharSet'] = this['config']['codePage'] == 65001 ? 'utf-8' : 'gbk';
	this['str']['Mode'] = 3;
	this['str']['Type'] = 2;	
	this['str']['CharSet'] = this['CharSet'];

	 //判断是否Log 执行Log()
	if( this['config']['accessLog'] ) this.log();
	
	$A('scriptURL') = this['config']['siteAddress'] + this['config']['scriptFileName'] ;
    $A('siteAddress') = this['config']['siteAddress'] ;
	
	$R.AddHeader("PoweredBy","MagicASP " + this['Static']['Version']);
	$R.Charset = this['CharSet'];
}


MagicASP.prototype.init = function( moduleList ) {
	
	var mod = m.getSegment(0);
	var act = m.getSegment(1);
	
	this['moduleList'] = moduleList ? moduleList : ['index'] ;	
	this['module'] = mod === false || mod == "" ? 'index' : mod;
	this['action'] = act === false || act == "" ? 'index' : act;
	
	//防止跨文件调用
	if( ( this['action']).match(/\.+/) ) {
		var detail = "This action is unsupported."
		this.throwError( 404, 'UNSUPPORT ACTION', detail); 
	}
	
	//防止出现未定义的模块
	if( this.arrayIndexOf(this['moduleList'], this['module'] )  === -1 ) {
		this.throwError( 404, 'UNDEFINED MODULE', 'Undefined module.' );
	}
}

MagicASP.prototype.run = function() {
	
	var runFilePath = this['module'] + '_' + this['action'] + '_' + this['config']['version'] + '.asp';
		
	if ( !this['fso']['FolderExists']( this.rootPath + 'runtime\\' ) ) { //创建runtime目录
		this['fso']['CreateFolder']( this.rootPath + 'runtime\\' );
	}
	
	runFilePath = this['rootPath'] + 'runtime\\' + runFilePath;
	
	if( this['config']['debug'] || !this['fso'].FileExists(runFilePath) ) {
					
		var scriptContent = "<% @LANGUAGE=\"VBSCRIPT\" CODEPAGE=\"" + this['config']['codePage'] + "\" %>\r\n";
		//scriptContent += this.addScriptTag( __UTILITY__, "JScript");
		scriptContent += this.include('module\\' + this['module'] + '.asp');
		scriptContent += this.include('action\\' + this['module'] + '\\' + this['action'] + '.asp');
		scriptContent = scriptContent.replace(/MagicASP_INCLUDE_MARK/g , "include");
		this.writeTextFile( runFilePath, scriptContent );

	}
	Server.Execute( runFilePath.replace( Server.Mappath('\\'), '' ) );
}

/*
 * 文件包含方法 [MagicASP关键核心]
 * @parameter filepath String 文件相对路径 
 * @return String 脚本生成的内容
 */
MagicASP.prototype.include = function( filePath ) {
	
	var fileName,
		scriptContent ,
		fileApplicationName ;

	//Application变量中保存的文件名称
	fileApplicationName = ( filePath.substring( 0, filePath.indexOf('.asp') ) ).split("\\").join("_") ;
	filePath = this['rootPath'] + filePath;
	filePath = filePath.replace(/\//g,"\\");

	if( ( Application(fileApplicationName) !== undefined ) && !this['config']['debug'] ) return Application(fileApplicationName);
		
	try{
		if( this['fso'].FileExists( filePath ) ){
			
			scriptContent = this.readTextFile( filePath );
			//去除每个文件的 <% RESPONSE.END %>			
			scriptContent = scriptContent.replace(/^<% RESPONSE\.END %>\r\n/, '');

			var self = this ;
			
			//将原有的include语法进行替换 注意virtual中路径为当前应用路径 自动进行替换
			scriptContent = scriptContent.replace(/<!-{2} ?#include ?virtual=\"(.*)\" ?-->/ig , function( $1, $2 ) {
				return self.include($2);
			});
			
			//将原有的include语法进行替换 file中路径为当前文件路径 自动进行替换
			scriptContent = scriptContent.replace(/<!-{2} ?#include ?file=\"(.*)\" ?-->/ig , function( $1, $2 ) { 
				$2 = $2.replace(/\//g,"\\");		
                var currentPath = filePath.replace(self['rootPath'],"");
				currentPath = currentPath.substring(0,currentPath.lastIndexOf("\\"));
				var path = currentPath + "\\" + $2;
				return self.include( path );	
			});
			
            //新增require语法 就是原本的include语法 一般不推荐此语法 因为是路径相关的 容易出错 建议使用在调试模式
			scriptContent = scriptContent.replace(/<!-{2} ?#require ?file=\"(.*)\" ?-->/ig , function( $1, $2 ) {
                $2 = $2.replace(/\//g,"\\");
				var currentPath = filePath.substring(0,filePath.lastIndexOf("\\"));
				var path = currentPath + "\\" + $2;
				path = self.getTruePath(path);
				path = path.replace($S.MapPath("\\")+"\\","");				
				return "\<\!-- #" + "MagicASP_INCLUDE_MARK " + "virtual=\"" +  path + "\" --\>";
			});

			//新增require语法 就是原本的include语法 一般不推荐此语法 因为是路径相关的 容易出错 建议使用在调试模式
			scriptContent = scriptContent.replace(/<!-{2} ?#require ?virtual=\"(.*)\" ?-->/ig , function( $1, $2 ) {
                $2 = $2.replace(/\//g,"\\");				
				var currentPath = $S.MapPath("\\") + "\\" + self['config']['directory'];
				var path = currentPath + $2;
				path = self.getTruePath(path);
				path = path.replace($S.MapPath("\\")+"\\","");				
				return "\<\!-- #" + "MagicASP_INCLUDE_MARK " + "virtual=\"" +  path + "\" --\>";
			});
			
			if( this['config']['debug'] ) {
                //为了方便调试 定位出错文件在哪
				scriptContent = "\r\n<% '-------- start@:" + filePath + " --------  %>\r\n" + scriptContent + "\r\n<% '-------- end@:" + filePath + " --------  %>\r\n";
			} else {
				//缓存已经读取的文件
                $A(fileApplicationName) = scriptContent;
            }
			
		} else {
			throw 'FILE_NOT_EXIST';
		}
	} catch( err ) {		
		if(err === 'FILE_NOT_EXIST'){
			var detail = (this['config']['debug'] ? filePath : ' ') + ' file not found ';
			this.throwError( 404 , "FILE NOT FOUND" , detail );
		} else {
			var detail = "Please Contact This Web Master, Thanks!";
			this.throwError( 500 , "CATCH A UNKOWN ERROR" , detail);
		}
	}
	return scriptContent;
}


/*
 * 通过ADODB Stream读取文本文件
 * @parameter filePath String 要读取的文件路径
 * @return String 文本内容
 */
MagicASP.prototype.readTextFile = function( filePath ) {
	
	this['str'].Open();	
	this['str'].LoadFromFile( filePath );

	if( this['str']['EOS'] ) {		
		var result = "";
	} else {
		var result = this['str'].ReadText();
		result = result.replace(/^\uFEFF/,'');
	}
	
	this['str'].Close();
	return result;
}

/*
 * 通过ADODB Stream写入文本文件
 * @parameter filePath String 文件路径
 * @parameter content String 写入文本内容
 * @return none
 */ 
MagicASP.prototype.writeTextFile = function( filePath , content ) {

	this['str'].Open();
	
	if ( this['fso'].FileExists( filePath ) ) {
		this['fso'].DeleteFile( filePath );
	}
	
	this['str'].WriteText(content);
	this['str'].SaveToFile(filePath,2);
	this['str'].Flush(); 
	this['str'].Close();
	
}

/*
 * 数组查找
 * @parameter array Array 需要被查找的数组
 * @parameter value String 被查找的字符串
 * @return Number 所在位置 未找到则返回-1
 */

MagicASP.prototype.arrayIndexOf = function( array , value ) { 
	var len = array.length ;
	if( !len ) return -1;
	for( var i = 0 ; i < len ; i++ )  { if (value == array[i]) return i };
	return -1;
}


/*
 * 获取真实的路径去除 ./ ../
 * @parameter path String 传进来的全路径
 * @return String 传出去的全路径
 */
MagicASP.prototype.getTruePath = function( path ) {

    path = path.replace(/\//g,"\\");
    var arr = path.split("\\");	
    var tmp = [];
    var cur = 0;
    for( var i in arr ) {
        if( arr[i] == ".." ) {
            cur--;			
        } else if( arr[i] != "." ) {
            tmp[cur] = arr[i];
            cur = tmp.length;
        }
    }
    return tmp.join("\\");
}

/*
 * 给脚本添加外部标签 
 * @parameter scriptContent String 文件路径
 * @parameter lang String 脚本类型 一般是JScript 跟 VBScript
 * @return 加好标签之后的脚本
 */
MagicASP.prototype.addScriptTag = function( scriptContent, lang ) {
	return '<' + 'script language="'+ lang +'" runat="server">\r\n' + scriptContent + '\r\n<\/script>';
}

/*
 * 抛出出错信息
 * @parameter httpCode String HTTP状态码
 * @parameter title String 出错标题
 * @parameter detail String 详细出错信息
 * @return none
 */
MagicASP.prototype.throwError = function(httpCode, title, detail){
	if( this["config"]["redirectAddress"] !== false ) $R.Redirect( this["config"]["redirectAddress"] );
	Response.Status = ( httpCode + this['Static']['HTTPStatuCode'][httpCode] );
	echo( this.replaceTemplate({
		'httpCode' : httpCode,
		'title' : title,
		'detail' : detail,
		'serverInfo' : String.prototype.concat('MagicASP ',this['Static']['Version'] ,' @ ',server_("Server_Name") ,' ',server_("Server_Software"))
	},this['Static']['errorTemplate']) );
	Response.End();
}

/*
 * HTML模板替换
 * @parameter contentArray String 参数列表
 * @parameter templateContent String 模板内容
 * @return 替换后的内容
 */
MagicASP.prototype.replaceTemplate = function ( contentObject, templateContent ) {
	for( var k in contentObject ) {
		templateContent = templateContent.replace(new RegExp('\\{\\$'+k+"\\}","ig"), contentObject[k] );
	}
	return templateContent;
}

/*
 * 将用户访问写入到日志
 */
MagicASP.prototype.log = function() {
	
	var accessLogFile = String.prototype.concat( this['rootPath'] + this['config']['logPath'] + 'access_' , this.getLogFileName() , '.log' );
	var file = this['fso'].OpenTextFile( accessLogFile, 8, true, -1 );	
	
	var logContent = this.replaceTemplate({
		
		remoteIP : server_("REMOTE_ADDR") ,
		time : (new Date()).toLocaleString().replace(/\(.+\)/ig,"") ,
		method : server_("REQUEST_METHOD") ,
		fullPath : get_() == '' ? server_("SCRIPT_NAME") : server_("SCRIPT_NAME")+ '?' + server_("QUERY_STRING") ,
		serverProtocol : server_("REQUEST_METHOD") ,
		referer : _string(server_("HTTP_REFERER")) == 'undefined' ? "" : server_("HTTP_REFERER") ,
		userAgent : server_("HTTP_USER_AGENT")
		
	} , this["Static"]["logTemplate"] );
		
    file.writeLine( logContent )    
    file.close();

}

/*
 * 根据时间获取日志文件名 
 */
MagicASP.prototype.getLogFileName = function() {
	
	var now = new Date();
    var day = now.getDate();	
    var month = now.getMonth() + 1;
	var year = now.getFullYear();
	
	if( day <= 9 ) day = String.prototype.concat( '0' , day );
	if( month <= 9 ) month = String.prototype.concat( '0' , month );
	
	return String.prototype.concat( year , month , day );
	
}

/*
 * 静态数据
 */
MagicASP.prototype.Static = { 
	
	/*
	 * 抛出错误的模板
	 */
	errorTemplate : '<!DOCTYPE HTML>\ 
<html>\
<head>\
	<title>{$httpCode} - {$title}</title>\
	<style>html, body{margin:20px;padding:0;font:13px/22px "Lucida Grande", "Lucida Sans Unicode", Helvetica, Arial, Verdana, sans-serif}</style>\
	<meta http-equiv="Content-Type" content="text/html; charset=gbk">\
</head>\
<body>\
	<fieldset>\
		<h1>{$httpCode} - {$title}</h1>\
		<h2>{$detail}</h2>\
		<div align="right">{$serverInfo}</div>\
	</fieldset>\
</body>\
</html>',

	/*
	 * 日志记录模板
	 */
	logTemplate : '{$remoteIP} [{$time}] "{$method} {$fullPath} {$serverProtocol}" "{$referer}" "{$userAgent}"',
		
	/*
	 * http状态码记录集
	 */
	HTTPStatuCode : {
		'400' : 'Bad Request',
		'401' : 'Unauthorized',
		'402' : 'Payment Required',
		'403' : 'Forbidden',
		'404' : 'Not Found',
		'405' : 'Method Not Allowed',
		'406' : 'Not Acceptable',
		'407' : 'Proxy Authentication Required',
		'408' : 'Request Timeout',
		'409' : 'Conflict',
		'410' : 'Gone',
		'411' : 'Length Required',
		'412' : 'Precondition Failed',
		'413' : 'Request Entity Too Large',
		'414' : 'Request-URI Too Long',
		'415' : 'Unsupported Media Type',
		'416' : 'Requested Range Not Satisfiable',
		'417' : 'Expectation Failed',
		'500' : 'Internal Server Error',
		'501' : 'Not Implemented',
		'502' : 'Bad Gateway',
		'503' : 'Service Unavailable',
		'504' : 'Gateway Timeout',
		'505' : 'HTTP Version Not Supported'
	} ,

	/*
	 * 产品版本信息
	 */
	Version : 'Beta-2012/02/18'
	
}


</script>