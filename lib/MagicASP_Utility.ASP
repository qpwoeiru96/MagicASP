<script Language="JavaScript" RunAt="Server">
/*
 * MagicASP 工具集
 */

var $S = Server ,
    $Q = Request,
    $R = Response,
    $A = Application;



_string = function(a) {
	if( typeof a == "undefined" ) {
		return '';
	}
	return new String( a );
}

_number = function() {
	return new Number( arguments[0] );
}


echo = function() { //模仿PHP echo
	for( var i = 0 ; i < arguments.length ; i++ ) {
		$R.Write( arguments[i] );
	}
}

_copy = function( source , target ) { //对象浅层次拷贝
	for( var key in source ) {
		target[key] = source[key] ;
	}	
	return target;
}

server_ = function(a) {
	return $Q.ServerVariables( a ) ;
}

cookie_ = function(a) {
    return $Q.Cookies(a);
}

get_ = function() {
	return arguments.length == 0 ? $Q.QueryString() : $Q.QueryString( arguments[0] ) ;
}

post_ = function(a) {
    return _Q.Form( a );
}

MagicASP_Utility = {
	
	kv : function( a, b ) {//可以在VBScript以及JScript共用的数据格式
		return [a,b];
	},
		
	getSegment : function(id) { //获取路径分段的函数 getSegment 不对就返回false 注意是false 可用vartype = 11判断是否有
	
		var segs = _string(get_("_"));		
		var arr = segs.length == 0 || segs == 'undefined' ? [] : segs.match(/[^\/]+/g);	
		
		if( id > arr.length - 1 ) {
			return false;
		} else {
			return arr[id];
		}
	
	},
	
	segments : function() { //创建分段 也可以自己构造 
	
		var seg = "";
		for( var i = 0 ; i < arguments.length ; i++ ) {
			seg += String.prototype.concat ( $S.URLEncode(arguments[i]) + '/' );
		}
		return seg;
	
	},
	
	createUrl : function( segments ) { //创建URL 
		
		var qs = '?_=' +  segments;
		
		for(var i = 1; i < arguments.length ; i++)
		{
			var arr = arguments[i];
			qs += "&" + $S.URLEncode(arr[0]) + "=" + $S.URLEncode(arr[1]);
		}
		return $A('scriptURL') + qs;
	},
	
	root : function() {
        return String.prototype.concat( $A('siteAddress') + (arguments.length !== 0 ? arguments[0] : "")  );
    },
	
	fixGETMethod : function( segments ) { //隐藏Form
		echo('<input type="hidden" name="_" value="'+segments+'"/>');
	}
}

m = MagicASP_Utility;
m.cu = m.createUrl;
m.fgm = m.fixGETMethod;
m.seg = m.segments;
m.gs = m.getSegment;

</script>