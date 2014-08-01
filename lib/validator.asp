<script type="text/javascript" language="javascript" runat="server">
/**
 *******************************************************************************************
 * VBScirpt ASP 数据验证类 进行验证的同时还能修改原始数据 
 * @author: QPWOEIRU96
 * @date: 2012年3月26日
 * @version: alpha
 * @site: http://sou.la/blog
 * @license: 你可以免费使用[但不应该出于商业目的或者从此谋得利益]，如果修改了源代码请告知我。
 ********************************************************************************************
 * @Code Sample[VBScript]
 * Dim V
 * Set V = New Validator '实例化
 * V.Method = "GET" '设置提交方式
 * Call V.add("Test|测试数据asd","Define$""360""|Required|Url|Max$8") '添加验证规则
 * Call V.add("AS|测试数据ad2","Required|Define$""360""|Url") '添加验证规则
 * '验证规则为： {POST表单名或者GET请求名}|{汉化文字}  后面是 {函数名}${参数}|{函数名}${参数} 
 * '其中 | 为分隔符 $为 函数名跟参数的分隔符
 * '请注意参数是直接加入到 Execute 字符串里面 并没有经过严格的数据验证跟类型验证 所以请写的仔细点
 * Call V.Check() '验证数据
 * 
 * If( V.Result ) Then
 *     Response.Write("验证通过，可以进行下一步处理。" & vbCrLF)
 * Else
 *     Response.Write("验证失败，请告知用户。" & vbCrLF)
 * End If
 * 
 * Dim Key

 * For Each Key In v.InfoArr
 *     Response.Write("键值是" & Key & ", 出错原因有: " & V.InfoArr.Item(Key) & vbCrLF)
 * Next
 * 
 * For Each Key In v.ResultData
 *     Response.Write("键值是" & Key & ", 处理后的数据为: " & v.ResultData.Item(key) & vbCrLF)
 * Next
 */


hasFunction = function(obj, func) { //一些验证方式必须使用到JScript 但是基本不影响全局 除了变量冲突
	if(typeof obj[func] === "unknown") return true; //如果存在此函数那么 类型为unkown
	else return false; //否则就是不存在
}

hasString = function(obj, str) {
	if(typeof obj[str] === "string") return true;
	else return false;
}
	
sprintf = function(str) {		
	var args = arguments;
	return str.replace(/%(\d)/g, function() {
		id = arguments[1];
		return typeof args[id] === arguments[0] ? "" : args[id];
	});		
}

vbUnescape = function( str ) {
	return typeof str == "undefined" ? "" : unescape(str);
}

removeScriptTag = function(a) {
	return a.replace(/<\/?script[^>]*(\/?)>/ig,"");
}

encodeText = function(a) {
	return a.replace(/>/ig, '&gt;').replace(/</ig, '&lt;');
}
</script>
<script language="vbscript" runat="server">
Class ValidatorFunctions

	Public Function Url(ByRef Obj)
		Dim Re
		Set Re = New RegExp
		Re.IgnoreCase = True
		Re.Global = True
		Re.Pattern = "^(http|https|ftp):\/\/[A-Za-z0-9]+\.[A-Za-z0-9]+[\/=\?%\-&_~`@[\]\':+!]*([^<>\""])*$"
		Url = Re.Test(CStr(Obj))
		Set Re = Nothing		
	End Function

	'参数中如果需要修改字符串 请使用ByRef 如果不修改请使用ByVal Define是Default的意思 因为Default在VBS中是关键字
	'Define并不参与验证 永远返回True 只是在你的值为空的时候输入默认值
	Public Function Define(ByRef Obj, ByVal DefaultValue)
		If ( Trim(Obj) = "" ) Then
			Obj = DefaultValue
		End If
		Define = True
	End Function
	
	Public Function isEmail(ByVal Obj)
		isEmail = Easp.Test(Obj, isEmail)
	End Function
	
	'需要函数 应该放在最前面 不应该与Define同时使用
	Public Function Required(ByVal Obj)
		If( Trim(Obj) = "" ) Then
			Required = False
		Else 
			Required = True
		End If
	End Function
	
	'这是一个长度计算 如果是最大值 应该是Max
	Public Function MaxLength(ByVal Obj, ByVal Length)
		If ( Len(Obj) > Length ) Then
			MaxLength = False
		Else
			MaxLength = True
		End If
	End Function
	
	Public Function MinLength(ByVal Obj, ByVal Length)
		If ( Len(Obj) < Length ) Then
			MinLength = False
		Else
			MinLength = True
		End If
	End Function
	
	Public Function IsNumber(ByVal Obj)
		If( IsNumeric(Obj) ) Then
			IsNumber = True
		Else
			IsNumber = False
		End If
	End Function
	
	Public Function IsUsername(ByVal Obj)
		IsUsername = Easp.Test(Obj, "^[0-9a-zA-Z]{4,}$")
	End Function
	
	Public Function IsAdminUserName(ByVal Obj)
		IsAdminUserName = Easp.Test(Obj, "^[0-9a-zA-Z]{2,}$")
	End Function
	
	Public Function ToNumber(ByRef Obj)
		Obj = IntVal(Obj)
		ToNumber = True
	End Function
	
	Public Function CBoolean(ByRef Obj)
		Obj = CBool(intval( Obj ))
		CBoolean = True
	End Function
	
	Public Function EasyRemoveTextXSS(ByRef Obj)
		Obj = encodeText(Obj)
		EasyRemoveTextXSS = True
	End Function
	
	Public Function isNotExist(ByVal Obj, ByVal TableName, ByVal FieldName) '必須事先過濾SQL注入漏洞
		sql = "SELECT TOP 1 * FROM " & TableName & " WHERE " & FieldName & " = '" & Obj & "'"
		Set Rs = conn.execute(sql)
		If( Rs.EOF ) Then
			isNotExist = True
		Else
			isNotExist = False
		End If
		
	End Function
	
	Public Function EasyRemoveTextAreaXSS(ByRef Obj)
		Obj = removeScriptTag(Obj)
		EasyRemoveTextAreaXSS = True
	End Function
	
End Class

Class ValidatorErrorString '验证提示类
	Public Required, Url, Define, MaxLength, MinLength, IsNumber, IsNotExist, isEmail, IsUsername, IsAdminUserName '先定义后使用 否则报错
	
	Private Sub Class_Initialize
		Required = "%1 不能为空。" '%1代表的是第一个参数 一般是中文名称 %2代表第二个参数 请跟函数对照
		Url = "%1 不是一个有效的URL地址。"
		Define = ""
		MaxLength = "%1 最长为 %2 个字符"
		MinLength = "%1 最短为 %2 个字符"
		IsNumber = "%1 必须为一个有效的数字"
		IsNotExist = "%1 已经存在于当前系统中，请换一个。"
		isEmail = "%1 不是一个有效的邮箱地址"
		IsUsername = "%1 不是一个有效的用户名"
		IsAdminUserName = "%1 不是一个有效的用户名"
	End Sub
	
End Class

Class Validator

	Private ValidatorObj, ValidatorStr
	Public Result, InfoArr, Data, ResultData, Method, isEncode
	
	Private Sub Class_Initialize
		Set Data = CreateObject("Scripting.Dictionary")
		Set ResultData = CreateObject("Scripting.Dictionary")
		Set InfoArr = CreateObject("Scripting.Dictionary")
		Set ValidatorObj = New ValidatorFunctions
		Set ValidatorStr = New ValidatorErrorString
		Method = "POST"		'设置提交方式为POST
		isEncode = True '是否经过URL编码
		Result = True '判断结果		
	End Sub
	
	Public Sub Add (ByVal Key, ByVal Val) '添加验证
		If ( Not isString(Key)  OR Not isString(Val) ) Then
			Call ThrowException("001", "Validator.Add(Type String, Type String)")
		End If
		Call Data.Add(Key, Val)
	End Sub
	
	Private Function isAvailableFunction( FuncName ) '利用JScript 进行判断是否存在函数
		isAvailableFunction = hasFunction(ValidatorObj, FuncName) 
	End Function
	
	Private Function isString(ByRef Obj) '判定是否为字符串
		If ( VarType(Obj) = 8 ) Then
			isString = True
		Else
			isString = False
		End If
	End Function
	
	Private Function ParseValidatorString(ByVal Str) '分析字符串获取验证函数跟参数到数组
		Dim Arr
		Arr = Split(Str, "|", -1, 1)
		ParseValidatorString = Arr
	End Function

	Private Function GetValue(ByVal Name)
		If( Method = "POST" ) Then
			If( VarType( Request.Form(Name) ) = 0 OR isEncode = False ) Then
				GetValue = Request.Form(Name)
			Else
				GetValue = vbUnescape(Request.Form(Name))
			End If
		Else
			GetValue = Request.QueryString(Name)
		End If
	End Function
	
	Private Function ParseKey(ByVal Key)
		Dim Arr
		Arr = Split(Key, "|", -1, 1)
		ParseKey = Arr
	End Function
	
	Public Function Check() '进行数据验证
	
		Dim Str, Arr, Query, Translate, TmpArr, Key, Counter, Re, Matches, Match
		For Each Key In Data
		
			Str = Data.Item(Key) '获取数据
			Arr = ParseValidatorString(Str)
			TmpArr = ParseKey(Key)
			Query = TmpArr(0) 'Form或者QueryString请求
			Translate = TmpArr(1) '中文翻译
			ResultData.Item(Query) = GetValue(Query)
			Dim Func, Args
			
			For Counter = LBound(Arr) To UBound(Arr) Step 1
				
				If( Trim(Arr(Counter)) = "" ) Then '碰到空的字符串则退出
					Exit For
				End If
				
				Set Re = New RegExp
				Re.Pattern = "^([\w]+)[$]?([\w\W]*)$"
				Re.Global = True
				Re.IgnoreCase = True
				Re.MultiLine = True
				Set Matches = re.Execute(Arr(Counter))   ' 执行搜索。
				Set Re = Nothing
				
				If( Matches.Count <> 1 ) Then
					Call ThrowException(002, "错误的验证参数。")
				End If
				
				Set Match = Matches(0)
				Func = Match.SubMatches(0)
				Args = Match.SubMatches(1)
				
				Set Matches = Nothing
				Set Match = Nothing
				
				If( Len(Args) <> 0 ) Then
					Args = ", " & Args
				End If
				
				If( Not isAvailableFunction(Func) ) Then
					Call ThrowException(003, "函数名称为[" & Func & "]不存在于验证函数类。")
				Else
					ValidatorTarget = ResultData.Item(Query)
					Call Execute("ValidatorFuncResult = ValidatorObj." & Func & "(ValidatorTarget" & Args & ")" )
					If(ValidatorFuncResult) Then
					Else
						Result = False						
						If( Not hasString(ValidatorStr, Func) ) Then
							Call ThrowException(004, "出错信息名为[" & Func & "]不存在于验证字符串类。")
						End If
						
						Call Execute("ValidatorTemp = sprintf(ValidatorStr." & Func & ", """ & Translate & """" & Args & ")")
						
						If( InfoArr.Item(Query) = "" ) Then
							InfoArr.Item(Query) = ValidatorTemp
						Else
							InfoArr.Item(Query) = InfoArr.Item(Query) & vbCrLf & ValidatorTemp
						End If
						
					End If
				End If
				
				ResultData.Item(Query) = ValidatorTarget '将验证数据返回到KV
				
			Next
			
		Next
		
	End Function
	
	Private Sub ThrowException(ByVal Code, ByVal Message)
		Call Print("发生异常，错误代码: " & Code & " , 错误原因：" & Message)
		Call Response.End()
	End Sub
	
	Private Sub Print(ByRef Str)
		Call Response.Write(Str & vbCrLf)
	End Sub
	
	Private Sub Class_Terminate   ' 设置 Terminate 事件。
    	Set Data = Nothing
		Set ResultData = Nothing
		Set InfoArr = Nothing
		Set ValidatorObj = Nothing
		Set ValidatorStr = Nothing
	End Sub
End Class

Dim ValidatorObj, ValidatorFuncResult, ValidatorTarget, ValidatorTemp, ValidatorStr

Set ValidatorObj = New ValidatorFunctions
Set ValidatorStr = New ValidatorErrorString
ValidatorFuncResult = False
ValidatorTarget = ""
ValidatorTemp = ""
</script>