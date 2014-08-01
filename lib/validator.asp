<script type="text/javascript" language="javascript" runat="server">
/**
 *******************************************************************************************
 * VBScirpt ASP ������֤�� ������֤��ͬʱ�����޸�ԭʼ���� 
 * @author: QPWOEIRU96
 * @date: 2012��3��26��
 * @version: alpha
 * @site: http://sou.la/blog
 * @license: ��������ʹ��[����Ӧ�ó�����ҵĿ�Ļ��ߴӴ�ı������]������޸���Դ�������֪�ҡ�
 ********************************************************************************************
 * @Code Sample[VBScript]
 * Dim V
 * Set V = New Validator 'ʵ����
 * V.Method = "GET" '�����ύ��ʽ
 * Call V.add("Test|��������asd","Define$""360""|Required|Url|Max$8") '�����֤����
 * Call V.add("AS|��������ad2","Required|Define$""360""|Url") '�����֤����
 * '��֤����Ϊ�� {POST��������GET������}|{��������}  ������ {������}${����}|{������}${����} 
 * '���� | Ϊ�ָ��� $Ϊ �������������ķָ���
 * '��ע�������ֱ�Ӽ��뵽 Execute �ַ������� ��û�о����ϸ��������֤��������֤ ������д����ϸ��
 * Call V.Check() '��֤����
 * 
 * If( V.Result ) Then
 *     Response.Write("��֤ͨ�������Խ�����һ������" & vbCrLF)
 * Else
 *     Response.Write("��֤ʧ�ܣ����֪�û���" & vbCrLF)
 * End If
 * 
 * Dim Key

 * For Each Key In v.InfoArr
 *     Response.Write("��ֵ��" & Key & ", ����ԭ����: " & V.InfoArr.Item(Key) & vbCrLF)
 * Next
 * 
 * For Each Key In v.ResultData
 *     Response.Write("��ֵ��" & Key & ", ����������Ϊ: " & v.ResultData.Item(key) & vbCrLF)
 * Next
 */


hasFunction = function(obj, func) { //һЩ��֤��ʽ����ʹ�õ�JScript ���ǻ�����Ӱ��ȫ�� ���˱�����ͻ
	if(typeof obj[func] === "unknown") return true; //������ڴ˺�����ô ����Ϊunkown
	else return false; //������ǲ�����
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

	'�����������Ҫ�޸��ַ��� ��ʹ��ByRef ������޸���ʹ��ByVal Define��Default����˼ ��ΪDefault��VBS���ǹؼ���
	'Define����������֤ ��Զ����True ֻ�������ֵΪ�յ�ʱ������Ĭ��ֵ
	Public Function Define(ByRef Obj, ByVal DefaultValue)
		If ( Trim(Obj) = "" ) Then
			Obj = DefaultValue
		End If
		Define = True
	End Function
	
	Public Function isEmail(ByVal Obj)
		isEmail = Easp.Test(Obj, isEmail)
	End Function
	
	'��Ҫ���� Ӧ�÷�����ǰ�� ��Ӧ����Defineͬʱʹ��
	Public Function Required(ByVal Obj)
		If( Trim(Obj) = "" ) Then
			Required = False
		Else 
			Required = True
		End If
	End Function
	
	'����һ�����ȼ��� ��������ֵ Ӧ����Max
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
	
	Public Function isNotExist(ByVal Obj, ByVal TableName, ByVal FieldName) '��������^�VSQLע��©��
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

Class ValidatorErrorString '��֤��ʾ��
	Public Required, Url, Define, MaxLength, MinLength, IsNumber, IsNotExist, isEmail, IsUsername, IsAdminUserName '�ȶ����ʹ�� ���򱨴�
	
	Private Sub Class_Initialize
		Required = "%1 ����Ϊ�ա�" '%1������ǵ�һ������ һ������������ %2����ڶ������� �����������
		Url = "%1 ����һ����Ч��URL��ַ��"
		Define = ""
		MaxLength = "%1 �Ϊ %2 ���ַ�"
		MinLength = "%1 ���Ϊ %2 ���ַ�"
		IsNumber = "%1 ����Ϊһ����Ч������"
		IsNotExist = "%1 �Ѿ������ڵ�ǰϵͳ�У��뻻һ����"
		isEmail = "%1 ����һ����Ч�������ַ"
		IsUsername = "%1 ����һ����Ч���û���"
		IsAdminUserName = "%1 ����һ����Ч���û���"
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
		Method = "POST"		'�����ύ��ʽΪPOST
		isEncode = True '�Ƿ񾭹�URL����
		Result = True '�жϽ��		
	End Sub
	
	Public Sub Add (ByVal Key, ByVal Val) '�����֤
		If ( Not isString(Key)  OR Not isString(Val) ) Then
			Call ThrowException("001", "Validator.Add(Type String, Type String)")
		End If
		Call Data.Add(Key, Val)
	End Sub
	
	Private Function isAvailableFunction( FuncName ) '����JScript �����ж��Ƿ���ں���
		isAvailableFunction = hasFunction(ValidatorObj, FuncName) 
	End Function
	
	Private Function isString(ByRef Obj) '�ж��Ƿ�Ϊ�ַ���
		If ( VarType(Obj) = 8 ) Then
			isString = True
		Else
			isString = False
		End If
	End Function
	
	Private Function ParseValidatorString(ByVal Str) '�����ַ�����ȡ��֤����������������
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
	
	Public Function Check() '����������֤
	
		Dim Str, Arr, Query, Translate, TmpArr, Key, Counter, Re, Matches, Match
		For Each Key In Data
		
			Str = Data.Item(Key) '��ȡ����
			Arr = ParseValidatorString(Str)
			TmpArr = ParseKey(Key)
			Query = TmpArr(0) 'Form����QueryString����
			Translate = TmpArr(1) '���ķ���
			ResultData.Item(Query) = GetValue(Query)
			Dim Func, Args
			
			For Counter = LBound(Arr) To UBound(Arr) Step 1
				
				If( Trim(Arr(Counter)) = "" ) Then '�����յ��ַ������˳�
					Exit For
				End If
				
				Set Re = New RegExp
				Re.Pattern = "^([\w]+)[$]?([\w\W]*)$"
				Re.Global = True
				Re.IgnoreCase = True
				Re.MultiLine = True
				Set Matches = re.Execute(Arr(Counter))   ' ִ��������
				Set Re = Nothing
				
				If( Matches.Count <> 1 ) Then
					Call ThrowException(002, "�������֤������")
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
					Call ThrowException(003, "��������Ϊ[" & Func & "]����������֤�����ࡣ")
				Else
					ValidatorTarget = ResultData.Item(Query)
					Call Execute("ValidatorFuncResult = ValidatorObj." & Func & "(ValidatorTarget" & Args & ")" )
					If(ValidatorFuncResult) Then
					Else
						Result = False						
						If( Not hasString(ValidatorStr, Func) ) Then
							Call ThrowException(004, "������Ϣ��Ϊ[" & Func & "]����������֤�ַ����ࡣ")
						End If
						
						Call Execute("ValidatorTemp = sprintf(ValidatorStr." & Func & ", """ & Translate & """" & Args & ")")
						
						If( InfoArr.Item(Query) = "" ) Then
							InfoArr.Item(Query) = ValidatorTemp
						Else
							InfoArr.Item(Query) = InfoArr.Item(Query) & vbCrLf & ValidatorTemp
						End If
						
					End If
				End If
				
				ResultData.Item(Query) = ValidatorTarget '����֤���ݷ��ص�KV
				
			Next
			
		Next
		
	End Function
	
	Private Sub ThrowException(ByVal Code, ByVal Message)
		Call Print("�����쳣���������: " & Code & " , ����ԭ��" & Message)
		Call Response.End()
	End Sub
	
	Private Sub Print(ByRef Str)
		Call Response.Write(Str & vbCrLf)
	End Sub
	
	Private Sub Class_Terminate   ' ���� Terminate �¼���
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