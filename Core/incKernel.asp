<%
'******************* 系统核心过程函数 *******************
' Powered By：Henreal Studio
' Update：Henreal SMCS V1.0.23 Build 20170831
' Website：http://www.henreal.com
' Weixin：Henreal-Net【恒锐网络科技】
' Tel：0831-8239995 / 13700999995
'----------------------------------------------------


'=========================================
'函数名：IsObjInstalled
'作  用：检查组件是否已经安装
'参  数：strClassString ----组件名
'返回值：True  ----已经安装
'       False ----没有安装
'=========================================
Function IsObjInstalled(strClassString)
    On Error Resume Next
    IsObjInstalled = False
    Err = 0
    Dim xTestObj
    Set xTestObj = Server.CreateObject(strClassString)
    If 0 = Err Then IsObjInstalled = True
    Set xTestObj = Nothing
    Err = 0
End Function

'=========================================
'函数名：ReplaceBadChar
'作  用：过滤非法的SQL字符
'参  数：strChar-----要过滤的字符
'返回值：过滤后的字符
'=========================================
Function ReplaceBadChar(strChar)
    If strChar = "" Or IsNull(strChar) Then ReplaceBadChar = "": Exit Function
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "+,',%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ",--"
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "")
    Next
    tempChar = Replace(tempChar, "@@", "@")
    ReplaceBadChar = tempChar
End Function

Function HR_CBool(strBool)
	HR_CBool = False
    If strBool = True Or LCase(Trim(strBool)) = "true" Or LCase(Trim(strBool)) = "yes" Or Trim(strBool) = "1" Then HR_CBool = True
End Function

Function HR_CLng(ByVal str1)
	HR_CLng = 0
	If IsNull(str1) Or IsEmpty(str1) Then Exit Function
	If str1 = "" Or str1 = "0" Then Exit Function
    If IsNumeric(str1) Then
		If str1 > -2147483648 And str1 < 2147483647 Then HR_CLng = CLng(str1)
	End If
End Function

Function HR_CDbl(ByVal str1)
	HR_CDbl = 0
	If IsNull(str1) Or IsEmpty(str1) Or str1 = "" Or str1 = "0" Then Exit Function
	str1 = CStr(str1)
	If str1 <> "" Then str1 = CDbl(FormatNumber(str1))
    If IsNumeric(str1) Then HR_CDbl = str1
End Function

Function HR_CDate(ByVal str1)
	HR_CDate = Now()
    If IsDate(str1) Then HR_CDate = CDate(str1)
End Function

Function HR_IsNull(ByVal str1)		'判断字符串是否为空，主要是解决Null的问题
	HR_IsNull = False
	If isNull(Trim(str1)) Or Trim(str1) = "" Or isEmpty(Trim(str1)) Then HR_IsNull = True
End Function

Function HR_IsNumeric(ByVal str1)
	HR_IsNumeric = False
	If str1 <> "" Then HR_IsNumeric = IsNumeric(Trim(str1))
End Function

'**************************************************
'函数名：JoinChar
'作  用：向地址中加入 ? 或 &
'参  数：strUrl  ----网址
'返回值：加了 ? 或 & 的网址
'**************************************************
Function JoinChar(ByVal strUrl)
    If strUrl = "" Or IsNull(strUrl) Then JoinChar = "" : Exit Function
    If InStr(strUrl, "?") < Len(strUrl) Then
        If InStr(strUrl, "?") > 1 Then
            If InStr(strUrl, "&") < Len(strUrl) Then
                JoinChar = strUrl & "&"
            Else
                JoinChar = strUrl
            End If
        Else
            JoinChar = strUrl & "?"
        End If
    Else
        JoinChar = strUrl
    End If
End Function

'=========================================
'函数名：SaveDate()
'作  用：保存日期（若格式不正确返回NULL）
'=========================================
Function SaveDate(ByVal str1)
	SaveDate = Null
	If Len(str1) > 0 Then
		If IsDate(str1) Then SaveDate = CDate(str1)
    End If
End Function

'==================================================
'函数名：HR_CSng　（将字符转为单精度数值）
'参  数：str1 ---- 字符
'返回值：如果传入的参数不是数值，返回0
'==================================================
Function HR_CSng(ByVal str1)
	HR_CSng = 0
	If IsNumeric(str1) Then HR_CSng = CSng(str1)
End Function

'==================================================
'函数名：GetMinID
'作  用：取某一表某一字段中的最小值
'参  数：SheetName ----查询表
'        FieldName ----查询字段
'==================================================
Function GetMinID(SheetName, FieldName)
    Dim mrs
    Set mrs = ConnSW.Execute("select min(" & FieldName & ") from " & SheetName & "")
    If IsNull(mrs(0)) Then
        GetMinID = 1
    Else
        GetMinID = mrs(0)
    End If
    Set mrs = Nothing
End Function

'==================================================
'函数名：GetNewID
'作  用：取某一表某一字段中的最大值+1
'参  数：SheetName ----查询表
'        FieldName ----查询字段
'返回值：该字段最大值+1
'==================================================
Function GetNewID(SheetName, FieldName)
    Dim mrs
    Set mrs = Conn.Execute("select max(" & FieldName & ") from " & SheetName & "")
    If IsNull(mrs(0)) Then
        GetNewID = 1
    Else
        GetNewID = mrs(0) + 1
    End If
    Set mrs = Nothing
End Function

'=========================================
'函数名：nohtml
'作  用：过滤html 元素
'参  数：str ---- 要过滤字符
'返回值：没有html 的字符
'=========================================
Public Function nohtml(ByVal str)
    If IsNull(str) Or Trim(str) = "" Then
        nohtml = ""
        Exit Function
    End If
    Dim re
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = "(\<.[^\<]*\>)"
    str = re.Replace(str, " ")
    re.Pattern = "(\<\/[^\<]*\>)"
    str = re.Replace(str, " ")
    Set re = Nothing
    
    str = Replace(str, "'", "")
    str = Replace(str, Chr(34), "")
    nohtml = str
End Function

'=========================================
'函数名：ReplaceBadUrl
'作  用：过滤非法Url地址函数
'=========================================
Public Function ReplaceBadUrl(ByVal strContent)
    Dim regEx, Matches, Match

    Set regEx = New RegExp
    regEx.IgnoreCase = True
    regEx.Global = True
    
    regEx.Pattern = "(a|%61|%41)(d|%64|%44)(m|%6D|4D)(i|%69|%49)(n|%6E|%4E)(\_|%5F)(.*?)(.|%2E)(a|%61|%41)(s|%73|%53)(p|%70|%50)"
    Set Matches = regEx.Execute(strContent)
    For Each Match In Matches
        strContent = Replace(strContent, Match.Value, "")
    Next
    regEx.Pattern = "(u|%75|%55)(s|%73|%53)(e|%65|%45)(r|%72|%52)(\_|%5F)(.*?)(.|%2E)(a|%61|%41)(s|%73|%53)(p|%70|%50)"
    Set Matches = regEx.Execute(strContent)
    For Each Match In Matches
        strContent = Replace(strContent, Match.Value, "")
    Next

    Set regEx = Nothing
    ReplaceBadUrl = strContent
End Function

'=========================================
'函数：FoundInArr
'作  用：检查一个数组中所有元素是否包含指定字符串
'参  数：strArr		----存储数据数据的字串
'		strToFind	----要查找的字符串
'		strSplit	----数组的分隔符
'返回值：True,False
'=========================================
Function FoundInArr(strArr, strToFind, strSplit)
    Dim arrTemp, i
    FoundInArr = False
    If InStr(strArr, strSplit) > 0 Then
        arrTemp = Split(strArr, strSplit)
        For i = 0 To UBound(arrTemp)
        If LCase(Trim(arrTemp(i))) = LCase(Trim(strToFind)) Then
            FoundInArr = True
            Exit For
        End If
        Next
    Else
        If LCase(Trim(strArr)) = LCase(Trim(strToFind)) Then
        FoundInArr = True
        End If
    End If
End Function

'=========================================
'函数名：strLength
'作  用：求字符串长度。汉字算两个字符，英文算一个字符。
'返回值：字符串长度
'=========================================
Function strLength(str)
    On Error Resume Next
    Dim WINNT_CHINESE
    WINNT_CHINESE = (Len("中国") = 2)
    If WINNT_CHINESE Then
        Dim l, t, c
        Dim i
        l = Len(str)
        t = l
        For i = 1 To l
            c = Asc(Mid(str, i, 1))
            If c < 0 Then c = c + 65536
            If c > 255 Then
                t = t + 1
            End If
        Next
        strLength = t
    Else
        strLength = Len(str)
    End If
    If Err.Number <> 0 Then Err.Clear
End Function

'=========================================
'函数名：DelUploadFiles(文件名)
'作  用：删除多个文件【支持通配符】
'返　回：删除文件数
'注　意：多个文件名用“|”分隔
'=========================================
Function DelUploadFiles(ByVal strFiles)
	Dim strF, arrF, strF1, iDelNum
	iDelNum = 0
	On Error Resume Next
	If Len(strFiles) > 0 Then
		arrF = Split(strFiles, "|")
		For k = 0 To Ubound(arrF)
			strF1 = Trim(arrF(k))
			If Instr(strF1, "/Upload/") > 0 Then
				strF1 = Server.MapPath(strF1)
				If fso.FileExists(strF1) Then
					fso.DeleteFile strF1
					If Instr(strF1, "_S") > 0 Then
						fso.DeleteFile Replace(strF1, "_S", "")
					End If
					iDelNum = iDelNum + 1
				End If
			ElseIf (Instr(strF1, "/Upload/") = 0) And (Instr(LCase(strF1), "http://") = 0) Then
				strF1 = Server.MapPath("/Upload/" & strF1)
				If fso.FileExists(strF1) Then
					fso.DeleteFile strF1
					If Instr(strF1, "_S") > 0 Then
						fso.DeleteFile Replace(strF1, "_S", "")
					End If
					iDelNum = iDelNum + 1
				End If
				If (Instr(strF1, "*.shtml") > 0) Or (Instr(strF1, "?.shtml")) > 0 Then fso.DeleteFile strF1		'----删除指定文件
			End If
		Next
	End If
	DelUploadFiles = HR_Clng(iDelNum)
End Function

'=========================================
'函数名：CreateMultiFolder(路径)
'作  用：建立目录（多级）
'=========================================
Function CreateMultiFolder(ByVal strPath)
    On Error Resume Next
    Dim strCreate
    If strPath = "" Or IsNull(strPath) Then CreateMultiFolder = False: Exit Function
    strPath = Replace(strPath, "\", "/")
    If Right(strPath, 1) <> "/" Then strPath = strPath & "/"
    Do While InStr(2, strPath, "/") > 1
        strCreate = strCreate & Left(strPath, InStr(2, strPath, "/") - 1)
        strPath = Mid(strPath, InStr(2, strPath, "/"))
        If Not fso.FolderExists(Server.MapPath(strCreate)) Then
            fso.CreateFolder Server.MapPath(strCreate)
        End If
        If Err Then Err.Clear: CreateMultiFolder = False: Exit Function
    Loop
    CreateMultiFolder = True
End Function

'=========================================
'函数名：DeleteAllFile(指定目录)
'作  用：删除指定目录的文件及子目录
'必要提前赋值 DelFileNum, DelFolderNum 取子目录数及文件数
'=========================================
Function DeleteAllFolder(iPath)
	Dim sFolder, x, f
	If FSO.FolderExists(iPath) Then
		Set sFolder = FSO.GetFolder(iPath)
			For Each x In sFolder.Files
				If FSO.FileExists(x.Path) Then
					FSO.DeleteFile x.Path, True
					DelFileNum = DelFileNum + 1
				End If
			Next
			For Each f In sFolder.SubFolders
				If FSO.FolderExists(f.Path) Then
					DeleteAllFolder(f.Path)
					DelFolderNum = DelFolderNum + 1
				End If
			Next
		Set sFolder = Nothing
		FSO.DeleteFolder iPath, True
	End If
	DeleteAllFolder = DelFileNum & "|" & DelFolderNum
End Function

'=========================================
'函数名：ReplaceUrlBadChar
'作  用：过滤Url中非法SQL字符
'=========================================
Function ReplaceUrlBadChar(strChar)
    If strChar = "" Or IsNull(strChar) Then
        ReplaceUrlBadChar = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "+,',--,(,),<,>,[,],{,},\,;," & Chr(34) & "," & Chr(0) & ""
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "")
    Next
    tempChar = Replace(tempChar, "@@", "@")
    ReplaceUrlBadChar = tempChar
End Function

'=========================================
'函数名：GetRndPassword
'作  用：得到随便密码（参数：密码长度）
'=========================================
Function GetRndPassword(PasswordLen)
    Dim Ran, i, strPassword
    strPassword = ""
    For i = 1 To PasswordLen
        Randomize
        Ran = CInt(Rnd * 2)
        Randomize
        If Ran = 0 Then
            Ran = CInt(Rnd * 25) + 97
            strPassword = strPassword & UCase(Chr(Ran))
        ElseIf Ran = 1 Then
            Ran = CInt(Rnd * 9)
            strPassword = strPassword & Ran
        ElseIf Ran = 2 Then
            Ran = CInt(Rnd * 25) + 97
            strPassword = strPassword & Chr(Ran)
        End If
    Next
    GetRndPassword = strPassword
End Function

'**************************************************
'函数名：GetSubStr
'作  用：截字符串，汉字一个算两个字符，英文算一个字符
'参  数：str   ----原字符串
'        strlen ----截取长度
'        bShowPoint ---- 是否显示省略号
'返回值：截取后的字符串
'**************************************************
Function GetSubStr(ByVal str, ByVal strlen, bShowPoint)
    If IsNull(str) Or str = ""  Then
        GetSubStr = ""
        Exit Function
    End If
    Dim l, t, c, i, strTemp
    str = Replace(Replace(Replace(Replace(str, "&nbsp;", " "), "&quot;", Chr(34)), "&gt;", ">"), "&lt;", "<")
    l = Len(str)
    t = 0
    strTemp = str
    strlen = HR_CLng(strlen)
    For i = 1 To l
        c = Abs(Asc(Mid(str, i, 1)))
        If c > 255 Then
            t = t + 2
        Else
            t = t + 1
        End If
        If t >= strlen Then
            strTemp = Left(str, i)
            Exit For
        End If
    Next
    str = Replace(Replace(Replace(Replace(str, " ", "&nbsp;"), Chr(34), "&quot;"), ">", "&gt;"), "<", "&lt;")
    strTemp = Replace(Replace(Replace(Replace(strTemp, " ", "&nbsp;"), Chr(34), "&quot;"), ">", "&gt;"), "<", "&lt;")
    If strTemp <> str And bShowPoint = True Then
        strTemp = strTemp & "…"
    End If
    GetSubStr = strTemp
End Function

'==================================================
'函数名：GetHttpPage
'作  用：获取网页源码
'参  数：HttpUrl ------网页地址
'==================================================
Function GetHttpPage(HttpUrl, Coding)
    On Error Resume Next
    If IsNull(HttpUrl) = True Or Len(HttpUrl) < 18 Or HttpUrl = "" Then
        GetHttpPage = ""
        Exit Function
    End If
    Dim Http
    Set Http = Server.CreateObject("MSXML2.XMLHTTP")
    Http.Open "GET", HttpUrl, False
    Http.Send
    If Http.Readystate <> 4 Then
        GetHttpPage = ""
        Exit Function
    End If
    If Coding = 1 Then
        GetHttpPage = BytesToBstr(Http.ResponseBody, "UTF-8")
    ElseIf Coding = 2 Then
        GetHttpPage = BytesToBstr(Http.ResponseBody, "Big5")
    Else
        GetHttpPage = BytesToBstr(Http.ResponseBody, "GB2312")
    End If
    
    Set Http = Nothing
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Function
'==================================================
'函数名：GetHttpStr【获取远程字符串_ServerXMLHTTP方式】
'参  数：fHttpUrl：远程地址　fStrCode：编码方式
'　　　　fLoopNum：循环次数　fOverTime：超时时间(秒)
'==================================================
Function GetHttpStr(fHttpUrl, fStrCode, fLoopNum, fOverTime)
	On Error Resume Next
	GetHttpStr = "连接远程数据失败！"
	Dim funOBJ, iFun : iFun = 0
	If HR_IsNull(fHttpUrl) = False And HR_Clng(fLoopNum) > 0 And HR_Clng(fOverTime) > 0 Then
		'Set funOBJ = Server.CreateObject("MSXML2.ServerXMLHTTP")
		Set funOBJ = Server.CreateObject("Msxml2.ServerXMLHTTP.6.0")
			funOBJ.SetTimeOuts fOverTime*1000, fOverTime*1000, 300000, 360000		'超时时间，后两位为提交数据和接受数据时间5分钟和6分钟
			funOBJ.Open "GET", fHttpUrl, True	'Get方式发起异步请求
			funOBJ.send
			Do While funOBJ.readyState <> 4		'若连接失败，则重试3次		
				funOBJ.waitForResponse(1000)	'1秒后重
				iFun = iFun + 1
				If iFun > fLoopNum Then Exit Do		'重试后退出
			Loop
			GetHttpStr = funOBJ.responseText	'返回字符串
		Set funOBJ = Nothing
	End If
	If Err.number <> 0 Then Err.Clear			'清除错误提示
End Function

'==================================================
'函数名：BytesToBstr
'作  用：将获取的二进制转换为中文
'参  数：Body ------要转换的变量
'参  数：Cset ------要转换的类型
'==================================================
Function BytesToBstr(Body, Cset)
   Dim Objstream
   Set Objstream = Server.CreateObject("Adodb.Stream")
	   Objstream.Type = 1
	   Objstream.Mode =3
	   Objstream.Open
	   Objstream.Write Body
	   Objstream.Position = 0
	   Objstream.Type = 2
	   Objstream.Charset = Cset
	   BytesToBstr = Objstream.ReadText 
	   Objstream.Close
   Set Objstream = nothing
End Function

'==================================================
'函数名：PostHttpPage
'作 用：登录
'==================================================
Function PostHttpPage(RefererUrl, PostUrl, PostData) 
	Dim xmlHttp, RetStr
	Set xmlHttp = CreateObject("MSXML2.XMLHTTP") 
		xmlHttp.Open "POST", PostUrl, False
		xmlHttp.setRequestHeader "Content-Length",Len(PostData) 
		xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		xmlHttp.setRequestHeader "Referer", RefererUrl
		xmlHttp.Send PostData 
		If Err.Number <> 0 Then 
			Set xmlHttp=Nothing
			PostHttpPage = "$False$"
			Exit Function
		End If
		PostHttpPage = bytesToBSTR(xmlHttp.responseBody, "UTF-8")
	Set xmlHttp = nothing
End Function

'==================================================
'函数名：UrlEncoding
'作 用：转换编码
'==================================================
Function UrlEncoding(DataStr)
	Dim StrReturn, Si, ThisChr, InnerCode, Hight8, Low8
	StrReturn = ""
	For Si = 1 To Len(DataStr)
		ThisChr = Mid(DataStr,Si,1)
		If Abs(Asc(ThisChr)) < &HFF Then
			StrReturn = StrReturn & ThisChr
		Else
			InnerCode = Asc(ThisChr)
			If InnerCode < 0 Then
				InnerCode = InnerCode + &H10000
			End If
			Hight8 = (InnerCode And &HFF00)\ &HFF
			Low8 = InnerCode And &HFF
			StrReturn = StrReturn & "%" & Hex(Hight8) & "%" & Hex(Low8)
		End If
	Next
	UrlEncoding = StrReturn
End Function

'==================================================
'函数名：GetBody
'作 用：截取字符串
'参 数：ConStr ------将要截取的字符串
'参 数：StartStr ------开始字符串
'参 数：OverStr ------结束字符串
'参 数：IncluL ------是否包含StartStr
'参 数：IncluR ------是否包含OverStr
'==================================================
Function GetBody(ConStr, StartStr, OverStr, IncluL, IncluR)
	If ConStr = "$False$" Or ConStr = "" Or IsNull(ConStr) = True Or StartStr = "" Or IsNull(StartStr) = True Or OverStr = "" Or IsNull(OverStr) = True Then
		GetBody = "$False$"
		Exit Function
	End If
	Dim ConStrTemp, Start, Over
	ConStrTemp = Lcase(ConStr)
	StartStr = Lcase(StartStr)
	OverStr = Lcase(OverStr)
	Start = InStrB(1, ConStrTemp, StartStr, VBBinaryCompare)
	If Start <= 0 then
		GetBody = "$False$"
		Exit Function
	Else
		If IncluL = False Then
			Start = Start + LenB(StartStr)
		End If
	End If
	Over = InStrB(Start, ConStrTemp, OverStr, VBBinaryCompare)
	If Over <= 0 Or Over <= Start Then
		GetBody = "$False$"
		Exit Function
	Else
		If IncluR = True Then
			Over = Over + LenB(OverStr)
		End If
	End If
	GetBody = MidB(ConStr, Start, Over-Start)
End Function

'**************************************************
'函数名：DelRightComma
'作  用：删除字符串（如："1,3,5,8"）右侧多余的逗号以消除SQL查询时出错的问题，Comma：逗号。
'参  数：str ---- 待处理的字符串
'**************************************************
Function DelRightComma(ByVal str)
    str = Trim(str)
    If Right(str, 1) = "," Then
        str = Left(str, Len(str) - 1)
    End If
    DelRightComma = str
End Function

'**************************************************
'函数名：FilterArrNull
'作  用：过滤数组空字符
'**************************************************
Function FilterArrNull(ByVal ArrString, ByVal CompartString)
    Dim arrContent, arrTemp, i
    If HR_IsNull(CompartString) Or HR_IsNull(ArrString) Then
        FilterArrNull = ArrString : Exit Function
    End If
    If InStr(ArrString, CompartString) = 0 Then
        FilterArrNull = ArrString : Exit Function
    Else
        arrContent = Split(ArrString, CompartString)
        For i = 0 To UBound(arrContent)
            If Trim(arrContent(i)) <> "" Then
                If arrTemp = "" Then
                    arrTemp = Trim(arrContent(i))
                Else
                    arrTemp = arrTemp & CompartString & Trim(arrContent(i))
                End If
            End If
        Next
    End If
    FilterArrNull = arrTemp
End Function

'**************************************************
'函数名：XmlText
'作  用：从语言包中读取指定节点的值
'参  数：iBigNode ---- 大节点
'        iSmallNode ---- 小节点
'        DefChar ---- 默认值
'返回值：语言包中指定节点的值
'**************************************************
Function XmlText(ByVal iBigNode, ByVal iSmallNode, ByVal DefChar)
    Dim LangRoot, LangSub
    If IsNull(iBigNode) Or IsNull(iSmallNode) Then
        XmlText = DefChar
    Else
        Set LangRoot = XmlDoc.getElementsByTagName(iBigNode)
        If LangRoot.Length = 0 Then
            XmlText = DefChar
        Else
            Set LangSub = LangRoot(0).getElementsByTagName(iSmallNode)
            If LangSub.Length = 0 Then
                XmlText = DefChar
            Else
                XmlText = LangSub(0).text
            End If
        End If
        Set LangRoot = Nothing
    End If
End Function
Function GetArrXMLDoc(iNodeBig, iNodeSmall, iReStr)
	Dim strFun, arrFun, TmpFun:TmpFun = ""
	iReStr = HR_Clng(iReStr)
	If Len(iNodeBig) > 0 And Len(iNodeSmall) > 0 And iReStr > 0 Then
		strFun = XmlText(iNodeBig, iNodeSmall, "")
		If Len(strFun) > 0 Then
			arrFun = Split(strFun, "|")
			If iReStr < Ubound(arrFun) + 2 Then TmpFun = arrFun(iReStr - 1)
		End If
	End If
	GetArrXMLDoc = TmpFun
End Function

'=========================================
'函数名：ShowErrMSG
'作  用：显示错误消息（1:返回 2:返回指定地址 3:显示消息内容）
'=========================================
Private Function ShowErrMSG(iStr, iShowType, iComUrl)
	Dim fStr
	If HR_Clng(iShowType) = 1 Then
		fStr = fStr & "<script type=""text/javascript"">"
		If Len(iStr) > 0 Then fStr = fStr & "window.alert('" & iStr & "');"
		If Len(iComUrl) > 0 Then
			fStr = fStr & "location.replace('" & iComUrl & "');"
		Else
			fStr = fStr & "javascript:history.go(-1)"
		End If
		fStr = fStr & "</script>"
	Else
		fStr = "<div class=""popMSG""><em><i class=""hr-icon""><i/><span>" & iStr & "</span></em>"
		If Len(iComUrl) > 0 Then fStr = fStr & "<a href=""" & iComUrl & """>【返回】</a>"
		fStr = fStr & "</div>"
	End If
	ShowErrMSG = fStr
End Function

'=====================================================================
'　函数名：ReadFromFile(iFileName, iCharSet, iType)
'　作  用：读取文件内容 【当文本为UTF-8时，iType必为1，否则会出现乱码】
'　参　数：iFileName ----- 文件路径（相对路径）
'		   iCharSet ------ 编码：GB2312|UTF-8，默认Unicode
'		   iType --------- 写入方式：ADO|FSO，默认为FSO　1:二进制方式
'=====================================================================
Function ReadFromFile(iFileName, iCharSet, iType)
	On Error Resume Next
	Dim str1, stmU, hf, TmpERR
	str1 = ""
	TmpERR = "<html><body><div style=""position:absolute;top:50px;text-align:center;height:auto;width:100%;border:0;"">"
	TmpERR = TmpERR & "<span style=""padding:10px;padding-left:50px;color:#F00;background:url(/Static/images/icon_nav_noti.png) left center no-repeat;background-size:auto 80%"">[@ErrMSG]</span></div>"
	TmpERR = TmpERR & "</body></html>"
	iType = HR_CLng(iType)
	If iFileName = "" Then
		ReadFromFile = Replace(TmpERR, "[@ErrMSG]", "未指定文件！")
		Exit Function
	End If
	If iCharSet = "" Then iCharSet = "UTF-8"
	If iType = 1 Then		'ADO方式【解决UTF-8的乱码】
		Set stmU = Server.CreateObject("Adodb.Stream")
			stmU.Type = 2
			stmU.Mode = 3
			stmU.CharSet = iCharSet
			stmU.Open
			stmU.loadfromfile Server.MapPath(iFileName)
			str1 = stmU.readtext
			stmU.Close
		Set stmU = Nothing
	Else
		If Not FSO.FileExists(Server.MapPath(iFileName)) Then
			ReadFromFile = TmpERR
			Exit Function
		End If
		Set hf = FSO.OpenTextFile(Server.MapPath(iFileName), 1, False)
			If Not hf.AtEndOfStream Then
				str1 = hf.ReadAll
			End If
			hf.Close
		Set hf = Nothing
	End If
	If Err Then str1 = Err.Description
	If Instr(str1, "文件无法被打开") > 0 Then str1 = Replace(TmpERR, "[@ErrMSG]", "请检查模板文件是否存在！[Check Template File!]")
	If str1 = "" Or IsNull(str1) Then str1 = TmpERR
	ReadFromFile = str1
End Function

'=====================================================================
'　函数名：WriteToFile(iFileName, iContentStr, iCharSet, iType)
'　作  用：写入文件至磁盘
'　参　数：iFileName ----- 文件路径（相对路径）
'		   iContentStr --- 文档内容
'		   iCharSet ------ 编码：GB2312|UTF-8，默认Unicode
'		   iType --------- 写入方式：ADO|FSO，默认为FSO　1:二进制方式
'=====================================================================
Function WriteToFile(iFileName, iContentStr, iCharSet, iType)
	On Error Resume Next
	Err.Clear
	WriteToFile = False
	iType = HR_CLng(iType)
	If Not(iCharSet <> "") Then iCharSet = "UTF-8"
	If iFileName = "" Or iContentStr = "" Then
		Exit Function
	End If
	
	If iType = 1 Then		'ADO方式
		Dim objStream
		Set objStream = Server.CreateObject("ADODB.Stream")
			With objStream
				.Open
				.Charset = iCharSet
				.Position = objStream.Size
				.WriteText = iContentStr						'模版+数据 写入内容
				.SaveToFile Server.MapPath(iFileName), 2		'生成文件路径
				.Close
			End With
			WriteToFile = True
		Set objStream = Nothing
	Else
		Dim hf
		Set hf = FSO.OpenTextFile(Server.MapPath(iFileName), 2, True)
			hf.Write iContentStr
			hf.Close
		Set hf = Nothing
		WriteToFile = True
	End If
	If Err Then
		Err.Clear
	Else
		WriteToFile = True
	End If		
End Function

'=========================================
'函数名：FormatDate
'作  用：格式化日期函数
'参  数：DateAndTime   ----原日期
'　　　　para ----日期格式
'返回值：输出格式化后的日期
'=========================================
Public Function FormatDate(DateAndTime, para)
	On Error Resume Next
	Dim y, m, d, h, mi, s, strDateTime
	FormatDate = DateAndTime
	If Not IsNumeric(para) Then Exit Function
	If Not IsDate(DateAndTime) Then Exit Function
	y = CStr(Year(DateAndTime))
	m = CStr(Month(DateAndTime))
	If Len(m) = 1 Then m = "0" & m
	d = CStr(Day(DateAndTime))
	If Len(d) = 1 Then d = "0" & d
	h = CStr(Hour(DateAndTime))
	If Len(h) = 1 Then h = "0" & h
	mi = CStr(Minute(DateAndTime))
	If Len(mi) = 1 Then mi = "0" & mi
	s = CStr(Second(DateAndTime))
	If Len(s) = 1 Then s = "0" & s
	Select Case para
		Case "1" strDateTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
		Case "2" strDateTime = y & "-" & m & "-" & d
		Case "3" strDateTime = y & "/" & m & "/" & d
		Case "4" strDateTime = y & "年" & m & "月" & d & "日"
		Case "5" strDateTime = m & "-" & d
		Case "6" strDateTime = m & "/" & d
		Case "7" strDateTime = m & "月" & d & "日"
		Case "8" strDateTime = y & "年" & m & "月"
		Case "9" strDateTime = y & "-" & m
		Case "10" strDateTime = y & "-" & m & "-" & d & " " & h & ":" & mi
		Case "11" strDateTime = y  & m & d
		Case "12" strDateTime = y  & m & d & h & mi & s
		Case "13" strDateTime = y & m
		Case "14" strDateTime = y
		Case "15" strDateTime = m & "-" & d & " " & h & ":" & mi
		Case Else strDateTime = DateAndTime
	End Select
	FormatDate = strDateTime
End Function

'=====================================================================
'函数名：Rndstr
'作  用：得到不重复的随机数
'参　数：istart -- 开始数字  iend -- 结束数字  isum -- 同时生成多少组
'=====================================================================
Function Rndstr(istart, iend, isum)
	Dim i, j, vntarray()
	Redim vntarray(iend-istart)
	j = istart
	For i = 0 To iend-istart
		vntarray(i) = j
		j = j + 1
	Next

	Dim vntarray2(), temp, x, y
	Redim vntarray2(isum-1)
	y = iend - istart + 1
	x = 0
	temp = vntarray
	Do While x < isum
		Dim a
		Randomize
		Vntarray2(x) = temp(int(rnd*y))
		a = "" & vntarray2(x) & ""
		temp = split(trim(replace(chr(32)&join(temp)&chr(32),a,"")))
		x = x + 1
		y = y - 1
	Loop
	rndstr = join(vntarray2)
End Function
Function GetRandomNumbers(iNumLen)		'取指定个数随机数字，最多15位
	iNumLen = HR_Clng(iNumLen)
	If iNumLen > 15 Then iNumLen = 15
	Dim strRndNum, iRndNum:iRndNum = 0
	If iNumLen > 0 Then
		Randomize
		strRndNum = 10 ^ (iNumLen - 1)
		iRndNum = Int((strRndNum*10 - strRndNum) * Rnd() + strRndNum)
	End If
	GetRandomNumbers = iRndNum
End Function

'=========================================
'函数名：GetRndString(StrLength) 
'作  用：取随机字符串
'参  数：StrLength-----字符串长度
'=========================================
Function GetRndString(StrLength) 
	Dim  RndSeed, iR, TmpRnd, StrFun
	RndSeed = "0123456789"
	StrLength = HR_Clng(StrLength)
	If StrLength = 0 Then Exit Function
	For iR = 1 To StrLength 
		Randomize
		TmpRnd = Round((Rnd * (Len(RndSeed) - 1)) + 1) 
		StrFun = StrFun & Mid(RndSeed, TmpRnd, 1) 
	Next 
	GetRndString = StrFun
End Function 

'============================================
'函数名：GetTypeName
'作  用：取指定表中字段记录（判断条件为INT字段）
'============================================
Function GetTypeName(SheetName,FieldName,FieldID,TmpID)
    Dim RsGet, arrTmp, strTmp
    If SheetName <> "" and FieldName <> "" and FieldID <> "" and TmpID <> "" Then
		Set RsGet = Server.CreateObject("ADODB.RecordSet")
			RsGet.Open("Select " & FieldName & " From " & SheetName & " where " & FieldID & "=" & TmpID ), Conn, 1, 1
			If RsGet.BOF And RsGet.EOF Then
				strTmp = ""
			Else
				strTmp = RsGet(0)
			End If
		Set RsGet = Nothing
    End If
    GetTypeName = strTmp
End Function

Function strGetTypeName(SheetName,FieldName,FieldID, TmpID)		'判断条件为字符字段
    Dim RsGet, arrTmp, strTmp
    If SheetName <> "" and FieldName <> "" and FieldID <> "" and TmpID <> "" Then
    Set RsGet = Server.CreateObject("ADODB.RecordSet")
        RsGet.Open("Select " & FieldName & " From " & SheetName & " where " & FieldID & "='" & TmpID & "'"), Conn, 1, 1
        If RsGet.BOF And RsGet.EOF Then
            strTmp = ""
        Else
            strTmp = RsGet(0)
        End If
    Set RsGet = Nothing
    End If
    strGetTypeName = strTmp
End Function

'=========================================
'函数名：ReadXmlData
'作  用：读取XML数据中的常量（DataType=1时取数组串，其它取字符串）
'=========================================
Function ReadXmlData(NodeName,DataType)

    Dim objXML, Root, NodeLis, NodeCount, i, Node, Cost
    Set objXML = CreateObject("Microsoft.XMLDOM")
        objXML.async = false
        objXML.load(Server.MapPath(XMLDataPath))
    If DataType = 1 Then
        Set Root = objXML.DocumentElement.SelectSingleNode(NodeName)		'设置节点
        Set NodeLis = Root.ChildNodes
            NodeCount = NodeLis.Length

            For i = 1 to NodeCount
            Set Node = NodeLis.NextNode()
            Set Cost = Node.Attributes.GetNamedItem("ID")
                If i = 1 Then
                    ReadXmlData = ReadXmlData & Cost.text & "@@" & Node.selectSingleNode("TmpValue").text
                Else
                    ReadXmlData = ReadXmlData & "||" & Cost.text & "@@" & Node.selectSingleNode("TmpValue").text
                End If
            Set Node = Nothing
            Set Cost = Nothing
            Next

        Set NodeLis = Nothing
        Set Root =Nothing
    Else
        Set Root = objXML.DocumentElement.SelectSingleNode(NodeName)		'读取指定节点的全部数据
            ReadXmlData = Root.text
        Set Root =Nothing
    End If
End Function

'=========================================
'函数名：UpdateXmlText
'作  用：更新XML指定节点的值
'参  数：iBigNode ---- 大节点
'        iSmallNode ---- 小节点
'        upChar ---- 新值
'=========================================
Function UpdateXmlText(ByVal iBigNode, ByVal iSmallNode, ByVal upChar)
	Dim LangRoot, LangSub, RootNode, NewBigNode, NewSmallNode
	UpdateXmlText = False
    If IsNull(iBigNode) Or IsNull(iSmallNode) Then
        UpdateXmlText = False:Exit Function
    End If
    Set LangRoot = XmlDoc.getElementsByTagName(iBigNode)
		If LangRoot.Length = 0 Then			'当大节点不存在时创建
			Set RootNode = XmlDoc.getElementsByTagName("Root")
				Set NewBigNode = XmlDoc.createElement(iBigNode)
					RootNode(0).appendChild(NewBigNode)
					NewBigNode.setAttribute "Value", ""
				Set NewBigNode = Nothing
			Set RootNode = Nothing
		End If
    Set LangRoot = Nothing
    
	Set LangRoot = XmlDoc.getElementsByTagName(iBigNode)
		Set LangSub = LangRoot(0).getElementsByTagName(iSmallNode)
			If LangSub.Length = 0 Then		'创建新的子节点
				Set NewSmallNode = XmlDoc.createElement(iSmallNode)
					LangRoot(0).appendChild NewSmallNode
					NewSmallNode.text = upChar
				Set NewSmallNode = Nothing
			Else
				LangSub(0).text = upChar
            End If
		Set LangSub = Nothing
		XmlDoc.Save(Server.MapPath(XMLDataPath))
		UpdateXmlText = True
	Set LangRoot = Nothing
End Function

Function UpdateXmlCDATA(ByVal iBigNode, ByVal iSmallNode, ByVal upChar)
    Dim LangRoot, LangSub, RootNode, NewBigNode, NewSmallNode
    UpdateXmlCDATA = False
	If IsNull(iBigNode) Or IsNull(iSmallNode) Then
        UpdateXmlCDATA = False:Exit Function
    End If
    Set LangRoot = XmlDoc.getElementsByTagName(iBigNode)
		If LangRoot.Length = 0 Then			'当大节点不存在时创建
			Set RootNode = XmlDoc.getElementsByTagName("Root")
				Set NewBigNode = XmlDoc.createElement(iBigNode)
					RootNode(0).appendChild(NewBigNode)
					NewBigNode.setAttribute "Value", ""
				Set NewBigNode = Nothing
			Set RootNode = Nothing
		End If
    Set LangRoot = Nothing
	Set LangRoot = XmlDoc.getElementsByTagName(iBigNode)
		Set LangSub = LangRoot(0).getElementsByTagName(iSmallNode)
			If LangSub.Length = 0 Then		'创建新的子节点
				Set NewSmallNode = XmlDoc.createElement(iSmallNode)
					LangRoot(0).appendChild NewSmallNode
					NewSmallNode.text = ""
					NewSmallNode.appendChild(XmlDoc.createCDATASection(upChar))
				Set NewSmallNode = Nothing
			Else
				LangSub(0).text = ""
				LangSub(0).appendChild(XmlDoc.createCDATASection(upChar))
            End If
		Set LangSub = Nothing
		XmlDoc.Save(Server.MapPath(XMLDataPath))
		UpdateXmlCDATA = True
	Set LangRoot = Nothing
End Function

'=================================================
'函数名：FilterJS()
'作  用：过滤非法JS字符
'参  数：strInput 需要过滤的内容
'=================================================
Public Function FilterJS(ByVal strInput)
    If IsNull(strInput) Or Trim(strInput) = "" Then
        FilterJS = ""
        Exit Function
    End If
    Dim RegEx
    Dim reContent
    Set RegEx = New RegExp    
    RegEx.IgnoreCase = True
    RegEx.Global = True
	RegEx.MultiLine = True

	' 替换掉HTML字符实体(Character Entities)名字和分号之间的空白字符，比如：&auml    ;替换成&auml;
	RegEx.Pattern = "(&#*\w+)[\x00-\x20]+;"
	strInput = RegEx.Replace(strInput, "$1;")

	' 将无分号结束符的数字编码实体规范成带分号的标准形式
	RegEx.Pattern = "(&#x*[0-9A-F]+);*"
	strInput = RegEx.Replace(strInput, "$1;")

	' 将&nbsp; &lt; &gt; &amp; &quot;字符实体中的 & 替换成 &amp; 以便在进行HtmlDecode时保留这些字符实体
	'RegEx.Pattern = "&(amp|lt|gt|nbsp|quot);"
	'strInput = RegEx.Replace(strInput, "&amp;$1;")

	' 将HTML字符实体进行解码，以消除编码字符对后续过滤的影响
	'strInput = HtmlDecode(strInput);

	' 将ASCII码表中前32个字符中的非打印字符替换成空字符串，保留 9、10、13、32，它们分别代表 制表符、换行符、回车符和空格。
	RegEx.Pattern = "[\x00-\x08\x0b-\x0c\x0e-\x19]"
	strInput = RegEx.Replace(strInput,  "")

	' 替换以on和xmlns开头的属性，动易系统的几个JS需要保留
	RegEx.Pattern = "on(load\s*=\s*""*'*resizepic\(this\)'*""*)"
	strInput = RegEx.Replace(strInput, "off$1") 
	RegEx.Pattern = "on(mousewheel\s*=\s*""*'*return\s*bbimg\(this\)'*""*)"
	strInput = RegEx.Replace(strInput, "off$1") 

	RegEx.Pattern = "(<[^>]+[\x00-\x20""'/])(on|xmlns)([^>]*)>"
	strInput = RegEx.Replace(strInput, "$1pe$3>") 

	RegEx.Pattern = "off(load\s*=\s*""*'*resizepic\(this\)'*""*)"
	strInput = RegEx.Replace(strInput, "on$1") 
	RegEx.Pattern = "off(mousewheel\s*=\s*""*'*return\s*bbimg\(this\)'*""*)"
	strInput = RegEx.Replace(strInput, "on$1") 


	' 替换javascript
	RegEx.Pattern = "([a-z]*)[\x00-\x20]*=[\x00-\x20]*([`'""]*)[\x00-\x20]*j[\x00-\x20]*a[\x00-\x20]*v[\x00-\x20]*a[\x00-\x20]*s[\x00-\x20]*c[\x00-\x20]*r[\x00-\x20]*i[\x00-\x20]*p[\x00-\x20]*t[\x00-\x20]*:"
	strInput = RegEx.Replace(strInput, "$1=$2nojavascript...")

	' 替换vbscript
	RegEx.Pattern = "([a-z]*)[\x00-\x20]*=[\x00-\x20]*([`'""]*)[\x00-\x20]*v[\x00-\x20]*b[\x00-\x20]*s[\x00-\x20]*c[\x00-\x20]*r[\x00-\x20]*i[\x00-\x20]*p[\x00-\x20]*t[\x00-\x20]*:"
	strInput = RegEx.Replace(strInput, "$1=$2novbscript...")

	' 替换expression
	RegEx.Pattern = "(<[^>]+)style[\x00-\x20]*=[\x00-\x20]*([`'""]*).*expression[\x00-\x20]*\([^>]*>"
	strInput = RegEx.Replace(strInput, "$1>")

	' 替换behaviour
	RegEx.Pattern = "(<[^>]+)style[\x00-\x20]*=[\x00-\x20]*([`'""]*).*behaviour[\x00-\x20]*\([^>]*>"
	strInput = RegEx.Replace(strInput, "$1>")

	' 替换script
	RegEx.Pattern = "(<[^>]+)style[\x00-\x20]*=[\x00-\x20]*([`'""]*).*s[\x00-\x20]*c[\x00-\x20]*r[\x00-\x20]*i[\x00-\x20]*p[\x00-\x20]*t[\x00-\x20]*:*[^>]*>"
	strInput = RegEx.Replace(strInput, "$1>")

	' 替换namespaced elements 不需要
	RegEx.Pattern = "</*\w+:\w[^>]*>"
	strInput = RegEx.Replace(strInput, "")

	Dim oldhtmlString
	oldhtmlString = ""
	Do while oldhtmlString <> strInput
		oldhtmlString = strInput
		'实行严格过滤
		RegEx.Pattern = "</*(applet|meta|xml|blink|link|style|script|iframe|frame|frameset|ilayer|layer|bgsound|title|base|embed|object)[^>]*>"
		strInput = RegEx.Replace(strInput, "")
		'过滤掉SHTML的Include包含文件漏洞
		RegEx.Pattern = "<!--\s*#include[^>]*>"
		strInput = RegEx.Replace(strInput, "")
		'If FilterLevel > 0 Then
		'	'实行严格过滤
		'	RegEx.Pattern = "</*(embed|object)[^>]*>"
		'	strInput = RegEx.Replace(strInput, "")
		'End If
	Loop
	Set RegEx = Nothing
    FilterJS = strInput
End Function

'=========================================
'函数名：Refresh
'作  用：等待特定时间后跳转到指定的网址
'参  数：url ---- 跳转网址
'        refreshTime ---- 等待跳转时间
'=========================================
Sub Refresh(url, refreshTime)
	Response.Write "<a Name='rsfreshurl' ID='rsfreshurl' href='"& url &"'></a>" & vbCrLf
	Response.Write "<script type=""text/javascript"">" & vbCrLf
	Response.Write "	function nextpage(){" & vbCrLf
	Response.Write "		var url = document.getElementById('rsfreshurl');" & vbCrLf
	Response.Write "		if (document.all) {" & vbCrLf
	Response.Write "			url.click();" & vbCrLf
	Response.Write "		}" & vbCrLf
	Response.Write "		else if (document.createEvent) {" & vbCrLf
	Response.Write "			var ev = document.createEvent('HTMLEvents');" & vbCrLf
	Response.Write "			ev.initEvent('click', false, true);" & vbCrLf
	Response.Write "			url.dispatchEvent(ev);" & vbCrLf
	Response.Write "		}" & vbCrLf
	Response.Write "	}" & vbCrLf
	Response.Write "	setTimeout(""nextpage();""," & refreshTime * 1000 & ");" & vbCrLf
	Response.Write "</script>" & vbCrLf
End Sub

'=====================================================================
'函数名：GetShowBit(vBit, ShowType)
'作  用：显示×或√
'参　数：vBit:True/False，ShowType:显示方式
'=====================================================================
Function GetShowBit(vBit, ShowType)
	Dim strTmp
	Select Case ShowType
		Case 1
			strTmp = "<b class=""hr-color-false"">否</b>"
			If HR_CBool(vBit) Then strTmp = "<b class=""hr-color-true"">是</b>"
		Case 2
			strTmp = "<i class=""hr-false""></i>"
			If HR_CBool(vBit) Then strTmp = "<i class=""hr-icon hr-true"">&#xe93e;</i>"
		Case Else
			strTmp = "<i class=""hr-icon hr-false"">&#xe960;</i>"
			If HR_CBool(vBit) Then strTmp = "<i class=""hr-icon hr-true"">&#xe95f;</i>"
	End Select
	GetShowBit = strTmp
End Function

'====================================================================
'函数名：ChkDataTable(iTableName)
'作  用：检查数据库中表是否存在（返回False/True）
'"Create Table 表名(主键名 int NOT NULL PRIMARY KEY,文本字段 nvarchar(255) NOT NULL,整形字段 int NULL,,时间字段 datetime NULL" ntext/numeric/money
'====================================================================
Function ChkDataTable(iTableName, IsCreate)
	On Error Resume Next
	Err.Clear
	ChkDataTable = False
	Dim rsChkTable
	Err.Clear
	iTableName = ReplaceBadChar(iTableName)
	Set rsChkTable = Server.CreateObject("ADODB.RecordSet")
		rsChkTable.Open("Select * From " & iTableName & ""), Conn, 1, 1
		If Err.Number <> 0 Then
			Err.Clear
			If HR_CBool(IsCreate) Then
				Conn.Execute("Create Table " & iTableName & "(ID int NOT NULL PRIMARY KEY,CreateTime datetime NULL,UpdateTime datetime NULL)")
			End If
		Else
			ChkDataTable = True
		End If
	Set rsChkTable = Nothing
End Function

'**************************************************
'函数名：HR_HTMLEncode
'作  用：将html 标记替换成 能在IE显示的HTML
'参  数：fString ---- 要处理的字符串
'返回值：处理后的字符串
'**************************************************
Function HR_HTMLEncode(ByVal fString)
    If IsNull(fString) Or Trim(fString) = "" Then
        HR_HTMLEncode = ""
        Exit Function
    End If
    fString = Replace(fString, ">", "&gt;")
    fString = Replace(fString, "<", "&lt;")

    fString = Replace(fString, Chr(32), "&nbsp;")
    fString = Replace(fString, Chr(9), "&nbsp;")
    fString = Replace(fString, Chr(34), "&quot;")
	fString = Replace(fString, Chr(39), "&#39;")
	fString = Replace(fString, Chr(9), "")
	fString = Replace(fString, "\", "＼")
    fString = Replace(fString, Chr(13), "")
    fString = Replace(fString, Chr(10), "<br />")

    HR_HTMLEncode = fString
End Function

'**************************************************
'函数名：HR_HtmlDecode
'作  用：还原Html标记,配合HR_HTMLEncode 使用
'参  数：fString ---- 要处理的字符串
'返回值：处理后的字符串
'**************************************************
Function HR_HtmlDecode(ByVal fString)
    If IsNull(fString) Or Trim(fString) = "" Then
        HR_HtmlDecode = ""
        Exit Function
    End If
    fString = Replace(fString, "&gt;", ">")
    fString = Replace(fString, "&lt;", "<")

    fString = Replace(fString, "&nbsp;", " ")
    fString = Replace(fString, "&quot;", Chr(34))
    fString = Replace(fString, "&#39;", Chr(39))
	'fString = Replace(fString, "</P><P> ", Chr(10) & Chr(10))
    fString = Replace(fString, "<br /><br />", Chr(10))

    HR_HtmlDecode = fString
End Function

'=========================================
'函数名：GetCurrentPath【返回当前路径】
'1：http://www.***.com/***/***.***
'2：http://www.***.com/***/
'3：/***/
'0：/***/***.***
'=========================================
Function GetCurrentPath(fType)
    Dim fPath:fPath = LCase(Request.ServerVariables("URL"))
	Dim rUrlLen:rUrlLen = InstrRev(fPath, "/")
	Dim lUrl:lUrl = Left(fPath, rUrlLen)
	Dim fFile:fFile = Mid(fPath, rUrlLen + 1)
    Select Case fType
		Case 1
			fPath = "http://" & Request.ServerVariables("Server_Name") & fPath
		Case 2
			fPath = "http://" & Request.ServerVariables("Server_Name") & lUrl
		Case 3
			fPath = lUrl
    End Select
    GetCurrentPath = fPath
End Function

'=====================================================================
'函数名：GetThisPageUrl		【返回当前页URL】
'用法：fDomain　是否带域名，fParam：1带参数,2仅返回文件名,3仅返回域名,5仅返回路径
'=====================================================================
Private Function GetThisPageUrl(fDomain, fParam)
	Dim strFun, fStrParam, fArrParam, fURL, fPath, fFileName
	Dim fPORT : fPORT = ":" & Request.ServerVariables("SERVER_PORT")			'取端口
	If fPORT = ":80" Or fPORT = ":443" Then fPORT = ""							'HTTPS时端口为443
	Dim fHttp : fHttp = "https://"
	If Request.ServerVariables("HTTPS") = "off" Then fHttp = "http://"			'判断HTTPS
	Dim fHost : fHost = Request.ServerVariables("HTTP_HOST")					'取域名
	If Instr(fHost, fPORT) = 0 And HR_IsNull(fPORT)=False Then fHost = fHost & fPORT

	fURL = Request.ServerVariables("HTTP_X_ORIGINAL_URL")				'仅IIS7 + Rewrite Module才能取此值
	If HR_IsNull(fURL) Then
		fURL = Trim(Request.ServerVariables("SCRIPT_NAME"))
		fStrParam = Trim(Request.QueryString())							'取请求的参数
	Else
		If Instr(fURL, "?") > 0 Then
			fArrParam = Split(fURL, "?") : fURL = fArrParam(0) : fStrParam = fArrParam(1)
		End If
	End If
	If HR_IsNull(fURL)=False Then					'取路径及文件名
		Dim fArrPath : fArrPath = Split(fURL, "/")
		fFileName = fArrPath(Ubound(fArrPath)) : fPath = Replace(fURL, fFileName, "")
	End If

	If HR_CBool(fDomain) Then strFun = fHttp & "" & fHost			'是否添加域名
	If HR_Clng(fParam)=5 Then		'仅返回路径
		strFun = strFun & fPath
	ElseIf HR_Clng(fParam)=2 Then		'仅返回文件名
		strFun = fFileName
	ElseIf HR_Clng(fParam)=3 Then		'仅返回域名
		strFun = fHttp & fHost
	Else
		strFun = strFun & fURL
	End If
	If HR_Clng(fParam)=1 And HR_IsNull(fStrParam) = False Then strFun = strFun & "?" & fStrParam		'返回文件名+参数
	GetThisPageUrl = strFun
End Function


'******************* 2017新增部分 *******************

'=====================================================
' 函数名：GetUserAgent()	【取用户来源】
' 返回：string (来源关键字)
'=====================================================
Function GetUserAgent()
	Dim strFun, funAgent
	funAgent = Request.ServerVariables("http_user_agent")
	If Instr(funAgent, "wxwork") Then		'企业微信
		strFun = "wxwork"
	ElseIf Instr(funAgent, "MicroMessenger") Then	'微信
		strFun = "weixin"
	ElseIf Instr(funAgent, "iPhone") Then
		strFun = "iPhone"
	ElseIf Instr(funAgent, "Android") Then
		strFun = "Android"
	ElseIf Instr(funAgent, "Trident") Then
		strFun = "IE"
	ElseIf Instr(funAgent, "Edge") Then
		strFun = "Edge"
	ElseIf Instr(funAgent, "iPad") Then
		strFun = "iPad"
	ElseIf Instr(funAgent, "AppleWebKit") Then
		strFun = "Chrome"
	Else
		strFun = "Other"
	End If
	GetUserAgent = strFun
End Function

'=====================================================================
'函数名：parseJSON		【解析JSON】
'返回值：string(JSON对象)
'用法：Dim obj, scriptCtrl : Set obj = parseJSON(json)
'=====================================================================
Function parseJSON(str)
	on error resume next
	If Not IsObject(scriptCtrl) Then
		Set scriptCtrl = Server.CreateObject("MSScriptControl.ScriptControl")
			scriptCtrl.Language = "JScript"
			scriptCtrl.AddCode "Array.prototype.get = function(x) { return this[x]; }; var result = null;"
	End If
	scriptCtrl.ExecuteStatement "result = " & str & ";"
	Set parseJSON = scriptCtrl.CodeObject.result
		If Not Err.Number=0 Then
			Err.Clear
			str = "{""code"":500,""msg"":""解析Excel数据出错"",""count"":0,""data"":[]}"
			Set parseJSON = Nothing
			scriptCtrl.ExecuteStatement "result = " & str & ";"
			Set parseJSON = scriptCtrl.CodeObject.result
		End If

End Function

'=====================================================================
'函数名：ToUnixTime		【把标准时间转换为UNIX时间戳】
'返回值：Number(10位bigint)
'=====================================================================
Function ToUnixTime(strTime, intTimeZone)
    If IsEmpty(strTime) or Not IsDate(strTime) Then strTime = Now
    If IsEmpty(intTimeZone) or Not isNumeric(intTimeZone) Then intTimeZone = 0
    ToUnixTime = DateAdd("h",-intTimeZone,strTime)
    ToUnixTime = DateDiff("s","1970-01-01 00:00:00", ToUnixTime)
End Function

'=====================================================================
'函数名：FromUnixTime		【把UNIX时间戳转换为标准时间】
'返回值：标准时间(DATETIME)
'=====================================================================
Function FromUnixTime(intTime, intTimeZone)
    If IsEmpty(intTime) or Not IsNumeric(intTime) Then
        FromUnixTime = Now()
        Exit Function
    End If
    If IsEmpty(intTime) or Not IsNumeric(intTimeZone) Then intTimeZone = 0
    FromUnixTime = DateAdd("s", intTime, "1970-01-01 00:00:00")
    FromUnixTime = DateAdd("h", intTimeZone, FromUnixTime)
End Function

'调用方法：
'示例：ToUnixTime("2009-12-05 12:52:25", +8)，返回值为1259988745 
'response.Write ToUnixTime("2009-12-05 12:52:25", +8)
'示例：FromUnixTime("1259988745", +8)，返回值2009-12-05 12:52:25 
'response.Write FromUnixTime("1259988745", +8)

'=====================================================================
'函数名：ChkDateExpired		【判断日期是否过期，参考当前时间】
'返回值：True/False
'=====================================================================
Function ChkDateExpired(fDateTime)
	ChkDateExpired = False
	If HR_IsNull(fDateTime) = False And IsDate(fDateTime) Then
		If DateDiff("s", fDateTime, Now()) <= 0 Then ChkDateExpired = True
	End If
End Function


'*********** 2022新增 ***********

'====================================================================
'函数名：GetTableDataQuery(表名, 指定字段, 指定行数, 查询条件)
'作  用：返回数据表查询结果【无字段则为全字段】
'====================================================================
Function GetTableDataQuery(fTable, fStrField, fRows, fQuery)
	Dim funArr, fField, sqlFun, rsFun, iFun, jFun
	If HR_IsNull(fTable) Or ChkDataTable(fTable, False)=False Then		'未指定表名或表不存在时返回空数组[二维数组]
		Redim funArr(0,0)
		funArr(0,0) = "NULL": GetTableDataQuery = funArr : Exit Function
	End If
	fRows = HR_CLng(fRows) : sqlFun = "Select"
	If fRows>0 Then sqlFun = "Select Top " & fRows

	If HR_IsNull(fStrField) = False Then
		sqlFun = sqlFun & " " & fStrField & " From " & fTable
	Else
		sqlFun = sqlFun & " * From " & fTable
	End If

	If HR_IsNull(fQuery) = False Then sqlFun = sqlFun & " Where " & fQuery			'增加查询条件
	Set rsFun = Server.CreateObject("ADODB.RecordSet")
		rsFun.Open(sqlFun), Conn, 1, 1
		Redim funArr(rsFun.Fields.count, 0) : iFun = 0
		For Each fField in rsFun.Fields			'数组第一行为字段名
			funArr(iFun, 0) = fField.Name : iFun = iFun + 1
		Next
		If Not(rsFun.BOF And rsFun.EOF) Then
			Redim Preserve funArr(rsFun.Fields.count, rsFun.RecordCount+1) : iFun = 1		'从第二行开始存数据
			Do While Not rsFun.EOF
				jFun = 0
				For Each fField in rsFun.Fields
					funArr(jFun, iFun) = fField.Value
					jFun = jFun + 1
				Next
				rsFun.MoveNext
				iFun = iFun + 1
			Loop
		Else
			Redim Preserve funArr(rsFun.Fields.count, 1)		'从第二行返回空数据
		End If
	Set rsFun = Nothing
	GetTableDataQuery = funArr
End Function
%>
