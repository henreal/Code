<%
Function ChkTeacherKPI(fuYGDM)		'检查员工KPI记录是否存在
	Dim rs2, fuYGXM, fuKSDM, fuKSMC, fuPRZC, fuYGXB
	If HR_Clng(fuYGDM) > 0 Then
		Set rs2 = Conn.Execute("Select * From HR_Teacher Where Cast(YGDM As Int)=" & HR_Clng(fuYGDM) )
			If Not(rs2.BOF And rs2.EOF) Then
				fuYGXM = Trim(rs2("YGXM"))
				fuKSDM = HR_Clng(rs2("KSDM"))
				fuYGXB = Trim(rs2("YGXB"))
				fuKSMC = Trim(rs2("KSMC"))
				fuPRZC = Trim(rs2("PRZC"))
			End If
		Set rs2 = Nothing
		Set rs2 = Server.CreateObject("ADODB.RecordSet")
			rs2.Open("Select * From HR_KPI Where YGDM=" & fuYGDM & " And scYear=" & DefYear), Conn, 1, 3
			If rs2.BOF And rs2.EOF Then
				rs2.AddNew
				rs2("ID") = GetNewID("HR_KPI", "ID")
				rs2("scYear") = DefYear
				rs2("YGDM") = HR_Clng(fuYGDM)
				rs2("YGXM") = fuYGXM
				rs2("YGXB") = Trim(fuYGXB)
				rs2("KSDM") = HR_Clng(fuKSDM)
				rs2("KSMC") = fuKSMC
				rs2("PRZC") = fuPRZC
				rs2.Update
			End If
		Set rs2 = Nothing
		Set rs2 = Server.CreateObject("ADODB.RecordSet")		'更新累计表
			rs2.Open("Select * From HR_KPI_SUM Where YGDM=" & fuYGDM & " And scYear=" & DefYear), Conn, 1, 3
			If rs2.BOF And rs2.EOF Then
				rs2.AddNew
				rs2("ID") = GetNewID("HR_KPI_SUM", "ID")
				rs2("scYear") = DefYear
				rs2("YGDM") = HR_Clng(fuYGDM)
				rs2("YGXM") = fuYGXM
				rs2("YGXB") = Trim(fuYGXB)
				rs2("KSDM") = HR_Clng(fuKSDM)
				rs2("KSMC") = fuKSMC
				rs2("PRZC") = fuPRZC
				rs2.Update
			End If
		Set rs2 = Nothing
	End If
End Function

Function UpdateTeacherKPI(fuItemID, fuYGDM, fuType)		'统计当前项目(列)
	Dim rs2, sql2, fuStr : fuStr = False
	Dim iSum, fuMaxSum, iScore, tTemplate, fuTable, fuField, iFormula, fuStrStuType, fuArrStuType, iFuStu
	If HR_Clng(fuItemID) > 0 And HR_Clng(fuYGDM) > 0 Then
		fuTable = "HR_Sheet_" & fuItemID
		tTemplate = GetTypeName("HR_Class", "Template", "ClassID", fuItemID)
		fuStrStuType = GetTypeName("HR_Class", "StudentType", "ClassID", fuItemID)		'取学生类别
		fuMaxSum = GetTypeName("HR_Class", "MaxScore", "ClassID", fuItemID)		'取最高分
		fuMaxSum = HR_CDbl(fuMaxSum)
		fuStrStuType = FilterArrNull(fuStrStuType, ",")

		If ChkTable(fuTable) Then			'检查表是否存在
			If HR_IsNull(fuStrStuType) Then			'无学生类别时直接统计
				sql2 = "Select * From " & fuTable & " Where Passed=" & HR_True & " And VA1=" & HR_Clng(fuYGDM) & " And scYear=" & DefYear
				Set rs2 = Server.CreateObject("ADODB.RecordSet")
					rs2.Open sql2, Conn, 1, 1
					iSum = 0 : iScore = 0 : iFormula = 0
					If Not(rs2.BOF And rs2.EOF) Then
						Do While Not rs2.EOF
							iFormula = GetRatioStutype(fuItemID, "", 0)		'无学生类别系数
							If tTemplate = "TempTableD" Then		'有级别无等级
								If Trim(rs2("VA7")) <> "" Then iFormula = GetRatioLevel(fuItemID, Trim(rs2("VA7")), "", 0)
							ElseIf tTemplate = "TempTableG" Then		'有级别无等级
								If Trim(rs2("VA6")) <> "" Then iFormula = GetRatioLevel(fuItemID, Trim(rs2("VA6")), "", 0)
							ElseIf tTemplate = "TempTableE" Then		'有级别及等级
								If Trim(rs2("VA7")) <> "" Then iFormula = GetRatioLevel(fuItemID, Trim(rs2("VA7")), Trim(rs2("VA8")), 0)
							ElseIf tTemplate = "TempTableF" Then		'有级别及等级
								If Trim(rs2("VA6")) <> "" Then iFormula = GetRatioLevel(fuItemID, Trim(rs2("VA6")), Trim(rs2("VA7")), 0)
							End If
							If iFormula <= 0 Then iFormula = 1
							iSum = iSum + HR_CDbl(rs2("VA3"))
							iScore = iScore + (HR_CDbl(rs2("VA3")) * iFormula)
							rs2.MoveNext
						Loop
					End If
					If iSum > fuMaxSum And fuMaxSum > 0 Then iSum = fuMaxSum		'如果有最高分限制
					fuField = "F" & fuItemID
					Conn.Execute("Update HR_KPI Set " & fuField & "=" & iScore & " Where YGDM=" & fuYGDM & " And scYear=" & DefYear)
					Conn.Execute("Update HR_KPI_SUM Set " & fuField & "=" & iSum & " Where YGDM=" & fuYGDM & " And scYear=" & DefYear)
					fuStr = True
				Set rs2 = Nothing
			Else
				fuArrStuType = Split(fuStrStuType, ",")
				For iFuStu = 0 To Ubound(fuArrStuType)
					fuField = "F" & fuItemID & "_" & GetStudentType(Trim(fuArrStuType(iFuStu)))		'取KPI字段名
					sql2 = "Select * From " & fuTable & " Where Passed=" & HR_True & " And VA1='" & Trim(fuYGDM) & "' And StudentType='" & Trim(fuArrStuType(iFuStu)) & "' And scYear=" & DefYear
					Set rs2 = Server.CreateObject("ADODB.RecordSet")
						rs2.Open sql2, Conn, 1, 1
						iSum = 0 : iScore = 0 : iFormula = 0
						If Not(rs2.BOF And rs2.EOF) Then
							If fuMaxSum > 0 Then			'当有上限值时（系数统一取值）
								iFormula = GetRatioStutype(fuItemID, Trim(fuArrStuType(iFuStu)), 0)		'有学生类别时取系数
								If HR_CDbl(iFormula) = 0 Then iFormula = 1		'当系数为0时，赋值为1
								Do While Not rs2.EOF
									iSum = iSum + HR_CDbl(rs2("VA3"))
									rs2.MoveNext
								Loop
								If iSum > fuMaxSum Then iSum = fuMaxSum
								iScore = iSum * iFormula		'累加后乘系数
							Else							'无上限时，按记录取系数
								Do While Not rs2.EOF
									iFormula = GetRatioStutype(fuItemID, Trim(fuArrStuType(iFuStu)), 0)		'有学生类别时取系数
									iSum = iSum + HR_CDbl(rs2("VA3"))
									iScore = iScore + (HR_CDbl(rs2("VA3")) * iFormula)
									rs2.MoveNext
								Loop
							End If
						End If

						Conn.Execute("Update HR_KPI Set " & fuField & "=" & iScore & " Where YGDM=" & fuYGDM & " And scYear=" & DefYear)
						Conn.Execute("Update HR_KPI_SUM Set " & fuField & "=" & iSum & " Where YGDM=" & fuYGDM & " And scYear=" & DefYear)
						fuStr = True
					Set rs2 = Nothing
				Next
			End If
		End If
	End If
	UpdateTeacherKPI = fuStr
End Function

Function UpdateTeacherTotalKPI(fuYGDM)		'总计员工业绩
	Dim rsFu, fTotal, fTotalSum, jFun
	fTotal = 0 : fTotalSum = 0
	If HR_Clng(fuYGDM) > 0 Then
		Set rsFu = Server.CreateObject("ADODB.RecordSet")
			rsFu.Open("Select * From HR_KPI Where YGDM=" & fuYGDM & " And scYear=" & DefYear), Conn, 1, 1
			If Not(rsFu.BOF And rsFu.EOF) Then
				For jFun = 12 To rsFu.Fields.Count - 1
					fTotal = fTotal + HR_CDbl(rsFu.Fields(jFun).value)
				Next
			End If
		Set rsFu = Nothing
		Set rsFu = Server.CreateObject("ADODB.RecordSet")
			rsFu.Open("Select * From HR_KPI_SUM Where YGDM=" & fuYGDM & " And scYear=" & DefYear), Conn, 1, 1
			If Not(rsFu.BOF And rsFu.EOF) Then
				For jFun = 12 To rsFu.Fields.Count - 1
					fTotalSum = fTotalSum + HR_CDbl(rsFu.Fields(jFun).value)
				Next
			End If
		Set rsFu = Nothing
		Conn.Execute("Update HR_KPI Set SumScore=" & fTotalSum & ", TotalScore=" & fTotal & " Where YGDM=" & fuYGDM & " And scYear=" & DefYear)
		Conn.Execute("Update HR_KPI_SUM Set SumScore=" & fTotalSum & ", TotalScore=" & fTotal & " Where YGDM=" & fuYGDM & " And scYear=" & DefYear)
	End If
End Function

Function UpdateKPIField()			'更新KPI表的列
	Dim rsFun, strFun, fStuType, iFun, arrField, tpField
	UpdateKPIField = False
	Set rsFun = Conn.Execute("Select * From HR_Class Order By ClassType,RootID,OrderID")			'取列
		If Not(rsFun.BOF And rsFun.EOF) Then
			Do While Not rsFun.EOF
				If rsFun("Child") = 0 Then	'有子类跳过
					If rsFun("StudentType") <> "" Then
						fStuType = Split(rsFun("StudentType"), ",")
						For iFun = 0 To Ubound(fStuType)
							strFun = strFun & "F" & rsFun("ClassID") & "_" & GetStudentType(fStuType(iFun)) & "||"
						Next
					Else
						strFun = strFun & "F" & rsFun("ClassID") & "||"
					End If
				End If
				rsFun.MoveNext
			Loop
		End If
	Set rsFun = Nothing
	If strFun <> "" Then
		strFun = FilterArrNull(strFun, "||")
		arrField = Split(strFun, "||")
		tpField = ""
		Set rsFun = Server.CreateObject("ADODB.RecordSet")		'取表中的字段
			rsFun.Open("Select * From HR_KPI"), Conn, 1, 1
			For iFun = 12 To rsFun.Fields.Count-1
				If iFun > 12 Then tpField = tpField & "@@"
				tpField = tpField & rsFun.Fields(iFun).name
			Next
		Set rsFun = Nothing

		For iFun = 0 To Ubound(arrField)
			If FoundInArr(tpField, arrField(iFun), "@@") = False Then
				Conn.Execute("alter table HR_KPI add " & arrField(iFun) & " Decimal(18,2) DEFAULT(0)")		'增加字段
				UpdateKPIField = True
			End If
		Next

		tpField = ""
		arrField = Split(strFun, "||")
		Set rsFun = Server.CreateObject("ADODB.RecordSet")		'取表中的字段
			rsFun.Open("Select * From HR_KPI_SUM"), Conn, 1, 1
			For iFun = 12 To rsFun.Fields.Count-1
				If iFun > 12 Then tpField = tpField & "@@"
				tpField = tpField & rsFun.Fields(iFun).name
			Next
		Set rsFun = Nothing
		For iFun = 0 To Ubound(arrField)
			If FoundInArr(tpField, arrField(iFun), "@@") = False Then
				Conn.Execute("alter table HR_KPI_SUM add " & arrField(iFun) & " Decimal(18,2) DEFAULT(0)")		'增加字段
				UpdateKPIField = True
			End If
		Next

	End If
End Function

'取指定学生类别系数值
Function GetRatioStutype(rClassID, rStuType, rType)
	Dim rFun, rArr, rArrStuType, rArrRatio, iR
	GetRatioStutype = 0
	Set rFun = Conn.Execute("Select * From HR_Class Where ClassID=" & HR_Clng(rClassID))
		If Not(rFun.BOF And rFun.EOF) Then
			If Trim(rFun("StudentType")) <> "" Then
				rArrStuType = Split(FilterArrNull(rFun("StudentType"), ","), ",")
				rArrRatio = Split(FilterArrNull(rFun("Ratio"), ","), ",")
				If Ubound(rArrStuType) <> Ubound(rArrRatio) Then Redim Preserve rArrRatio(Ubound(rArrStuType))
				For iR = 0 To Ubound(rArrRatio)
					If Trim(rStuType) = Trim(rArrStuType(iR)) Then GetRatioStutype = HR_CDbl(rArrRatio(iR))
				Next
			Else
				GetRatioStutype = HR_CDbl(rFun("Ratio"))
			End If
		End If
	Set rFun = Nothing
End Function

'取指定级别系数值，参数：项目ID, 级别名, 等级名, 查询条件
Function GetRatioLevel(rClassID, rLevel, rGrade, rType)
	Dim rFun, rArr, iR, tRatio, rsFor
	GetRatioLevel = 0
	Set rFun = Conn.Execute("Select * From HR_ItemModel Where ClassID=" & HR_Clng(rClassID) & " And FieldName='" & rLevel & "'")
		If Not(rFun.BOF And rFun.EOF) Then
			tRatio = HR_CDbl(rFun("Formula"))
			If Trim(rGrade) <> "" Then		'如果该级别下等级未设置，系统会以级别的分值计算
				Set rsFor = Conn.Execute("Select * From HR_ItemGrade Where ClassID=" & HR_Clng(rClassID) & " And LevelID=" & HR_Clng(rFun("ID")) & " And Grade='" & Trim(rGrade) & "'")
					If Not(rsFor.BOF And rsFor.EOF) Then
						tRatio = HR_CDbl(rsFor("Ratio"))
					End If
				Set rsFor = Nothing
			End If
		End If
	Set rFun = Nothing
	If HR_CDbl(tRatio) > 0 Then GetRatioLevel = tRatio
End Function

'重置员工汇总
Function ResetTeacherKPI(rYGDM, rItemID)
	Dim rsFun, rStr, rFun : rYGDM = HR_Clng(rYGDM)
	If rYGDM > 0 Then
		Set rsFun = Server.CreateObject("ADODB.RecordSet")		'取表中的字段
			rsFun.Open("Select * From HR_KPI_SUM"), Conn, 1, 1
			For rFun = 12 To rsFun.Fields.Count-1
				If rFun > 12 Then rStr = rStr & "@@"
				rStr = rStr & rsFun.Fields(rFun).name
			Next
		Set rsFun = Nothing
	End If
	ResetTeacherKPI = rStr
End Function
%>