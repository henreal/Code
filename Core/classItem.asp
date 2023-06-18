<%
Class ItemClass
	Public itemName
	Private itemTable
	
	Private Sub Class_Initialize()
		itemName = ""
		itemTable = ""
	End Sub

	Public Property Get itemID(id)
		id = HR_Clng(id)
		sql = "Select Top 1 * From HR_Class Where ClassID=1000"
		Set rs = Server.CreateObject("ADODB.RecordSet")
			rs.Open(sql), conn, 1, 1
			itemName = Trim(rs("ClassName"))
			Response.Write itemName &"<br>"
		Set rs = Nothing
	End Property
	
End Class
%>