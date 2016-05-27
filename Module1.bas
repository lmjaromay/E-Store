Attribute VB_Name = "Module1"
Public con As New ADODB.Connection
Public prevcon As New ADODB.Connection
Public dbcon As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rsd As New ADODB.Recordset
Public rst As New ADODB.Recordset
Public user1 As String
Public timein1 As String
Public PrevMonth As String
Public sMonth As String
Public sDate As String
Public sSet
Public sYesterday As String
Public sPrevDay As String
Public dbCreated As Boolean

Public iTotalSale As Double

Public Sub CreateTable()
    Dim MyADOCmd As New ADODB.Command

    Dim MySQL As String

    Set rs = New ADODB.Recordset
    Set MyADOCmd = New ADODB.Command
    Set MyADOCmd.ActiveConnection = con
    MyADOCmd.CommandText = "CREATE TABLE [" & sMonth & sDate & "]" _
        & " (ID int ,Primary Key(ID), ProdCode Text, ProdName Text," _
        & " ProdType Text, BegInv Single, Refill Int, TotalBegInv Int," _
        & " BegInvVal Single, EndInv Int ,EndInvVal Single, TotalSoldItem Int," _
        & " RetailPrice Single, TotalSale Single)"
    rs.Open MyADOCmd
    'rs.Close
    Set rs = Nothing
End Sub

Public Function IsExistingTable( _
      ByVal Database As String, _
      ByVal TableName As String _
   ) As Boolean

   Dim ConnectString As String
   Dim ADOXConnection As Object
   Dim ADODBConnection As Object
   Dim Table As Variant

   ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;data source=" & Database
   Set ADOXConnection = CreateObject("ADOX.Catalog")
   Set ADODBConnection = CreateObject("ADODB.Connection")
   ADODBConnection.Open ConnectString
   ADOXConnection.ActiveConnection = ADODBConnection
   For Each Table In ADOXConnection.Tables
      If LCase(Table.Name) = LCase(TableName) Then
         IsExistingTable = True
         Exit For
      End If
   Next
   ADODBConnection.Close

End Function



Public Sub dbCompute()
Set rst = New ADODB.Recordset
With rst
    .Open "select * from " & sMonth + sDate & "", con, 3, 3
    .MoveFirst
    
    Do While Not .EOF
    .Update
        .Fields("TotalBegInv") = .Fields("BegInv") + .Fields("Refill")
        .Fields("BegInvVal") = .Fields("TotalBegInv") * .Fields("RetailPrice")
        .Fields("EndInvVal") = .Fields("EndInv") * .Fields("RetailPrice")
        .Fields("TotalSoldItem") = .Fields("TotalBegInv") - .Fields("EndInv")
        .Fields("TotalSale") = .Fields("TotalSoldItem") * .Fields("RetailPrice")
        .Update
        
        .MoveNext
    Loop
    .Close
End With
Set rst = Nothing
End Sub

