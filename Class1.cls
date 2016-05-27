VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub Class_Initialize()
DoEvents

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                               CONNECTION For Login and Loghis (storedb.mdb)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set dbcon = New ADODB.Connection
With dbcon
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\storedb.mdb;Persist Security Info=False;Jet OLEDB:Database Password=atienza01"
    .CursorLocation = adUseClient
    .Open
End With

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                               CONNECTION For Store Inventory (curMonthdb and prevMonthdb)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set con = New ADODB.Connection

    CurMonth = Format$(Date, "MMMM")
    sMonth = CurMonth
    sDate = Format$(Date, "dd")


If Dir$("" & App.Path & "\" & CurMonth & ".mdb") <> "" Then
    
    ' To Check if Table is Existing or Not''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim result As Boolean
         result = IsExistingTable("" & App.Path & "\" & CurMonth & ".mdb;Persist Security Info=False;Jet OLEDB:Database Password=atienza01", "" & sMonth & sDate & "")
    If result Then
        With con
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & CurMonth & ".mdb;Persist Security Info=False;Jet OLEDB:Database Password=atienza01"
        .CursorLocation = adUseClient
        .Open
        End With
    Else
        With con
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & CurMonth & ".mdb;Persist Security Info=False;Jet OLEDB:Database Password=atienza01"
        .CursorLocation = adUseClient
        .Open
        End With
        Call CreateTable
        
        Set rs = New ADODB.Recordset
        Set rsd = New ADODB.Recordset
        
        sYesterday = Val(sDate) - 1
        
        ' Dim result2 As Boolean
        'result2 = IsExistingTable("" & App.Path & "\" & CurMonth & ".mdb", "" & sMonth & sYesterday & "")
        ' If result2 Then
        '    MsgBox "Existing"
        ' Else
        '    MsgBox "Missing" & sMonth & sYesterday
            
        ' End If
        rs.Open "select * from " & sMonth & sYesterday, con, 3, 3
        rsd.Open "select * from " & sMonth & sDate, con, 3, 3
        rs.MoveFirst
        
        Do While Not rs.EOF

        
        rsd.AddNew
        rsd.Fields("ID") = rs.Fields("ID")
        rsd.Fields("ProdCode") = rs.Fields("ProdCode")
        rsd.Fields("ProdName") = rs.Fields("ProdName")
        rsd.Fields("ProdType") = rs.Fields("ProdType")
        rsd.Fields("BegInv") = rs.Fields("EndInv")
        rsd.Fields("Refill") = 0
        rsd.Fields("TotalBegInv") = rs.Fields("EndInv")
        rsd.Fields("EndInv") = rs.Fields("EndInv")
        rsd.Fields("RetailPrice") = rs.Fields("RetailPrice")
        rsd.Fields("TotalSoldItem") = 0
        rsd.Fields("TotalSale") = 0
        rsd.Update
       
        rs.MoveNext
        
        Loop
        
        Call dbCompute
        
        Set rs = Nothing
        Set rsd = Nothing
        
    End If
    
    
    
    
Else
    
    ' Create database named after current Month''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set db = DBEngine.Workspaces(0).CreateDatabase("" & App.Path & "\" & CurMonth & ".mdb", dbLangGeneral & ";pwd=atienza01")
    Set db = Nothing
    
    With con
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & CurMonth & ".mdb;Persist Security Info=False;Jet OLEDB:Database Password=atienza01"
    .CursorLocation = adUseClient
    .Open
    End With
    
    'After Creating the database, we need to create a table'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call CreateTable

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '                               CONNECTING to PrevMonthDB
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set prevcon = New ADODB.Connection
    
    'Determines the Previous Month'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case CurMonth
        Case "January"
             PrevMonth = "December"
        Case "February"
             PrevMonth = "January"
        Case "March"
             PrevMonth = "February"
        Case "April"
             PrevMonth = "March"
        Case "May"
             PrevMonth = "April"
        Case "June"
             PrevMonth = "May"
        Case "July"
             PrevMonth = "June"
        Case "August"
             PrevMonth = "July"
        Case "September"
             PrevMonth = "August"
        Case "October"
             PrevMonth = "September"
        Case "November"
             PrevMonth = "October"
        Case "December"
             PrevMonth = "November"
    End Select
    
    'Check if PrevMonth Exists''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Dir$("" & App.Path & "\" & PrevMonth & ".mdb") <> "" Then
        With prevcon
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & PrevMonth & ".mdb;Persist Security Info=False;Jet OLEDB:Database Password=atienza01"
        .CursorLocation = adUseClient
        .Open
        End With
        MsgBox "Existing"
        
    Else
        MsgBox ("" & PrevMonth & "")
    End If

    Select Case PrevMonth
        Case "January"
             sPrevDay = "31"
        Case "February"
             sPrevDay = "28"
        Case "March"
             sPrevDay = "31"
        Case "April"
             sPrevDay = "30"
        Case "May"
             sPrevDay = "31"
        Case "June"
             sPrevDay = "30"
        Case "July"
             sPrevDay = "31"
        Case "August"
             sPrevDay = "31"
        Case "September"
             sPrevDay = "30"
        Case "October"
             sPrevDay = "31"
        Case "November"
             sPrevDay = "30"
        Case "December"
             sPrevDay = "31"
    End Select
    
    Set rs = New ADODB.Recordset
    Set rsd = New ADODB.Recordset
        PrevMonth = "April"
        sPrevDay = "30"
        rs.Open "select * from " & PrevMonth & sPrevDay, prevcon, 3, 3
        rsd.Open "select * from " & sMonth & sDate, con, 3, 3
        rs.MoveFirst
        
        Do While Not rs.EOF

        
        rsd.AddNew
        rsd.Fields("ID") = rs.Fields("ID")
        rsd.Fields("ProdCode") = rs.Fields("ProdCode")
        rsd.Fields("ProdName") = rs.Fields("ProdName")
        rsd.Fields("ProdType") = rs.Fields("ProdType")
        rsd.Fields("BegInv") = rs.Fields("EndInv")
        rsd.Fields("Refill") = 0
        rsd.Fields("TotalBegInv") = rs.Fields("EndInv")
        rsd.Fields("EndInv") = rs.Fields("EndInv")
        rsd.Fields("TotalSoldItem") = 0
        rsd.Fields("RetailPrice") = rs.Fields("RetailPrice")
        rsd.Fields("TotalSale") = 0
        rsd.Update
       
        rs.MoveNext
        
        Loop
        
        'Call dbCompute
        
        Set rs = Nothing
        Set rsd = Nothing
    
    dbCreated = True
End If
End Sub

