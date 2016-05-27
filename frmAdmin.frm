VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmAdmin 
   BackColor       =   &H00404040&
   Caption         =   "E-Store : Administrator"
   ClientHeight    =   7425
   ClientLeft      =   1410
   ClientTop       =   2190
   ClientWidth     =   17595
   Icon            =   "frmAdmin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   17595
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10440
      TabIndex        =   9
      Top             =   360
      Width           =   2055
   End
   Begin VB.ComboBox cboSearchBy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmAdmin.frx":08CA
      Left            =   7560
      List            =   "frmAdmin.frx":08DD
      TabIndex        =   8
      Text            =   "Product Name"
      Top             =   360
      Width           =   2655
   End
   Begin VB.ComboBox cboSort 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmAdmin.frx":0920
      Left            =   15120
      List            =   "frmAdmin.frx":0933
      TabIndex        =   7
      Text            =   "Product Name"
      Top             =   360
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   " Product Information "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   6495
      Left            =   240
      TabIndex        =   14
      Top             =   720
      Width           =   6975
      Begin VB.TextBox txtPtype 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1920
         TabIndex        =   2
         Top             =   2160
         Width           =   3735
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4440
         TabIndex        =   22
         Top             =   5040
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4440
         TabIndex        =   21
         Top             =   4200
         Width           =   1815
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2640
         TabIndex        =   20
         Top             =   5040
         Width           =   1815
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2640
         TabIndex        =   19
         Top             =   4200
         Width           =   1815
      End
      Begin VB.TextBox txtPcode 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1920
         TabIndex        =   0
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox txtPname 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1920
         TabIndex        =   1
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox txtRemStock 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1320
         TabIndex        =   3
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox txtRetailPrice 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         TabIndex        =   4
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   840
         TabIndex        =   5
         Top             =   5040
         Width           =   1815
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   840
         TabIndex        =   10
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Product Type :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Remaining Stock:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Retail Price :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3480
         TabIndex        =   15
         Top             =   2880
         Width           =   1575
      End
   End
   Begin VB.TextBox txtpw 
      Height          =   285
      Left            =   15600
      TabIndex        =   13
      Text            =   "Text13"
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtun 
      Height          =   285
      Left            =   15600
      TabIndex        =   12
      Text            =   "Text13"
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6375
      Left            =   7560
      TabIndex        =   6
      Top             =   840
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11245
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Product Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Product Name"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Product Type"
         Object.Width           =   3352
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Retail Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Stock"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   15960
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Search :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   10560
      TabIndex        =   28
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Search By :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7560
      TabIndex        =   27
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sort By :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   15120
      TabIndex        =   26
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblUser 
      Height          =   375
      Left            =   15720
      TabIndex        =   11
      Top             =   -120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblBtn 
      Caption         =   "Label5"
      Height          =   255
      Left            =   15840
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogout 
         Caption         =   "Logout"
      End
   End
   Begin VB.Menu mnuManage 
      Caption         =   "Manage"
      Begin VB.Menu mnuTransactions 
         Caption         =   "Transactions"
      End
      Begin VB.Menu mnuUser 
         Caption         =   "User Accounts"
      End
      Begin VB.Menu mnuProducts 
         Caption         =   "Products"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboSort_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ListView1.SetFocus
End If
End Sub

Private Sub cboSort_LostFocus()
ListView1.ListItems.Clear
Set rs = New ADODB.Recordset
If cboSort.Text = "Product Code" Then
    With rs
    .Open "select * from " & sMonth + sDate & " order by ProdCode", con, 3, 3
    Do While Not .EOF
        ListView1.ListItems.Add , , !ID & ""
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "" & !ProdCode
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = "" & !ProdName
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = "" & !ProdType
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !RetailPrice
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !BegInv
        
        .MoveNext
    Loop
    .Close
    End With
ElseIf cboSort.Text = "Product Name" Then
    With rs
    .Open "select * from " & sMonth + sDate & " order by ProdName", con, 3, 3
    Do While Not .EOF
        ListView1.ListItems.Add , , !ID & ""
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "" & !ProdCode
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = "" & !ProdName
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = "" & !ProdType
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !RetailPrice
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !BegInv
        
        .MoveNext
    Loop
    .Close
    End With
ElseIf cboSort.Text = "Product Type" Then
    With rs
    .Open "select * from " & sMonth + sDate & " order by ProdType", con, 3, 3
    Do While Not .EOF
        ListView1.ListItems.Add , , !ID & ""
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "" & !ProdCode
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = "" & !ProdName
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = "" & !ProdType
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !RetailPrice
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !BegInv
        
        .MoveNext
    Loop
    .Close
    End With
ElseIf cboSort.Text = "Retail Price" Then
    With rs
    .Open "select * from " & sMonth + sDate & " order by RetailPrice", con, 3, 3
    Do While Not .EOF
        ListView1.ListItems.Add , , !ID & ""
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "" & !ProdCode
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = "" & !ProdName
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = "" & !ProdType
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !RetailPrice
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !BegInv
        
        .MoveNext
    Loop
    .Close
    End With
ElseIf cboSort.Text = "Stock" Then
    With rs
    .Open "select * from " & sMonth + sDate & " order by BegInv", con, 3, 3
    Do While Not .EOF
        ListView1.ListItems.Add , , !ID & ""
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "" & !ProdCode
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = "" & !ProdName
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = "" & !ProdType
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !RetailPrice
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !BegInv
        
        .MoveNext
    Loop
    .Close
    End With
End If
Set rs = Nothing

'Call dbase
End Sub


Private Sub cmdAdd_Click()
txtPcode.Text = ""
txtPname.Text = ""
txtPtype.Text = ""
txtRemStock.Text = ""
txtRetailPrice.Text = ""


ListView1.Enabled = False
txtPcode.Enabled = True
txtPname.Enabled = True
txtPtype.Enabled = True
txtRemStock.Enabled = True
txtRetailPrice.Enabled = True
cmdCancel.Enabled = True
cmdSave.Enabled = True
cmdEdit.Enabled = False
cmdDelete.Enabled = False
cmdClear.Enabled = False
cmdAdd.Enabled = False
lblBtn.Caption = cmdAdd.Caption
txtPcode.SetFocus
txtSearch.Enabled = False
cboSearchBy.Enabled = False
cboSort.Enabled = False
End Sub

Private Sub cmdCancel_Click()
txtPcode.Enabled = False
txtPname.Enabled = False
txtPtype.Enabled = False
txtRemStock.Enabled = False
txtRetailPrice.Enabled = False
cmdCancel.Enabled = False
cmdSave.Enabled = False
cmdEdit.Enabled = True
cmdDelete.Enabled = True
cmdClear.Enabled = True
cmdAdd.Enabled = True
ListView1.Enabled = True
Call ListSelect


End Sub

Private Sub cmdClear_Click()
Set rs = New ADODB.Recordset
Set rsd = New ADODB.Recordset
If MsgBox("This will Delete All Data in the Database, Are you sure you want to Delete All?", vbCritical + vbYesNo, " E-Store Admin Panel") = vbYes Then

With rs
    .Open "select * from " & sMonth + sDate & "", con, 3, 3
    Do While Not .EOF
        .Delete
        .MoveNext
    Loop
    .Close
End With
 rsd.Open "select * from " & sMonth + sDate & "", con, 3, 3

    With rsd
        .AddNew
        .Fields("ProdCode") = "0"
        .Fields("ProdName") = "0"
        .Fields("ProdType") = "0"
        .Fields("BegInv") = "0"
        .Fields("RetailPrice") = "0"
        .Update
        .Close
    End With
Else
    ListView1.SetFocus
    Call ListSelect
End If
Set rsd = Nothing
Set rs = Nothing
Call dbase
cmdEdit.Enabled = False
End Sub

Private Sub cmdDelete_Click()
Set rsd = New ADODB.Recordset
rsd.Open "select * from " & sMonth + sDate & " where ID = " & ListView1.SelectedItem, con, 3, 3
If MsgBox("Are you sure you want to delete this data?", vbQuestion + vbYesNo, "E-Store Admin Panel") = vbYes Then
With rsd
    .Delete
End With
Else
ListView1.SetFocus
End If
Call dbase
rsd.Close
Set rsd = Nothing
cmdEdit.Enabled = False
End Sub

Private Sub cmdEdit_Click()
txtPcode.Enabled = True
txtPname.Enabled = True
txtPtype.Enabled = True
txtRemStock.Enabled = True
txtRetailPrice.Enabled = True
cmdCancel.Enabled = True
cmdSave.Enabled = True
cmdEdit.Enabled = False
cmdDelete.Enabled = False
cmdClear.Enabled = False
cmdAdd.Enabled = False
lblBtn.Caption = cmdEdit.Caption
txtPcode.SetFocus
ListView1.Enabled = False
txtSearch.Enabled = False
cboSearchBy.Enabled = False
cboSort.Enabled = False
End Sub

Private Sub mnuLogout_Click()
If MsgBox("Do you want to switch to Another User?", vbQuestion + vbYesNoCancel, "Logging Out...") = vbYes Then
Set rs = New ADODB.Recordset
rs.Open "select * from LogHis where Uname = '" & txtun & "' And Timein = '" & txtpw & "'", dbcon, 3, 3
If rs.EOF = False Then
    With rs
        .Fields("Timeout") = Time
        .Update
    End With
Else
    MsgBox ("Error")
End If
rs.Close
Set rs = Nothing
Unload Me
frmLogin.Show
ElseIf vbNo Then
Set rs = New ADODB.Recordset
rs.Open "select * from LogHis where Uname = '" & txtun & "' And Timein = '" & txtpw & "'", dbcon, 3, 3
If rs.EOF = False Then
    With rs
        .Fields("Timeout") = Time
        .Update
    End With
Else
    MsgBox ("Error")
End If
rs.Close
Set rs = Nothing
End
Else
ListView1.SetFocus
End If

End Sub

Private Sub cmdSave_Click()
Set rsd = New ADODB.Recordset
If txtPcode.Text = "" Or txtPname.Text = "" Or txtPtype.Text = "" Or txtRemStock.Text = "" Or txtRetailPrice.Text = "" Then
    If MsgBox("Please Fill-Up all empty fields!", vbInformation + vbOKOnly, "Save Failed") = vbOK Then: txtPcode.SetFocus
Else
If lblBtn.Caption = "Edit" Then
    y = ListView1.SelectedItem
    rsd.Open "select * from " & sMonth + sDate & " where ID = " & y, con, 3, 3
        With rsd
            .Update
            .Fields("ProdCode") = txtPcode
            .Fields("ProdName") = txtPname
            .Fields("ProdType") = txtPtype
            .Fields("BegInv") = txtRemStock
            .Fields("RetailPrice") = txtRetailPrice
            .Update
            .Close
        End With
ElseIf lblBtn.Caption = "Add" Then
      
    rsd.Open "select * from " & sMonth + sDate & "", con, 3, 3
    With rsd
        .AddNew
        .Fields("ID") = Int((Rnd * 1000) + 500)
        .Fields("ProdCode") = txtPcode
        .Fields("ProdName") = txtPname
        .Fields("ProdType") = txtPtype
        .Fields("BegInv") = txtRemStock
        .Fields("RetailPrice") = txtRetailPrice
        .Update
        .Close
    End With

        
End If

txtPcode.Enabled = False
txtPname.Enabled = False
txtPtype.Enabled = False
txtRemStock.Enabled = False
txtRetailPrice.Enabled = False
cmdCancel.Enabled = False
cmdSave.Enabled = False
cmdEdit.Enabled = False
cmdDelete.Enabled = True
cmdClear.Enabled = True
cmdAdd.Enabled = True
ListView1.Enabled = True
txtSearch.Enabled = True
cboSearchBy.Enabled = True
cboSort.Enabled = True

Call dbase
End If
Set rsd = Nothing
ListView1.SetFocus
End Sub

Private Sub Form_Load()
Set Connect = New Class1
txtun.Text = user1
txtpw.Text = timein1
Call dbase

End Sub

Private Sub dbase()
ListView1.ListItems.Clear
Set rs = New ADODB.Recordset
With rs
    .Open "select * from " & sMonth + sDate & "", con, 3, 3
    Do While Not .EOF
        ListView1.ListItems.Add , , !ID & ""
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "" & !ProdCode
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = "" & !ProdName
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = "" & !ProdType
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !RetailPrice
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !BegInv
        
        .MoveNext
    Loop
    .Close
End With
Set rs = Nothing
End Sub

Private Sub ListSelect()

txtID.Text = ListView1.SelectedItem
txtPcode.Text = ListView1.SelectedItem.SubItems(1)
txtPname.Text = ListView1.SelectedItem.SubItems(2)
txtPtype.Text = ListView1.SelectedItem.SubItems(3)
txtRetailPrice.Text = ListView1.SelectedItem.SubItems(4)
txtRemStock.Text = ListView1.SelectedItem.SubItems(5)
End Sub

Private Sub ListDeselect()

txtID.Text = ""
txtPcode.Text = ""
txtPname.Text = ""
txtPtype.Text = ""
txtRetailPrice.Text = ""
txtRemStock.Text = ""
End Sub

Private Sub ListView1_Click()
Call ListSelect
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
cmdEdit.Enabled = True
cmdDelete.Enabled = True
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 25 Then
   Call ListSelect
End If
End Sub

Private Sub cboSort_GotFocus()
cmdEdit.Enabled = False
cmdDelete.Enabled = False
Call ListDeselect
End Sub
Private Sub cboSearchby_GotFocus()
cmdEdit.Enabled = False
cmdDelete.Enabled = False
Call ListDeselect
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
Call ListSelect
End Sub

Private Sub mnuTransactions_Click()
frmTransaction.Show
Unload Me
End Sub

Private Sub txtSearch_GotFocus()
cmdEdit.Enabled = False
cmdDelete.Enabled = False
Call ListDeselect
End Sub
Private Sub txtSearch_Change()

Set rs = New ADODB.Recordset
Select Case cboSearchBy.Text
Case "Product Code"
ListView1.ListItems.Clear
With rs
    .Open "select * from " & sMonth + sDate & " where ProdCode like '" & txtSearch & "%' order by " & sMonth + sDate & ".ProdCode asc;", con, 3, 3
    Do While Not .EOF
        ListView1.ListItems.Add , , !ID & ""
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "" & !ProdCode
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = "" & !ProdName
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = "" & !ProdType
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !RetailPrice
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !BegInv
        
        .MoveNext
    Loop
    .Close
End With
Case "Product Name"
ListView1.ListItems.Clear
With rs
    .Open "select * from " & sMonth + sDate & " where ProdName like '" & txtSearch & "%' order by " & sMonth + sDate & ".ProdName asc;", con, 3, 3
    Do While Not .EOF
        ListView1.ListItems.Add , , !ID & ""
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "" & !ProdCode
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = "" & !ProdName
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = "" & !ProdType
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !RetailPrice
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !BegInv
        
        .MoveNext
    Loop
    .Close
End With
Case "Product Type"
ListView1.ListItems.Clear
With rs
    .Open "select * from " & sMonth + sDate & " where ProdType like '" & txtSearch & "%' order by " & sMonth + sDate & ".ProdType asc;", con, 3, 3
    Do While Not .EOF
        ListView1.ListItems.Add , , !ID & ""
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "" & !ProdCode
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = "" & !ProdName
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = "" & !ProdType
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !RetailPrice
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !BegInv
        
        .MoveNext
    Loop
    .Close
End With
Case "Retail Price"
ListView1.ListItems.Clear
With rs
    .Open "select * from " & sMonth + sDate & " where RetailPrice like '" & txtSearch & "%' order by " & sMonth + sDate & ".RetailPrice asc;", con, 3, 3
    Do While Not .EOF
        ListView1.ListItems.Add , , !ID & ""
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "" & !ProdCode
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = "" & !ProdName
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = "" & !ProdType
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !RetailPrice
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !BegInv
        
        .MoveNext
    Loop
    .Close
End With
Case "Stock"
ListView1.ListItems.Clear
With rs
    .Open "select * from " & sMonth + sDate & " where BegInv like '" & txtSearch & "%' order by " & sMonth + sDate & ".BegInv asc;", con, 3, 3
    Do While Not .EOF
        ListView1.ListItems.Add , , !ID & ""
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "" & !ProdCode
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = "" & !ProdName
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = "" & !ProdType
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !RetailPrice
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !BegInv
        
        .MoveNext
    Loop
    .Close
End With
End Select
Set rs = Nothing

End Sub
