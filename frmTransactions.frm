VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmTransaction 
   BackColor       =   &H00404040&
   Caption         =   "Form1"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   13875
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   720
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FFFFFF&
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
      Left            =   8160
      TabIndex        =   14
      Top             =   120
      Width           =   2175
   End
   Begin VB.ComboBox cboSearchBy 
      BackColor       =   &H00FFFFFF&
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
      ItemData        =   "frmTransactions.frx":0000
      Left            =   5760
      List            =   "frmTransactions.frx":0016
      TabIndex        =   13
      Text            =   "Transaction Date"
      Top             =   120
      Width           =   2295
   End
   Begin VB.ComboBox cboSort 
      BackColor       =   &H00FFFFFF&
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
      ItemData        =   "frmTransactions.frx":0064
      Left            =   11400
      List            =   "frmTransactions.frx":007A
      TabIndex        =   12
      Text            =   "Transaction Date"
      Top             =   120
      Width           =   2295
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6615
      Left            =   5760
      TabIndex        =   1
      Top             =   600
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   11668
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cashier"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Total Sales"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Actual Cash"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Credit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ShortExcess"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   5175
      Begin VB.TextBox txtES 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   17
         Top             =   4920
         Width           =   2895
      End
      Begin VB.TextBox txtTime 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   15
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   1920
         TabIndex        =   6
         Top             =   4200
         Width           =   2895
      End
      Begin VB.TextBox txtActualCash 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   5
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox txtTotalSale 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   4
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox txtCashier 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   3
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   2
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cashier :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Time :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Short/Excess :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Credit :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Cash :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Sale :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Set Connect = New Class1
Call dbase
Call ListSelect

End Sub

Private Sub dbase()
ListView1.ListItems.Clear
Set rst = New ADODB.Recordset
With rst
    .Open "Select * from Trans", dbcon, 3, 3
    Do While Not .EOF
        ListView1.ListItems.Add , , !ID & ""
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "" & !TransDate
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = "" & !TransDate
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = "" & !Cashier
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !TotalSale
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !ActualCash
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(6) = "" & !Credit
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(7) = "" & !ExcessOrShort
        .MoveNext
    Loop
    .Close
End With
Set rst = Nothing
End Sub

Private Sub ListSelect()

txtID.Text = ListView1.SelectedItem
txtDate.Text = ListView1.SelectedItem.SubItems(1)
txtTime.Text = ListView1.SelectedItem.SubItems(2)
txtCashier.Text = ListView1.SelectedItem.SubItems(3)
txtTotalSale.Text = ListView1.SelectedItem.SubItems(4)
txtActualCash.Text = ListView1.SelectedItem.SubItems(5)
txtCredit.Text = ListView1.SelectedItem.SubItems(6)
txtES.Text = ListView1.SelectedItem.SubItems(7)


End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 38 Then
Call ListSelect
End If
End Sub


