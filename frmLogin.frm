VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5070
   ClientLeft      =   6180
   ClientTop       =   2850
   ClientWidth     =   8505
   ForeColor       =   &H80000014&
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "User Account Login"
      Height          =   3135
      Left            =   1920
      TabIndex        =   0
      Top             =   2040
      Width           =   4695
      Begin VB.CommandButton cmdLogin 
         Caption         =   "Login"
         Height          =   495
         Left            =   1440
         TabIndex        =   6
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   3120
         TabIndex        =   5
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtPword 
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
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtUname 
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
         Left            =   1800
         TabIndex        =   2
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblError 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Username or Password is Incorrect!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   -240
         TabIndex        =   9
         Top             =   0
         Width           =   4935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Username :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Nookers IC"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.1.4"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E - S T O R E"
      BeginProperty Font 
         Name            =   "Concorde"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   1560
      TabIndex        =   7
      Top             =   960
      Width           =   5535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404040&
      Height          =   1935
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   8775
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdLogin_Click()


Dim auth As String

Set rs = New ADODB.Recordset
rs.Open "select * from Login where Uname = '" & txtUname & "' and Pword = '" & txtPword & "'", dbcon, 3, 3
If rs.EOF = True Then
    lblError.Visible = True
    Label5.BackStyle = 1
    txtUname.SetFocus
    txtUname.Text = ""
    txtPword.Text = ""
Else
    user1 = txtUname.Text
    timein1 = Time
    If MsgBox("Login Successful! Welcome " + user1 + "!", vbInformation + vbOKOnly, "Logged In...") = vbOK Then
        frmLogin.Hide
        Set rsd = New ADODB.Recordset
        rsd.Open "select * from LogHis", dbcon, 3, 3
        auth = rs.Fields("AuthLvl")
        With rsd
            .AddNew
            .Fields("Uname") = user1
            .Fields("TransDate") = Date
            .Fields("Timein") = timein1
            .Fields("AuthLvl") = auth
            .Update
        End With
        rsd.Close
        Set rsd = Nothing
        Set rsd = New ADODB.Recordset
        rsd.Open "select * from Login where Uname = '" & txtUname & "'", dbcon, 3, 3
        If rsd.EOF = False Then
            If rsd.Fields("AuthLvl") = "Administrator" Then
                frmGate.Show
            ElseIf rsd.Fields("AuthLvl") = "Cashier" Then
                frmCashier.Show
            End If
            Me.Hide
        End If
        'rsd.Close
        Set rsd = Nothing
        'rs.Close
        lblError.Visible = False
        Label5.BackStyle = Transparent
    End If
    
End If

Set rs = Nothing
txtUname.Text = ""
txtPword.Text = ""

End Sub

Private Sub Form_Load()
Set Connect = New Class1


End Sub


Private Sub txtPword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdLogin_Click
End If
End Sub


