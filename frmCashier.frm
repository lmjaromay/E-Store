VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmCashier 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   Caption         =   "E-Store : Cashier"
   ClientHeight    =   10470
   ClientLeft      =   120
   ClientTop       =   480
   ClientWidth     =   20160
   ForeColor       =   &H00808080&
   Icon            =   "frmCashier.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10470
   ScaleWidth      =   20160
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H00404040&
      Height          =   2655
      Left            =   12480
      TabIndex        =   46
      Top             =   0
      Width           =   7575
      Begin VB.TextBox txtActualCash 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         TabIndex        =   56
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         TabIndex        =   54
         Top             =   1920
         Width           =   2775
      End
      Begin VB.OptionButton optSecSet 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         Caption         =   "Second Set"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6000
         TabIndex        =   51
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton optFirstSet 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         Caption         =   "First Set"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6000
         TabIndex        =   50
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "Submit"
         Height          =   615
         Left            =   5760
         TabIndex        =   49
         Top             =   1920
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "frmCashier.frx":08CA
         Left            =   5640
         List            =   "frmCashier.frx":08F8
         TabIndex        =   48
         Text            =   "All"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblTotalSale 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         TabIndex        =   55
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Credit :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   -1680
         TabIndex        =   53
         Top             =   2040
         Width           =   3735
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Cash :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   -720
         TabIndex        =   52
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Total Sale :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   -600
         TabIndex        =   47
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      Height          =   2655
      Left            =   7800
      TabIndex        =   40
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton cmdRefill 
         Caption         =   "Refill"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   43
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtStockRefill 
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
         Left            =   840
         TabIndex        =   42
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Critical"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   840
         TabIndex        =   45
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Status :"
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
         Left            =   1800
         TabIndex        =   44
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Refill :"
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
         Left            =   240
         TabIndex        =   41
         Top             =   600
         Width           =   975
      End
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
      ItemData        =   "frmCashier.frx":0981
      Left            =   17760
      List            =   "frmCashier.frx":0994
      TabIndex        =   36
      Text            =   "Product Name"
      Top             =   3120
      Width           =   2295
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
      ItemData        =   "frmCashier.frx":09D7
      Left            =   7800
      List            =   "frmCashier.frx":09EA
      TabIndex        =   35
      Text            =   "Product Name"
      Top             =   3120
      Width           =   2655
   End
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
      Height          =   420
      Left            =   10680
      TabIndex        =   34
      Top             =   3120
      Width           =   2535
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6255
      Left            =   7800
      TabIndex        =   0
      Top             =   3720
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   11033
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
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Product Code"
         Object.Width           =   1236
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Product Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Product Type"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Beginning Inventory"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Refill"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Total Beginning Inventory"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "BIV"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Ending Inventory"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "EIV"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "Total Sold Item"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Retail Price"
         Object.Width           =   1765
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Total Sale"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   " Product Information "
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
      Height          =   6375
      Left            =   360
      TabIndex        =   9
      Top             =   3600
      Width           =   7215
      Begin VB.TextBox txtEndInvVal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5280
         TabIndex        =   26
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox txtTotalSale 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   720
         TabIndex        =   24
         Top             =   5160
         Width           =   6015
      End
      Begin VB.TextBox txtRetailPrice 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         TabIndex        =   22
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox txtSoldItems 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5280
         TabIndex        =   20
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox txtEndInv 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         TabIndex        =   18
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox txtBegInvVal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5280
         TabIndex        =   16
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtTotalBegInv 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5280
         TabIndex        =   14
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtRefill 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         TabIndex        =   12
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtBegInv 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Ending Inventory Value :"
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
         Left            =   3600
         TabIndex        =   27
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Item Sold :"
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
         Left            =   720
         TabIndex        =   25
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total No. Of Sold Items :"
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
         Left            =   3600
         TabIndex        =   23
         Top             =   3000
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
         Left            =   360
         TabIndex        =   21
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ending Inventory :"
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
         Left            =   360
         TabIndex        =   19
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Beginning Inventory Value :"
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
         Height          =   495
         Left            =   3600
         TabIndex        =   17
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Beg. Inv. :"
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
         Left            =   3600
         TabIndex        =   15
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Refill :"
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
         Left            =   360
         TabIndex        =   13
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Beginning Inventory :"
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
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   " Store Inventory Update "
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
      Height          =   3495
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   7215
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
         Height          =   495
         Left            =   1920
         TabIndex        =   32
         Top             =   1800
         Width           =   4575
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "RECORD"
         Height          =   615
         Left            =   4080
         TabIndex        =   8
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox txtRemStock 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1920
         TabIndex        =   6
         Top             =   2520
         Width           =   1815
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
         Height          =   495
         Left            =   1920
         TabIndex        =   4
         Top             =   1080
         Width           =   4575
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
         Height          =   495
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label13 
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
         TabIndex        =   33
         Top             =   1920
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
         Left            =   240
         TabIndex        =   7
         Top             =   2640
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
         TabIndex        =   5
         Top             =   1200
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
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.TextBox txtun 
      Height          =   285
      Left            =   7800
      TabIndex        =   28
      Text            =   "Text13"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtpw 
      Height          =   285
      Left            =   7800
      TabIndex        =   29
      Text            =   "Text13"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   30
      Top             =   10095
      Width           =   20160
      _ExtentX        =   35560
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Enabled         =   0   'False
            Text            =   "Cashier :"
            TextSave        =   "Cashier :"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Cashier Account"
            TextSave        =   "Cashier Account"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Enabled         =   0   'False
            Text            =   "Date :"
            TextSave        =   "Date :"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            TextSave        =   "26/05/2016"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   8520
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label16 
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
      Left            =   17760
      TabIndex        =   39
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label15 
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
      Left            =   7800
      TabIndex        =   38
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label14 
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
      Left            =   10680
      TabIndex        =   37
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuSysOver 
         Caption         =   "System Override"
      End
      Begin VB.Menu mnuDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogout 
         Caption         =   "Logout"
      End
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "Transaction"
      Begin VB.Menu mnuCredit 
         Caption         =   "Credit"
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "Load"
      End
   End
End
Attribute VB_Name = "frmCashier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdRecord_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdRefill_Click
End If
End Sub

Private Sub cmdSubmit_Click()
    Dim iSet As Integer
    iSet = 1
    If optFirstSet.Value = True Then
        Set rst = New ADODB.Recordset
        With rst
            .Open "select * from Trans", dbcon, 3, 3
            .AddNew
            .Fields("TransDate") = Date
            .Fields("TransTime") = Time
            .Fields("Cashier") = user1
            .Fields("Set") = iSet
            .Fields("TotalSale") = "" & iTotalSale
            .Fields("ActualCash") = "" & txtActualCash.Text
            .Fields("Credit") = "" & txtCredit.Text
            .Fields("ExcessOrShort") = iTotalSale - Val(txtActualCash.Text)
            .Update
            .Close
        End With
        Set rst = Nothing

    ElseIf optSecSet.Value = True Then
        iSet = 2
        Set rst = New ADODB.Recordset
        With rst
            .Open "select * from Trans where TransDate = " & sMonth & sDate & " And Set = 1", dbcon, 3, 3
            iTotalSale = iTotalSale - .Fields("TotalSale")
            .Close
        End With
        Set rst = Nothing
        
        Set rst = New ADODB.Recordset
        With rst
            .Open "select * from Trans", dbcon, 3, 3
            .AddNew
            .Fields("TransDate") = Date
            .Fields("TransTime") = Time
            .Fields("Cashier") = user1
            .Fields("Set") = iSet
            .Fields("TotalSale") = "" & iTotalSale
            .Fields("ActualCash") = "" & txtActualCash.Text
            .Fields("Credit") = "" & txtCredit.Text
            .Fields("ExcessOrShort") = iTotalSale - Val(txtActualCash.Text)
            .Update
            .Close
        End With
        Set rst = Nothing
    Else
        If MsgBox("Please Select the Current Set", vbInformation, "E-Store : Submitting...") = vbOK Then: frmCashier.SetFocus
        Exit Sub
    End If

Call dbTotal
txtActualCash.Text = ""
txtCredit.Text = ""
MsgBox "Set " & iSet & " Total Sale Submitted", vbInformation, "E-Store : Submitting..."
End Sub

Private Sub cmdSysOver_Click()
Dim sauthlvl As String
sauthlvl = "Administrator"
Set rst = New ADODB.Recordset

With rst
    .Open "select * from Login where Pword = '" & txtAdminPass & "' And AuthLvl = '" & sauthlvl & "'", dbcon, 3, 3
    If .EOF = False Then
        MsgBox "System Override Granted!", vbInformation, "E-Store : Overriding..."
        fmeSysOver.Visible = False
    Else
        MsgBox "System Override Failed!", vbInformation, "E-Store : Overriding..."
    End If
    .Close
End With
Set rst = Nothing
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

Private Sub cmdRecord_Click()
If txtRemStock.Text <> "" Then
    Set rs = New ADODB.Recordset
    With rs
        .Open "select * from " & sMonth + sDate & " where ID =" & txtID.Text, con, 3, 3
        .Update
        .Fields("EndInv") = txtRemStock.Text
        .Update
        .Close
    End With
    Set rs = Nothing
    Call dbCompute
    If txtSearch.Text = "" Then
        Call dbase
    Else
        Call txtSearch_Change
    End If
    Call dbTotal
    
    txtRemStock.Text = ""
Else
    MsgBox "Empty Field!", vbInformation, "E-Store : Recording"
End If
End Sub

Private Sub cmdRefill_Click()
If txtStockRefill.Text <> "" Then
    Set rs = New ADODB.Recordset
    If MsgBox("Are you sure you want to refill " + txtPname.Text + "?", vbQuestion + vbOKCancel, "E-Store : Refilling...") = vbOK Then
        With rs
            .Open "select * from " & sMonth & sDate & " where ID = " & txtID, con, 3, 3
            .Fields("Refill") = .Fields("Refill") + txtStockRefill.Text
            .Fields("EndInv") = .Fields("EndInv") + txtStockRefill.Text
            .Fields("TotalBegInv") = .Fields("BegInv") + .Fields("Refill")
            .Fields("BegInvVal") = .Fields("TotalBegInv") * .Fields("RetailPrice")
            .Fields("EndInvVal") = .Fields("EndInv") * .Fields("RetailPrice")
            .Fields("TotalSoldItem") = .Fields("TotalBegInv") - .Fields("EndInv")
            .Fields("TotalSale") = .Fields("TotalSoldItem") * .Fields("RetailPrice")
            .Update
            .Close
        End With
        Call dbTotal
        ListView1.Refresh
    
        Call txtSearch_Change
    Else
        txtStockRefill.SetFocus
    End If
    Set rs = Nothing
    
    txtStockRefill.Text = ""
Else
    MsgBox "Empty Field!", vbInformation, "E-Store : Refilling..."
End If
End Sub

Private Sub Form_Load()
Set Connect = New Class1
txtun.Text = user1
txtpw.Text = timein1
StatusBar1.Panels.Item(2).Text = user1
Call dbase
Call dbTotal
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 38 Then
Call ListSelect
Call Status
End If
End Sub

Private Sub dbase()
ListView1.ListItems.Clear
Set rs = New ADODB.Recordset
With rs
    .Open "select * from " & sMonth + sDate & " order by ProdName", con, 3, 3
        Do While Not .EOF
        ListView1.ListItems.Add , , !ID & ""
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "" & !ProdCode
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = "" & !ProdName
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = "" & !ProdType
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !BegInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !Refill
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(6) = "" & !TotalBegInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(7) = "" & FormatNumber(!BegInvVal, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(8) = "" & !EndInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(9) = "" & FormatNumber(!EndInvVal, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(10) = "" & !TotalSoldItem
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(11) = "" & FormatNumber(!RetailPrice, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(12) = "" & FormatNumber(!TotalSale, 2, True, True, True)
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
txtBegInv.Text = ListView1.SelectedItem.SubItems(4)
txtRefill.Text = ListView1.SelectedItem.SubItems(5)
txtTotalBegInv.Text = Val(txtBegInv.Text) + Val(txtRefill.Text) 'ListView1.SelectedItem.SubItems(6)
txtBegInvVal.Text = ListView1.SelectedItem.SubItems(7)
txtEndInv.Text = ListView1.SelectedItem.SubItems(8)
txtEndInvVal.Text = ListView1.SelectedItem.SubItems(9)
txtSoldItems.Text = ListView1.SelectedItem.SubItems(10)
txtRetailPrice.Text = ListView1.SelectedItem.SubItems(11)
txtTotalSale.Text = ListView1.SelectedItem.SubItems(12)

End Sub


Private Sub ListView1_Click()
Call ListSelect
Call Status
End Sub

Private Sub mnuSysOver_Click()
    fmeSysOver.Visible = True
End Sub


Private Sub txtActualCash_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case Is < 32
Case 48 To 57
Case 46
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub txtCredit_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case Is < 32
Case 48 To 57
Case 46
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub txtRemStock_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
Case 13
    Call cmdRecord_Click
Case Is < 32
Case 48 To 57
Case 46
Case Else
KeyAscii = 0
End Select
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
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !BegInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !Refill
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(6) = "" & !TotalBegInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(7) = "" & FormatNumber(!BegInvVal, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(8) = "" & !EndInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(9) = "" & FormatNumber(!EndInvVal, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(10) = "" & !TotalSoldItem
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(11) = "" & FormatNumber(!RetailPrice, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(12) = "" & FormatNumber(!TotalSale, 2, True, True, True)
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
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !BegInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !Refill
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(6) = "" & !TotalBegInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(7) = "" & FormatNumber(!BegInvVal, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(8) = "" & !EndInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(9) = "" & FormatNumber(!EndInvVal, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(10) = "" & !TotalSoldItem
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(11) = "" & FormatNumber(!RetailPrice, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(12) = "" & FormatNumber(!TotalSale, 2, True, True, True)
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
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !BegInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !Refill
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(6) = "" & !TotalBegInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(7) = "" & FormatNumber(!BegInvVal, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(8) = "" & !EndInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(9) = "" & FormatNumber(!EndInvVal, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(10) = "" & !TotalSoldItem
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(11) = "" & FormatNumber(!RetailPrice, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(12) = "" & FormatNumber(!TotalSale, 2, True, True, True)
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
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !BegInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !Refill
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(6) = "" & !TotalBegInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(7) = "" & FormatNumber(!BegInvVal, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(8) = "" & !EndInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(9) = "" & FormatNumber(!EndInvVal, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(10) = "" & !TotalSoldItem
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(11) = "" & FormatNumber(!RetailPrice, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(12) = "" & FormatNumber(!TotalSale, 2, True, True, True)
        .MoveNext
    Loop
    .Close
End With
Case "Stock"
ListView1.ListItems.Clear
With rs
    .Open "select * from " & sMonth + sDate & " where BegInv like '" & txtSearch & "%' order by " & sMonth + sDate & ".TotalBegInv asc;", con, 3, 3
    Do While Not .EOF
        ListView1.ListItems.Add , , !ID & ""
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "" & !ProdCode
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = "" & !ProdName
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = "" & !ProdType
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !BegInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !Refill
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(6) = "" & !TotalBegInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(7) = "" & FormatNumber(!BegInvVal, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(8) = "" & !EndInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(9) = "" & FormatNumber(!EndInvVal, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(10) = "" & !TotalSoldItem
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(11) = "" & FormatNumber(!RetailPrice, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(12) = "" & FormatNumber(!TotalSale, 2, True, True, True)
        .MoveNext
    Loop
    .Close
End With
End Select
Set rs = Nothing

End Sub


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
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !BegInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !Refill
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(6) = "" & !TotalBegInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(7) = "" & FormatNumber(!BegInvVal, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(8) = "" & !EndInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(9) = "" & FormatNumber(!EndInvVal, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(10) = "" & !TotalSoldItem
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(11) = "" & FormatNumber(!RetailPrice, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(12) = "" & FormatNumber(!TotalSale, 2, True, True, True)
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
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !BegInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !Refill
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(6) = "" & !TotalBegInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(7) = "" & FormatNumber(!BegInvVal, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(8) = "" & !EndInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(9) = "" & FormatNumber(!EndInvVal, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(10) = "" & !TotalSoldItem
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(11) = "" & FormatNumber(!RetailPrice, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(12) = "" & FormatNumber(!TotalSale, 2, True, True, True)
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
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !BegInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !Refill
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(6) = "" & !TotalBegInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(7) = "" & FormatNumber(!BegInvVal, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(8) = "" & !EndInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(9) = "" & FormatNumber(!EndInvVal, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(10) = "" & !TotalSoldItem
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(11) = "" & FormatNumber(!RetailPrice, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(12) = "" & FormatNumber(!TotalSale, 2, True, True, True)
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
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !BegInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !Refill
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(6) = "" & !TotalBegInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(7) = "" & FormatNumber(!BegInvVal, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(8) = "" & !EndInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(9) = "" & FormatNumber(!EndInvVal, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(10) = "" & !TotalSoldItem
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(11) = "" & FormatNumber(!RetailPrice, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(12) = "" & FormatNumber(!TotalSale, 2, True, True, True)
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
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !BegInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !Refill
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(6) = "" & !TotalBegInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(7) = "" & FormatNumber(!BegInvVal, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(8) = "" & !EndInv
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(9) = "" & FormatNumber(!EndInvVal, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(10) = "" & !TotalSoldItem
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(11) = "" & FormatNumber(!RetailPrice, 2, True, True, True)
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(12) = "" & FormatNumber(!TotalSale, 2, True, True, True)
        .MoveNext
    Loop
    .Close
    End With
End If
Set rs = Nothing

'Call dbase
End Sub


Private Sub dbTotal()
    Set rsd = New ADODB.Recordset
    
    iTotalSale = 0
    With rsd
        .Open "select * from " & sMonth & sDate, con, 3, 3
        .MoveFirst
        Do While Not .EOF
        iTotalSale = iTotalSale + .Fields("TotalSale")
        .MoveNext
        Loop
        .Close
    End With
    Set rsd = Nothing
        Set rs = New ADODB.Recordset
        With rs
            .Open "select * from Trans where TransDate = '" & sMonth & sDate & "' and Set = 1", dbcon, 3, 3
            
            If .EOF = True Then
                lblTotalSale.Caption = "Php. " & FormatNumber(iTotalSale, 2, True, True, True)
            ElseIf .EOF = False Then
                iTotalSale = iTotalSale - .Fields("TotalSale")
                lblTotalSale.Caption = "Php. " & FormatNumber(iTotalSale, 2, True, True, True)
            End If
            .Close
        End With
    Set rs = Nothing
End Sub

Private Sub Status()
    If Val(txtTotalBegInv.Text) = 0 Then
        lblStatus.Caption = "Depleted"
        lblStatus.ForeColor = &H404040
    ElseIf Val(txtTotalBegInv.Text) < 5 Then
        lblStatus.Caption = "Critical"
        lblStatus.ForeColor = &HFF&
    ElseIf Val(txtTotalBegInv.Text) > 5 Then
        lblStatus.Caption = "Sufficient"
        lblStatus.ForeColor = &HFF00&
    End If
End Sub


Private Sub txtStockRefill_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
Case 13
    Call cmdRefill_Click
Case Is < 32
Case 48 To 57
Case 46
Case Else
KeyAscii = 0
End Select
End Sub
