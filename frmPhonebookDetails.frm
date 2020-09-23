VERSION 5.00
Begin VB.Form frmPhonebookDetails 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Phonebook Details"
   ClientHeight    =   4575
   ClientLeft      =   2910
   ClientTop       =   1815
   ClientWidth     =   4440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPhonebookDetails.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   4440
   Begin VB.TextBox PrimaryKey 
      Height          =   315
      Left            =   120
      TabIndex        =   18
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtMobilePhoneNumber 
      BackColor       =   &H00FFFFFF&
      DataField       =   "strCitizenshipDesc"
      DataSource      =   "Data1"
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1515
      TabIndex        =   6
      Top             =   3240
      Width           =   2775
   End
   Begin VB.TextBox txtEMail 
      BackColor       =   &H00FFFFFF&
      DataField       =   "strCitizenshipDesc"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1515
      TabIndex        =   7
      Top             =   3675
      Width           =   2775
   End
   Begin VB.TextBox txtFaxNo 
      BackColor       =   &H00FFFFFF&
      DataField       =   "strCitizenshipDesc"
      DataSource      =   "Data1"
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1515
      TabIndex        =   5
      Top             =   2820
      Width           =   2775
   End
   Begin VB.TextBox txtAddress 
      BackColor       =   &H00FFFFFF&
      DataField       =   "strCitizenshipDesc"
      DataSource      =   "Data1"
      Height          =   900
      IMEMode         =   3  'DISABLE
      Left            =   1515
      TabIndex        =   3
      Top             =   1395
      Width           =   2775
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3285
      TabIndex        =   9
      Top             =   4110
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2235
      TabIndex        =   8
      Top             =   4110
      Width           =   975
   End
   Begin VB.TextBox txtLastName 
      BackColor       =   &H00FFFFFF&
      DataField       =   "strCitizenshipDesc"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1515
      TabIndex        =   0
      Top             =   165
      Width           =   2775
   End
   Begin VB.TextBox txtMiddleName 
      BackColor       =   &H00FFFFFF&
      DataField       =   "strCitizenshipDesc"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1515
      TabIndex        =   2
      Top             =   975
      Width           =   2775
   End
   Begin VB.TextBox txtTelNo 
      BackColor       =   &H00FFFFFF&
      DataField       =   "strCitizenshipDesc"
      DataSource      =   "Data1"
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1515
      TabIndex        =   4
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox txtFirstName 
      BackColor       =   &H00FFFFFF&
      DataField       =   "strCitizenshipDesc"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1515
      TabIndex        =   1
      Top             =   555
      Width           =   2775
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C00000&
      Caption         =   "Cellular #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   180
      TabIndex        =   17
      Top             =   3240
      Width           =   1320
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C00000&
      Caption         =   "E-Mail Address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   180
      TabIndex        =   16
      Top             =   3675
      Width           =   1320
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C00000&
      Caption         =   "Fax Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   180
      TabIndex        =   15
      Top             =   2820
      Width           =   1200
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C00000&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   180
      TabIndex        =   14
      Top             =   1425
      Width           =   1200
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C00000&
      Caption         =   "Middle Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   13
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C00000&
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   12
      Top             =   180
      Width           =   1200
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C00000&
      Caption         =   "Telephone #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   11
      Top             =   2400
      Width           =   1200
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C00000&
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   10
      Top             =   585
      Width           =   1200
   End
End
Attribute VB_Name = "frmPhonebookDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Module by  Raymond Tan Chiu
'For your comments and suggestions you can conntact me at:
'           raymondchiu@eudoramail.com
'           rchiu@unionbankph.com
'           raymondchiu@edsamail.com.ph
'           (63)(917)376-1894
'           (63)(032)340-8471
'           (63)(032)254-7500
'Development Date   : 06-19-2001
'Description        : Shows a ADO data control binds with datagrid.  And a global connection of ADO
'Components         : Datagrid, Adodc(ADO Data Control), Datagrid

Public mblnAddmode As Boolean
Public mblnUpdated As Boolean
Public mstrLastName As String
Public mstrFirstName As String
Public mstrMiddleName As String
Public mstrAddress As String
Public mstrTelNo As String
Public mstrCellNo As String
Public mstrFaxNo As String
Public mstrEmail As String
Public mstrPrimaryKey As String

Public Sub IncrementRecord()
Dim Rs As ADODB.Recordset
Dim strSQL As String

strSQL = "SELECT ContactID FROM tblPhonebook ORDER BY ContactID DESC"
Set Rs = gadoConn.Execute(strSQL)
If Not Rs.EOF Then
    Me.PrimaryKey = Format(Rs!ContactID + 1, "0000")
Else
    Me.PrimaryKey = Format(1001, "0000")
End If
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub cmdOk_Click()
On Error Resume Next
If Me.txtLastName = "" Then
    MsgBox "User Login should not be Blank  !!", vbInformation, "User Login Error"
    Exit Sub
End If
    mblnUpdated = True
With Me
    .mstrPrimaryKey = Trim(.PrimaryKey)
    .mstrFirstName = Trim(.txtFirstName)
    .mstrLastName = Trim(.txtLastName)
    .mstrMiddleName = Trim(.txtMiddleName)
    .mstrAddress = Trim(.txtAddress)
    .mstrTelNo = Trim(.txtTelNo)
    .mstrCellNo = Trim(.txtMobilePhoneNumber)
    .mstrFaxNo = Trim(.txtFaxNo)
    .mstrEmail = Trim(.txtEMail)
End With
    
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case vbKeyEscape
        Call cmdCancel_Click
End Select
End Sub

Private Sub Form_Load()
    Me.Top = frmPhonebook.Top
    Me.Left = frmPhonebook.Left
End Sub
