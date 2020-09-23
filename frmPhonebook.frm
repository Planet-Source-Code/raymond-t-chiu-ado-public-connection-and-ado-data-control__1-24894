VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPhonebook 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Phonebook"
   ClientHeight    =   4575
   ClientLeft      =   2910
   ClientTop       =   1830
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPhonebook.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6855
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6240
      Picture         =   "frmPhonebook.frx":030A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   8
      ToolTipText     =   "Click on me to send e-mail to author..."
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Refresh"
      Height          =   600
      Left            =   5160
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPhonebook.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Search"
      Top             =   3840
      Width           =   750
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Close"
      Height          =   600
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPhonebook.frx":0CD6
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Search"
      Top             =   3840
      Width           =   750
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Erase"
      Height          =   600
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPhonebook.frx":0DD0
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Search"
      Top             =   3840
      Width           =   750
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&New"
      Height          =   600
      Left            =   3480
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPhonebook.frx":0ED2
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Search"
      Top             =   3840
      Width           =   750
   End
   Begin VB.CommandButton cmdSearch 
      Height          =   315
      Left            =   3960
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPhonebook.frx":0FD4
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Search"
      Top             =   120
      Width           =   360
   End
   Begin VB.TextBox txtSearchValue 
      BackColor       =   &H00FFFFFF&
      DataField       =   "strCitizenshipDesc"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin MSDataGridLib.DataGrid dbgPhonebook 
      Bindings        =   "frmPhonebook.frx":10D6
      Height          =   3135
      Left            =   90
      TabIndex        =   1
      ToolTipText     =   "Double Click to select the data"
      Top             =   600
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "TelephoneNo"
         Caption         =   "Telephone Number"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   2
         ScrollBars      =   2
         BeginProperty Column00 
            ColumnWidth     =   3495.118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2594.835
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adcPhonebook 
      Height          =   330
      Left            =   120
      Top             =   3840
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Phonebook..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Search:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmPhonebook"
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

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdClose_Click()
On Error Resume Next
    Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim SqlString As String
Dim Rs As New ADODB.Recordset

    If MsgBox("Confirm deletion of this record.  Delete record now?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Delete") = vbYes Then
        'delete record
            SqlString = "Delete from tblPhonebook where ContactID ='" & Me.adcPhonebook.Recordset("ContactID") & "'"
            gadoConn.Execute SqlString
            Call RefreshGrid
    End If
End Sub

Private Sub cmdNew_Click()
Dim Rs As New ADODB.Recordset
Dim strSQL As String
   
'On Error Resume Next

    With frmPhonebookDetails
        .mblnAddmode = True
        .mblnUpdated = False
        .IncrementRecord
        .Show vbModal
        
        If .mblnUpdated Then
            strSQL = "SELECT * from tblPhonebook where 1<>1"
                        
            Rs.Open strSQL, gadoConn, adOpenKeyset, adLockOptimistic
            
            Rs.AddNew
            
            Rs!ContactID = Trim(.mstrPrimaryKey)
            Rs!LastName = Trim(.mstrLastName)
            Rs!FirstName = Trim(.mstrFirstName)
            Rs!MiddleName = Trim(.mstrMiddleName)
            Rs!TelephoneNo = Trim(.mstrTelNo)
            Rs!FaxNo = (.mstrFaxNo)
            Rs!EMailAddress = (.mstrEmail)
            Rs!MobilePhoneNo = (.mstrCellNo)
            Rs!Address = (.mstrAddress)
            
            Rs.Update
            
            Rs.Close
            Set Rs = Nothing
            Call RefreshGrid
        End If
    End With
End Sub

Private Sub cmdRefresh_Click()
    Call RefreshGrid
    Me.txtSearchValue = ""
End Sub

Private Sub cmdSearch_Click()
Dim Wherestring As String, Sstring As String
'On Error Resume Next
    Sstring = Trim(txtSearchValue)
    Wherestring = " WHERE FirstName LIKE '" + Trim(Sstring) + "%' "
    Call RefreshGrid(Wherestring)
End Sub

Private Sub dbgPhonebook_DblClick()
Dim Rs As New ADODB.Recordset
Dim strSQL As String
   
On Error GoTo ErrorHandler
   
    With frmPhonebookDetails
        .mblnAddmode = True
        .mblnUpdated = False
        .PrimaryKey = Me.adcPhonebook.Recordset("ContactID")
        .txtFirstName = Me.adcPhonebook.Recordset("FirstName")
        .txtLastName = Me.adcPhonebook.Recordset("LastName")
        .txtMiddleName = Me.adcPhonebook.Recordset("MiddleName")
        .txtTelNo = Me.adcPhonebook.Recordset("TelephoneNo")
        .txtFaxNo = Me.adcPhonebook.Recordset("FaxNo")
        .txtEMail = Me.adcPhonebook.Recordset("EMailAddress")
        .txtMobilePhoneNumber = Me.adcPhonebook.Recordset("MobilePhoneNo")
        .txtAddress = Me.adcPhonebook.Recordset("Address")
        
        .Show vbModal
        
        If .mblnUpdated Then
            strSQL = "SELECT * from tblPhonebook where ContactID = '" & Me.adcPhonebook.Recordset("ContactID") & "'"
                        
            Rs.Open strSQL, gadoConn, adOpenKeyset, adLockOptimistic
            
            Rs!ContactID = Trim(.mstrPrimaryKey)
            Rs!LastName = Trim(.mstrLastName)
            Rs!FirstName = Trim(.mstrFirstName)
            Rs!MiddleName = Trim(.mstrMiddleName)
            Rs!TelephoneNo = Trim(.mstrTelNo)
            Rs!FaxNo = (.mstrFaxNo)
            Rs!EMailAddress = (.mstrEmail)
            Rs!MobilePhoneNo = (.mstrCellNo)
            Rs!Address = (.mstrAddress)
            
            Rs.Update
            
            Rs.Close
            Set Rs = Nothing
            Call RefreshGrid
        End If
    End With
Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 94
            Resume Next
        Case Else
            Dim intAns As Integer
            intAns = MsgBox("Error saving record.  " & Trim(Err.Description), vbExclamation + vbAbortRetryIgnore + vbDefaultButton2, "Save Education Error")
            
            Select Case intAns
                Case vbAbort
                    Exit Sub
                Case vbRetry
                    Resume
                Case vbIgnore
                    Resume Next
            End Select
    End Select
End Sub

Private Sub dbgPhonebook_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case vbKeyDelete
        Call cmdDelete_Click
    Case vbKeyReturn
        Call dbgPhonebook_DblClick
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case vbKeyEscape
        Unload Me
End Select
End Sub

Private Sub Form_Load()
    Call RefreshGrid
End Sub

Private Sub RefreshGrid(Optional ByVal WhereStr As String)
Dim strSQL As String
Dim Rs As New ADODB.Recordset
    
On Error Resume Next

    strSQL = "SELECT ContactID, [FirstName] + ' ' + [MiddleName] + ' ' + [LastName] AS Name, *"
    strSQL = strSQL + " FROM tblPhonebook"
    If WhereStr <> "" Then
        strSQL = strSQL + WhereStr
    End If
    strSQL = strSQL + " ORDER BY FirstName ASC"
    
    Rs.Open strSQL, gadoConn, adOpenKeyset, adLockOptimistic
    
    Set Me.adcPhonebook.Recordset = Rs
    adcPhonebook.Refresh
    dbgPhonebook.ReBind
End Sub

Private Sub Label1_DblClick()
    MsgBox "Smart Minds Software Development Company", vbInformation, "Raymond T. Chiu"
End Sub

Private Sub Picture1_Click()
    ShellExecute hwnd, "open", "mailto:raymondchiu@eudoramail.com", vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub txtSearchValue_DblClick()
    Me.txtSearchValue = ""
End Sub

Private Sub txtSearchValue_KeyPress(KeyAscii As Integer)
    cmdSearch_Click
End Sub
