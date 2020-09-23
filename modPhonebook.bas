Attribute VB_Name = "modPhonebook"
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

Public gadoConn As ADODB.Connection

Sub Main()
'starting point of all ITC Applications
    On Error Resume Next
    If ConnectToDatabase = False Then
        MsgBox "Database not found!, Please place the database with the same path as the .exe file", vbExclamation, "Database not found!"
    End If
    frmPhonebook.Show
End Sub

Public Function ConnectToDatabase() As Boolean
Dim strConnect As String
Dim Path As String

    On Error GoTo ConnectError
    Path = App.Path & "\Phonebook.mdb"
    strConnect = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=false;"
    strConnect = strConnect & "Data Source=" & Path & ";Jet OLEDB:Database password = Pure;"
                
    Set gadoConn = New ADODB.Connection
    gadoConn.CursorLocation = adUseServer
    
    gadoConn.Open strConnect
        
    ConnectToDatabase = True
    
    Exit Function

ConnectError:
    MsgBox Error$, vbExclamation, "Connection Error"
    ConnectToDatabase = False
End Function
