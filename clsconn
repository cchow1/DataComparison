Option Explicit
Private p_strProvider As String
Private p_strConn As String
Private p_strProperties As String
Private p_con As ADODB.Connection

Private Sub Class_Initialize() 'Use initialization to make sure there's Always the private connection object
    Set p_con = New ADODB.Connection
End Sub
Private Sub Class_Terminate() 'Clean up after yourself
    Set p_con = Nothing
End Sub

'Properties needed:
Property Let ConnProvider(strCP As String)
    p_strProvider = strCP
End Property
Property Let ConnString(strCS As String)
    p_strConn = strCS
End Property
Property Let ConnProperties(strCPP As String)
    p_strProperties = strCPP
End Property

Private Sub OpenConnection()
'Takes the variables, builds a connectionstring, creates the connection
    Dim conStr As String
    If p_strProvider = "" Or p_strConn = "" Or p_strProperties = "" Then
        MsgBox "Connection parameters were not provided."
        Exit Sub
    Else
        conStr = "Data Source=" & p_strConn & "; Extended Properties='" & p_strProperties & "'"
        With p_con
            .Provider = p_strProvider
            .ConnectionString = conStr
            .CursorLocation = adUseClient
            .Open
        End With
    End If
End Sub
Public Function GetConnectionObject() As ADODB.Connection
'Builds and then exposes the connection object from within the class to the outside world
    OpenConnection
    Set GetConnectionObject = p_con
    Debug.Print GetConnectionObject
End Function
