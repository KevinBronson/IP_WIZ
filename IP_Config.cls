VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IP_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'============================================================================
'           Copyright© 2002 Kevin Bronson. All rights reserved.
'============================================================================

    '==========================================
    '       GENERAL OBJECT HIERARCHY
    '==========================================
    'IP_Config >> Connections >> IP Stacks
    '                         >> DNS Servers
    '==========================================
    '==========================================
    
Dim strTitle As String
Dim strHostName As String
Dim strPrimary_DNS_Suffix As String
Dim strNodeType As String
Dim blnIP_RoutingEnabled As Boolean
Dim blnWINS_Proxy_Enabled As Boolean
Dim colConnections As New Collection
Dim strIP_ConfigResults As String
Dim strMainRawData As String
Dim objDefaultConnection As New Connection
Dim strLastError As String


'==================================================================================
'                               PUBLIC PROPERTIES
'==================================================================================

'LastError
Public Property Get LastError() As String
    LastError = strLastError
End Property

'IP_Config_Results
Public Property Get IP_Config_Results()
    IP_Config_Results = Replace(strIP_ConfigResults, "±", vbCrLf)
End Property

'Connection
Public Property Get Connection(Optional lngIndex As Long) As Connection
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If lngIndex = 0 Then
        Set Connection = objDefaultConnection
    Else
        Set Connection = colConnections(lngIndex)
    End If
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, "IP_Config", "Connection")
End Property

'Title
Public Property Get Title() As String
    Title = strTitle
End Property

'HostName
Public Property Get HostName() As String
    HostName = strHostName
End Property

'Primary_DNS_Suffix
Public Property Get Primary_DNS_Suffix() As String
    Primary_DNS_Suffix = strPrimary_DNS_Suffix
End Property

'NodeType
Public Property Get NodeType() As String
    NodeType = strNodeType
End Property

'IP_RoutingEnabled
Public Property Get IP_RoutingEnabled() As Boolean
    IP_RoutingEnabled = blnIP_RoutingEnabled
End Property

'WINS_Proxy_Enabled
Public Property Get WINS_Proxy_Enabled() As Boolean
    WINS_Proxy_Enabled = blnWINS_Proxy_Enabled
End Property

'Refresh
Public Sub Refresh()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Call Reset
    Call GetIP_ConfigData
    Call ParseData
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, "IP_Config", "Refresh")
End Sub

'Reset
Private Sub Reset()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    'Reset all objects
    Call ClearCollection(colConnections)
    If Not objDefaultConnection Is Nothing Then Set objDefaultConnection = Nothing
    Set objDefaultConnection = New Connection
    'Reset all Properties and variables
    strTitle = ""
    strHostName = ""
    strPrimary_DNS_Suffix = ""
    strNodeType = ""
    blnIP_RoutingEnabled = False
    blnWINS_Proxy_Enabled = False
    strIP_ConfigResults = ""
    strMainRawData = ""
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, "IP_Config", "Reset")
End Sub

'ClearCollection
Private Sub ClearCollection(objCol As Collection)
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If Not objCol Is Nothing Then Set objCol = Nothing
    Set objCol = New Collection
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, "IP_Config", "ClearCollection")
End Sub

'Class_Initialize
Private Sub Class_Initialize()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Call Refresh
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, "IP_Config", "Class_Initialize")
End Sub

'GetIP_ConfigData
Private Sub GetIP_ConfigData()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Const strBatchFile = "C:\IP_WIZ_ipconfig.bat"
    Const strDataFile = "C:\IP_WIZ_ipconfig.txt"
    Dim objShell As New Shell32.Shell
    Dim FS As New Scripting.FileSystemObject
    Dim TS As Scripting.TextStream
    'Create Batch file
    Set TS = FS.OpenTextFile(strBatchFile, ForWriting, True)
    TS.Write "ipconfig /all > " & strDataFile
    TS.Close
    'Run Batch File
    objShell.Open strBatchFile
    Rest 300 'Gives time to write to disk
    'Gather ipconfig Results
    strIP_ConfigResults = FS.OpenTextFile(strDataFile, ForReading).ReadAll
    'Clean up
    FS.DeleteFile strBatchFile, True
    FS.DeleteFile strDataFile, True
    If Not objShell Is Nothing Then Set objShell = Nothing
    If Not FS Is Nothing Then Set FS = Nothing
    If Not TS Is Nothing Then Set TS = Nothing
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, "IP_Config", "GetIP_ConfigData")
End Sub

'ParseData
Private Sub ParseData()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Dim a As Long
    Dim arrMajorParts() As String
    Dim arrVariables() As String
    Call CleanString
    
    'Divide into major part and hand out strings to objects
    'REM- Cariage returns have been replaced with "±"
    arrMajorParts = Split(strIP_ConfigResults, "±±±±")
    
    'Debug.Print Join(arrMajorParts, "==============" & vbCrLf)
    strTitle = Replace(arrMajorParts(0), "±", "")
    strMainRawData = arrMajorParts(1) 'First part of IPCONFIG data.. the header
    For a = 3 To UBound(arrMajorParts) Step 2
        'Every other part will be connection data...
        Dim objTemp As New Connection
        objTemp.Name = arrMajorParts(a - 1)
        objTemp.RawData = arrMajorParts(a)
        colConnections.ADD objTemp
        Set objTemp = Nothing
        objDefaultConnection.Count = colConnections.Count
        Dim d As Long
        For d = 1 To colConnections.Count
            Dim objTempConn As Connection
            Set objTempConn = colConnections(d)
            objTempConn.Count = colConnections.Count
        Next
        If Not objTempConn Is Nothing Then Set objTempConn = Nothing
    Next
    arrVariables = Split(strMainRawData, "±")
    'Assign Values to Public Properties
    For a = 0 To UBound(arrVariables) Step 2
        Select Case arrVariables(a)
            Case "Host Name"
                strHostName = arrVariables(a + 1)
            Case "Primary DNS Suffix"
                strPrimary_DNS_Suffix = arrVariables(a + 1)
            Case "Node Type"
                strNodeType = arrVariables(a + 1)
            Case "IP Routing Enabled"
                blnIP_RoutingEnabled = (arrVariables(a + 1) = "Yes")
            Case "WINS Proxy Enabled"
                blnWINS_Proxy_Enabled = (arrVariables(a + 1) = "Yes")
        End Select
    Next
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, "IP_Config", "ParseData")
End Sub

'CleanString
Private Sub CleanString()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Dim arrVariables() As String
    Dim a As Long
    strIP_ConfigResults = Replace(strIP_ConfigResults, " .", "")
    strIP_ConfigResults = Replace(strIP_ConfigResults, vbCrLf, "±")
    strIP_ConfigResults = Replace(strIP_ConfigResults, Chr(9), "")
    strIP_ConfigResults = Replace(strIP_ConfigResults, Chr(13), "")
    Do While InStr(strIP_ConfigResults, "  ")
        strIP_ConfigResults = Replace(strIP_ConfigResults, "  ", " ")
    Loop
    strIP_ConfigResults = Replace(strIP_ConfigResults, "±±", "±±±±")
    'Split into variable names and data
    arrVariables = Split(strIP_ConfigResults, " : ")
    'Debug.Print Join(arrVariables, "========" & vbCrLf)
    For a = 0 To UBound(arrVariables)
        arrVariables(a) = Trim(arrVariables(a))
        'Clean variable name
        If Right(arrVariables(a), 1) = "." Then
            arrVariables(a) = Left(arrVariables(a), Len(arrVariables(a)) - 1)
        End If
    Next
    strIP_ConfigResults = Join(arrVariables, "±")
    'Debug.Print strIP_ConfigResults
    If Right(strIP_ConfigResults, 1) = "±" Then strIP_ConfigResults = Left(strIP_ConfigResults, Len(strIP_ConfigResults) - 1)
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, "IP_Config", "CleanString")
End Sub















