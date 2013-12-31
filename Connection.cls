VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Connection"
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
    
Public Name As String

Dim strRawData As String
Dim strConnection_Specific_DNS_Suffix As String
Dim strMedia_State As String
Dim strDescription As String
Dim strPhysical_Address As String
Dim blnDHCP_Enabled As Boolean
Dim strDefault_Gateway As String
Dim colIP_Stacks As New Collection
Dim objDefaultIP_Stack As New IP_Stack
Dim strPrimary_WINS_Server As String
Dim strSecondary_WINS_Server As String
Dim colDNS_Servers As New Collection
Dim objDefaultDNS_Server As New DNS_Server
Dim lngCount As Long
Dim strLastError As String


'==================================================================================
'                               PUBLIC PROPERTIES
'==================================================================================

'LastError
Public Property Get LastError() As String
    LastError = strLastError
End Property

'Count
Public Property Get Count() As Long
    Count = lngCount
End Property

Friend Property Let Count(lngInput As Long)
    lngCount = lngInput
End Property

'Primary_WINS_Server
Public Property Get Primary_WINS_Server() As String
    Primary_WINS_Server = strPrimary_WINS_Server
End Property

Friend Property Let Primary_WINS_Server(strInput As String)
    strPrimary_WINS_Server = strInput
End Property

'Secondary_WINS_Server
Public Property Get Secondary_WINS_Server() As String
    Secondary_WINS_Server = strSecondary_WINS_Server
End Property

Friend Property Let Secondary_WINS_Server(strInput As String)
    strSecondary_WINS_Server = strInput
End Property

'DNS_Server
Public Property Get DNS_Server(Optional Index As Long) As DNS_Server
    If Index = 0 Then
        Set DNS_Server = objDefaultDNS_Server
    Else
        Set DNS_Server = colDNS_Servers(Index)
    End If
End Property

'IP_Stack
Public Property Get IP_Stack(Optional Index As Long) As IP_Stack
    If Index = 0 Then
        Set IP_Stack = objDefaultIP_Stack
    Else
        Set IP_Stack = colIP_Stacks(Index)
    End If
End Property

'Connection_Specific_DNS_Suffix
Public Property Get Connection_Specific_DNS_Suffix() As String
    Connection_Specific_DNS_Suffix = strConnection_Specific_DNS_Suffix
End Property

'Description
Public Property Get Description() As String
    Description = strDescription
End Property

'Media_State
Public Property Get Media_State() As String
    Media_State = strMedia_State
End Property

'Physical_Address
Public Property Get Physical_Address() As String
    Physical_Address = strPhysical_Address
End Property

'DHCP_Enabled
Public Property Get DHCP_Enabled() As Boolean
    DHCP_Enabled = blnDHCP_Enabled
End Property

'Default_Gateway
Public Property Get Default_Gateway() As String
    Default_Gateway = strDefault_Gateway
End Property

'RawData
Friend Property Get RawData() As String
    RawData = strRawData
End Property

Friend Property Let RawData(strInput As String)
    strRawData = strInput
    Call ParseData
End Property

'ParseData
Private Sub ParseData()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Dim arrVariables() As String
    Dim a As Long
    
    arrVariables = Split(strRawData, "±")
    
    'Make sure not even number
    If UBound(arrVariables) Mod 2 = 0 Then ReDim Preserve arrVariables(UBound(arrVariables) + 1)
    
    For a = 0 To UBound(arrVariables) Step 2
        If arrVariables(a) <> "" Then
            arrVariables(a) = Trim(arrVariables(a))
        End If
        Select Case arrVariables(a)
            Case "Connection-specific DNS Suffix"
                strConnection_Specific_DNS_Suffix = arrVariables(a + 1)
            Case "Description"
                strDescription = arrVariables(a + 1)
            Case "Media State"
                strMedia_State = arrVariables(a + 1)
            Case "Physical Address"
                strPhysical_Address = arrVariables(a + 1)
            Case "DHCP Enabled"
                blnDHCP_Enabled = (arrVariables(a + 1) = "Yes")
            Case "IP Address"
                Dim objTemp As New IP_Stack
                colIP_Stacks.ADD objTemp
                objTemp.IP_Address = arrVariables(a + 1)
                objTemp.Subnet_Mask = arrVariables(a + 3)
                Dim b As Long
                For b = 1 To colIP_Stacks.Count
                    'REM- Must take object out of collection and use it directly
                    ' or else will get error from 'Friend Poperty Let Count' inside
                    ' the IP_Stack Object
                    Set objTemp = colIP_Stacks(b)
                    objTemp.Count = colIP_Stacks.Count
                Next
                objDefaultIP_Stack.Count = colIP_Stacks.Count
                
                If Not objTemp Is Nothing Then Set objTemp = Nothing
            'Case "Subnet Mask"
                'strSubnet_Mask = arrVariables(a + 1)
            Case "Default Gateway"
                strDefault_Gateway = arrVariables(a + 1)
            Case "DNS Servers"
                'There could be zero or more values here
                Do While IS_IP_Address(arrVariables(a + 1))
                    Dim objDNS_Server As New DNS_Server
                    Dim c As Long
                    objDNS_Server.IP_Address = arrVariables(a + 1)
                    colDNS_Servers.ADD objDNS_Server
                    objDNS_Server.Count = colDNS_Servers.Count
                    objDefaultDNS_Server.Count = colDNS_Servers.Count
                    For c = 1 To colDNS_Servers.Count
                        Set objDNS_Server = colDNS_Servers(c)
                        objDNS_Server.Count = colDNS_Servers.Count
                    Next
                    If Not objDNS_Server Is Nothing Then Set objDNS_Server = Nothing
                    arrVariables(a + 1) = ""
                    If UBound(arrVariables) >= a + 2 Then
                        If IS_IP_Address(arrVariables(a + 2)) Then a = a + 1
                    End If
                Loop
            Case "Primary WINS Server"
                strPrimary_WINS_Server = arrVariables(a + 1)
            Case "Secondary WINS Server"
                strSecondary_WINS_Server = arrVariables(a + 1)
        End Select
    Next
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, "Connection", "ParseData")
End Sub













