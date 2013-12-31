VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DNS_Server"
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
    
Dim lngCount As Long
Dim strIP_Address As String
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

'IP_Address
Public Property Get IP_Address() As String
    IP_Address = strIP_Address
End Property

Friend Property Let IP_Address(strInput As String)
    strIP_Address = strInput
End Property

















