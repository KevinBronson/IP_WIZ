Attribute VB_Name = "Global"
Option Explicit

'============================================================================
'           Copyright© 2002 Kevin Bronson. All rights reserved.
'============================================================================

'±


'===========================================================================
'                          Application Settings
'===========================================================================

'Application Settings
Public gblSoundON As Boolean

Public objCompanyDisplay As CompanyDisplay






'===========================================================================
'                               FUNCTIONS
'===========================================================================

'IsWholeEvenNumber
Public Function IsWholeEvenNumber(dblInput As Double) As Boolean
    IsWholeEvenNumber = ((dblInput Mod 2) = 0) And Int(dblInput) = dblInput
End Function

'IS_IP_Address
Public Function IS_IP_Address(strHost As String) As Boolean
    If strHost Like "*.*.*.*" Then
        Dim arrTemp() As String
        arrTemp = Split(strHost, ".")
        If UBound(arrTemp) = 3 Then
            Dim a As Long
            IS_IP_Address = True
            For a = 0 To 3
                If IsNumeric(arrTemp(a)) Then
                    
                Else
                    IS_IP_Address = False
                    Exit For
                End If
            Next
        End If
    End If
End Function

'IS_Valid_IP_Address
Public Function IS_Valid_IP_Address(strHost As String) As Boolean
    If IS_IP_Address(strHost) Then
        Dim arrTemp() As String
        Dim a As Long
        arrTemp = Split(strHost, ".")
        IS_Valid_IP_Address = True
        For a = 0 To 3
            If CLng(arrTemp(a)) > -1 And CLng(arrTemp(a)) < 256 Then
                
            Else
                IS_Valid_IP_Address = False
                Exit For
            End If
        Next
    End If
End Function














