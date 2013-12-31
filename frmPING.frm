VERSION 5.00
Begin VB.Form frmPING 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7830
      TabIndex        =   3
      Top             =   1710
      Width           =   2085
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   2250
      TabIndex        =   2
      Top             =   1710
      Width           =   2085
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   90
      TabIndex        =   1
      Top             =   1710
      Width           =   2085
   End
   Begin IP_WIZ.VisualPingControl vpc1 
      Height          =   1635
      Left            =   0
      TabIndex        =   0
      Top             =   90
      Width           =   10545
      _extentx        =   18600
      _extenty        =   2884
   End
End
Attribute VB_Name = "frmPING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'============================================================================
'           Copyright© 2002 Kevin Bronson. All rights reserved.
'============================================================================


'PING
Public Sub PING(strHost As String)
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If strHost = "" Then
        Me.Hide
        MsgBox "Invalid IP Address."
        Unload Me
    Else
        vpc1.Caption = "Pinging " & strHost & "..."
        Me.Caption = vpc1.Caption
        vpc1.ComputerNameOrIP = strHost
        Call cmdStart_Click
    End If
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "PING")
End Sub

'cmdCancel_Click
Private Sub cmdCancel_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    vpc1.StopTest
    Unload Me
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "cmdCancel_Click")
End Sub

'cmdStart_Click
Private Sub cmdStart_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    vpc1.StartTest
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "cmdStart_Click")
End Sub

'cmdStop_Click
Private Sub cmdStop_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    vpc1.StopTest
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "cmdStop_Click")
End Sub

'Form_Load
Private Sub Form_Load()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    vpc1.DetailsButtonVisible = True
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "Form_Load")
End Sub

'Form_Unload
Private Sub Form_Unload(Cancel As Integer)
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    vpc1.StopTest
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "Form_Unload")
End Sub

'vpc1_DetailsButtonClicked
Private Sub vpc1_DetailsButtonClicked()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    frmDetails.Show
    Set frmDetails.objVisualPingControl = vpc1
    frmDetails.ShowDetails
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "vpc1_DetailsButtonClicked")
End Sub





