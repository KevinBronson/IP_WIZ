VERSION 5.00
Begin VB.Form frmDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Details"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDetails 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   5145
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   7935
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   465
      Left            =   1170
      TabIndex        =   1
      Top             =   2880
      Width           =   2445
   End
End
Attribute VB_Name = "frmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'============================================================================
'           Copyright© 2002 Kevin Bronson. All rights reserved.
'============================================================================

Public WithEvents objVisualPingControl As VisualPingControl
Attribute objVisualPingControl.VB_VarHelpID = -1

'cmdClose_Click
Private Sub cmdClose_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Unload Me
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "cmdClose_Click")
End Sub

'objVisualPingControl_DetailsUpdate
Private Sub objVisualPingControl_DetailsUpdate()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    ShowDetails
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "objVisualPingControl_DetailsUpdate")
End Sub

'ShowDetails
Public Sub ShowDetails()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    txtDetails = objVisualPingControl.Details
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "ShowDetails")
End Sub




