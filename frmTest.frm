VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "frmTest"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin IP_WIZ.VisualPingControl VisualPingControl1 
      Height          =   1365
      Left            =   2520
      TabIndex        =   5
      Top             =   3870
      Width           =   3345
      _extentx        =   5900
      _extenty        =   2408
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   825
      Left            =   450
      TabIndex        =   4
      Top             =   2610
      Width           =   1275
   End
   Begin VB.CommandButton cmdCollapseAll 
      Caption         =   "Collapse All"
      Height          =   465
      Left            =   4950
      TabIndex        =   3
      Top             =   360
      Width           =   1725
   End
   Begin VB.CommandButton cmdExpandAll 
      Caption         =   "Expand All"
      Height          =   465
      Left            =   2700
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   465
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   2085
   End
   Begin IP_WIZ.IP_Config_Tree IP_Config_Tree1 
      Height          =   1005
      Left            =   180
      TabIndex        =   0
      Top             =   1170
      Width           =   7845
      _extentx        =   13838
      _extenty        =   7488
   End
   Begin VB.Image Image1 
      Height          =   1005
      Left            =   630
      Top             =   4050
      Width           =   915
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'============================================================================
'           Copyright© 2002 Kevin Bronson. All rights reserved.
'============================================================================

Private Sub cmdCollapseAll_Click()
    IP_Config_Tree1.CollapseAll
End Sub

Private Sub cmdExpandAll_Click()
    IP_Config_Tree1.ExpandAll
End Sub

Private Sub cmdRefresh_Click()
    IP_Config_Tree1.Refresh
End Sub





