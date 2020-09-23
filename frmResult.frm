VERSION 5.00
Begin VB.Form frmResult 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "..."
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   Icon            =   "frmResult.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.PictureBox picRes 
      AutoSize        =   -1  'True
      Height          =   1470
      Index           =   1
      Left            =   180
      Picture         =   "frmResult.frx":000C
      ScaleHeight     =   1410
      ScaleWidth      =   1410
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.PictureBox picRes 
      AutoSize        =   -1  'True
      Height          =   1470
      Index           =   0
      Left            =   180
      Picture         =   "frmResult.frx":04BE
      ScaleHeight     =   1410
      ScaleWidth      =   1410
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Label lblRes 
      Caption         =   "..."
      Height          =   675
      Left            =   1860
      TabIndex        =   2
      Top             =   360
      Width           =   3315
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Unload Me
End Sub

