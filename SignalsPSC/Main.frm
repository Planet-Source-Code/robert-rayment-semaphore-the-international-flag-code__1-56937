VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   " SIGNALS"
   ClientHeight    =   1785
   ClientLeft      =   105
   ClientTop       =   -345
   ClientWidth     =   1575
   DrawWidth       =   2
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   119
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   1710
      Left            =   60
      TabIndex        =   0
      Top             =   -15
      Width           =   1470
      Begin VB.CommandButton cmdSignal 
         Caption         =   "SEMAPHORE"
         Height          =   450
         Index           =   0
         Left            =   135
         TabIndex        =   3
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton cmdSignal 
         Caption         =   "FLAG CODE"
         Height          =   450
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   705
         Width           =   1200
      End
      Begin VB.CommandButton cmdSignal 
         Caption         =   "EXIT"
         Height          =   390
         Index           =   2
         Left            =   135
         TabIndex        =   1
         Top             =   1170
         Width           =   1200
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' SEMAPHORE & INTERNATIONAL FLAG CODE
' by  Robert Rayment

' Form1 (Main.frm)

Option Explicit

Private Sub cmdSignal_Click(Index As Integer)
   Select Case Index
   Case 0   ' SEMAPHORE
      frmSema.Show vbModal
   Case 1   ' FLAG CODE
      frmFlags.Show vbModal
   Case 2   ' EXIT
      Form_Unload 0
   End Select
End Sub

Private Sub Form_Load()
   frmSema.Hide
   frmFlags.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Form As Form
   ' Make sure all forms cleared
   For Each Form In Forms
      Unload Form
      Set Form = Nothing
   Next Form
   End
End Sub
