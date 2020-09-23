VERSION 5.00
Begin VB.Form frmFlags 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000016&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " INTERNATIONAL FLAG CODE"
   ClientHeight    =   4455
   ClientLeft      =   150
   ClientTop       =   0
   ClientWidth     =   4665
   ControlBox      =   0   'False
   DrawWidth       =   2
   Icon            =   "Flags.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   297
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   311
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkHideText 
      Caption         =   "Hide text"
      Height          =   255
      Left            =   270
      TabIndex        =   13
      Top             =   4065
      Width           =   1200
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      Left            =   2340
      Max             =   50
      Min             =   1
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2055
      Value           =   1
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Height          =   1500
      Left            =   2700
      TabIndex        =   4
      Top             =   45
      Width           =   1650
      Begin VB.CommandButton cmdSSS 
         Caption         =   "CLOSE"
         Height          =   285
         Index           =   2
         Left            =   195
         TabIndex        =   7
         Top             =   1050
         Width           =   1215
      End
      Begin VB.CommandButton cmdSSS 
         Caption         =   "CANCEL"
         Height          =   285
         Index           =   1
         Left            =   195
         TabIndex        =   6
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton cmdSSS 
         Caption         =   "SEND"
         Height          =   285
         Index           =   0
         Left            =   195
         TabIndex        =   5
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.CheckBox chkHide 
      BackColor       =   &H80000016&
      Caption         =   "Hide captions"
      Height          =   285
      Left            =   225
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   1605
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Text            =   "Flags.frx":0442
      Top             =   2400
      Width           =   4185
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   255
      ScaleHeight     =   103.125
      ScaleMode       =   0  'User
      ScaleWidth      =   112.8
      TabIndex        =   1
      Top             =   195
      Width           =   1440
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   405
         Left            =   -180
         Top             =   750
         Width           =   450
      End
   End
   Begin VB.PictureBox picSrc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9285
      Left            =   5385
      Picture         =   "Flags.frx":0448
      ScaleHeight     =   617
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   485
      TabIndex        =   0
      Top             =   -5505
      Width           =   7305
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000016&
      Caption         =   "The flags would be strung out along a line."
      Height          =   450
      Left            =   225
      TabIndex        =   12
      Top             =   1455
      Width           =   2250
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000016&
      Caption         =   "s"
      Height          =   165
      Left            =   4200
      TabIndex        =   11
      Top             =   2085
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3780
      TabIndex        =   10
      Top             =   2070
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000016&
      Caption         =   "Delay"
      Height          =   240
      Left            =   1800
      TabIndex        =   9
      Top             =   2085
      Width           =   510
   End
End
Attribute VB_Name = "frmFlags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmFlags (Flags.frm)
' Intenational Flag Code by Robert Rayment

Option Explicit

' Flag Bitmap
' 40 (5 x 8) images each W 77 x H 97 pixels

Private Declare Function BitBlt Lib "gdi32" _
   (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
   ByVal nWidth As Long, ByVal nHeight As Long, _
   ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim aDone As Boolean
Dim aNum As Boolean
Dim Delay As Long

Private Sub Form_Load()
Dim k As Long, N As Long
Dim Cul As Long
   Cul = vbWhite
   N = 36: Disp N  ' Blank
   With Shape1
      .FillColor = Cul
      .BorderColor = Cul
      .FillStyle = 1
   End With
   chkHide.Value = 1
   HScroll1.Value = 5
   PIC.Width = 96
   PIC.Height = 77
   Show
   'Text1.Text = "INTERNATIONAL FLAG CODE" & vbCrLf & "BY" & vbCrLf & "ROBERT RAYMENT"
   Text1.Text = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" & vbCr & "0123456789"
End Sub

Private Sub cmdSSS_Click(Index As Integer)
Dim A$
Dim AscCode As Long
Dim k As Long, N As Long
   
   N = 36: Disp N  ' Blank

   Select Case Index
   Case 0   ' SEND
      AscCode = 0
      aNum = False
      aDone = False
      A$ = UCase$(Text1.Text)
      For k = 1 To Len(A$)
         DoEvents
         If aDone Then Exit For
         AscCode = Asc(Mid$(A$, k, 1))
         Select Case AscCode
         Case Is <= 47, 58 To 64, Is >= 123
            N = 36: Disp N  ' Blank
         Case 48 To 57
            N = AscCode - 22: Disp N  ' 0 = 26 Numbers
         Case 65 To 90
            N = AscCode - 65: Disp N   ' Characters
         Case Else
            N = 36: Disp N  ' Blank
         End Select
      Next k
      
      If k = Len(A$) + 1 Then
         N = 36: Disp N  ' Blank
      End If
   
   Case 1   ' CANCEL
      aDone = True
      aNum = False
      N = 36: Disp N ' Blank
      
   Case 2   ' Exit
      aDone = True
      Me.Hide
   End Select
End Sub

Private Sub Disp(N As Long)
Dim ix As Long, iy As Long
   If N = 36 Then  ' ie Blank. Blink it.
      iy = (37 \ 5) * 77
      ix = (37 Mod 5) * 97
      BitBlt PIC.hDC, 0, 0, 97, 77, picSrc.hDC, ix, iy, vbSrcCopy
      PIC.Refresh
      Sleep 60
   End If

   iy = (N \ 5) * 77
   ix = (N Mod 5) * 97
   BitBlt PIC.hDC, 0, 0, 97, 77, picSrc.hDC, ix, iy, vbSrcCopy
   PIC.Refresh
   Sleep Delay * 100
End Sub

Private Sub chkHide_Click()
   If chkHide.Value Then
      Shape1.FillStyle = 0
   Else
      Shape1.FillStyle = 1
   End If
End Sub

Private Sub chkHideText_Click()
   If chkHideText.Value Then
      Text1.ForeColor = vbWhite
   Else
      Text1.ForeColor = 0
   End If
End Sub

Private Sub HScroll1_Change()
Dim D$
Dim zNum As Single
   aDone = True
   Delay = HScroll1.Value
   zNum = Delay / 10
   D$ = Str$(zNum)
   If zNum < 1 Then D$ = "0" & D$
   Label2 = D$
End Sub

Private Sub HScroll1_Scroll()
Dim D$
Dim zNum As Single
   aDone = True
   Delay = HScroll1.Value
   zNum = Delay / 10
   D$ = Str$(zNum)
   If zNum < 1 Then D$ = "0" & D$
   Label2 = D$
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub
