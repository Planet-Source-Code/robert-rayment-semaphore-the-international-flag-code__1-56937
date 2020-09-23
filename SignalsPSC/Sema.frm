VERSION 5.00
Begin VB.Form frmSema 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000016&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  SEMAPHORE"
   ClientHeight    =   4560
   ClientLeft      =   150
   ClientTop       =   0
   ClientWidth     =   4590
   ControlBox      =   0   'False
   DrawWidth       =   2
   Icon            =   "Sema.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   304
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   306
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkHideText 
      Caption         =   "Hide text"
      Height          =   225
      Left            =   225
      TabIndex        =   13
      Top             =   4155
      Width           =   1200
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      Left            =   2340
      Max             =   50
      Min             =   1
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2055
      Value           =   1
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Height          =   1920
      Left            =   2700
      TabIndex        =   4
      Top             =   45
      Width           =   1650
      Begin VB.CommandButton cmdSSS 
         Caption         =   "CLOSE"
         Height          =   285
         Index           =   3
         Left            =   210
         TabIndex        =   8
         Top             =   1470
         Width           =   1215
      End
      Begin VB.CommandButton cmdSSS 
         Caption         =   "CANCEL"
         Height          =   285
         Index           =   2
         Left            =   195
         TabIndex        =   7
         Top             =   1050
         Width           =   1215
      End
      Begin VB.CommandButton cmdSSS 
         Caption         =   "SEND"
         Height          =   285
         Index           =   1
         Left            =   195
         TabIndex        =   6
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton cmdSSS 
         Caption         =   "ATTENTION"
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
      Height          =   1635
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Text            =   "Sema.frx":0442
      Top             =   2400
      Width           =   4125
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1860
      Left            =   210
      ScaleHeight     =   167.75
      ScaleMode       =   0  'User
      ScaleWidth      =   186
      TabIndex        =   1
      Top             =   135
      Width           =   2355
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   1
         Left            =   405
         Top             =   1620
         Width           =   585
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   405
         Index           =   0
         Left            =   -15
         Top             =   1410
         Width           =   585
      End
   End
   Begin VB.PictureBox picSrc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   20205
      Left            =   5385
      Picture         =   "Sema.frx":0448
      ScaleHeight     =   1345
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   617
      TabIndex        =   0
      Top             =   -5505
      Width           =   9285
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000016&
      Caption         =   "s"
      Height          =   165
      Left            =   4200
      TabIndex        =   12
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
      TabIndex        =   11
      Top             =   2070
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000016&
      Caption         =   "Delay"
      Height          =   240
      Left            =   1800
      TabIndex        =   10
      Top             =   2085
      Width           =   510
   End
End
Attribute VB_Name = "frmSema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmSema (Sema.frm)
' Semaphore by Robert Rayment

Option Explicit

' Semaphore Bitmap
' 44 (4 x 11) images each W 155 x H 122 pixels

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
   Cul = RGB(205, 205, 205)
   N = 28: Disp N  ' Ready/Finish
   With Shape1(0)
      .FillColor = Cul
      .BorderColor = Cul
      .FillStyle = 1
   End With
   With Shape1(1)
      .FillColor = Cul
      .BorderColor = Cul
      .FillStyle = 1
   End With
   chkHide.Value = 1
   HScroll1.Value = 5
   PIC.Width = 157
   PIC.Height = 124
   Show
   Text1.Text = "SEMAPHORE" & vbCrLf & "BY" & vbCrLf & "ROBERT RAYMENT"
End Sub

Private Sub cmdSSS_Click(Index As Integer)
Dim A$
Dim AscCode As Long
Dim k As Long, N As Long
   
   N = 28: Disp N  ' Ready/Finish

   Select Case Index
   Case 0   ' ATTENTION
      aNum = False
      aDone = False
      For k = 1 To 3
         If aDone Then Exit For
         N = 26: Disp N
         N = 30: Disp N
      Next k
      aDone = True
      aNum = False
      N = 28: Disp N  ' Ready/Finish
   
   Case 1   ' SEND
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
            N = 27: Disp N  ' Break
         Case 48 To 57
            If Not aNum Then
               aNum = True
               N = 29: Disp N ' Number
            End If
            If AscCode = 48 Then
               N = 41
            Else
               N = AscCode - 17
            End If
            Disp N     ' Numbers
         
         Case 65 To 90
            If aNum Then
               aNum = False
               N = 42: Disp N ' Alphabet
            End If
            N = AscCode - 65: Disp N   ' Characters
         Case Else
            N = 27: Disp N  ' Break
         End Select
      Next k
      If k = Len(A$) + 1 Then
         N = 28: Disp N  ' Ready/Finish
      End If
   
   Case 2   ' CANCEL
      aDone = True
      aNum = False
      N = 31: Disp N ' Cancel
      
   Case 3   ' Exit
      aDone = True
      Me.Hide
   End Select
End Sub

Private Sub Disp(N As Long)
Dim ix As Long, iy As Long
   If N <> 26 And N <> 30 Then  ' ie Not Attention
      iy = (43 \ 4) * 122
      ix = (43 Mod 4) * 155
      BitBlt PIC.hDC, 0, 0, 155, 122, picSrc.hDC, ix, iy, vbSrcCopy
      PIC.Refresh
      Sleep 60
   End If
   
   iy = (N \ 4) * 122
   ix = (N Mod 4) * 155
   BitBlt PIC.hDC, 0, 0, 155, 122, picSrc.hDC, ix, iy, vbSrcCopy
   PIC.Refresh
   Sleep Delay * 100
End Sub

Private Sub chkHide_Click()
   If chkHide.Value Then
      Shape1(0).FillStyle = 0
      Shape1(1).FillStyle = 0
   Else
      Shape1(0).FillStyle = 1
      Shape1(1).FillStyle = 1
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
