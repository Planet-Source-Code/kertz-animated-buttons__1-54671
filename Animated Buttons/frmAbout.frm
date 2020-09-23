VERSION 5.00
Object = "{34F681D0-3640-11CF-9294-00AA00B8A733}#1.0#0"; "danim.dll"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "About Monad"
   ClientHeight    =   5850
   ClientLeft      =   2295
   ClientTop       =   1605
   ClientWidth     =   5865
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmAbout.frx":0442
   ScaleHeight     =   4037.773
   ScaleMode       =   0  'User
   ScaleWidth      =   5507.539
   ShowInTaskbar   =   0   'False
   Begin VB.Label Lblexit 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5280
      TabIndex        =   8
      Top             =   360
      Width           =   135
   End
   Begin DirectAnimationCtl.DAViewerControl DAV2 
      Height          =   810
      Left            =   3960
      TabIndex        =   6
      ToolTipText     =   "This is Animated Button 2"
      Top             =   3360
      Visible         =   0   'False
      Width           =   810
      OpaqueForHitDetect=   -1  'True
      UpdateInterval  =   0.033
   End
   Begin VB.Label lblcaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Animated Buttons!"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   1200
      TabIndex        =   7
      Top             =   360
      Width           =   3855
   End
   Begin DirectAnimationCtl.DAViewerControl DAV1 
      Height          =   810
      Left            =   2400
      TabIndex        =   5
      ToolTipText     =   "This is Animated Button 2"
      Top             =   3360
      Visible         =   0   'False
      Width           =   810
      OpaqueForHitDetect=   -1  'True
      UpdateInterval  =   0.033
   End
   Begin DirectAnimationCtl.DAViewerControl DAV 
      Height          =   810
      Left            =   840
      TabIndex        =   4
      ToolTipText     =   "This is Animated button 1..."
      Top             =   3360
      Visible         =   0   'False
      Width           =   810
      OpaqueForHitDetect=   -1  'True
      UpdateInterval  =   0.033
   End
   Begin VB.Label lbl5 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":1EB9
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   4935
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2004 by Amal R S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   4815
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "Created By :- R S Amal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Version - 1.0.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
   Begin MSForms.Image Img 
      Height          =   810
      Left            =   840
      Top             =   3360
      Width           =   810
      AutoSize        =   -1  'True
      BorderStyle     =   0
      SizeMode        =   1
      SpecialEffect   =   6
      Size            =   "1429;1429"
      Picture         =   "frmAbout.frx":1F4B
      PictureAlignment=   4
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Img1 
      Height          =   825
      Left            =   2400
      Top             =   3360
      Width           =   825
      AutoSize        =   -1  'True
      BorderStyle     =   0
      SizeMode        =   1
      SpecialEffect   =   6
      Size            =   "1455;1455"
      Picture         =   "frmAbout.frx":74D5
      PictureAlignment=   4
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Img2 
      Height          =   810
      Left            =   3960
      Top             =   3360
      Width           =   810
      AutoSize        =   -1  'True
      BorderStyle     =   0
      SizeMode        =   1
      SpecialEffect   =   6
      Size            =   "1429;1429"
      Picture         =   "frmAbout.frx":8870
      PictureAlignment=   4
      VariousPropertyBits=   19
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ipath As String
Dim currentanimation, otheranimation As Control

Private Sub DAV_Click()
MsgBox "Wow you clicked it!!!. Do you like it? It's not free!(just kidding). If you vote you can have it!", vbExclamation, "Animated Button - Created by Amal R S"
End Sub

Private Sub DAV1_Click()
MsgBox "Wow you clicked it again!!!. Do you like it? It's not free!(just kidding). If you vote you can have it!", vbExclamation, "Animated Button - Created by Amal R S"
End Sub

Private Sub DAV2_Click()
MsgBox "Wow you clicked it again and again!!!.But this time good bye!", vbExclamation, "Aniamated Buttons"
End
End Sub

'***********************
Private Sub Form_Load()
'remeber the visiblity property of DAV,DAV1,DAV2 must be False
'do not delete or modify the animation files in the application folder
On Error GoTo errh
'set path
ipath = App.Path
'import animation
DAV.Image = ImportImage(ipath & "\TECH-EYE.gif")
DAV1.Image = ImportImage(ipath & "\RUN-MAN.gif")
DAV2.Image = ImportImage(ipath & "\E-MAIL12.gif")
'defining control size
DAV.Width = Img.Width
DAV1.Width = Img1.Width
DAV2.Width = Img2.Width
DAV.Height = Img.Height
DAV1.Height = Img1.Height
DAV2.Height = Img2.Height
DAV.Top = Img.Top
DAV1.Top = Img1.Top
DAV2.Top = Img2.Top
DAV.Left = Img.Left
DAV1.Left = Img1.Left
DAV2.Left = Img2.Left
'start animation
DAV.Start
DAV1.Start
DAV2.Start
errh:
If Err.Number <> 0 Then
MsgBox Err.Description, vbCritical, Form1.Caption
End
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'hide animation
DAV.Visible = False
DAV1.Visible = False
DAV2.Visible = False

End Sub

Private Sub Img_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'show animation
DAV.Visible = True
End Sub


Private Sub Img1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'show animation
DAV1.Visible = True
End Sub

Private Sub Img2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'show animation
DAV2.Visible = True
End Sub

Private Sub Lblexit_Click()
MsgBox "If you like it please vote for me...", vbInformation, "animated buttons"
End
End Sub
Private Sub lbl2_Click()
'hide animation
DAV.Visible = False
DAV1.Visible = False
DAV2.Visible = False
End Sub

Private Sub lbl3_Click()
'hide animation
DAV.Visible = False
DAV1.Visible = False
DAV2.Visible = False
End Sub



Private Sub lbl5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'hide animation
DAV.Visible = False
DAV1.Visible = False
DAV2.Visible = False
End Sub

Private Sub lblcaption_Click()
'hide animation
DAV.Visible = False
DAV1.Visible = False
DAV2.Visible = False
End Sub
