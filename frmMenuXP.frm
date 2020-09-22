VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmMenuXP 
   AutoRedraw      =   -1  'True
   Caption         =   " Test MenuXP ..."
   ClientHeight    =   3495
   ClientLeft      =   2250
   ClientTop       =   1605
   ClientWidth     =   6555
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMenuXP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3495
   ScaleWidth      =   6555
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Please, right click on the form area to show menu ..."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   15
      TabIndex        =   0
      Top             =   2955
      Width           =   6525
   End
   Begin ComctlLib.ImageList imgMain 
      Left            =   5985
      Top             =   2250
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuXP.frx":1042
            Key             =   "save"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuXP.frx":1584
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuXP.frx":1AC6
            Key             =   "prv"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMenuXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Call CHookWnd(Me.hwnd, True)
    
    Call CInitMenu
    Call CSetupMenu(Me.hwnd)

    Me.Show

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  Dim pt As POINTAPI
    
    If Button = vbRightButton Then
     
        pt.x = Me.ScaleX(x, vbTwips, vbPixels)
        pt.y = Me.ScaleY(y, vbTwips, vbPixels)
        ClientToScreen Me.hwnd, pt
        
        'Me.Print Caps(1, 1), Caps(1, 5)
        TrackPopupMenuEx Caps(1, 5), TPM_LEFTALIGN Or TPM_TOPALIGN Or TPM_LEFTBUTTON, pt.x, pt.y, Me.hwnd, ByVal 0&
    
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CHookWnd(Me.hwnd, False)
End Sub
