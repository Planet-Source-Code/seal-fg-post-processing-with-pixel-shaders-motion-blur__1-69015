VERSION 5.00
Begin VB.Form wndSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Effect Settings"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2295
   Icon            =   "wndSettigns.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   145
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   153
   Begin VB.CheckBox chkUV 
      Caption         =   "Fix Texture UV Coords"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1935
   End
   Begin VB.CheckBox chkHelp 
      Caption         =   "Display Help"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox chkWalls 
      Caption         =   "Render Walls Mesh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CheckBox chkMotion 
      Caption         =   "Entrie Motion Blur Effect"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "x16 Anisotropic Filtering"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   975
   End
End
Attribute VB_Name = "wndSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub btnCancel_Click()
  
  Unload Me
  
End Sub


Private Sub btnOk_Click()
    
    effFilter = chkFilter.Value
    effMotion = chkMotion.Value
    effUV = chkUV.Value
    shwWalls = chkWalls.Value
    shwHelp = chkHelp.Value
    
    ppScreenQuad.memClear
    If effUV = 1 Then
      ppScreenQuad.objCreate5Tap confDevice.BackBufferWidth, confDevice.BackBufferWidth
    Else
      ppScreenQuad.objCreate
    End If
    
    Unload Me

End Sub


Private Sub Form_Load()
  
  Move wndRender.Left + wndRender.Width - 1000 - wndSettings.Width, wndRender.Top + wndRender.Height - 1000 - wndSettings.Height
 
  chkFilter.Value = effFilter
  chkMotion.Value = effMotion
  chkUV.Value = effUV
  chkWalls.Value = shwWalls
  chkHelp.Value = shwHelp
 
  Show
  
End Sub

