VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About SDS"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3390
      TabIndex        =   1
      Top             =   2700
      Width           =   1125
   End
   Begin VB.TextBox txtAbout 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   180
      Width           =   4395
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Set_About()

Dim sAbout As String

    sAbout = "~ About SDS - Secure Document Shredder ~" & vbCrLf & _
    vbCrLf & "SDS is provided as a free utility for disposing of secure " & _
    "documents in a safe and timely manner." & _
    vbCrLf & vbCrLf & "NSPowertools.com offers several free software components to " & _
    "the public as a service." & _
    vbCrLf & vbCrLf & "If you would like to contact us about our products, email: support@nspowertools.com. " & _
    "For more information on our products or services, feel free to " & _
    "email us, or visit our website at www.nspowertools.com." & _
    vbCrLf & vbCrLf & "Thank you for using NSP SDS.." & _
    vbCrLf & "The NSPowertools team"
        
    txtAbout.Text = ""
    txtAbout.Text = sAbout
        
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set_About
End Sub
