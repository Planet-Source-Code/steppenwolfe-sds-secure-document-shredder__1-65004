VERSION 5.00
Begin VB.Form frmDialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Destroy this File?"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4200
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDialog.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4200
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkOverride 
      Caption         =   "Do not display this dialog again"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   1530
      Width           =   3315
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "Cancel"
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
      Index           =   1
      Left            =   2460
      TabIndex        =   1
      Top             =   900
      Width           =   1185
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "Proceed"
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
      Index           =   0
      Left            =   540
      TabIndex        =   0
      Top             =   900
      Width           =   1185
   End
   Begin VB.Label lblCommon 
      Caption         =   "The Selected File will be Permanently Destroyed! Are you sure you want to Proceed?"
      Height          =   480
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Top             =   330
      Width           =   3570
   End
   Begin VB.Label lblCommon 
      AutoSize        =   -1  'True
      Caption         =   "Do you want to Proceed?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   90
      Width           =   2055
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bQuery            As Boolean
Private m_sRegPath          As String

Public Property Get p_Query() As Boolean
    p_Query = m_bQuery
End Property

Private Sub chkOverride_Click()

    If chkOverride.Value = 1 Then
        With New clsLightning
            .Write_String HKEY_CURRENT_USER, m_sRegPath, "bactdg", "True"
        End With
        frmMain.p_Dialog = True
    Else
        With New clsLightning
            .Write_String HKEY_CURRENT_USER, m_sRegPath, "bactdg", "False"
        End With
        frmMain.p_Dialog = False
    End If
    
End Sub

Private Sub cmdQuery_Click(Index As Integer)
    
    Select Case Index
    Case 0
        m_bQuery = True
    Case 1
        m_bQuery = False
    End Select
    Me.Hide
    
End Sub

Private Sub Form_Load()
    m_sRegPath = "Software\" + App.ProductName
End Sub
