VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SDS - Secure Document Shredder"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   5730
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fmCommon 
      Caption         =   "Deletion Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   1
      Left            =   90
      TabIndex        =   5
      Top             =   1200
      Width           =   3375
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   60
         ScaleHeight     =   1575
         ScaleWidth      =   3225
         TabIndex        =   6
         Top             =   180
         Width           =   3225
         Begin VB.CheckBox chkAttributes 
            Caption         =   "Reset Attributes"
            Height          =   225
            Left            =   1710
            TabIndex        =   15
            Top             =   690
            Value           =   1  'Checked
            Width           =   1515
         End
         Begin VB.CheckBox chkScatter 
            Caption         =   "Scatter Write"
            Height          =   255
            Left            =   1710
            TabIndex        =   10
            Top             =   360
            Width           =   1365
         End
         Begin VB.OptionButton optPasses 
            Caption         =   "16 Passes"
            Height          =   225
            Index           =   2
            Left            =   60
            TabIndex        =   9
            Top             =   900
            Width           =   1305
         End
         Begin VB.OptionButton optPasses 
            Caption         =   "8 Passes"
            Height          =   225
            Index           =   1
            Left            =   60
            TabIndex        =   8
            Top             =   630
            Width           =   1305
         End
         Begin VB.OptionButton optPasses 
            Caption         =   "4 Passes"
            Height          =   225
            Index           =   0
            Left            =   60
            TabIndex        =   7
            Top             =   360
            Value           =   -1  'True
            Width           =   1305
         End
         Begin MSComctlLib.ProgressBar pbProgress 
            Height          =   165
            Left            =   150
            TabIndex        =   11
            Top             =   1380
            Visible         =   0   'False
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   291
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label lblCommon 
            AutoSize        =   -1  'True
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   2
            Left            =   180
            TabIndex        =   14
            Top             =   1200
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Label lblCommon 
            AutoSize        =   -1  'True
            Caption         =   "Advanced Options"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   1710
            TabIndex        =   13
            Top             =   90
            Width           =   1350
         End
         Begin VB.Label lblCommon 
            AutoSize        =   -1  'True
            Caption         =   "Overwrite Passes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   60
            TabIndex        =   12
            Top             =   90
            Width           =   1320
         End
      End
   End
   Begin VB.Frame fmCommon 
      Caption         =   "File Selection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   180
      Width           =   5565
      Begin VB.TextBox txtFile 
         Height          =   345
         Left            =   150
         TabIndex        =   4
         Top             =   360
         Width           =   4695
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4980
         TabIndex        =   3
         Top             =   360
         Width           =   435
      End
   End
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   3165
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Waiting.."
            TextSave        =   "Waiting.."
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdShred 
      Caption         =   "Shred It!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4230
      TabIndex        =   0
      Top             =   2340
      Width           =   1305
   End
   Begin MSComDlg.CommonDialog cdFile 
      Left            =   2790
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuMain 
      Caption         =   "About"
      Begin VB.Menu mnuSub 
         Caption         =   "About Us"
         Index           =   0
      End
      Begin VB.Menu mnuSub 
         Caption         =   "Visit Us"
         Index           =   1
      End
      Begin VB.Menu mnuSub 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuSub 
         Caption         =   "Exit"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                              ByVal lpOperation As String, _
                                                                              ByVal lpFile As String, _
                                                                              ByVal lpParameters As String, _
                                                                              ByVal lpDirectory As String, _
                                                                              ByVal nShowCmd As Long) As Long

Private WithEvents cShred   As clsShredder
Attribute cShred.VB_VarHelpID = -1
Private m_bDialog           As Boolean
Private m_sRegPath          As String

Public Property Let p_Dialog(PropVal As Boolean)
    m_bDialog = PropVal
End Property

Private Sub cmdSelect_Click()
'/* get the file

On Error GoTo Handler

    With cdFile
        .DialogTitle = "Select a File"
        .CancelError = True
        .DefaultExt = ".txt"
        .InitDir = Left$(App.Path, 3)
        .ShowOpen
        txtFile.Text = .FileName
    End With

    If Len(txtFile.Text) > 0 Then
        stBar.Panels(1).Text = "File: " + txtFile.Text + " selected.."
    End If
    
Handler:

End Sub

Private Sub cmdShred_Click()
'/* shred it

Dim sPath       As String
Dim lPasses     As Long

    Select Case True
    Case optPasses(0).Value
        lPasses = 4
    Case optPasses(1).Value
        lPasses = 8
    Case optPasses(2).Value
        lPasses = 16
    End Select
    
    '/* user proofing/warning
    sPath = txtFile.Text
    If Len(sPath) = 0 Then
        MsgBox "Please supply a file name before proceeding!", _
        vbInformation, "No File Selected!"
        Exit Sub
    ElseIf InStr(1, sPath, ".") = 0 Then
        MsgBox "File name is invalid! Please select a proper file before proceeding!", _
        vbInformation, "Invalid Input!"
        Exit Sub
    ElseIf Not cShred.File_Exists(sPath) Then
        MsgBox "File does not exist! Please select a proper file before proceeding!", _
        vbInformation, "Invalid Input!"
        Exit Sub
    Else
        If Not m_bDialog Then
            Dim bRes As Integer
            frmDialog.Show vbModal, Me
            If Not frmDialog.p_Query Then
                stBar.Panels(1).Text = "Deletion has been Aborted.."
                Exit Sub
            End If
        End If
    End If
    
    stBar.Panels(1).Text = "Building Encryption Blocks.."
    '/* call into class
    With cShred
        '/* reset attributes to normal
        .p_Attributes = (chkAttributes.Value = 1)
        '/* number of passes
        .p_Passes = lPasses
        '/* file path
        .p_SourceFile = sPath
        '/* core
        .File_Shred
    End With

End Sub

Private Sub cShred_eSCompComplete()
'/* document destroyed

Dim sFile   As String
    
    sFile = Mid$(txtFile.Text, InStrRev(txtFile.Text, Chr$(92)) + 1)
    pbProgress.Visible = False
    lblCommon(2).Visible = False
    stBar.Panels(1).Text = "File: " + sFile + " has been destroyed!"
    txtFile.Text = ""
    
End Sub

Private Sub cShred_eSCompPMax(lMax As Long)
'/* progress max

    pbProgress.Max = lMax
    pbProgress.Visible = True
    lblCommon(2).Visible = True

End Sub

Private Sub cShred_eSCompPTick(lCnt As Long)
'/* progress tick

Dim lTick   As Long

    lTick = (lCnt / pbProgress.Max) * 100
    pbProgress.Value = lCnt
    lblCommon(2).Caption = CStr(lTick) + "%"
    stBar.Panels(1).Text = "Completed pass " + CStr(lCnt) + " of " + CStr(pbProgress.Max)
    
End Sub

Private Sub mnuSub_Click(Index As Integer)
'/* menu calls

    Select Case Index
    Case 0
        frmAbout.Show vbModal, Me
    Case 1
        Site_Launch
    Case 3
        Unload Me
    End Select
    
End Sub

Private Sub Site_Launch()
'/* launch a url

Dim sLink As String

    sLink = "http://www.planetsourcecode.com"
    ShellExecute Me.hwnd, "open", sLink, 0&, 0&, &H1


End Sub
 
Private Sub Form_Load()

    Set cShred = New clsShredder
    
    '/* get worning dialog show status
    m_sRegPath = "Software\" + App.ProductName

    With New clsLightning
        If .Key_Exist(HKEY_CURRENT_USER, m_sRegPath) Then
            m_bDialog = CBool(.Read_String(HKEY_CURRENT_USER, m_sRegPath, "bactdg"))
        End If
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set cShred = Nothing

End Sub
