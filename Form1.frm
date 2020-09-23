VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visual Basic 6.0 Project Scanner"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   4860
      TabIndex        =   5
      Top             =   5940
      Width           =   1035
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   330
      Left            =   3720
      TabIndex        =   4
      Top             =   5940
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   180
      TabIndex        =   3
      Top             =   6000
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Settings"
      Height          =   5295
      Left            =   180
      TabIndex        =   1
      Top             =   480
      Width           =   5715
      Begin VB.CheckBox ChVBW 
         Caption         =   "Disable .vbw files"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label Label7 
         Caption         =   "HKEY_CLASSES_ROOT\VisualBasic.ProjectGroup"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   1920
         Width           =   3735
      End
      Begin VB.Label Label6 
         Caption         =   "HKEY_CLASSES_ROOT\VisualBasic.Project"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   1740
         Width           =   3735
      End
      Begin VB.Label Label5 
         Caption         =   $"Form1.frx":0000
         Height          =   1035
         Left            =   300
         TabIndex        =   10
         Top             =   720
         Width           =   5115
      End
      Begin VB.Label Label4 
         Caption         =   "Note:This will only intercept shelled projects, it will not work for projects opened from within VB6 IDE"
         Height          =   435
         Left            =   300
         TabIndex        =   9
         Top             =   2220
         Width           =   5115
      End
      Begin VB.Label Label3 
         Caption         =   "How it works:"
         Height          =   255
         Left            =   300
         TabIndex        =   8
         Top             =   2820
         Width           =   2715
      End
      Begin VB.Label Label2 
         Caption         =   $"Form1.frx":012A
         Height          =   1215
         Left            =   300
         TabIndex        =   7
         Top             =   3120
         Width           =   5115
      End
      Begin VB.Label Label1 
         Caption         =   $"Form1.frx":02C4
         Height          =   615
         Left            =   300
         TabIndex        =   6
         Top             =   4500
         Width           =   5175
      End
   End
   Begin VB.CheckBox ChEnable 
      Caption         =   "Enable Scanning"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1635
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form is a standard "Options" form
'The action happens in the module
Private Sub ChEnable_Click()
    ChVBW.Enabled = ChEnable.Value = 1
    CheckEnabled
End Sub

Private Sub ChVBW_Click()
    CheckEnabled
End Sub

Private Sub cmdApply_Click()
    SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "DisableVBW", ChVBW.Value
    EnableScanning ChEnable.Value
    CheckEnabled
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If cmdApply.Enabled Then cmdApply_Click
    Unload Me
End Sub

Private Sub Form_Load()
    ChEnable.Value = ScannerEnabled
    ChVBW.Value = VBWScanEnabled
    ChVBW.Enabled = ChEnable.Value = 1
    CheckEnabled
End Sub

Public Sub CheckEnabled()
    cmdApply.Enabled = False
    If ChEnable.Value <> ScannerEnabled Then cmdApply.Enabled = True
    If ChVBW.Value <> VBWScanEnabled Then cmdApply.Enabled = True
End Sub
