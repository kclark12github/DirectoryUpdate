VERSION 5.00
Begin VB.Form frmPreferences 
   Caption         =   "Preferences"
   ClientHeight    =   2172
   ClientLeft      =   72
   ClientTop       =   360
   ClientWidth     =   6516
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2172
   ScaleWidth      =   6516
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   432
      Left            =   3540
      TabIndex        =   9
      Top             =   1500
      Width           =   1272
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   432
      Left            =   1920
      TabIndex        =   8
      Top             =   1500
      Width           =   1272
   End
   Begin VB.TextBox txtLogFile 
      Height          =   288
      Left            =   2100
      TabIndex        =   6
      Top             =   1020
      Width           =   4212
   End
   Begin VB.TextBox txtDateModified 
      Height          =   288
      Left            =   2100
      TabIndex        =   4
      Top             =   720
      Width           =   2112
   End
   Begin VB.TextBox txtTarget 
      Height          =   288
      Left            =   2100
      TabIndex        =   2
      Top             =   420
      Width           =   4212
   End
   Begin VB.TextBox txtSource 
      Height          =   288
      Left            =   2100
      TabIndex        =   0
      Top             =   120
      Width           =   4212
   End
   Begin VB.Label lblLogFile 
      AutoSize        =   -1  'True
      Caption         =   "Secondary &Log File:"
      Height          =   192
      Left            =   540
      TabIndex        =   7
      Top             =   1080
      Width           =   1440
   End
   Begin VB.Label lblDateModified 
      AutoSize        =   -1  'True
      Caption         =   "Default &Date Modified:"
      Height          =   192
      Left            =   420
      TabIndex        =   5
      Top             =   780
      Width           =   1584
   End
   Begin VB.Label lblTarget 
      AutoSize        =   -1  'True
      Caption         =   "Default &Target Directory:"
      Height          =   192
      Left            =   264
      TabIndex        =   3
      Top             =   480
      Width           =   1740
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      Caption         =   "Default &Source Directory:"
      Height          =   192
      Left            =   228
      TabIndex        =   1
      Top             =   180
      Width           =   1776
   End
End
Attribute VB_Name = "frmPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
   Me.Hide
End Sub

Private Sub cmdOK_Click()
   SaveSetting frmMain.AppName, "Preferences", "Source", txtSource.Text
   SaveSetting frmMain.AppName, "Preferences", "Target", txtTarget.Text
   SaveSetting frmMain.AppName, "Preferences", "DateModified", txtDateModified.Text
   SaveSetting frmMain.AppName, "Preferences", "Log", txtLogFile.Text
   Me.Hide
End Sub

Private Sub Form_Load()
   txtDateModified.Text = GetSetting(frmMain.AppName, "Preferences", "DateModified", Now())
   txtSource.Text = GetSetting(frmMain.AppName, "Preferences", "Source", "")
   txtTarget.Text = GetSetting(frmMain.AppName, "Preferences", "Target", "")
   txtLogFile.Text = GetSetting(frmMain.AppName, "Preferences", "Log", "C:\" & frmMain.AppName & ".log")
End Sub
