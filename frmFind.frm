VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4395
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2475
      TabIndex        =   7
      Top             =   1755
      Width           =   1230
   End
   Begin VB.CheckBox chkCaseSensitive 
      Caption         =   "Case Sensitive"
      Height          =   195
      Left            =   1170
      TabIndex        =   5
      Top             =   1350
      Width           =   1500
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Start"
      Height          =   375
      Left            =   1170
      TabIndex        =   4
      Top             =   1755
      Width           =   1230
   End
   Begin VB.TextBox txtFilemask 
      Height          =   330
      Left            =   1125
      TabIndex        =   3
      Text            =   "*.bas;*.frm;*.ctl"
      Top             =   900
      Width           =   3075
   End
   Begin VB.TextBox txtSearchString 
      Height          =   375
      Left            =   1125
      TabIndex        =   0
      Top             =   405
      Width           =   3075
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seperate Multiple Search Items With A Dash ( - )"
      Height          =   240
      Index           =   2
      Left            =   405
      TabIndex        =   6
      Top             =   90
      Width           =   4110
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Text:"
      Height          =   240
      Index           =   1
      Left            =   90
      TabIndex        =   2
      Top             =   495
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File Mask:"
      Height          =   240
      Index           =   0
      Left            =   270
      TabIndex        =   1
      Top             =   945
      Width           =   825
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSearch_Click()

    If Len(txtSearchString.Text) > 0 Then
        Me.Hide   'hide this form
        DoEvents
        With frmMain
            .SearchString = txtSearchString.Text                     'located in frmMain declarations
            .SearchMethod = IIf(chkCaseSensitive.Value = 1, 0, 1)  'located in frmMain declarations
            .FileMask = txtFilemask.Text                           'located in frmMain declarations
            .cmdStartSearch.Enabled = False
            .cmdCancelSearch.Enabled = True
            .StartSearch                                           'public Sub (method) In frmMain
        End With
    Else
        MsgBox "Invalid Search String", vbCritical
    End If

End Sub

Private Sub cmdCancel_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    txtSearchString.Text = GetSetting(App.Title, "Settings", "LastSearch", "")
    txtFilemask.Text = GetSetting(App.Title, "Settings", "LastMask", txtFilemask.Text)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    SaveSetting App.Title, "Settings", "LastSearch", txtSearchString.Text
    SaveSetting App.Title, "Settings", "LastMask", txtFilemask.Text

End Sub

