VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   3555
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   435
      Left            =   270
      TabIndex        =   2
      Top             =   1710
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   767
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog dlgDest 
      Left            =   3660
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgSource 
      Left            =   3690
      Top             =   1890
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Execute Package"
      Height          =   435
      Left            =   570
      TabIndex        =   1
      Top             =   900
      Width           =   2505
   End
   Begin MSComDlg.CommonDialog dlgPkg 
      Left            =   3690
      Top             =   1890
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdMakePkg 
      Caption         =   "Create DTS Package"
      Height          =   465
      Left            =   570
      TabIndex        =   0
      Top             =   180
      Width           =   2475
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents pkgClass As clsDTSUpdate

Private Sub cmdMakePkg_Click()

Dim strSourceMDB As String
Dim strDestMDB As String

With dlgSource
  .InitDir = App.Path & "\Database"
  .DialogTitle = "Find Source Database"
  .Filter = "Access Databases (*.mdb)|*.mdb"
  .ShowOpen
End With

If Len(dlgSource.FileName) = 0 Then Exit Sub

strSourceMDB = dlgSource.FileName

With dlgDest
  .InitDir = App.Path & "\Database"
  .DialogTitle = "Find Destination Database"
  .Filter = "Access Databases (*.mdb)|*.mdb"
  .ShowOpen
End With

If Len(dlgDest.FileName) = 0 Then Exit Sub

strDestMDB = dlgDest.FileName

Call pkgClass.TransferVolumeFile(strSourceMDB, strDestMDB, App.Path & "\Database")

End Sub

Private Sub Command3_Click()
  
  ProgressBar1.Value = 0
  ProgressBar1.Visible = True
  
  If pkgClass.ExecutePackage = True Then
    MsgBox "Package Successful"
    ProgressBar1.Visible = False
  Else
    MsgBox "Package Unsuccessful"
    ProgressBar1.Visible = False
  End If
  
End Sub

Private Sub Form_Load()
Set pkgClass = New clsDTSUpdate
ProgressBar1.Visible = False
End Sub



Private Sub pkgClass_ErrorOccurred(ByVal pErr As Long, ByVal pSource As String, ByVal pDescription As String)
  MsgBox "Error Occurred: " & pDescription
End Sub

Private Sub pkgClass_PercentDone(ByVal percent As Integer)
  ProgressBar1.Value = percent
End Sub
