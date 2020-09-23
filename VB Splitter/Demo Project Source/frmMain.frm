VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "VB Splitter Demo"
   ClientHeight    =   6060
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8505
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuDemo 
      Caption         =   "&Demo"
      Begin VB.Menu mnuBasics 
         Caption         =   "&The Basics"
      End
      Begin VB.Menu mnuFeatures 
         Caption         =   "&Features"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private frmBasics As New frmDemoBasics
Private frmDemoFeatures As New frmDemoFeatures

Private Sub mnuBasics_Click()
  frmBasics.Show
  frmBasics.SetFocus
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuFeatures_Click()
  frmDemoFeatures.Show
  frmDemoFeatures.SetFocus
End Sub
