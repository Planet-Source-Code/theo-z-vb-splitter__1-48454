VERSION 5.00
Object = "*\A..\ActiveX Control Source\VB Splitter.vbp"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDemoBasics 
   Caption         =   "The Basics"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   10380
   WindowState     =   2  'Maximized
   Begin VBSplitter.Splitter Splitter1 
      Height          =   7035
      Left            =   225
      TabIndex        =   0
      Top             =   240
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   12409
      Begin VB.PictureBox Picture20 
         Height          =   4645
         Left            =   8956
         ScaleHeight     =   4590
         ScaleWidth      =   405
         TabIndex        =   21
         Top             =   484
         Width           =   472
      End
      Begin VB.PictureBox Picture19 
         Height          =   5129
         Left            =   9488
         ScaleHeight     =   5070
         ScaleWidth      =   315
         TabIndex        =   20
         Top             =   0
         Width           =   382
      End
      Begin VB.PictureBox Picture18 
         Height          =   424
         Left            =   8660
         ScaleHeight     =   360
         ScaleWidth      =   705
         TabIndex        =   19
         Top             =   0
         Width           =   768
      End
      Begin VB.PictureBox Picture17 
         Height          =   5421
         Left            =   8660
         ScaleHeight     =   5355
         ScaleWidth      =   180
         TabIndex        =   18
         Top             =   484
         Width           =   236
      End
      Begin VB.PictureBox Picture16 
         Height          =   716
         Left            =   8956
         ScaleHeight     =   660
         ScaleWidth      =   855
         TabIndex        =   17
         Top             =   5189
         Width           =   914
      End
      Begin VB.PictureBox Picture15 
         Height          =   319
         Left            =   1889
         ScaleHeight     =   255
         ScaleWidth      =   7380
         TabIndex        =   16
         Top             =   6307
         Width           =   7437
      End
      Begin VB.PictureBox Picture14 
         Height          =   349
         Left            =   1889
         ScaleHeight     =   285
         ScaleWidth      =   7920
         TabIndex        =   15
         Top             =   6686
         Width           =   7981
      End
      Begin VB.PictureBox Picture13 
         Height          =   661
         Left            =   9386
         ScaleHeight     =   600
         ScaleWidth      =   420
         TabIndex        =   14
         Top             =   5965
         Width           =   484
      End
      Begin VB.PictureBox Picture12 
         Height          =   282
         Left            =   1361
         ScaleHeight     =   225
         ScaleWidth      =   7905
         TabIndex        =   13
         Top             =   5965
         Width           =   7965
      End
      Begin VB.PictureBox Picture11 
         Height          =   728
         Left            =   1361
         ScaleHeight     =   675
         ScaleWidth      =   405
         TabIndex        =   12
         Top             =   6307
         Width           =   468
      End
      Begin VB.PictureBox Picture10 
         Height          =   4720
         Left            =   496
         ScaleHeight     =   4665
         ScaleWidth      =   285
         TabIndex        =   11
         Top             =   1861
         Width           =   346
      End
      Begin VB.PictureBox Picture9 
         Height          =   5628
         Left            =   902
         ScaleHeight     =   5565
         ScaleWidth      =   345
         TabIndex        =   10
         Top             =   953
         Width           =   399
      End
      Begin VB.PictureBox Picture8 
         Height          =   394
         Left            =   496
         ScaleHeight     =   330
         ScaleWidth      =   750
         TabIndex        =   9
         Top             =   6641
         Width           =   805
      End
      Begin VB.PictureBox Picture7 
         Height          =   5174
         Left            =   0
         ScaleHeight     =   5115
         ScaleWidth      =   375
         TabIndex        =   8
         Top             =   1861
         Width           =   436
      End
      Begin VB.PictureBox Picture6 
         Height          =   848
         Left            =   0
         ScaleHeight     =   795
         ScaleWidth      =   780
         TabIndex        =   7
         Top             =   953
         Width           =   842
      End
      Begin VB.PictureBox Picture5 
         Height          =   311
         Left            =   405
         ScaleHeight     =   255
         ScaleWidth      =   7590
         TabIndex        =   6
         Top             =   281
         Width           =   7645
      End
      Begin VB.PictureBox Picture4 
         Height          =   592
         Left            =   8110
         ScaleHeight     =   525
         ScaleWidth      =   435
         TabIndex        =   5
         Top             =   0
         Width           =   490
      End
      Begin VB.PictureBox Picture3 
         Height          =   221
         Left            =   0
         ScaleHeight     =   165
         ScaleWidth      =   7995
         TabIndex        =   4
         Top             =   0
         Width           =   8050
      End
      Begin VB.PictureBox Picture2 
         Height          =   612
         Left            =   0
         ScaleHeight     =   555
         ScaleWidth      =   285
         TabIndex        =   3
         Top             =   281
         Width           =   345
      End
      Begin VB.PictureBox Picture1 
         Height          =   241
         Left            =   405
         ScaleHeight     =   180
         ScaleWidth      =   8130
         TabIndex        =   2
         Top             =   652
         Width           =   8195
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   4952
         Left            =   1361
         TabIndex        =   1
         Top             =   953
         Width           =   7239
         _ExtentX        =   12779
         _ExtentY        =   8731
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         FileName        =   "The Basics.rtf"
         TextRTF         =   $"frmDemoBasics.frx":0000
      End
   End
End
Attribute VB_Name = "frmDemoBasics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
  Splitter1.Activate
End Sub
