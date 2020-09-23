VERSION 5.00
Object = "*\A..\ActiveX Control Source\VB Splitter.vbp"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.OCX"
Begin VB.Form frmDemoFeatures 
   Caption         =   "Features"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   10575
   WindowState     =   2  'Maximized
   Begin VBSplitter.Splitter Splitter2 
      Height          =   360
      Index           =   1
      Left            =   8550
      TabIndex        =   41
      Top             =   8145
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   635
   End
   Begin VBSplitter.Splitter Splitter2 
      Height          =   285
      Index           =   0
      Left            =   7635
      TabIndex        =   40
      Top             =   8130
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   503
   End
   Begin VB.Timer tmrEvents 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1000
      Left            =   555
      Top             =   7440
   End
   Begin VBSplitter.Splitter Splitter1 
      Height          =   7905
      Left            =   3255
      TabIndex        =   29
      Top             =   45
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   13944
      LiveUpdate      =   0   'False
      MarginLeft      =   3250
      Begin VB.TextBox Text1 
         Height          =   5265
         Left            =   60
         TabIndex        =   39
         Text            =   "TextBox Sample"
         Top             =   75
         Width           =   1335
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   6600
         Left            =   1500
         TabIndex        =   38
         Top             =   60
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   11642
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         FileName        =   "Features.rtf"
         TextRTF         =   $"frmDemoFeatures.frx":0000
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1110
         Left            =   1545
         TabIndex        =   37
         Top             =   6705
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1958
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Column 1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Column 2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Column 3"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Column 4"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Column 5"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   4860
         Left            =   6135
         TabIndex        =   36
         Top             =   1755
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   8573
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         HotTracking     =   -1  'True
         Appearance      =   1
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1605
         Left            =   6120
         Picture         =   "frmDemoFeatures.frx":27CD
         Stretch         =   -1  'True
         Top             =   60
         Width           =   1065
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2265
         Left            =   75
         Picture         =   "frmDemoFeatures.frx":4EC1
         Stretch         =   -1  'True
         Top             =   5475
         Width           =   1305
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   8085
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   3210
      Begin MSComDlg.CommonDialog cdlColor 
         Left            =   45
         Top             =   7530
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.TabStrip tabFeatures 
         Height          =   315
         Left            =   30
         TabIndex        =   3
         Top             =   210
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   556
         TabWidthStyle   =   1
         MultiRow        =   -1  'True
         Style           =   2
         HotTracking     =   -1  'True
         Separators      =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Properties"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Methods"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Events"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame fraFeatures 
         Height          =   7035
         Index           =   1
         Left            =   45
         TabIndex        =   25
         Top             =   510
         Width           =   3030
         Begin VB.CommandButton cmdActivate 
            Caption         =   "Call"
            Height          =   285
            Left            =   2100
            TabIndex        =   47
            Top             =   180
            Width           =   660
         End
         Begin VB.CommandButton cmdMoveSplitter 
            Caption         =   "Call"
            Height          =   285
            Left            =   2100
            TabIndex        =   46
            Top             =   675
            Width           =   660
         End
         Begin VB.TextBox txtMoveTo 
            Height          =   300
            Left            =   1020
            TabIndex        =   44
            Top             =   1380
            Width           =   1725
         End
         Begin VB.Label Label14 
            Caption         =   "Activate"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   225
            TabIndex        =   48
            ToolTipText     =   $"frmDemoFeatures.frx":7A20
            Top             =   255
            Width           =   1590
         End
         Begin VB.Label Label15 
            Caption         =   "MoveSplitter"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   225
            TabIndex        =   45
            ToolTipText     =   "Move a splitter to the specified x- or y- (depending on its Orientation property) coordinate"
            Top             =   720
            Width           =   1590
         End
         Begin VB.Label lblIdSplitter 
            Caption         =   "(click the spliter)"
            Height          =   255
            Left            =   1005
            TabIndex        =   43
            Top             =   1080
            Width           =   1800
         End
         Begin VB.Label Label4 
            Caption         =   "MoveTo:"
            Height          =   255
            Left            =   225
            TabIndex        =   42
            Top             =   1470
            Width           =   810
         End
         Begin VB.Label Label3 
            Caption         =   "IdSplitter:"
            Height          =   255
            Left            =   225
            TabIndex        =   26
            Top             =   1080
            Width           =   705
         End
      End
      Begin VB.Frame fraFeatures 
         Height          =   7035
         Index           =   0
         Left            =   45
         TabIndex        =   1
         Top             =   510
         Width           =   3030
         Begin VB.TextBox txtSize 
            Height          =   285
            Left            =   1545
            TabIndex        =   24
            ToolTipText     =   "Returns/sets the size of the splitters"
            Top             =   4305
            Width           =   1230
         End
         Begin VB.TextBox txtMarginTop 
            Height          =   285
            Left            =   1545
            TabIndex        =   23
            ToolTipText     =   "Returns/sets the top margin of the control from its container"
            Top             =   3915
            Width           =   1230
         End
         Begin VB.TextBox txtMarginRight 
            Height          =   285
            Left            =   1545
            TabIndex        =   22
            ToolTipText     =   "Returns/sets the right margin of the control from its container"
            Top             =   3510
            Width           =   1230
         End
         Begin VB.TextBox txtMarginLeft 
            Height          =   285
            Left            =   1545
            TabIndex        =   21
            ToolTipText     =   "Returns/sets the left margin of the control from its container"
            Top             =   3105
            Width           =   1230
         End
         Begin VB.TextBox txtMarginBottom 
            Height          =   285
            Left            =   1545
            TabIndex        =   20
            ToolTipText     =   "Returns/sets the bottom margin of the control from its container"
            Top             =   2715
            Width           =   1230
         End
         Begin VB.ComboBox cboLiveUpdate 
            Height          =   315
            ItemData        =   "frmDemoFeatures.frx":7AA8
            Left            =   1545
            List            =   "frmDemoFeatures.frx":7AB2
            Style           =   2  'Dropdown List
            TabIndex        =   13
            ToolTipText     =   "Returns/sets a value that determines whether the controls should be resized as the splitter is moved"
            Top             =   2250
            Width           =   1230
         End
         Begin VB.ComboBox cboEnable 
            Height          =   315
            ItemData        =   "frmDemoFeatures.frx":7AC3
            Left            =   1545
            List            =   "frmDemoFeatures.frx":7ACD
            Style           =   2  'Dropdown List
            TabIndex        =   11
            ToolTipText     =   "Returns a value that determines whether the splitters is movable"
            Top             =   1440
            Width           =   1230
         End
         Begin VB.ComboBox cboFillContainer 
            Height          =   315
            ItemData        =   "frmDemoFeatures.frx":7ADE
            Left            =   1545
            List            =   "frmDemoFeatures.frx":7AE8
            Style           =   2  'Dropdown List
            TabIndex        =   9
            ToolTipText     =   $"frmDemoFeatures.frx":7AF9
            Top             =   1845
            Width           =   1230
         End
         Begin VB.ComboBox cboClipCursor 
            Height          =   315
            ItemData        =   "frmDemoFeatures.frx":7B97
            Left            =   1545
            List            =   "frmDemoFeatures.frx":7BA1
            Style           =   2  'Dropdown List
            TabIndex        =   8
            ToolTipText     =   $"frmDemoFeatures.frx":7BB2
            Top             =   1035
            Width           =   1230
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Size:"
            Height          =   195
            Left            =   240
            TabIndex        =   19
            ToolTipText     =   "Returns/sets the size of the splitters"
            Top             =   4380
            Width           =   345
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "MarginTop:"
            Height          =   195
            Left            =   240
            TabIndex        =   18
            ToolTipText     =   "Returns/sets the top margin of the control from its container"
            Top             =   3960
            Width           =   810
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "MarginRight:"
            Height          =   195
            Left            =   240
            TabIndex        =   17
            ToolTipText     =   "Returns/sets the right margin of the control from its container"
            Top             =   3555
            Width           =   900
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "MarginLeft:"
            Height          =   195
            Left            =   240
            TabIndex        =   16
            ToolTipText     =   "Returns/sets the left margin of the control from its container"
            Top             =   3150
            Width           =   795
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "MarginBottom:"
            Height          =   195
            Left            =   240
            TabIndex        =   15
            ToolTipText     =   "Returns/sets the bottom margin of the control from its container"
            Top             =   2745
            Width           =   1020
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Live Update:"
            Height          =   195
            Left            =   240
            TabIndex        =   14
            ToolTipText     =   "Returns/sets a value that determines whether the controls should be resized as the splitter is moved"
            Top             =   2325
            Width           =   915
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Enable:"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            ToolTipText     =   "Returns a value that determines whether the splitters is movable"
            Top             =   1515
            Width           =   540
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fill Container:"
            Height          =   195
            Left            =   240
            TabIndex        =   10
            ToolTipText     =   $"frmDemoFeatures.frx":7C50
            Top             =   1920
            Width           =   945
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Clip Cursor:"
            Height          =   195
            Left            =   255
            TabIndex        =   7
            ToolTipText     =   $"frmDemoFeatures.frx":7CEE
            Top             =   1110
            Width           =   795
         End
         Begin VB.Label lblBackColor 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1545
            TabIndex        =   6
            ToolTipText     =   "Returns/sets the background color used to display the splitters"
            Top             =   615
            Width           =   1215
         End
         Begin VB.Label lblActiveColor 
            BackColor       =   &H00404040&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1545
            TabIndex        =   5
            ToolTipText     =   "Returns/sets the background color used to display the splitter when the user moves it in none live update mode"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Active Color:"
            Height          =   195
            Left            =   240
            TabIndex        =   4
            ToolTipText     =   "Returns/sets the background color used to display the splitter when the user moves it in none live update mode"
            Top             =   285
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Back Color:"
            Height          =   195
            Left            =   240
            TabIndex        =   2
            ToolTipText     =   "Returns/sets the background color used to display the splitters"
            Top             =   690
            Width           =   825
         End
      End
      Begin VB.Frame fraFeatures 
         Height          =   7035
         Index           =   2
         Left            =   45
         TabIndex        =   27
         Top             =   510
         Width           =   3030
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "Moving"
            Height          =   255
            Index           =   6
            Left            =   225
            TabIndex        =   35
            ToolTipText     =   "Occurs when a splitter is being moved by the user"
            Top             =   2520
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "Moved"
            Height          =   255
            Index           =   5
            Left            =   225
            TabIndex        =   34
            ToolTipText     =   "Occurs when the user is finished moving a splitter"
            Top             =   2150
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "MouseUp"
            Height          =   255
            Index           =   4
            Left            =   225
            TabIndex        =   33
            ToolTipText     =   "Occurs when the user releases a mouse button over a splitter without previously moving the splitter"
            Top             =   1780
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "MouseMove"
            Height          =   255
            Index           =   3
            Left            =   225
            TabIndex        =   32
            ToolTipText     =   "Occurs when the user moves the mouse over a splitter without moving the splitter"
            Top             =   1410
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "MouseDown"
            Height          =   255
            Index           =   2
            Left            =   225
            TabIndex        =   31
            ToolTipText     =   "Occurs when the user presses a mouse button over a splitter"
            Top             =   1040
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "DblClick"
            Height          =   255
            Index           =   1
            Left            =   225
            TabIndex        =   30
            ToolTipText     =   "Occurs when the user presses and then realeses a mouse button and then presses and releases it again over a splitter"
            Top             =   670
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "Click"
            Height          =   255
            Index           =   0
            Left            =   225
            TabIndex        =   28
            ToolTipText     =   "Occurs when the user presses and then releases a mouse button over a splitter"
            Top             =   300
            Width           =   2580
         End
      End
   End
End
Attribute VB_Name = "frmDemoFeatures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const conFraProperties = 0
Private Const conFraMethods = 1
Private Const conFraEvents = 2

Private Const conTabProperties = 1
Private Const conTabMethods = 2
Private Const conTabEvents = 3

Private Const conLblEventClick = 0
Private Const conLblEventDblClick = 1
Private Const conLblEventMouseDown = 2
Private Const conLblEventMouseMove = 3
Private Const conLblEventMouseUp = 4
Private Const conLblEventMoved = 5
Private Const conLblEventMoving = 6

Private lngSelectedFrame As Long

Private Sub cboClipCursor_Click()
  Splitter1.ClipCursor = CBool(cboClipCursor)
End Sub

Private Sub cboEnable_Click()
  Splitter1.Enable = CBool(cboEnable)
End Sub

Private Sub cboFillContainer_Click()
  Splitter1.FillContainer = CBool(cboFillContainer)
End Sub

Private Sub cboLiveUpdate_Click()
  Splitter1.LiveUpdate = CBool(cboLiveUpdate)
End Sub

Private Sub ClearEvent(ByVal lngId As Long)
  lblEvents(lngId).BorderStyle = vbBSNone
  lblEvents(lngId).Font.Bold = False
  tmrEvents(lngId).Enabled = False
End Sub

Private Sub cmdActivate_Click()
  On Error GoTo ErrorHandler
  
  Splitter1.Activate
  Exit Sub
  
ErrorHandler:
  ShowErrMessage
End Sub

Private Sub cmdMoveSplitter_Click()
  On Error GoTo ErrorHandler
  
  Splitter1.MoveSplitter IdSplitter:=CLng(lblIdSplitter), _
                         MoveTo:=CLng(txtMoveTo)
  Exit Sub
  
ErrorHandler:
  ShowErrMessage
End Sub

Private Sub Form_Load()
  Dim Index As Long
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim nodi As Node
  Dim nodj As Node

  For i = lblEvents.LBound + 1 To lblEvents.UBound
    Load tmrEvents(i)
  Next
  
  For i = 1 To 10
    Set nodi = TreeView1.Nodes.Add(, , , "Node " & CStr(i))
    For j = 1 To 6
      Set nodj = TreeView1.Nodes.Add(nodi.Index, tvwChild, , _
                                     "Node " & CStr(i) & "." & CStr(j))
      For k = 1 To 3
        TreeView1.Nodes.Add nodj.Index, tvwChild, , _
                            "Node " & CStr(i) & "." & CStr(j) & "." & CStr(k)
      Next
    Next
  Next
    
  For i = 1 To 10
    ListView1.ListItems.Add , , "Item " & CStr(i) & ".1"
    For j = 1 To ListView1.ColumnHeaders.Count - 1
      ListView1.ListItems(i).SubItems(j) = "Item " & CStr(i) & "." & CStr(j + 1)
    Next
  Next
  
  Me.Show
  tabFeatures_Click
End Sub

Private Sub Form_Resize()
  Dim lngNewHeight As Long

  fraMain.Height = Me.ScaleHeight
  lngNewHeight = Me.ScaleHeight - (tabFeatures.Top + tabFeatures.Height)
  If lngNewHeight > 0 Then
    fraFeatures(lngSelectedFrame).Height = lngNewHeight
  Else
    fraFeatures(lngSelectedFrame).Height = 0
  End If
  
  Splitter1.Activate
End Sub

Private Sub HighlightEvent(ByVal lngId As Long)
  If fraFeatures(conFraEvents).Visible Then
    lblEvents(lngId).Font.Bold = True
    lblEvents(lngId).BorderStyle = vbFixedSingle
    tmrEvents(lngId).Enabled = True
  End If
End Sub

Private Sub InitFeatures()
  If lngSelectedFrame = conFraProperties Then
    With Splitter1
      lblActiveColor.BackColor = .ActiveColor
      lblBackColor.BackColor = .BackColor
      cboClipCursor = CStr(.ClipCursor)
      cboEnable = CStr(.Enable)
      cboFillContainer = CStr(.FillContainer)
      cboLiveUpdate = CStr(.LiveUpdate)
      txtMarginBottom = CStr(.MarginBottom)
      txtMarginLeft = CStr(.MarginLeft)
      txtMarginRight = CStr(.MarginRight)
      txtMarginTop = CStr(.MarginTop)
      txtSize = CStr(.Size)
    End With
  End If
End Sub

Private Sub lblActiveColor_Click()
  cdlColor.Flags = cdlCCRGBInit
  cdlColor.Color = lblActiveColor.BackColor
  cdlColor.ShowColor
  lblActiveColor.BackColor = cdlColor.Color
  Splitter1.ActiveColor = lblActiveColor.BackColor
End Sub

Private Sub lblBackColor_Click()
  cdlColor.Flags = cdlCCRGBInit
  cdlColor.Color = lblBackColor.BackColor
  cdlColor.ShowColor
  lblBackColor.BackColor = cdlColor.Color
  Splitter1.BackColor = lblBackColor.BackColor
End Sub

Private Sub ShowErrMessage()
  MsgBox Err.Description, vbCritical + vbOKOnly
End Sub

Private Sub ShowSelectedFrame()
  Dim i As Integer
  
  For i = fraFeatures.LBound To fraFeatures.UBound
    fraFeatures(i).Visible = False
  Next
  fraFeatures(lngSelectedFrame).Height = Me.ScaleHeight - _
                                         (tabFeatures.Top + tabFeatures.Height)
  fraFeatures(lngSelectedFrame).Visible = True
End Sub

Private Sub Splitter1_Click(ByVal IdSplitter As Long)
  HighlightEvent conLblEventClick
  
  If fraFeatures(conFraMethods).Visible Then lblIdSplitter = CStr(IdSplitter)
End Sub

Private Sub Splitter1_DblClick(ByVal IdSplitter As Long)
  HighlightEvent conLblEventDblClick
End Sub

Private Sub Splitter1_MouseDown( _
              ByVal IdSplitter As Long, _
              ByVal Button As Integer, ByVal Shift As Integer, _
              ByVal X As Single, ByVal Y As Single)
  HighlightEvent conLblEventMouseDown
End Sub

Private Sub Splitter1_MouseMove( _
              ByVal IdSplitter As Long, _
              ByVal Button As Integer, ByVal Shift As Integer, _
              ByVal X As Single, ByVal Y As Single)
  HighlightEvent conLblEventMouseMove
End Sub

Private Sub Splitter1_MouseUp( _
              ByVal IdSplitter As Long, _
              ByVal Button As Integer, ByVal Shift As Integer, _
              ByVal X As Single, ByVal Y As Single)
  HighlightEvent conLblEventMouseUp
End Sub

Private Sub Splitter1_Moved( _
              ByVal IdSplitter As Long, ByVal Shift As Integer, _
              ByVal X As Single, ByVal Y As Single)
  HighlightEvent conLblEventMoved
End Sub

Private Sub Splitter1_Moving( _
              ByVal IdSplitter As Long, ByVal Shift As Integer, _
              ByVal X As Single, ByVal Y As Single)
  HighlightEvent conLblEventMoving
End Sub

Private Sub tabFeatures_Click()
  Select Case tabFeatures.SelectedItem.Index
    Case conTabProperties
      lngSelectedFrame = conFraProperties
    Case conTabMethods
      lngSelectedFrame = conFraMethods
    Case conTabEvents
      lngSelectedFrame = conFraEvents
  End Select
  InitFeatures
  ShowSelectedFrame
  
  RichTextBox1.SetFocus
End Sub

Private Sub tmrEvents_Timer(Index As Integer)
  ClearEvent Index
End Sub

Private Sub txtMarginBottom_Validate(Cancel As Boolean)
  Dim lngOldValue As Long
  
  On Error GoTo ErrorHandler

  lngOldValue = Splitter1.MarginBottom
  Splitter1.MarginBottom = CLng(txtMarginBottom)
  Exit Sub

ErrorHandler:
  ShowErrMessage
  txtMarginBottom = lngOldValue
  Cancel = True
End Sub

Private Sub txtMarginLeft_Validate(Cancel As Boolean)
  Dim lngOldValue As Long
  
  On Error GoTo ErrorHandler

  lngOldValue = Splitter1.MarginLeft
  Splitter1.MarginLeft = CLng(txtMarginLeft)
  Exit Sub

ErrorHandler:
  ShowErrMessage
  txtMarginLeft = lngOldValue
  Cancel = True
End Sub

Private Sub txtMarginRight_Validate(Cancel As Boolean)
  Dim lngOldValue As Long
  
  On Error GoTo ErrorHandler

  lngOldValue = Splitter1.MarginRight
  Splitter1.MarginRight = CLng(txtMarginRight)
  Exit Sub

ErrorHandler:
  ShowErrMessage
  txtMarginRight = lngOldValue
  Cancel = True
End Sub

Private Sub txtMarginTop_Validate(Cancel As Boolean)
  Dim lngOldValue As Long
  
  On Error GoTo ErrorHandler

  lngOldValue = Splitter1.MarginTop
  Splitter1.MarginTop = CLng(txtMarginTop)
  Exit Sub

ErrorHandler:
  ShowErrMessage
  txtMarginTop = lngOldValue
  Cancel = True
End Sub

Private Sub txtSize_Validate(Cancel As Boolean)
  Dim lngOldValue As Long
  
  On Error GoTo ErrorHandler

  lngOldValue = Splitter1.Size
  Splitter1.Size = CLng(txtSize)
  txtSize = CStr(Splitter1.Size)
  Exit Sub
  
ErrorHandler:
  ShowErrMessage
  txtSize = lngOldValue
  Splitter1.Size = lngOldValue
  Cancel = True
End Sub

