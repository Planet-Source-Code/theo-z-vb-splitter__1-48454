VERSION 5.00
Begin VB.UserControl Splitter 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   FillColor       =   &H00404040&
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ctlSplitter.ctx":0000
   Begin VB.PictureBox picSplitter 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   3390
      Index           =   9999
      Left            =   105
      ScaleHeight     =   3390
      ScaleWidth      =   180
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   180
   End
End
Attribute VB_Name = "Splitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An ActiveX control to allow the user to resize docked controls at run time"
'*******************************************************************************
'** File Name  : ctlSplitter.ctl                                              **
'** Language   : Visual Basic 6.0                                             **
'** Author     : Theo Zacharias (theo_yz@yahoo.com)                           **
'** Description: An ActiveX control to allow the user to resize docked        **
'**              controls at run-time                                         **
'** Members    :                                                              **
'**   * Properties : ActivateColor (r/w), BackColor (r/w), ClipCursor (r/w),  **
'**                  Enable (r/w), FillContainer (r/w), LiveUpdate (r/w),     **
'**                  MarginBottom (r/w), MarginLeft (r/w), MarginRight (r/w), **
'**                  MarginTop (r/w), Size (r/w)                              **
'**   * Methods    : Activate, MoveSplitters                                  **
'**   * Events     : Click, DblClick, MouseDown, MouseMove, MouseUp, Moved,   **
'**                  Moving                                                   **
'** Last modified on September 12, 2003                                       **
'*******************************************************************************

Option Explicit

'--- Property Variables

' A collection to represents virtual controls and splitters
Private mControls As clsControls
Private mSplitters As clsSplitters

' Property variables which appears in the property page
Private mlngActiveColor As OLE_COLOR
Private mlngBackColor As OLE_COLOR
Private mblnFillContainer As Boolean
Private mlngMarginBottom As Long
Private mlngMarginLeft As Long
Private mlngMarginRight As Long
Private mlngMarginTop As Long

'--- PropBag Names
Private Const mconActiveColor As String = "ActiveColor"
Private Const mconBackColor As String = "BackColor"
Private Const mconClipCursor As String = "ClipCursor"
Private Const mconEnable As String = "Enable"
Private Const mconFillContainer As String = "FillContainer"
Private Const mconLiveUpdate As String = "LiveUpdate"
Private Const mconMarginBottom As String = "MarginBottom"
Private Const mconMarginLeft As String = "MarginLeft"
Private Const mconMarginRight As String = "MarginRight"
Private Const mconMarginTop As String = "MarginTop"
Private Const mconSize As String = "Size"

'--- Property Default Values
Private Const mconDefaultActiveColor As Long = vbBlack
Private Const mconDefaultFillContainer As Boolean = True
Private Const mconDefaultMarginBottom As Long = 0
Private Const mconDefaultMarginLeft As Long = 0
Private Const mconDefaultMarginRight As Long = 0
Private Const mconDefaultMarginTop As Long = 0

'--- Other Variables
Private mblnDrag As Boolean                         'indicating whether the user
                                                    '   is dragging the splitter
Private mblnSplitterMoved As Boolean          'indicating whether a splitter has
                                              '      just been moved by the user
Private mblnVisibleSave As Boolean              'to restore the Visible property
                                                '      of the control's instance
Private mlngDragStart As Long    'the x- or y- coordinate (depends on the active
                                 '                Splitter 's orientation) where
                                 '                    the user strats to drag it
Private mposPrev As mdlAPI.POINTAPI  'previous mouse pointer coordinate relative
                                     '   to the splitter (note: this variable is
                                     '        used to make sure the custom event
                                     '                  MouseMove works properly
                                     
'-------------------------------
' ActiveX Control Custom Events
'-------------------------------

'Description: Occurs when the user presses and then realeses a mouse button over
'             a splitter
'Argument   : IdSplitter (a value that uniquely identifies a splitter)
Public Event Click(ByVal IdSplitter As Long)

'Description: Occurs when the user presses and then realeses a mouse button and
'             then presses and releases it again over a splitter
'Argument   : IdSplitter (a value that uniquely identifies a splitter)
Public Event DblClick(ByVal IdSplitter As Long)

'Description: Occurs when the user presses a mouse button over a splitter
'Arguments  : IdSplitter, Button, Shift, X, Y (see reference for MouseDown event in
'             MSDN for the description of the arguments)
Public Event MouseDown(ByVal IdSplitter As Long, _
                       ByVal Button As Integer, ByVal Shift As Integer, _
                       ByVal x As Single, ByVal y As Single)

'Description: Occurs when the user moves the mouse over a splitter without
'             moving the splitter
'Arguments  : IdSplitter, Button, Shift, X, Y (see reference for MouseMove event in
'             MSDN for the description of the arguments)
Public Event MouseMove(ByVal IdSplitter As Long, _
                       ByVal Button As Integer, ByVal Shift As Integer, _
                       ByVal x As Single, ByVal y As Single)

'Description: Occurs when the user releases a mouse button over a splitter
'             without previously moving the splitter
'Arguments  : IdSplitter, Button, Shift, X, Y (see reference for MouseUp event in
'             MSDN for the description of the arguments)
Public Event MouseUp(ByVal IdSplitter As Long, _
                     ByVal Button As Integer, ByVal Shift As Integer, _
                     ByVal x As Single, ByVal y As Single)

'Description: Occurs when the user is finished moving a splitter
'Arguments  : IdSplitter, Button, Shift, X, Y (see reference for MouseUp event in
'             MSDN for the description of the arguments)
Public Event Moved(ByVal IdSplitter As Long, ByVal Shift As Integer, _
                   ByVal x As Single, ByVal y As Single)

'Description: Occurs when a splitter is being moved by the user
'Arguments  : IdSplitter, Button, Shift, X, Y (see reference for MouseMove event in
'             MSDN for the description of the arguments)
Public Event Moving(ByVal IdSplitter As Long, ByVal Shift As Integer, _
                    ByVal x As Single, ByVal y As Single)

'--------------------------------------------
' ActiveX Control Constructor and Destructor
'--------------------------------------------

Private Sub UserControl_Initialize()
  Set mControls = New clsControls
  Set mSplitters = New clsSplitters
End Sub

Private Sub UserControl_Terminate()
  Set mControls = Nothing
  Set mSplitters = Nothing
End Sub

'-----------------------------------
' ActiveX Control Properties Events
'-----------------------------------

Private Sub UserControl_InitProperties()
  mlngActiveColor = mconDefaultActiveColor
  mlngBackColor = Ambient.BackColor
  mblnFillContainer = mconDefaultFillContainer
  mlngMarginBottom = mconDefaultMarginBottom
  mlngMarginLeft = mconDefaultMarginLeft
  mlngMarginRight = mconDefaultMarginRight
  mlngMarginTop = mconDefaultMarginTop
  
  mlngBackColor = Ambient.BackColor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
    mlngActiveColor = .ReadProperty(Name:=mconActiveColor, _
                                    DefaultValue:=mconDefaultActiveColor)
    mlngBackColor = .ReadProperty(Name:=mconBackColor, _
                                  DefaultValue:=Ambient.BackColor)
    mSplitters.ClipCursor = _
      .ReadProperty(Name:=mconClipCursor, _
                    DefaultValue:=mSplitters.DefaultClipCursor)
    mSplitters.Enable = .ReadProperty(Name:=mconEnable, _
                                      DefaultValue:=mSplitters.DefaultEnable)
    mblnFillContainer = .ReadProperty(Name:=mconFillContainer, _
                                      DefaultValue:=mconDefaultFillContainer)
    mSplitters.LiveUpdate = _
      .ReadProperty(Name:=mconLiveUpdate, _
                    DefaultValue:=mSplitters.DefaultLiveUpdate)
    mlngMarginBottom = .ReadProperty(Name:=mconMarginBottom, _
                                     DefaultValue:=mconDefaultMarginBottom)
    mlngMarginLeft = .ReadProperty(Name:=mconMarginLeft, _
                                   DefaultValue:=mconDefaultMarginLeft)
    mlngMarginRight = .ReadProperty(Name:=mconMarginRight, _
                                    DefaultValue:=mconDefaultMarginRight)
    mlngMarginTop = .ReadProperty(Name:=mconMarginTop, _
                                  DefaultValue:=mconDefaultMarginTop)
    mSplitters.Size = .ReadProperty(Name:=mconSize, _
                                    DefaultValue:=mSplitters.DefaultSize)
  End With
  
  gstrControlName = Ambient.DisplayName
  
  ' Hide the ActiveX control when initializing the controls in it to reduce the
  '   flickering
  If Ambient.UserMode Then
    mblnVisibleSave = Extender.Visible
    Extender.Visible = False
  End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    .WriteProperty Name:=mconActiveColor, _
                   Value:=mlngActiveColor, DefaultValue:=mconDefaultActiveColor
    .WriteProperty Name:=mconBackColor, Value:=mlngBackColor, _
                   DefaultValue:=Ambient.BackColor
    .WriteProperty Name:=mconClipCursor, Value:=mSplitters.ClipCursor, _
                   DefaultValue:=mSplitters.DefaultClipCursor
    .WriteProperty Name:=mconEnable, Value:=mSplitters.Enable, _
                   DefaultValue:=mSplitters.DefaultEnable
    .WriteProperty Name:=mconFillContainer, Value:=mblnFillContainer, _
                   DefaultValue:=mconDefaultFillContainer
    .WriteProperty Name:=mconLiveUpdate, Value:=mSplitters.LiveUpdate, _
                   DefaultValue:=mSplitters.DefaultLiveUpdate
    .WriteProperty Name:=mconMarginBottom, Value:=mlngMarginBottom, _
                   DefaultValue:=mconDefaultMarginBottom
    .WriteProperty Name:=mconMarginLeft, Value:=mlngMarginLeft, _
                   DefaultValue:=mconDefaultMarginLeft
    .WriteProperty Name:=mconMarginRight, Value:=mlngMarginRight, _
                   DefaultValue:=mconDefaultMarginRight
    .WriteProperty Name:=mconMarginTop, Value:=mlngMarginTop, _
                   DefaultValue:=mconDefaultMarginTop
    .WriteProperty Name:=mconSize, Value:=mSplitters.Size, _
                   DefaultValue:=mSplitters.DefaultSize
  End With
End Sub

'-------------------------------
' Others ActiveX Control Events
'-------------------------------

' Purpose    : Raises custom event Click
' Effect     : As specified
Private Sub picSplitter_Click(Index As Integer)
  RaiseEvent Click(CLng(Index))
End Sub

' Purpose    : Raises custom event DblClick
' Effect     : As specified
Private Sub picSplitter_DblClick(Index As Integer)
  RaiseEvent DblClick(CLng(Index))
End Sub

' Purpose    : Initializes all things needed to move the splitter at run-time
'              and raises custom event MouseDown
' Assumption : Picture Box control picSplitter(Index) which represents the
'              splitter exits
' Effects    : * mblnDrag = true
'              * mlngDragStart = x or y (see the codes)
'              * Control picSplitter(Index) is in front of the other controls
'              * If the splitter's LiveUpdate property is false, then the
'                picSpliter(Index) BackColor property has been set to the
'                splitter's ActiveColor property
'              * If the splitter's ClipCursor property is true, then the mouse
'                pointer has been confined based on the splitter's MinXc, MinYc,
'                MaxXc and MaxYc property value
'              * Custom event MouseDown has been raised
' Inputs     : Index, Button, Shift, X, Y
' Note       : Notes that this procedure may confine the mouse pointer to
'              certain area in the screen. If you call this procedure, don't
'              forget to free the mouse pointer afterwards using
'              mdlAPI.ClipCursorClear function.
Private Sub picSplitter_MouseDown(Index As Integer, Button As Integer, _
                                  Shift As Integer, x As Single, y As Single)
  Dim uposCursor As mdlAPI.POINTAPI                  'another variable needed to
                                                     ' confine the mouse pointer
  Dim urecClipCursor As mdlAPI.RECT          'the rectangle area where the mouse
                                             '         pointer would be confined
  
  RaiseEvent MouseDown(CLng(Index), Button, Shift, x, y)
  
  If Button = vbLeftButton Then
    mblnDrag = True
    Select Case mSplitters(Index).Orientation
      Case orHorizontal
        mlngDragStart = y
      Case orVertical
        mlngDragStart = x
    End Select
    picSplitter(Index).ZOrder
    
    If Not mSplitters(Index).LiveUpdate Then
      picSplitter(Index).BackColor = mlngActiveColor
    End If
    
    If mSplitters(Index).ClipCursor Then
      mdlAPI.GetCursorPos uposCursor
      uposCursor.x = (uposCursor.x * Screen.TwipsPerPixelX) - _
                     (picSplitter(Index).Left + x)
      uposCursor.y = (uposCursor.y * Screen.TwipsPerPixelY) - _
                     (picSplitter(Index).Top + y)
      With urecClipCursor
        Select Case mSplitters(Index).Orientation
          Case orHorizontal
            .Top = (uposCursor.y + mSplitters(Index).MinYc) \ _
                   Screen.TwipsPerPixelY
            .Right = (uposCursor.x + mSplitters(Index).Right) \ _
                     Screen.TwipsPerPixelX
            .Bottom = (uposCursor.y + mSplitters(Index).MaxYc) \ _
                      Screen.TwipsPerPixelY
            .Left = (uposCursor.x + mSplitters(Index).Left) \ _
                    Screen.TwipsPerPixelX
          Case orVertical
            .Top = (uposCursor.y + mSplitters(Index).Top) \ _
                   Screen.TwipsPerPixelY
            .Right = (uposCursor.x + mSplitters(Index).MaxXc) \ _
                     Screen.TwipsPerPixelX
            .Bottom = (uposCursor.y + mSplitters(Index).Bottom) \ _
                      Screen.TwipsPerPixelY
            .Left = (uposCursor.x + mSplitters(Index).MinXc) \ _
                    Screen.TwipsPerPixelX
        End Select
      End With
      mdlAPI.ClipCursor urecClipCursor
    End If
  End If
  
  mposPrev.x = x
  mposPrev.y = y
End Sub

' Purpose    : Moves the splitter at run-time and raises custom event MouseMove
'              or Moving
' Assumption : The picSplitter_MouseMove procedure has been called
' Effects    : * If the user moves the splitter, custom event Moving has been
'                raised
'              * Otherwise, custom event MouseMove has been raised
'              * Other effect, as specified
' Inputs     : Index, Button, Shift, x, y
Private Sub picSplitter_MouseMove(Index As Integer, Button As Integer, _
                                  Shift As Integer, x As Single, y As Single)
Attribute picSplitter_MouseMove.VB_Description = "Moves the splitter at run time"
  Dim blnSplitterMoved As Boolean      'indicating whether the splitter is moved
  Dim lngPos As Long                          'dummy variable to determine where
                                              '       the splitter will be moved
  
  Select Case mSplitters(Index).Orientation
    Case orHorizontal
      blnSplitterMoved = mblnDrag And (y <> mlngDragStart)
    Case orVertical
      blnSplitterMoved = mblnDrag And (x <> mlngDragStart)
  End Select
  If blnSplitterMoved = True Then mblnSplitterMoved = True
  
  If Not mblnDrag And Not blnSplitterMoved And _
     ((x <> mposPrev.x) Or (y <> mposPrev.y)) Then
    RaiseEvent MouseMove(CLng(Index), Button, Shift, x, y)
  ElseIf blnSplitterMoved Then
    RaiseEvent Moving(CLng(Index), Shift, x, y)
  End If
  
  If blnSplitterMoved Then
    Select Case mSplitters(Index).Orientation
      Case orHorizontal
        lngPos = picSplitter(Index).Top + (y - mlngDragStart)
        If (mSplitters(Index).MinYc <= lngPos) And _
           (lngPos + picSplitter(Index).Height <= mSplitters(Index).MaxYc) Then
          picSplitter(Index).Top = lngPos
        End If
        If mSplitters(Index).LiveUpdate Then
          MoveSplitter IdSplitter:=CLng(Index), _
                       MoveTo:=picSplitter(Index).Top + _
                               (picSplitter(Index).Height \ 2)
        End If
      Case orVertical
        lngPos = picSplitter(Index).Left + (x - mlngDragStart)
        If (mSplitters(Index).MinXc <= lngPos) And _
           (lngPos + picSplitter(Index).Width <= mSplitters(Index).MaxXc) Then
          picSplitter(Index).Left = lngPos
        End If
        If mSplitters(Index).LiveUpdate Then
          MoveSplitter IdSplitter:=CLng(Index), _
                       MoveTo:=picSplitter(Index).Left + _
                               (picSplitter(Index).Width \ 2)
        End If
    End Select
  End If
  
  mposPrev.x = x
  mposPrev.y = y
End Sub

' Purpose    : Ends the run-time splitter move action and raises custom event
'              MouseUp or Moved
' Assumption : Picture Box control picSplitter(Index) which represents the
'              splitter exits
' Effects    : * mblnDrag = false
'              * Control picSplitter(Index) is in front of the other controls
'              * If the splitter's LiveUpdate property is false, then the
'                picSpliter(Index) BackColor property has been set to the
'                splitter's BackColor property
'              * The splitters minimum and maximum x- and y- coordinates have
'                been adjusted
'              * If the splitter's ClipCursor property is true, then the mouse
'                pointer has been freed from confinement
'              * If the splitter was moved then custom event Moved has been
'                raised, otherwise, custom event MouseUp has been raised
' Inputs     : Index, Button, Shift, x, y
Private Sub picSplitter_MouseUp(Index As Integer, Button As Integer, _
                                Shift As Integer, x As Single, y As Single)
Attribute picSplitter_MouseUp.VB_Description = "Ends the run time splitter move action"
  If mblnSplitterMoved Then
    RaiseEvent Moved(CLng(Index), Shift, x, y)
    mblnSplitterMoved = False
  Else
    RaiseEvent MouseUp(CLng(Index), Button, Shift, x, y)
  End If
  
  mblnDrag = False
  If Not mSplitters(Index).LiveUpdate Then
    picSplitter(Index).BackColor = mlngBackColor
  End If
  Select Case mSplitters(Index).Orientation
    Case orHorizontal
      MoveSplitter IdSplitter:=CLng(Index), _
                   MoveTo:=picSplitter(Index).Top + _
                           (picSplitter(Index).Height \ 2)
    Case orVertical
      MoveSplitter IdSplitter:=CLng(Index), _
                   MoveTo:=picSplitter(Index).Left + _
                           (picSplitter(Index).Width \ 2)
  End Select
  If mSplitters(Index).ClipCursor Then mdlAPI.ClipCursorClear
End Sub

' Purpose    : Adjusts the components inside the control to agree with the
'              control's size
' Effect     : See the codes below and see the effects of BuildSplitter and
'              Stretch procedures
Private Sub UserControl_Resize()
Attribute UserControl_Resize.VB_Description = "Adjusts the components inside the control to agree with the control's size"
  If ContainedControls.Count > 0 Then
    If (mControls.Count = 0) Or Not Ambient.UserMode Then
      BuildSplitter
      If Ambient.UserMode Then Extender.Visible = mblnVisibleSave
    Else
      Stretch
    End If
  End If
End Sub

'----------------------------
' ActiveX Control Properties
'----------------------------

' Purpose    : Sets the background color used to display the splitter when the
'              user moves it in none live update mode
' Effect     : As specified
' Input      : lngActiveColor (the new ActiveColor property value)
Public Property Let ActiveColor(lngActiveColor As OLE_COLOR)
Attribute ActiveColor.VB_Description = "Returns/sets the background color used to display the splitter when the user moves it in none live update mode"
  mlngActiveColor = lngActiveColor
  PropertyChanged mconActiveColor
End Property

' Purpose    : Returns the background color used to display the splitter when
'              the user moves it in none live update mode
' Return     : As specified
Public Property Get ActiveColor() As OLE_COLOR
  ActiveColor = mlngActiveColor
End Property

' Purpose    : Sets the background color used to display the splitters
' Effect     : As specified
' Input      : lngBackColor (the new BackColor property value)
Public Property Let BackColor(lngBackColor As OLE_COLOR)
Attribute BackColor.VB_Description = "Returns/sets the background color used to display the splitters"
  mlngBackColor = lngBackColor
  UserControl.BackColor = mlngBackColor
  Refresh
  PropertyChanged mconBackColor
End Property

' Purpose    : Returns the background color used to display the splitters
' Return     : As specified
Public Property Get BackColor() As OLE_COLOR
  BackColor = mlngBackColor
End Property

' Purpose    : Sets a value that determines whether the mouse pointer is
'              confined to its minimum and maximum x- and y-coordinate when
'              the user moves the splitter
' Effect     : As specified
' Input      : blnClipCursor (the new ClipCursor property value)
Public Property Let ClipCursor(ByVal blnClipCursor As Boolean)
Attribute ClipCursor.VB_Description = "Returns/sets a value that determines whether the mouse pointer is confined to its minimum and maximum x- and y-coordinate when the user moves the splitter"
  mSplitters.ClipCursor = blnClipCursor
  PropertyChanged mconClipCursor
End Property

' Purpose    : Returns a value that determines whether the mouse pointer is
'              confined to its minimum and maximum x- and y-coordinate when
'              the user moves the splitter
' Return     : As specified
Public Property Get ClipCursor() As Boolean
  ClipCursor = mSplitters.ClipCursor
End Property

' Purpose    : Sets a value that determines whether the splitters is movable
' Effect     : As specified
' Input      : blnEnable (the new Enable property value)
Public Property Let Enable(ByVal blnEnable As Boolean)
Attribute Enable.VB_Description = "Returns or sets a value that determines whether the splitters is movable"
  mSplitters.Enable = blnEnable
  PropertyChanged mconEnable
  
  Refresh
End Property

' Purpose    : Returns a value that determines whether the splitters is movable
' Return     : As specified
Public Property Get Enable() As Boolean
  Enable = mSplitters.Enable
End Property

' Purpose    : Sets a value that determines whether the control will
'              automatically adjust its size to fill-up its container with
'              respect to the margin properties
' Effect     : As specified
' Input      : blnFillContainer (the new FillContainer property value)
Public Property Let FillContainer(ByVal blnFillContainer As Boolean)
Attribute FillContainer.VB_Description = "Returns/sets a value that determines whether the control will automatically adjust its size to fill its container with respect to the margin properties"
  mblnFillContainer = blnFillContainer
  PropertyChanged mconFillContainer

  Activate
End Property

' Purpose    : Returns a value that determines whether the control will
'              automatically adjust its size to fill-up its container with
'              respect to the margin properties
' Return     : As specified
Public Property Get FillContainer() As Boolean
  FillContainer = mblnFillContainer
End Property

' Purpose    : Sets a value that determines whether the controls should be
'              resized as the splitter is moved
' Effect     : As specified
' Input      : blnLiveUpdate (the new LiveUpdate property value)
Public Property Let LiveUpdate(ByVal blnLiveUpdate As Boolean)
Attribute LiveUpdate.VB_Description = "Returns/sets a value that determines whether the controls should be resized as the splitter is moved"
  mSplitters.LiveUpdate = blnLiveUpdate
  PropertyChanged mconLiveUpdate
End Property

' Purpose    : Returns a value that determines whether the controls should be
'              resized as the splitter is moved
' Return     : As specified
Public Property Get LiveUpdate() As Boolean
  LiveUpdate = mSplitters.LiveUpdate
End Property

' Purpose    : Sets the bottom margin of the control from its container
' Effect     : As specified
' Input      : lngMarginBottom (the new MarginBottom property value)
Public Property Let MarginBottom(ByVal lngMarginBottom As Long)
Attribute MarginBottom.VB_Description = "Returns/sets the bottom margin of the control from its container"
  mlngMarginBottom = lngMarginBottom
  PropertyChanged mconMarginBottom
  
  Activate
End Property

' Purpose    : Returns the bottom margin of the control from its container
' Return     : As specified
Public Property Get MarginBottom() As Long
  MarginBottom = mlngMarginBottom
End Property

' Purpose    : Sets the left margin of the control from its container
' Effect     : As specified
' Input      : lngMarginLeft (the new MarginLeft property value)
Public Property Let MarginLeft(ByVal lngMarginLeft As Long)
Attribute MarginLeft.VB_Description = "Returns/sets the left margin of the control from its container"
  mlngMarginLeft = lngMarginLeft
  PropertyChanged mconMarginLeft
  
  Activate
End Property

' Purpose    : Returns the left margin of the control from its container
' Return     : As specified
Public Property Get MarginLeft() As Long
  MarginLeft = mlngMarginLeft
End Property

' Purpose    : Sets the right margin of the control from its container
' Effect     : As specified
' Input      : lngMarginRight (the new MarginRight property value)
Public Property Let MarginRight(ByVal lngMarginRight As Long)
Attribute MarginRight.VB_Description = "Returns/sets the right margin of the control from its container"
  mlngMarginRight = lngMarginRight
  PropertyChanged mconMarginRight

  Activate
End Property

' Purpose    : Returns the right margin of the control from its container
' Return     : As specified
Public Property Get MarginRight() As Long
  MarginRight = mlngMarginRight
End Property

' Purpose    : Sets the top margin of the control from its container
' Effect     : As specified
' Input      : lngMarginTop (the new MarginTop property value)
Public Property Let MarginTop(ByVal lngMarginTop As Long)
Attribute MarginTop.VB_Description = "Returns/sets the top margin of the control from its container"
  mlngMarginTop = lngMarginTop
  PropertyChanged mconMarginTop
  
  Activate
End Property

' Purpose    : Returns the top margin of the control from its container
' Return     : As specified
Public Property Get MarginTop() As Long
  MarginTop = mlngMarginTop
End Property

' Purpose    : Sets the size of the splitters
' Effects    : * If Size is smaller than the splitters' minimum size then the
'                splitters' size has been set to their minimum size
'              * Otherwise, as specified
'              * The control has been rebuilt
' Input      : lngSize (the new Size property value)
Public Property Let Size(ByVal lngSize As Long)
Attribute Size.VB_Description = "Returns/sets the size of the splitters"
  Dim blnNeedToShow As Boolean

  If lngSize >= mSplitters.MinimumSize Then
    mSplitters.Size = lngSize
  Else
    mSplitters.Size = mSplitters.MinimumSize
  End If
  PropertyChanged mconSize

  blnNeedToShow = (mControls.Count = 0) And Ambient.UserMode
  BuildSplitter
  If blnNeedToShow Then Extender.Visible = mblnVisibleSave
End Property

' Purpose    : Returns the size of the splitters
' Return     : As specified
Public Property Get Size() As Long
  Size = mSplitters.Size
End Property

'-------------------------
' ActiveX Control Methods
'-------------------------

' Purposes   : Activates and resize the control to meet its container size with
'              respect to the control's margin property and FillContainer
'              property
' Assumption : The parent of the control has ScaleWidth and ScaleHeight property
' Effect     : As specified
' Note       : This is the main method of the control. This method should be
'              called whenever its container is loaded. Also this method should
'              be called everytime its container's size is changed so that the
'              FillContainer property would work. If the container is forms,
'              this method should be called in the form's resize event.
Public Sub Activate()
Attribute Activate.VB_Description = "Activates and resize the control to meet its container size with respect to the control's margin property and FillContainer property"
  Dim lngWidth As Long                             'the new width of the control
  Dim lngHeight As Long                           'the new height of the control

  If mblnFillContainer Then
    lngWidth = UserControl.Parent.ScaleWidth - mlngMarginRight - mlngMarginLeft
    If lngWidth < 0 Then lngWidth = 0
    lngHeight = UserControl.Parent.ScaleHeight - _
                mlngMarginBottom - mlngMarginTop
    If lngHeight < 0 Then lngHeight = 0
    Extender.Move mlngMarginLeft, mlngMarginTop, lngWidth, lngHeight
  Else
    If mControls.Count = 0 Then UserControl_Resize
  End If
End Sub

' Purpose    : Move splitter IdSplitter to x- or y- (depending on its Orientation
'              property) coordinate MoveTo
' Assumption : Meet the assumption in Refresh procedure
' Effects    : * As specified
'              * All other effected splitters and controls' minimum and maximum
'                x- and y- coordinates have been adjusted
' Input      : * IdSplitter (a value that uniquely identifies a splitter)
'              * MoveTo (an x- or y- coordinate where the splitter will be moved)
Public Sub MoveSplitter(IdSplitter As Long, MoveTo As Long)
Attribute MoveSplitter.VB_Description = "Move splitter Index to x- or y- (depending on its Orientation property) coordinate MoveTo"
  Dim lngId As Long       'to determines the new friend control for the splitter
  Dim oid As clsId                     'for enumerating all Id in Ids collection
  Dim oid2 As clsId                    'for enumerating all Id in Ids collection
  
  With mSplitters(IdSplitter)
    Select Case .Orientation
      Case orHorizontal
        '-- If the destination coordinate is beyond the splitter's minimum or
        '   maximum value, generates a custom run-time error
        If (MoveTo < .MinYc) Or (MoveTo > .MaxYc) Then _
          RaiseError udeErrNumber:=errMoveSplitter, strSource:="MoveSplitter"
              
        '-- Move the splitter
        .Yc = MoveTo
        
        '-- Resize the controls and splitters that effected by the splitter
        '   movement
        For Each oid In .IdsCtlTop
          mControls(oid).Bottom = .Top
        Next
        For Each oid In .IdsCtlBottom
          mControls(oid).Top = .Bottom
        Next
        For Each oid In .IdsSplTop
          mSplitters(oid).Bottom = .Top
        Next
        For Each oid In .IdsSplBottom
          mSplitters(oid).Top = .Bottom
        Next
        
        '-- Finalizes the splitter movement by adjusting the minimum and maximum
        '   y- coordinates of the splitters above or below the active splitter
        If Not mblnDrag Then
          For Each oid In .IdsCtlTop
            If mControls(oid).IdSplTop <> gconUninitialized Then
              lngId = gconUninitialized
              For Each oid2 In mSplitters(mControls(oid).IdSplTop).IdsCtlBottom
                If lngId = gconUninitialized Then
                  lngId = oid2
                ElseIf mControls(oid2).Height - mControls(oid2).MinHeight < _
                       mControls(lngId).Height - mControls(lngId).MinHeight Then
                  lngId = oid2
                End If
              Next
              mSplitters(mControls(oid).IdSplTop).MaxYc = _
                mControls(lngId).Bottom - mControls(lngId).MinHeight
              mSplitters(mControls(oid).IdSplTop).IdCtlFriendBottom = lngId
            End If
          Next
          For Each oid In .IdsCtlBottom
            If mControls(oid).IdSplBottom <> gconUninitialized Then
              lngId = gconUninitialized
              For Each oid2 In mSplitters(mControls(oid).IdSplBottom).IdsCtlTop
                If lngId = gconUninitialized Then
                  lngId = oid2
                ElseIf mControls(oid2).Height - mControls(oid2).MinHeight < _
                       mControls(lngId).Height - mControls(lngId).MinHeight Then
                  lngId = oid2
                End If
              Next
              mSplitters(mControls(oid).IdSplBottom).MinYc = _
                mControls(lngId).Top + mControls(lngId).MinHeight
              mSplitters(mControls(oid).IdSplBottom).IdCtlFriendTop = lngId
            End If
          Next
        End If
      Case orVertical
        '-- If the destination coordinate is beyond the splitter's minimum or
        '   maximum value, generates a custom run-time error
        If (MoveTo < .MinXc) Or (MoveTo > .MaxXc) Then _
          RaiseError udeErrNumber:=errMoveSplitter, strSource:="MoveSplitter"
        
        ' Move the splitter
        .Xc = MoveTo
        
        '-- Resize the controls and splitters that effected by the splitter
        '   movement
        For Each oid In .IdsCtlLeft
          mControls(oid).Right = .Left
        Next
        For Each oid In .IdsCtlRight
          mControls(oid).Left = .Right
        Next
        For Each oid In .IdsSplLeft
          mSplitters(oid).Right = .Left
        Next
        For Each oid In .IdsSplRight
          mSplitters(oid).Left = .Right
        Next
        
        '-- Finalizes the splitter movement by adjusting the minimum and maximum
        '   x- coordinates of the splitters above or below the active splitter
        If Not mblnDrag Then
          For Each oid In .IdsCtlLeft
            If mControls(oid).IdSplLeft <> gconUninitialized Then
              lngId = gconUninitialized
              For Each oid2 In mSplitters(mControls(oid).IdSplLeft).IdsCtlRight
                If lngId = gconUninitialized Then
                  lngId = oid2
                ElseIf mControls(oid2).Width - mControls(oid2).MinWidth < _
                       mControls(lngId).Width - mControls(lngId).MinWidth Then
                  lngId = oid2
                End If
              Next
              mSplitters(mControls(oid).IdSplLeft).MaxXc = _
                mControls(lngId).Right - mControls(lngId).MinWidth
              mSplitters(mControls(oid).IdSplLeft).IdCtlFriendRight = lngId
            End If
          Next
          For Each oid In .IdsCtlRight
            If mControls(oid).IdSplRight <> gconUninitialized Then
              lngId = gconUninitialized
              For Each oid2 In mSplitters(mControls(oid).IdSplRight).IdsCtlLeft
                If lngId = gconUninitialized Then
                  lngId = oid2
                ElseIf mControls(oid2).Width - mControls(oid2).MinWidth < _
                       mControls(lngId).Width - mControls(lngId).MinWidth Then
                  lngId = oid2
                End If
              Next
              mSplitters(mControls(oid).IdSplRight).MinXc = _
                mControls(lngId).Left + mControls(lngId).MinWidth
              mSplitters(mControls(oid).IdSplRight).IdCtlFriendLeft = lngId
            End If
          Next
        End If
    End Select
  End With
  Refresh
End Sub

'----------------------------------
' Private Functions and Procedures
'----------------------------------

' Purpose    : Returns the adjusted height of control ctl
' Inputs     : * ctl
'              * octl (the virtual control of control ctl)
' Note       : This function is used to avoid flickering effect in LiveUpdate
'              mode for list box control or other controls that inherit it
Private Function AdjustedHeight(ByVal ctl As Control, _
                                ByVal octl As clsControl) As Long
Attribute AdjustedHeight.VB_Description = "Returns the adjusted height of control ctl"
  Dim lngAdjustedHeight                                          'returned value
  Dim lngHeightFactor As Long           'the height of each item in the list box
  
  If (TypeOf ctl Is ListBox) Or _
     (TypeOf ctl Is DirListBox) Or (TypeOf ctl Is FileListBox) Then
    lngHeightFactor = _
      mdlAPI.SendMessage(ctl.hWnd, mdlAPI.LB_GETITEMHEIGHT, 0&, 0&) * _
      Screen.TwipsPerPixelY
    lngAdjustedHeight = (((octl.Height - octl.MinHeight) \ lngHeightFactor) * _
                         lngHeightFactor) + octl.MinHeight
  Else
    lngAdjustedHeight = octl.Height
  End If
  AdjustedHeight = lngAdjustedHeight
End Function

' Purpose    : Build virtual controls and splitters and applies it to the real
'              controls and splitters
' Effect     : * If successed, as specified
'              * Otherwise, the custom error message has been raised
Private Sub BuildSplitter()
Attribute BuildSplitter.VB_Description = "Build virtual controls and splitters and applies it to the real controls and splitters"
  Dim i As Long       'for iterating all control in ContainedControls collection
  Dim octl As clsControl                   'for enumerating all virtual controls
                                           '              in Controls collection
  Dim oid As clsId                     'for enumerating all Id in Ids collection
  Dim ospl As clsSplitter                 'for enumerating all virtual splitters
                                          '              in Splitters collection
  
  '-- VB Splitter control can't have another VB Splitter control inside it
  If IsSelfContained Then
    RaiseError udeErrNumber:=errSelfContained, strSource:="Init"
  End If
 
  On Error GoTo ErrorHandler
  
  '-- Make the virtual controls solid and fill-up VB Splitter control's
  '   container
  Set mControls = New clsControls
  With mControls
    For i = 0 To ContainedControls.Count - 1
      .Add cctl:=ContainedControls, IdCtl:=i
    Next
    .Left = 0
    .Top = 0
    .Right = UserControl.ScaleWidth
    .Bottom = UserControl.ScaleHeight
    .RemoveHeaps
    .Compact
    .RemoveHoles
    .Stretch
  End With

  '-- Build virtual splitters and place it as the virtual controls' "border"
  With mSplitters
    .Left = 0
    .Top = 0
    .Right = UserControl.ScaleWidth
    .Bottom = UserControl.ScaleHeight
    .Clear
    For Each octl In mControls
      .Add octl:=octl
    Next
    For Each ospl In mSplitters
      ospl.IdsSplTop.RemoveDeleted lngLastPos:=.Count
      ospl.IdsSplRight.RemoveDeleted lngLastPos:=.Count
      ospl.IdsSplBottom.RemoveDeleted lngLastPos:=.Count
      ospl.IdsSplLeft.RemoveDeleted lngLastPos:=.Count
    Next
    For Each ospl In mSplitters
      Select Case ospl.Orientation
        Case orHorizontal
          For Each oid In ospl.IdsSplTop
            .Item(oid).Bottom = .Item(oid).Bottom - (ospl.Height \ 2)
          Next
          For Each oid In ospl.IdsSplBottom
            .Item(oid).Top = .Item(oid).Top + (ospl.Height \ 2)
          Next
        Case orVertical
          For Each oid In ospl.IdsSplLeft
            .Item(oid).Right = .Item(oid).Right - (ospl.Width \ 2)
          Next
          For Each oid In ospl.IdsSplRight
            .Item(oid).Left = .Item(oid).Left + (ospl.Width \ 2)
          Next
      End Select
    Next
  End With
  
  '-- Creates the new PictureBox control instances to represent the splitter
  For Each ospl In mSplitters
    On Error Resume Next
    Load picSplitter(ospl)
    On Error GoTo ErrorHandler
    picSplitter(ospl).MousePointer = vbCustom
    Select Case ospl.Orientation
      Case orHorizontal
        picSplitter(ospl).MouseIcon = _
          LoadResPicture(gconCurHSplitter, vbResCursor)
      Case orVertical
        picSplitter(ospl).MouseIcon = _
          LoadResPicture(gconCurVSplitter, vbResCursor)
    End Select
    picSplitter(ospl).Visible = True
  Next
  
ErrorHandler:
  If (Err.Number <> 0) And (Not Ambient.UserMode) Then
    Resume Next
  ElseIf (Err.Number <> 0) Or (Not IsSolid) Or (Not mControls.IsValid) Then
    RaiseError udeErrNumber:=errBuildSplitters, strSource:="Init"
  End If
  
  Refresh
End Sub

' Purpose    : Returns a valid x- and y- coordinate scale
' Return     : As specified
Private Sub GetValidStretchScale(ByRef sngXScale As Single, _
                                 ByRef sngYScale As Single)
Attribute GetValidStretchScale.VB_Description = "Returns a valid x- and y- coordinate scale"
  Dim octl As clsControl                   'for enumerating all virtual controls
                                           '              in Controls collection
  
  sngXScale = UserControl.ScaleWidth / mControls.Width
  sngYScale = UserControl.ScaleHeight / mControls.Height
  For Each octl In mControls
    If octl.Width * sngXScale < octl.MinWidth Then sngXScale = 1
    If octl.Height * sngYScale < octl.MinHeight Then sngYScale = 1
  Next
End Sub

' Purpose    : Returns a value indicating whether this VB Splitter control
'              instance contains another VB Splitter Controls instance
' Return     : As specified
Private Function IsSelfContained() As Boolean
Attribute IsSelfContained.VB_Description = "Returns a value indicating whether this VB Splitter control instance contains another VB Splitter Controls instance"
  Dim blnIsSelfContained As Boolean                              'returned value
  Dim ctl As Control                            'for enumerating all controls in
                                                '   ContainedControls collection
  
  blnIsSelfContained = False
  For Each ctl In ContainedControls
    If TypeOf ctl Is Splitter Then
      blnIsSelfContained = True
      Exit For
    End If
  Next
  IsSelfContained = blnIsSelfContained
End Function

' Purpose    : Returns a value indicating whether the virtual controls and
'              splitters are solid
' Return     : As specified
' Note       : See VB Splitter's documention for the definition of "solid"
Private Function IsSolid() As Boolean
Attribute IsSolid.VB_Description = "Returns a value indicating whether the virtual controls and splitters are solid"
  Dim lngExtent As Long      'total extent of the virtual controls and splitters
  Dim lngSplTopHeight As Long         'the height of the virtual splitter on the
                                      '       top-side of the current enumerated
                                      '                          virtual control
  Dim lngSplRightWidth As Long         'the width of the virtual splitter on the
                                       '    right-side of the current enumerated
                                       '                         virtual control
  Dim lngSplBottomHeight As Long      'the height of the virtual splitter on the
                                      '    bottom-side of the current enumerated
                                      '                          virtual control
  Dim lngSplLeftWidth As Long          'the width of the virtual splitter on the
                                       '     left-side of the current enumerated
                                       '                         virtual control
  Dim octl As clsControl                   'for enumerating all virtual controls
                                           '              in Controls collection
  
  lngExtent = 0
  For Each octl In mControls
    If octl.IdSplTop <> gconUninitialized Then
      lngSplTopHeight = mSplitters(octl.IdSplTop).Height
    Else
      lngSplTopHeight = 0
    End If
    If octl.IdSplRight <> gconUninitialized Then
      lngSplRightWidth = mSplitters(octl.IdSplRight).Width
    Else
      lngSplRightWidth = 0
    End If
    If octl.IdSplBottom <> gconUninitialized Then
      lngSplBottomHeight = mSplitters(octl.IdSplBottom).Height
    Else
      lngSplBottomHeight = 0
    End If
    If octl.IdSplLeft <> gconUninitialized Then
      lngSplLeftWidth = mSplitters(octl.IdSplLeft).Width
    Else
      lngSplLeftWidth = 0
    End If
    lngExtent = lngExtent + ((octl.Width + (lngSplLeftWidth \ 2) + _
                             (lngSplRightWidth \ 2)) * _
                            (octl.Height + (lngSplTopHeight \ 2) + _
                             (lngSplBottomHeight \ 2)))
  Next
  IsSolid = (lngExtent = (mControls.Right - mControls.Left) * _
                         (mControls.Bottom - mControls.Top))
End Function

' Purpose    : Applies the virtual controls and splitters to their real controls
'              and splitter
' Effect     : As specified
Private Sub Refresh()
  Const conErrHeightReadOnly = 383
  
  Dim blnNeedRefresh As Boolean                            'to reduce flickering
  Dim lngHeight As Long                    'adjusted height for list box control
  Dim lngErrNumber As Long                      'for the control with r/o height
  Dim octl As clsControl                   'for enumerating all virtual controls
                                           '              in Controls collection
  Dim ospl As clsSplitter                 'for enumerating all virtual splitters
                                          '              in Splitters collection
  
  On Error GoTo ErrorHandler
  
  '-- Applies all virtuals splitters to their real splitters
  For Each ospl In mSplitters
    With picSplitter(ospl)
      .Move ospl.Left, ospl.Top, ospl.Width, ospl.Height
      .BackColor = mlngBackColor
      .Enabled = ospl.Enable
      .ZOrder
    End With
  Next
  
  '-- Applies all virtuals controls to their real controls
  blnNeedRefresh = False
  For Each octl In mControls
    lngHeight = AdjustedHeight(ctl:=ContainedControls(octl.Id), octl:=octl)
    With ContainedControls(octl.Id)
      If (.Left <> octl.Left) Or (.Top <> octl.Top) Or _
         (.Width <> octl.Width) Or (.Height <> lngHeight) Then
        .Move octl.Left, octl.Top, octl.Width, lngHeight
        If lngErrNumber = conErrHeightReadOnly Then
          .Move octl.Left, octl.Top, octl.Width
          lngErrNumber = 0
        End If
        blnNeedRefresh = True
      End If
    End With
  Next
  
  If blnNeedRefresh Then UserControl.Refresh
  Exit Sub
  
ErrorHandler:
  If Err.Number = conErrHeightReadOnly Then
    lngErrNumber = Err.Number
    Resume Next
  Else
    Err.Raise Err.Number
  End If
End Sub

' Purpose    : Stretchs the controls and splitters to fill-up their container
' Effect     : As specified
Private Sub Stretch()
Attribute Stretch.VB_Description = "Stretchs the controls and splitters to fill-up their container"
  Dim octl As clsControl                   'for enumerating all virtual controls
                                           '              in Controls collection
  Dim oid As clsId                     'for enumerating all Id in Ids collection
  Dim ospl As clsSplitter                 'for enumerating all virtual splitters
                                          '              in Splitters collection
  Dim sngXScale As Single                          'a valid x-coorindate's scale
  Dim sngYScale As Single                          'a valid y-coordinate's scale
  
  GetValidStretchScale sngXScale, sngYScale
  
  '-- Stretch the virtual splitters
  If (Abs(sngXScale - 1) > 0.001) Or (Abs(sngYScale - 1) > 0.001) Then
    mSplitters.Width = mSplitters.Width * sngXScale
    mSplitters.Height = mSplitters.Height * sngYScale
    For Each ospl In mSplitters
      With ospl
        Select Case .Orientation
          Case orHorizontal
            .Xc = CLng(.Xc * sngXScale)
            .Yc = CLng((.Top * sngYScale) + ((.Height * sngYScale) / 2))
            .Width = CLng(.Width * sngXScale)
            .MinYc = CLng((mControls(ospl.IdCtlFriendTop).Top * sngYScale) + _
                          mControls(ospl.IdCtlFriendTop).MinHeight)
            .MaxYc = CLng((mControls(ospl.IdCtlFriendBottom).Bottom * _
                           sngYScale) - _
                          mControls(ospl.IdCtlFriendBottom).MinHeight)
          Case orVertical
            .Xc = CLng((.Left * sngXScale) + ((.Width * sngXScale) / 2))
            .Yc = CLng(.Yc * sngYScale)
            .Height = CLng(.Height * sngYScale)
            .MinXc = CLng((mControls(ospl.IdCtlFriendLeft).Left * sngXScale) + _
                          mControls(ospl.IdCtlFriendLeft).MinWidth)
            .MaxXc = CLng((mControls(ospl.IdCtlFriendRight).Right * _
                           sngXScale) - _
                          mControls(ospl.IdCtlFriendRight).MinWidth)
        End Select
      End With
    Next
    For Each ospl In mSplitters
      Select Case ospl.Orientation
        Case orHorizontal
          For Each oid In ospl.IdsSplTop
            mSplitters(oid).Bottom = ospl.Top
          Next
          For Each oid In ospl.IdsSplBottom
            mSplitters(oid).Top = ospl.Bottom
          Next
        Case orVertical
          For Each oid In ospl.IdsSplLeft
            mSplitters(oid).Right = ospl.Left
          Next
          For Each oid In ospl.IdsSplRight
            mSplitters(oid).Left = ospl.Right
          Next
      End Select
    Next
    
    '-- Stretch the virtual controls
    mControls.Width = mControls.Width * sngXScale
    mControls.Height = mControls.Height * sngYScale
    For Each octl In mControls
      If octl.IdSplTop = gconUninitialized Then
        octl.Top = mControls.Top
        If octl.IdSplLeft <> gconUninitialized Then _
          mSplitters(octl.IdSplLeft).Top = octl.Top
        If octl.IdSplRight <> gconUninitialized Then _
          mSplitters(octl.IdSplRight).Top = octl.Top
      Else
        octl.Top = mSplitters(octl.IdSplTop).Bottom
      End If
      If octl.IdSplRight = gconUninitialized Then
        octl.Right = mControls.Right
        If octl.IdSplTop <> gconUninitialized Then _
          mSplitters(octl.IdSplTop).Right = octl.Right
        If octl.IdSplBottom <> gconUninitialized Then _
          mSplitters(octl.IdSplBottom).Right = octl.Right
      Else
        octl.Right = mSplitters(octl.IdSplRight).Left
      End If
      If octl.IdSplBottom = gconUninitialized Then
        octl.Bottom = mControls.Bottom
        If octl.IdSplLeft <> gconUninitialized Then _
          mSplitters(octl.IdSplLeft).Bottom = octl.Bottom
        If octl.IdSplRight <> gconUninitialized Then _
          mSplitters(octl.IdSplRight).Bottom = octl.Bottom
      Else
        octl.Bottom = mSplitters(octl.IdSplBottom).Top
      End If
      If octl.IdSplLeft = gconUninitialized Then
        octl.Left = mControls.Left
        If octl.IdSplTop <> gconUninitialized Then _
          mSplitters(octl.IdSplTop).Left = octl.Left
        If octl.IdSplBottom <> gconUninitialized Then _
        mSplitters(octl.IdSplBottom).Left = octl.Left
      Else
        octl.Left = mSplitters(octl.IdSplLeft).Right
      End If
      If octl.Height < 0 Then Stop
    Next
    
    Refresh
  End If
End Sub
