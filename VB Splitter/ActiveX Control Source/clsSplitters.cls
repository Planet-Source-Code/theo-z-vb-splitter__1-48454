VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSplitters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A class module to implement custom collection of class clsSplitter"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************************************************************
'** File Name  : clsSplitters.cls                                             **
'** Language   : Visual Basic 6.0                                             **
'** Author     : Theo Zacharias (theo_yz@yahoo.com)                           **
'** Description: A class module to implement custom collection of class       **
'**              clsSplitter                                                  **
'** Usage      : VB Splitter control uses PictureBox control to for the       **
'**              splitters. Instead of using PictureBox properties directly,  **
'**              VB Splitter control saves the necessary properties plus      **
'**              several custom properties in a virtual splitter.             **
'**              Furthermore, VB Splitter control works with this virtual     **
'**              splitter (will be called splitter next) to manipulate its    **
'**              position and then write it back to the PictureBox control.   **
'** Members    :                                                              **
'**   * Properties: Bottom (r/w), ClipCursor (r/w), Count (r/o),              **
'**                 DefaultClipCursor (r/o), DefaultEnable (r/o),             **
'**                 DefaultLiveUpdate (r/o), DefaultSize (r/o),               **
'**                 Enable (r/w), Height (r/w), Item (def. r/o), Left (r/w),  **
'**                 LiveUpdate (r/w), MinimumSize (r/o), Right (r/w),         **
'**                 Size (r/w), Top (r/w), Width (r/w)                        **
'**   * Methods   : Add, Clear, Remove                                        **
'** Last modified on September 12, 2003                                       **
'*******************************************************************************

Option Explicit

'--- Custom Collection Class Variable
Private mcolSplitters As Collection

'--- Property Variables

' Property variables which appears in the property page
Private mblnClipCursor As Boolean
Private mblnEnable As Boolean
Private mblnLiveUpdate As Boolean
Private mlngSize As Long

' Splitters frame area
Private mlngLeft As Long
Private mlngTop As Long
Private mlngRight As Long
Private mlngBottom As Long

'--- Property Default Values
Private Const mconDefaultClipCursor As Boolean = False
Private Const mconDefaultEnable As Boolean = True
Private Const mconDefaultLiveUpdate As Boolean = True
Private Const mconDefaultSize As Long = 60

'--- Other Variables
Private mlngMinimumSize As Long                     'the minimum splitters' size

'----------------------------------
' Class Constructor and Destructor
'----------------------------------

Private Sub Class_Initialize()
  Set mcolSplitters = New Collection
  
  mblnClipCursor = mconDefaultClipCursor
  mblnEnable = mconDefaultEnable
  mblnLiveUpdate = mconDefaultLiveUpdate
  mlngMinimumSize = IIf(Screen.TwipsPerPixelX > Screen.TwipsPerPixelY, _
                        Screen.TwipsPerPixelX, Screen.TwipsPerPixelY)
  mlngSize = mconDefaultSize
End Sub

Private Sub Class_Terminate()
  Set mcolSplitters = Nothing
End Sub

'-----------------------------------------
' Collection Class Properties and Methods
'-----------------------------------------

' Purpose    : Adds new items to the collection
' Effect     : * If does not exist, the top-, right-, bottom- and left-side
'                splitters of control octl have been added and their necessary
'                properties have been initialized
'              * Control octl have been shrunk to give space for the splitters
'                around it
'              * This method may left several splitters' id marked deleted that
'                should be removed after all the splitters have been added
' Input      : octl (the control that trigger the splitters addition)
Public Sub Add(ByRef octl As clsControl)
Attribute Add.VB_Description = "Adds new items to the collection"
  Dim lngPos As Long                   'a temporary variable used in determining
                                       '               Splitter 's movement area
  Dim ospl As clsSplitter           'for enumerating all items in the collection
  Dim osplNew As clsSplitter 'the new item which will be added to the collection
  
  '-- If the splitters for control octl already exist, assigns the splitters to
  '   controls octl's IdSplTop, IdSpRight, IdSplBottom or IdSplLeft property and
  '   adjust the splitters' size
  For Each ospl In Me
    If (ospl.Orientation = orHorizontal) And (octl.Top = ospl.Yc) Then
      octl.IdSplTop = ospl.Id
      If octl.Left < ospl.Left Then ospl.Left = octl.Left
      If octl.Right > ospl.Right Then ospl.Right = octl.Right
    End If
    If (ospl.Orientation = orVertical) And (octl.Right = ospl.Xc) Then
      octl.IdSplRight = ospl.Id
      If octl.Top < ospl.Top Then ospl.Top = octl.Top
      If octl.Bottom > ospl.Bottom Then ospl.Bottom = octl.Bottom
    End If
    If (ospl.Orientation = orHorizontal) And (octl.Bottom = ospl.Yc) Then
      octl.IdSplBottom = ospl.Id
      If octl.Left < ospl.Left Then ospl.Left = octl.Left
      If octl.Right > ospl.Right Then ospl.Right = octl.Right
    End If
    If (ospl.Orientation = orVertical) And (octl.Left = ospl.Xc) Then
      octl.IdSplLeft = ospl.Id
      If octl.Top < ospl.Top Then ospl.Top = octl.Top
      If octl.Bottom > ospl.Bottom Then ospl.Bottom = octl.Bottom
    End If
  Next
    
  '-- If the splitters for control octl doesn't exist, create a new splitter,
  '   add it to the collection, and assigns it to controls octl's IdSplTop,
  '   IdSplRight, IdSplBottom or IdSplLeft
  If (octl.IdSplTop = gconUninitialized) And (octl.Top <> mlngTop) Then
    Set osplNew = New clsSplitter
    With osplNew
      .Id = mcolSplitters.Count
      .Orientation = orHorizontal
      .Xc = octl.Xc
      .Yc = octl.Top
      .Width = octl.Width
      .Height = mlngSize
      .ClipCursor = mblnClipCursor
      .Enable = mblnEnable
      .LiveUpdate = mblnLiveUpdate
      mcolSplitters.Add Item:=osplNew, Key:=CStr(.Id)
      octl.IdSplTop = .Id
    End With
    Set osplNew = Nothing
  End If
  If (octl.IdSplRight = gconUninitialized) And (octl.Right <> mlngRight) Then
    Set osplNew = New clsSplitter
    With osplNew
      .Id = mcolSplitters.Count
      .Orientation = orVertical
      .Xc = octl.Right
      .Yc = octl.Yc
      .Width = mlngSize
      .Height = octl.Height
      .ClipCursor = mblnClipCursor
      .Enable = mblnEnable
      .LiveUpdate = mblnLiveUpdate
      mcolSplitters.Add Item:=osplNew, Key:=CStr(.Id)
      octl.IdSplRight = .Id
    End With
    Set osplNew = Nothing
  End If
  If (octl.IdSplBottom = gconUninitialized) And (octl.Bottom <> mlngBottom) Then
    Set osplNew = New clsSplitter
    With osplNew
      .Id = mcolSplitters.Count
      .Orientation = orHorizontal
      .Xc = octl.Xc
      .Yc = octl.Bottom
      .Width = octl.Width
      .Height = mlngSize
      .ClipCursor = mblnClipCursor
      .Enable = mblnEnable
      .LiveUpdate = mblnLiveUpdate
      mcolSplitters.Add Item:=osplNew, Key:=CStr(.Id)
      octl.IdSplBottom = .Id
    End With
    Set osplNew = Nothing
  End If
  If (octl.IdSplLeft = gconUninitialized) And (octl.Left <> mlngLeft) Then
    Set osplNew = New clsSplitter
    With osplNew
      .Id = mcolSplitters.Count
      .Orientation = orVertical
      .Xc = octl.Left
      .Yc = octl.Yc
      .Width = mlngSize
      .Height = octl.Height
      .ClipCursor = mblnClipCursor
      .Enable = mblnEnable
      .LiveUpdate = mblnLiveUpdate
      mcolSplitters.Add Item:=osplNew, Key:=CStr(.Id)
      octl.IdSplLeft = .Id
    End With
    Set osplNew = Nothing
  End If
  
  '-- Shrinks the control octl's size to give space for the splitters
  octl.Left = octl.Left + IIf(octl.Left <> mlngLeft, mlngSize \ 2, 0)
  octl.Top = octl.Top + IIf(octl.Top <> mlngTop, mlngSize \ 2, 0)
  octl.Right = octl.Right - IIf(octl.Right <> mlngRight, mlngSize \ 2, 0)
  octl.Bottom = octl.Bottom - IIf(octl.Bottom <> mlngBottom, mlngSize \ 2, 0)
  
  '-- Assigns all controls and splitters' id that effected by splitter
  '   octl.IdSplTop, octl.IdSplRight, octl.IdSplBottom and octl.IdSplLeft to the
  '   related splitter
  If octl.IdSplTop <> gconUninitialized Then
    With Me(octl.IdSplTop)
      .IdsCtlBottom.Add octl.Id
      .IdsSplBottom.Add octl.IdSplRight
      If .IdsSplTop.IsExist(octl.IdSplRight) Then
        .IdsSplTop.MarkDeleted octl.IdSplRight
        .IdsSplBottom.MarkDeleted octl.IdSplRight
      End If
      .IdsSplBottom.Add octl.IdSplLeft
      If .IdsSplTop.IsExist(octl.IdSplLeft) Then
        .IdsSplTop.MarkDeleted octl.IdSplLeft
        .IdsSplBottom.MarkDeleted octl.IdSplLeft
      End If
      lngPos = octl.Bottom - octl.MinHeight
      If lngPos < .MaxYc Then
        .MaxYc = lngPos
        .IdCtlFriendBottom = octl.Id
      End If
    End With
  End If
  If octl.IdSplRight <> gconUninitialized Then
    With Me(octl.IdSplRight)
      .IdsCtlLeft.Add octl.Id
      .IdsSplLeft.Add octl.IdSplTop
      If .IdsSplRight.IsExist(octl.IdSplTop) Then
        .IdsSplRight.MarkDeleted octl.IdSplTop
        .IdsSplLeft.MarkDeleted octl.IdSplTop
      End If
      .IdsSplLeft.Add octl.IdSplBottom
      If .IdsSplRight.IsExist(octl.IdSplBottom) Then
        .IdsSplRight.MarkDeleted octl.IdSplBottom
        .IdsSplLeft.MarkDeleted octl.IdSplBottom
      End If
      lngPos = octl.Left + octl.MinWidth
      If lngPos > .MinXc Then
        .MinXc = lngPos
        .IdCtlFriendLeft = octl.Id
      End If
    End With
  End If
  If octl.IdSplBottom <> gconUninitialized Then
    With Me(octl.IdSplBottom)
      .IdsCtlTop.Add octl.Id
      .IdsSplTop.Add octl.IdSplRight
      If .IdsSplBottom.IsExist(octl.IdSplRight) Then
        .IdsSplBottom.MarkDeleted octl.IdSplRight
        .IdsSplTop.MarkDeleted octl.IdSplRight
      End If
      .IdsSplTop.Add octl.IdSplLeft
      If .IdsSplBottom.IsExist(octl.IdSplLeft) Then
        .IdsSplBottom.MarkDeleted octl.IdSplLeft
        .IdsSplTop.MarkDeleted octl.IdSplLeft
      End If
      lngPos = octl.Top + octl.MinHeight
      If lngPos > .MinYc Then
        .MinYc = lngPos
        .IdCtlFriendTop = octl.Id
      End If
    End With
  End If
  If octl.IdSplLeft <> gconUninitialized Then
    With Me(octl.IdSplLeft)
      .IdsCtlRight.Add octl.Id
      .IdsSplRight.Add octl.IdSplTop
      If .IdsSplLeft.IsExist(octl.IdSplTop) Then
        .IdsSplLeft.MarkDeleted octl.IdSplTop
        .IdsSplRight.MarkDeleted octl.IdSplTop
      End If
      .IdsSplRight.Add octl.IdSplBottom
      If .IdsSplLeft.IsExist(octl.IdSplBottom) Then
        .IdsSplLeft.MarkDeleted octl.IdSplBottom
        .IdsSplRight.MarkDeleted octl.IdSplBottom
      End If
      lngPos = octl.Right - octl.MinWidth
      If lngPos < .MaxXc Then
        .MaxXc = lngPos
        .IdCtlFriendRight = octl.Id
      End If
    End With
  End If
End Sub

' Purpose    : Removes all items in the collection
' Effect     : As specified
Public Sub Clear()
Attribute Clear.VB_Description = "Removes all items in the collection"
  Dim ospl As clsSplitter           'for enumerating all items in the collection
  
  For Each ospl In Me
    Remove ospl
  Next
End Sub

' Purpose    : Returns the number of items in the collection
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of items in the collection"
  Count = mcolSplitters.Count
End Property

' Purpose    : Returns an item in the collection which has key IdSpl
' Assumption : Key IdSpl exists in the collection
' Input      : IdSpl
' Return     : As specified
Public Property Get Item(ByVal IdSpl As Long) As clsSplitter
Attribute Item.VB_Description = "Returns an item in the collection which has key IdSpl"
Attribute Item.VB_UserMemId = 0
  Set Item = mcolSplitters.Item(CStr(IdSpl))
End Property

' Purpose    : Enables For Each ... Next enumeration
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Enables For Each ... Next enumeration"
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mcolSplitters.[_NewEnum]
End Property

' Purpose    : Removes an item which has key IdSpl from the collection
' Assumption : Key IdSpl exists in the collection
' Effect     : As specified
' Input      : IdSpl
Public Sub Remove(ByVal IdSpl As Long)
Attribute Remove.VB_Description = "Removes an item which has key IdSpl from the collection"
  mcolSplitters.Remove CStr(IdSpl)
End Sub

'------------------
' Class Properties
'------------------

' Purpose    : Sets the distance between the bottom edge of the splitters frame
'              area and the top edge of VB Splitter control
' Effect     : As specified
' Input      : lngBottom (the new Bottom propety value)
Public Property Let Bottom(ByVal lngBottom As Long)
Attribute Bottom.VB_Description = "Returns or sets the distance between the bottom edge of the splitters frame area and the top edge of VB Splitter control"
  mlngBottom = lngBottom
End Property

' Purpose    : Returns the distance between the bottom edge of the splitters
'              frame area and the top edge of VB Splitter control
' Return     : As specified
Public Property Get Bottom() As Long
  Bottom = mlngBottom
End Property

' Purpose    : Returns the default ClipCursor property
Public Property Get DefaultClipCursor() As Boolean
Attribute DefaultClipCursor.VB_Description = "Returns the default ClipCursor property"
  DefaultClipCursor = mconDefaultClipCursor
End Property

' Purpose    : Returns the default Enable property
Public Property Get DefaultEnable() As Boolean
Attribute DefaultEnable.VB_Description = "Returns the default Enable property"
  DefaultEnable = mconDefaultEnable
End Property

' Purpose    : Returns the default LiveUpdate property
Public Property Get DefaultLiveUpdate() As Boolean
Attribute DefaultLiveUpdate.VB_Description = "Returns the default LiveUpdate property"
  DefaultLiveUpdate = mconDefaultLiveUpdate
End Property

' Purpose    : Returns the default Size property
Public Property Get DefaultSize() As Long
Attribute DefaultSize.VB_Description = "Returns the default Size property"
  DefaultSize = mconDefaultSize
End Property

' Purpose    : Sets a value that determines whether the mouse pointer is
'              confined to the items' region
' Effects    : * The ClipCursor property of all items in the collection has been
'                set to blnClipCursor
'              * As specified
' Input      : blnClipCursor (the new ClipCursor property value)
Public Property Let ClipCursor(ByVal blnClipCursor As Boolean)
Attribute ClipCursor.VB_Description = "Returns or sets a value that determines whether the mouse pointer is confined to the items' region\r\n"
  Dim ospl As clsSplitter           'for enumerating all items in the collection
  
  For Each ospl In Me
    ospl.ClipCursor = blnClipCursor
  Next
  mblnClipCursor = blnClipCursor
End Property

' Purpose    : Returns a value that determines whether the mouse pointer is
'              confined to the items' region
' Return     : As specified
' Note       : Each item has its own ClipCursor properties. This property only
'              return the last ClipCursor property value set by the users. If
'              the users set the property individually for each item, this
'              property may not return the correct ClipCursor property value.
Public Property Get ClipCursor() As Boolean
  ClipCursor = mblnClipCursor
End Property

' Purpose    : Sets a value that determines whether the splitter is movable
' Effects    : * The Enable property of all items in the collection has been
'                set to blnEnable
'              * As specified
' Input      : blnEnable (the new Enable property value)
Public Property Let Enable(ByVal blnEnable As Boolean)
Attribute Enable.VB_Description = "Returns or sets a value that determines whether the items are movable"
  Dim ospl As clsSplitter           'for enumerating all items in the collection
  
  For Each ospl In Me
    ospl.Enable = blnEnable
  Next
  mblnEnable = blnEnable
End Property

' Purpose    : Returns a value that determines whether the items are movable
' Return     : As specified
' Note       : Each item has its own Enable properties. This property only
'              return the last Enable property value set by the users. If the
'              users set the property individually for each item, this property
'              may not return the correct Enable property value.
Public Property Get Enable() As Boolean
  Enable = mblnEnable
End Property

' Purpose    : Sets the height of the splitters frame area
' Effect     : As specified
' Input      : lngHeight(the new Height property value)
Public Property Let Height(ByVal lngHeight As Long)
Attribute Height.VB_Description = "Returns or sets the height of the splitters frame area"
  mlngBottom = mlngTop + lngHeight
End Property

' Purpose    : Returns the height of the splitters frame area
' Return     : As specified
Public Property Get Height() As Long
  Height = mlngBottom - mlngTop
End Property

' Purpose    : Sets the distance between the left edge of the splitters frame
'              area and the left edge of VB Splitter control
' Effect     : As specified
' Input      : lngLeft (the new Left propety value)
Public Property Let Left(ByVal lngLeft As Long)
Attribute Left.VB_Description = "Returns or sets the distance between the left edge of the splitters frame area and the left edge of VB Splitter control"
  mlngLeft = lngLeft
End Property

' Purpose    : Returns the distance between the left edge of the splitters frame
'              area and the left edge of VB Splitter control
' Return     : As specified
Public Property Get Left() As Long
  Left = mlngLeft
End Property

' Purpose    : Sets a value that determines whether the controls should be
'              resized as the item is moved
' Effects    : * The LiveUpdate property of all items in the collection has been
'                set to blnLiveUpdate
'              * As specified
' Input      : blnLiveUpdate (the new LiveUpdate property value)
Public Property Let LiveUpdate(ByVal blnLiveUpdate As Boolean)
Attribute LiveUpdate.VB_Description = "Returns or sets a value that determines whether the controls should be resized as the  item is moved"
  Dim ospl As clsSplitter           'for enumerating all items in the collection
  
  For Each ospl In Me
    ospl.LiveUpdate = blnLiveUpdate
  Next
  mblnLiveUpdate = blnLiveUpdate
End Property

' Purpose    : Returns a value that determines whether the controls should be
'              resized as the item is moved
' Return     : As specified
' Note       : Each item has its own LiveUpdate properties. This property only
'              return the last LiveUpdate property value set by the users. If
'              the users set the property individually for each item, this
'              property may not return the correct LiveUpdate property value.
Public Property Get LiveUpdate() As Boolean
  LiveUpdate = mblnLiveUpdate
End Property

' Purpose    : Returns the minimum items' size
Public Property Get MinimumSize() As Long
Attribute MinimumSize.VB_Description = "Returns the minimum items' size"
  MinimumSize = mlngMinimumSize
End Property

' Purpose    : Sets the distance between the right edge of the splitters frame
'              area and the left edge of VB Splitter control
' Effect     : As specified
' Input      : lngRight (the new Right propety value)
Public Property Let Right(ByVal lngRight As Long)
Attribute Right.VB_Description = "Returns or sets the distance between the right edge of the splitters frame area and the left edge of VB Splitter control"
  mlngRight = lngRight
End Property

' Purpose    : Returns the distance between the right edge of the splitters
'              frame area and the left edge of VB Splitter control
' Return     : As specified
Public Property Get Right() As Long
  Right = mlngRight
End Property

' Purpose    : Sets the items' size
' Effects    : * The Height or Width property (depend on the Orientation
'                property) of all items in the collection have been set to
'                lngSize
'              * As specified
' Input      : lngSize (the new Size property value)
Public Property Let Size(ByVal lngSize As Long)
Attribute Size.VB_Description = "Returns or sets the items' size"
  Dim ospl As clsSplitter           'for enumerating all items in the collection

  For Each ospl In Me
    Select Case ospl.Orientation
      Case orHorizontal
        ospl.Height = lngSize
      Case orVertical
        ospl.Width = lngSize
    End Select
  Next
  mlngSize = lngSize
End Property

' Purpose    : Returns the items' size
' Return     : As specified
' Note       : Each item has its own Size properties. This property only return
'              the last Size property value set by the users. If the users set
'              the property individually for each item, this property may not
'              return the correct Size property value.
Public Property Get Size() As Long
  Size = mlngSize
End Property

' Purpose    : Sets the distance between the top edge of the splitters frame
'              area and the top edge of VB Splitter control
' Effect     : As specified
' Input      : lngTop (the new Top propety value)
Public Property Let Top(ByVal lngTop As Long)
Attribute Top.VB_Description = "Returns or sets the distance between the top edge of the splitters frame area and the top edge of VB Splitter control"
  mlngTop = lngTop
End Property

' Purpose    : Returns the distance between the top edge of the splitters frame
'              area and the top edge of VB Splitter control
' Return     : As specified
Public Property Get Top() As Long
  Top = mlngTop
End Property

' Purpose    : Sets the width of the splitters frame area
' Effect     : As specified
' Input      : lngWidth(the new Width property value)
Public Property Let Width(ByVal lngWidth As Long)
Attribute Width.VB_Description = "Returns or sets the width of the splitters frame area"
  mlngRight = mlngLeft + lngWidth
End Property

' Purpose    : Returns the width of the splitters frame area
' Return     : As specified
Public Property Get Width() As Long
  Width = mlngRight - mlngLeft
End Property
