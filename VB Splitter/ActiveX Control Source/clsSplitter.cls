VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A class module to represent a ""virtual"" splitter"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************************************************************
'** File Name  : clsSplitter.cls                                              **
'** Language   : Visual Basic 6.0                                             **
'** Author     : Theo Zacharias (theo_yz@yahoo.com)                           **
'** Description: A class module to represent a virtual splitter               **
'** Usage      : VB Splitter control uses PictureBox control to for the       **
'**              splitters. Instead of using PictureBox properties directly,  **
'**              VB Splitter control saves the necessary properties plus      **
'**              several custom properties in a virtual splitter.             **
'**              Furthermore, VB Splitter control works with this virtual     **
'**              splitter (will be called splitter next) to manipulate its    **
'**              position and then write it back to the PictureBox control.   **
'** Members    :                                                              **
'**   * Properties: Bottom (r/w), ClipCursor (r/w), Enable (r/w),             **
'**                 Height (r/w), Id (def. r/w), IdCtlFriendBottom (r/w),     **
'**                 IdCtlFriendLeft (r/w), IdCtlFriendRight (r/w),            **
'**                 IdCtlFriendTop (r/w), IdsCtlBottom (r/o),                 **
'**                 IdsCtlLeft (r/o), IdsCtlRight (r/o), IdsCtlTop (r/o),     **
'**                 IdsSplBottom (r/o), IdsSplLeft (r/o), IdsSplRight (r/o),  **
'**                 IdsSplTop (r/o), Left (r/o), LiveUpdate (r/w),            **
'**                 MaxXc (r/o), MaxYc (r/o), MinXc (r/o), MinYc (r/o),       **
'**                 Orientation (r/o), Right (r/o), Top (r/o),                **
'**                 Width (r/o), Xc (r/o), Yc (r/o)                           **
'**   * Methods   : -                                                         **
'** Note       : * Left and Right properties value are generated from Xc and  **
'**                Width properties value                                     **
'**              * Top and Bottom properties value are generated from Yc and  **
'**                Height properties value                                    **
'** Last modified on September 12, 2003                                       **
'*******************************************************************************

Option Explicit

'--- Public Type Declaration
Public Enum genmSplitterOrientation
  orHorizontal
  orVertical
End Enum

'--- Property Variables

Private mId As Long
Private mudeOrientation As genmSplitterOrientation

' Property variables which appears in the property page
Private mblnClipCursor As Boolean
Private mblnEnable As Boolean
Private mblnLiveUpdate As Boolean

' Position and Size
Private mlngXc As Long
Private mlngYc As Long
Private mlngWidth As Long
Private mlngHeight As Long

' Splitter's moving region
Private mlngMinXc As Long
Private mlngMinYc As Long
Private mlngMaxXc As Long
Private mlngMaxYc As Long

' The controls which is effected by the splitter's motion
Private mIdsCtlTop As clsIds
Private mIdsCtlRight As clsIds
Private mIdsCtlBottom As clsIds
Private mIdsCtlLeft As clsIds

' The splitters which is effected by the splitter's motion
Private mIdsSplTop As clsIds
Private mIdsSplRight As clsIds
Private mIdsSplBottom As clsIds
Private mIdsSplLeft As clsIds

' Friends Controls' Id
' Note: Friend controls are the nearest control around the splitter
Private mIdCtlFriendTop As Long
Private mIdCtlFriendRight As Long
Private mIdCtlFriendBottom As Long
Private mIdCtlFriendLeft As Long

'----------------------------------
' Class Constructor and Destructor
'----------------------------------

Private Sub Class_Initialize()
  ' Initialize the MinXc and MinYc property to the minimum value and the MaxXc
  '   and MaxYc tot the maximum value to make sure that these values will be
  '   replaced with the real minimum and maximum values
  mlngMinXc = 0
  mlngMinYc = 0
  mlngMaxXc = gconLngInfinite
  mlngMaxYc = gconLngInfinite
  
  Set mIdsCtlTop = New clsIds
  Set mIdsCtlRight = New clsIds
  Set mIdsCtlBottom = New clsIds
  Set mIdsCtlLeft = New clsIds
  
  Set mIdsSplTop = New clsIds
  Set mIdsSplRight = New clsIds
  Set mIdsSplBottom = New clsIds
  Set mIdsSplLeft = New clsIds
  
  mIdCtlFriendTop = gconUninitialized
  mIdCtlFriendRight = gconUninitialized
  mIdCtlFriendBottom = gconUninitialized
  mIdCtlFriendLeft = gconUninitialized
End Sub

Private Sub Class_Terminate()
  Set mIdsCtlTop = Nothing
  Set mIdsCtlRight = Nothing
  Set mIdsCtlBottom = Nothing
  Set mIdsCtlLeft = Nothing
  
  Set mIdsSplTop = Nothing
  Set mIdsSplRight = Nothing
  Set mIdsSplBottom = Nothing
  Set mIdsSplLeft = Nothing
End Sub

'------------------
' Class Properties
'------------------

' Purpose    : Sets the distance between the internal bottom edge of the
'              splitter and the top edge of VB Splitter control
' Effect     : As specified
' Input      : lngBottom (the new Bottom property value)
Public Property Let Bottom(ByVal lngBottom As Long)
  Dim lngTop As Long
  
  lngTop = Top
  mlngHeight = lngBottom - lngTop
  mlngYc = CLng((lngTop + lngBottom) / 2)
End Property

' Purpose    : Returns the distance between the internal bottom edge of the
'              splitter and the top edge of VB Splitter control
' Return     : As specified
Public Property Get Bottom() As Long
Attribute Bottom.VB_Description = "Returns the distance between the internal bottom edge of the splitter and the top edge of VB Splitter control"
  Bottom = CLng(mlngYc + (mlngHeight / 2))
End Property

' Purpose    : Sets a value that determines whether the mouse pointer is
'              confined to the splitter's region
' Effect     : As specified
' Input      : blnClipCursor (the new ClipCursor property value)
Public Property Let ClipCursor(ByVal blnClipCursor As Boolean)
  mblnClipCursor = blnClipCursor
End Property

' Purpose    : Returns a value that determines whether the mouse pointer is
'              confined to the splitter's region
' Return     : As specified
Public Property Get ClipCursor() As Boolean
Attribute ClipCursor.VB_Description = "Returns or sets a value that determines whether the mouse pointer is confined to the splitter's region"
  ClipCursor = mblnClipCursor
End Property

' Purpose    : Sets a value that determines whether the splitter is movable
' Effect     : As specified
' Input      : blnEnable (the new Enable property value)
Public Property Let Enable(ByVal blnEnable As Boolean)
  mblnEnable = blnEnable
End Property

' Purpose    : Returns a value that determines whether the splitter is movable
' Return     : As specified
Public Property Get Enable() As Boolean
Attribute Enable.VB_Description = "Returns or sets a value that determines whether the splitter is movable"
  Enable = mblnEnable
End Property

' Purpose    : Sets the height of the splitter
' Effect     : As specified
' Input      : lngHeight(the new Height property value)
Public Property Let Height(ByVal lngHeight As Long)
  mlngHeight = lngHeight
End Property

' Purpose    : Returns the height of the splitter
' Return     : As specified
Public Property Get Height() As Long
Attribute Height.VB_Description = "Returns the height of the splitter"
  Height = mlngHeight
End Property

' Purpose    : Sets the id of the splitter
' Effect     : As specified
' Input      : lngId (the new Id property value)
Public Property Let Id(ByVal lngId As Long)
  mId = lngId
End Property

' Purpose    : Returns the id of the splitter
' Return     : As specified
Public Property Get Id() As Long
Attribute Id.VB_Description = "Returns the id of the splitter"
Attribute Id.VB_UserMemId = 0
  Id = mId
End Property

' Purpose    : Sets the id of the splitter's nearest bottom-side control
' Effect     : As specified
' Input      : lngIdCtlFriendBottom (the new IdCtlFriendBottom property values)
Public Property Let IdCtlFriendBottom(ByVal lngIdCtlFriendBottom As Long)
Attribute IdCtlFriendBottom.VB_Description = "Returns or sets the id of the splitter's nearest bottom-side control"
  mIdCtlFriendBottom = lngIdCtlFriendBottom
End Property

' Purpose    : Returns the id of the splitter's nearest bottom-side control
' Return     : As specified
Public Property Get IdCtlFriendBottom() As Long
  IdCtlFriendBottom = mIdCtlFriendBottom
End Property

' Purpose    : Sets the id of the splitter's nearest left-side control
' Effect     : As specified
' Input      : lngIdCtlFriendLeft (the new IdCtlFriendLeft property values)
Public Property Let IdCtlFriendLeft(ByVal lngIdCtlFriendLeft As Long)
Attribute IdCtlFriendLeft.VB_Description = "Returns or sets the id of the splitter's nearest left-side control"
  mIdCtlFriendLeft = lngIdCtlFriendLeft
End Property

' Purpose    : Returns the id of the splitter's nearest left-side control
' Return     : As specified
Public Property Get IdCtlFriendLeft() As Long
  IdCtlFriendLeft = mIdCtlFriendLeft
End Property

' Purpose    : Sets the id of the splitter's nearest right-side control
' Effect     : As specified
' Input      : lngIdCtlFriendRight (the new IdCtlFriendRight property values)
Public Property Let IdCtlFriendRight(ByVal lngIdCtlFriendRight As Long)
Attribute IdCtlFriendRight.VB_Description = "Returns or sets the id of the splitter's nearest right-side control"
  mIdCtlFriendRight = lngIdCtlFriendRight
End Property

' Purpose    : Returns the id of the splitter's nearest right-side control
' Return     : As specified
Public Property Get IdCtlFriendRight() As Long
  IdCtlFriendRight = mIdCtlFriendRight
End Property

' Purpose    : Sets the id of the splitter's nearest top-side control
' Effect     : As specified
' Input      : lngIdCtlFriendTop (the new IdCtlFriendTop property values)
Public Property Let IdCtlFriendTop(ByVal lngIdCtlFriendTop As Long)
Attribute IdCtlFriendTop.VB_Description = "Returns or sets the id of the splitter's nearest top-side control"
  mIdCtlFriendTop = lngIdCtlFriendTop
End Property

' Purpose    : Returns the id of the splitter's nearest top-side control
' Return     : As specified
Public Property Get IdCtlFriendTop() As Long
  IdCtlFriendTop = mIdCtlFriendTop
End Property

' Purpose    : Sets the ids of the controls which are effected by the splitter's
'              down-motion
' Effect     : As specified
' Input      : oidsCtlBottom (the new IdsCtlBottom property values)
Public Property Set IdsCtlBottom(oidsCtlBottom As clsIds)
  Set mIdsCtlBottom = oidsCtlBottom
End Property

' Purpose    : Returns the ids of the controls which are effected by the
'              splitter's down-motion
' Return     : As specified
Public Property Get IdsCtlBottom() As clsIds
Attribute IdsCtlBottom.VB_Description = "Returns the ids of the controls which are effected by the splitter's down-motion"
  Set IdsCtlBottom = mIdsCtlBottom
End Property

' Purpose    : Sets the ids of the controls which are effected by the splitter's
'              left-motion
' Effect     : As specified
' Input      : oidsCtlLeft (the new IdsCtlLeft property values)
Public Property Set IdsCtlLeft(oidsCtlLeft As clsIds)
  Set mIdsCtlLeft = oidsCtlLeft
End Property

' Purpose    : Returns the ids of the controls which are effected by the
'              splitter's left-motion
' Return     : As specified
Public Property Get IdsCtlLeft() As clsIds
Attribute IdsCtlLeft.VB_Description = "Returns the ids of the controls which are effected by the splitter's left-motion"
  Set IdsCtlLeft = mIdsCtlLeft
End Property

' Purpose    : Sets the ids of the controls which are effected by the splitter's
'              right-motion
' Effect     : As specified
' Input      : oidsCtlRight (the new IdsCtlRight property values)
Public Property Set IdsCtlRight(oidsCtlRight As clsIds)
  Set mIdsCtlRight = oidsCtlRight
End Property

' Purpose    : Returns the ids of the controls which are effected by the
'              splitter's right-motion
' Return     : As specified
Public Property Get IdsCtlRight() As clsIds
Attribute IdsCtlRight.VB_Description = "Returns the ids of the controls which are effected by the splitter's right-motion"
  Set IdsCtlRight = mIdsCtlRight
End Property

' Purpose    : Sets the ids of the controls which are effected by the splitter's
'              up-motion
' Effect     : As specified
' Input      : oidsCtlTop (the new IdsCtlTop property values)
Public Property Set IdsCtlTop(oidsCtlTop As clsIds)
  Set mIdsCtlTop = oidsCtlTop
End Property

' Purpose    : Returns the ids of the controls which are effected by the
'              splitter's up-motion
' Return     : As specified
Public Property Get IdsCtlTop() As clsIds
Attribute IdsCtlTop.VB_Description = "Returns the ids of the controls which are effected by the splitter's up-motion"
  Set IdsCtlTop = mIdsCtlTop
End Property

' Purpose    : Sets the ids of the splitters which are effected by the
'              splitter's down-motion
' Effect     : As specified
' Input      : oidsSplBottom (the new IdsSplBottom property values)
Public Property Set IdsSplBottom(oidsSplBottom As clsIds)
  Set mIdsSplBottom = oidsSplBottom
End Property

' Purpose    : Returns the ids of the splitters which are effected by the
'              splitter's down-motion
' Return     : As specified
Public Property Get IdsSplBottom() As clsIds
Attribute IdsSplBottom.VB_Description = "Returns the ids of the splitters which are effected by the splitter's down-motion"
  Set IdsSplBottom = mIdsSplBottom
End Property

' Purpose    : Sets the ids of the splitters which are effected by the
'              splitter's left-motion
' Effect     : As specified
' Input      : oidsSplLeft (the new IdsSplLeft property values)
Public Property Set IdsSplLeft(oidsSplLeft As clsIds)
  Set mIdsSplLeft = oidsSplLeft
End Property

' Purpose    : Returns the ids of the splitters which are effected by the
'              splitter's left-motion
' Return     : As specified
Public Property Get IdsSplLeft() As clsIds
Attribute IdsSplLeft.VB_Description = "Returns the ids of the splitters which are effected by the splitter's left-motion"
  Set IdsSplLeft = mIdsSplLeft
End Property

' Purpose    : Sets the ids of the splitters which are effected by the
'              splitter's right-motion
' Effect     : As specified
' Input      : oidsSplRight (the new IdsSplRight property values)
Public Property Set IdsSplRight(oidsSplRight As clsIds)
  Set mIdsSplRight = oidsSplRight
End Property

' Purpose    : Returns the ids of the splitters which are effected by the
'              splitter's right-motion
' Return     : As specified
Public Property Get IdsSplRight() As clsIds
Attribute IdsSplRight.VB_Description = "Returns the ids of the splitters which are effected by the splitter's right-motion"
  Set IdsSplRight = mIdsSplRight
End Property

' Purpose    : Sets the ids of the splitters which are effected by the
'              splitter's up-motion
' Effect     : As specified
' Input      : oidsSplTop (the new IdsSplTop property values)
Public Property Set IdsSplTop(oidsSplTop As clsIds)
  Set mIdsSplTop = oidsSplTop
End Property

' Purpose    : Returns the ids of the splitters which are effected by the
'              splitter's up-motion
' Return     : As specified
Public Property Get IdsSplTop() As clsIds
Attribute IdsSplTop.VB_Description = "Returns the ids of the splitters which are effected by the splitter's up-motion"
  Set IdsSplTop = mIdsSplTop
End Property

' Purpose    : Sets the distance between the internal left edge of the splitter
'              and the left edge of VB Splitter control
' Effect     : As specified
' Input      : lngLeft (the new Left property value)
Public Property Let Left(ByVal lngLeft As Long)
  Dim lngRight As Long
  
  lngRight = Right
  mlngWidth = lngRight - lngLeft
  mlngXc = CLng((lngLeft + lngRight) / 2)
End Property

' Purpose    : Returns the distance between the internal left edge of the
'              splitter and the left edge of VB Splitter control
' Return     : As specified
Public Property Get Left() As Long
Attribute Left.VB_Description = "Returns the distance between the internal left edge of the splitter and the left edge of VB Splitter control"
  Left = CLng(Xc - (mlngWidth / 2))
End Property

' Purpose    : Sets a value that determines whether the controls should be
'              resized as the splitter is moved
' Effect     : As specified
' Input      : blnLiveUpdate (the new LiveUpdate property value)
Public Property Let LiveUpdate(ByVal blnLiveUpdate As Boolean)
  mblnLiveUpdate = blnLiveUpdate
End Property

' Purpose    : Returns a value that determines whether the controls should be
'              resized as the splitter is moved
' Return     : As specified
Public Property Get LiveUpdate() As Boolean
Attribute LiveUpdate.VB_Description = "Returns or sets a value that determines whether the controls should be resized a the splitter is moved"
  LiveUpdate = mblnLiveUpdate
End Property

' Purpose    : Sets the maximum x-coordinate of the splitter
' Effect     : As specified
' Input      : lngMaxXc (the new MaxXc property value)
Public Property Let MaxXc(ByVal lngMaxXc As Long)
  mlngMaxXc = lngMaxXc
End Property

' Purpose    : Returns the maximum x-coordinate of the splitter
' Return     : As specified
Public Property Get MaxXc() As Long
Attribute MaxXc.VB_Description = "Returns the maximum x-coordinate of the splitter"
  MaxXc = mlngMaxXc
End Property

' Purpose    : Sets the maximum y-coordinate of the splitter
' Effect     : As specified
' Input      : lngMaxYc (the new MaxYc property value)
Public Property Let MaxYc(ByVal lngMaxYc As Long)
  mlngMaxYc = lngMaxYc
End Property

' Purpose    : Returns the maximum y-coordinate of the splitter
' Return     : As specified
Public Property Get MaxYc() As Long
Attribute MaxYc.VB_Description = "Returns the maximum y-coordinate of the splitter"
  MaxYc = mlngMaxYc
End Property

' Purpose    : Sets the minimum x-coordinate of the splitter
' Effect     : As specified
' Input      : lngMinXc (the new MinXc property value)
Public Property Let MinXc(ByVal lngMinXc As Long)
  mlngMinXc = lngMinXc
End Property

' Purpose    : Returns the minimum x-coordinate of the splitter
' Return     : As specified
Public Property Get MinXc() As Long
Attribute MinXc.VB_Description = "Returns the minimum x-coordinate of the splitter"
  MinXc = mlngMinXc
End Property

' Purpose    : Sets the minimum y-coordinate of the splitter
' Effect     : As specified
' Input      : lngMinYc (the new MinYc property value)
Public Property Let MinYc(ByVal lngMinYc As Long)
  mlngMinYc = lngMinYc
End Property

' Purpose    : Returns the minimum y-coordinate of the splitter
' Return     : As specified
Public Property Get MinYc() As Long
Attribute MinYc.VB_Description = "Returns the minimum y-coordinate of the splitter"
  MinYc = mlngMinYc
End Property

' Purpose    : Sets the splitter motion's orientation
' Effect     : As specified
' Input      : udeOrientation (the new Orientation property value)
Public Property Let Orientation(ByVal udeOrientation As genmSplitterOrientation)
  mudeOrientation = udeOrientation
End Property

' Purpose    : Returns the splitter motion's orientation
' Return     : As specified
Public Property Get Orientation() As genmSplitterOrientation
Attribute Orientation.VB_Description = "Returns the splitter motion's orientation"
  Orientation = mudeOrientation
End Property

' Purpose    : Sets the distance between the internal right edge of the splitter
'              and the left edge of VB Splitter control
' Effect     : As specified
' Input      : lngLeft (the new Left property value)
Public Property Let Right(ByVal lngRight As Long)
  Dim lngLeft As Long
  
  lngLeft = Left
  mlngWidth = lngRight - lngLeft
  mlngXc = CLng((lngLeft + lngRight) / 2)
End Property

' Purpose    : Returns the distance between the internal right edge of the
'              splitter and the left edge of VB Splitter control
' Return     : As specified
Public Property Get Right() As Long
Attribute Right.VB_Description = "Returns the distance between the internal right edge of the splitter and the left edge of VB Splitter control"
  Right = CLng(mlngXc + (mlngWidth / 2))
End Property

' Purpose    : Sets the distance between the internal top edge of the splitter
'              and the top edge of VB Splitter control
' Effect     : As specified
' Input      : lngLeft (the new Left property value)
Public Property Let Top(ByVal lngTop As Long)
  Dim lngBottom As Long
  
  lngBottom = Bottom
  mlngHeight = lngBottom - lngTop
  mlngYc = CLng((lngTop + lngBottom) / 2)
End Property

' Purpose    : Returns the distance between the internal top edge of the
'              splitter and the top edge of VB Splitter control
' Return     : As specified
Public Property Get Top() As Long
Attribute Top.VB_Description = "Returns the distance between the internal top edge of the splitter and the top edge of VB Splitter control"
  Top = CLng(mlngYc - (mlngHeight / 2))
End Property

' Purpose    : Sets the width of the splitter
' Effect     : As specified
' Input      : lngWidth (the new Width property value)
Public Property Let Width(ByVal lngWidth As Long)
  mlngWidth = lngWidth
End Property

' Purpose    : Returns the width of the splitter
' Return     : As specified
Public Property Get Width() As Long
Attribute Width.VB_Description = "Returns the width of the splitter"
  Width = mlngWidth
End Property

' Purpose    : Sets the x-coordinate of the splitter's center
' Effect     : As specified
' Input      : lngXc (the new Xc property value)
Public Property Let Xc(ByVal lngXc As Long)
  mlngXc = lngXc
End Property

' Purpose    : Returns the x-coordinate of the splitter's center
' Return     : As specified
Public Property Get Xc() As Long
Attribute Xc.VB_Description = "Returns the x-coordinate of the splitter's center"
  Xc = mlngXc
End Property

' Purpose    : Sets the y-coordinate of the splitter's center
' Effect     : As specified
' Input      : lngYc (the new Yc property value)
Public Property Let Yc(ByVal lngYc As Long)
  mlngYc = lngYc
End Property

' Purpose    : Returns the y-coordinate of the splitter's center
' Return     : As specified
Public Property Get Yc() As Long
Attribute Yc.VB_Description = "Returns the y-coordinate of the splitter's center"
  Yc = mlngYc
End Property
