VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A class module to represent a ""virtual"" control"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************************************************************
'** File Name  : clsControl.cls                                               **
'** Language   : Visual Basic 6.0                                             **
'** Author     : Theo Zacharias (theo_yz@yahoo.com)                           **
'** Description: A class module to represent a virtual control                **
'** Usage      : Instead of using a real control (like TextBox control)       **
'**              properties directly, VB Splitter control saves the necessary **
'**              properties plus several custom properties in a virtual       **
'**              control. Furthermore, VB Splitter control works with this    **
'**              virtual control (will be called control next) to             **
'**              manipulate its position and size and then write it back to   **
'**              the related real control.                                    **
'** Members    :                                                              **
'**   * Properties: Bottom (r/w), Height (r/w), Id (def. r/w),                **
'**                 IdCtlFriendBottom (r/w), IdCtlFriendLeft (r/w),           **
'**                 IdCtlFriendRight (r/w), IdCtlFriendTop (r/w),             **
'**                 IdSplBottom (r/w), IdSplLeft (r/w), IdSplRight (r/w),     **
'**                 IdSplTop (r/w), Left (r/w), MinHeight (r/w),              **
'**                 MinWidth (r/w), Right (r/w), Top (r/w), Width (r/w),      **
'**                 Xc (r/o), Yc (r/o)                                        **
'**   * Methods   : -                                                         **
'** Note       : * Height and Yc properties value are generated from Bottom   **
'**                and Top properties value                                   **
'**              * Width and Xc properties value are generated from Left and  **
'**                Right properties value                                     **
'** Last modified on September 11, 2003                                       **
'*******************************************************************************

Option Explicit

'--- Property Variables

Private mId As Long

' Position and Size
Private mlngLeft As Long
Private mlngTop As Long
Private mlngRight As Long
Private mlngBottom As Long

' Minimum Size
Private mlngMinWidth As Long
Private mlngMinHeight As Long

' Friends Controls' Id
' Note: Friend controls are the nearest control around the control
Private mIdCtlFriendTop As Long
Private mIdCtlFriendRight As Long
Private mIdCtlFriendBottom As Long
Private mIdCtlFriendLeft As Long

' Friend Splitters' Id
' Note: Friend splitters (called splitters in this class) are the splitters
'       around the control
Private mIdSplTop As Long
Private mIdSplRight As Long
Private mIdSplBottom As Long
Private mIdSplLeft As Long

'-------------------
' Class Constructor
'-------------------

Private Sub Class_Initialize()
  mlngMinWidth = 0
  mlngMinHeight = 0
  
  mIdCtlFriendTop = gconUninitialized
  mIdCtlFriendRight = gconUninitialized
  mIdCtlFriendBottom = gconUninitialized
  mIdCtlFriendLeft = gconUninitialized
  
  mIdSplTop = gconUninitialized
  mIdSplRight = gconUninitialized
  mIdSplBottom = gconUninitialized
  mIdSplLeft = gconUninitialized
End Sub

'------------------
' Class Properties
'------------------

' Purpose    : Sets the distance between the internal bottom edge of the control
'              and the top edge of VB Splitter control
' Effect     : As specified
' Input      : lngBottom (the new Bottom property value)
Public Property Let Bottom(ByVal lngBottom As Long)
  mlngBottom = lngBottom
End Property

' Purpose    : Returns the distance between the internal bottom edge of the
'              control and the top edge of VB Splitter control
' Return     : As specified
Public Property Get Bottom() As Long
Attribute Bottom.VB_Description = "Returns the distance between the internal bottom edge of the control and the top edge of Splitter control"
  Bottom = mlngBottom
End Property

' Purpose    : Sets the height of the control
' Effect     : As specified
' Input      : lngHeight(the new Height property value)
Public Property Let Height(ByVal lngHeight As Long)
  mlngBottom = mlngTop + lngHeight
End Property

' Purpose    : Returns the height of the control
' Return     : As specified
Public Property Get Height() As Long
Attribute Height.VB_Description = "Returns the height of the control"
  Height = mlngBottom - mlngTop
End Property

' Purpose    : Sets the id of the control
' Effect     : As specified
' Input      : lngId (the new Id property value)
Public Property Let Id(ByVal lngId As Long)
  mId = lngId
End Property

' Purpose    : Returns the id of the control
' Return     : As specified
Public Property Get Id() As Long
Attribute Id.VB_Description = "Returns the id of the control"
Attribute Id.VB_UserMemId = 0
  Id = mId
End Property

' Purpose    : Sets the control's bottom-side friend control's Id
' Effect     : As specified
' Input      : lngIdCtlFriendBottom (the new IdCtlFriendBottom property value)
Public Property Let IdCtlFriendBottom(ByVal lngIdCtlFriendBottom As Long)
Attribute IdCtlFriendBottom.VB_Description = "Returns or sets the control's bottom-side friend control's Id"
  mIdCtlFriendBottom = lngIdCtlFriendBottom
End Property

' Purpose    : Returns the control's bottom-side friend control's Id
' Return     : As specified
Public Property Get IdCtlFriendBottom() As Long
  IdCtlFriendBottom = mIdCtlFriendBottom
End Property

' Purpose    : Sets the control's left-side friend control's Id
' Effect     : As specified
' Input      : lngIdCtlFriendLeft (the new IdCtlFriendLeft property value)
Public Property Let IdCtlFriendLeft(ByVal lngIdCtlFriendLeft As Long)
Attribute IdCtlFriendLeft.VB_Description = "Returns or sets the control's left-side friend control's Id"
  mIdCtlFriendLeft = lngIdCtlFriendLeft
End Property

' Purpose    : Returns the control's left-side friend control's Id
' Return     : As specified
Public Property Get IdCtlFriendLeft() As Long
  IdCtlFriendLeft = mIdCtlFriendLeft
End Property

' Purpose    : Sets the control's right-side friend control's Id
' Effect     : As specified
' Input      : lngIdCtlFriendRight (the new IdCtlFriendRight property value)
Public Property Let IdCtlFriendRight(ByVal lngIdCtlFriendRight As Long)
Attribute IdCtlFriendRight.VB_Description = "Returns or sets the control's right-side friend control's Id"
  mIdCtlFriendRight = lngIdCtlFriendRight
End Property

' Purpose    : Returns the control's right-side friend control's Id
' Return     : As specified
Public Property Get IdCtlFriendRight() As Long
  IdCtlFriendRight = mIdCtlFriendRight
End Property

' Purpose    : Sets the control's top-side friend control's Id
' Input      : lngIdCtlFriendTop (the new IdCtlFriendTop property value)
Public Property Let IdCtlFriendTop(ByVal lngIdCtlFriendTop As Long)
Attribute IdCtlFriendTop.VB_Description = "Returns or sets the control's top-side friend control's Id"
  mIdCtlFriendTop = lngIdCtlFriendTop
End Property

' Purpose    : Returns the control's top-side friend control's Id
Public Property Get IdCtlFriendTop() As Long
  IdCtlFriendTop = mIdCtlFriendTop
End Property

' Purpose    : Sets the control's bottom-side friend splitter's Id
' Effect     : As specified
' Input      : lngIdSplFriendBottom (the new IdSplFriendBottom property value)
Public Property Let IdSplBottom(ByVal lngIdSplBottom As Long)
  mIdSplBottom = lngIdSplBottom
End Property

' Purpose    : Returns the control's bottom-side friend splitter's Id
' Return     : As specified
Public Property Get IdSplBottom() As Long
Attribute IdSplBottom.VB_Description = "Returns the control's bottom-side friend splitter's Id"
  IdSplBottom = mIdSplBottom
End Property

' Purpose    : Sets the control's left-side friend splitter's Id
' Effect     : As specified
' Input      : lngIdSplFriendLeft (the new IdSplFriendLeft property value)
Public Property Let IdSplLeft(ByVal lngIdSplLeft As Long)
  mIdSplLeft = lngIdSplLeft
End Property

' Purpose    : Returns the control's left-side friend splitter's Id
' Return     : As specified
Public Property Get IdSplLeft() As Long
Attribute IdSplLeft.VB_Description = "Returns the control's left-side friend splitter's Id"
  IdSplLeft = mIdSplLeft
End Property

' Purpose    : Sets the control's right-side friend splitter's Id
' Effect     : As specified
' Input      : lngIdSplFriendRight (the new IdSplFriendRight property value)
Public Property Let IdSplRight(ByVal lngIdSplRight As Long)
  mIdSplRight = lngIdSplRight
End Property

' Purpose    : Returns the control's right-side friend splitter's Id
' Return     : As specified
Public Property Get IdSplRight() As Long
Attribute IdSplRight.VB_Description = "Returns the control's right-side friend splitter's Id"
  IdSplRight = mIdSplRight
End Property

' Purpose    : Sets the control's top-side friend splitter's Id
' Effect     : As specified
' Input      : lngIdSplFriendTop (the new IdSplFriendTop property value)
Public Property Let IdSplTop(ByVal lngIdSplTop As Long)
  mIdSplTop = lngIdSplTop
End Property

' Purpose    : Returns the control's top-side friend splitter's Id
' Return     : As specified
Public Property Get IdSplTop() As Long
Attribute IdSplTop.VB_Description = "Returns the control's top-side friend splitter's Id"
  IdSplTop = mIdSplTop
End Property

' Purpose    : Sets the distance between the internal left edge of the control
'              and the left edge of VB Splitter control
' Effect     : As specified
' Input      : lngLeft (the new Left property value)
Public Property Let Left(ByVal lngLeft As Long)
  mlngLeft = lngLeft
End Property

' Purpose    : Returns the distance between the internal left edge of the
'              control and the left edge of VB Splitter control
' Return     : As specified
Public Property Get Left() As Long
Attribute Left.VB_Description = "Returns the distance between the internal left edge of the control and the left edge of Splitter control"
  Left = mlngLeft
End Property

' Purpose    : Sets the minimum height of the control
' Effect     : As specified
' Input      : lngMinHeight(the new MinHeight propety value)
Public Property Let MinHeight(ByVal lngMinHeight As Long)
  mlngMinHeight = lngMinHeight
End Property

' Purpose    : Returns the minimum height of the control
' Return     : As specified
Public Property Get MinHeight() As Long
Attribute MinHeight.VB_Description = "Returns or sets the minimum height of the control"
  MinHeight = mlngMinHeight
End Property

' Purpose    : Sets the minimum width of the control
' Effect     : As specified
' Input      : lngMinWidth(the new MinWidth propety value)
Public Property Let MinWidth(ByVal lngMinWidth As Long)
  mlngMinWidth = lngMinWidth
End Property

' Purpose    : Returns the minimum width of the control
' Return     : As specified
Public Property Get MinWidth() As Long
Attribute MinWidth.VB_Description = "Returns or sets the minimum width of the control"
  MinWidth = mlngMinWidth
End Property

' Purpose    : Sets the distance between the internal right edge of the control
'              and the left edge of VB Splitter control
' Effect     : As specified
' Input      : lngRight (the new Right property value)
Public Property Let Right(ByVal lngRight As Long)
  mlngRight = lngRight
End Property

' Purpose    : Returns the distance between the internal right edge of the
'              control and the left edge of VB Splitter control
' Return     : As specified
Public Property Get Right() As Long
Attribute Right.VB_Description = "Returns the distance between the internal right edge of the control and the left edge of Splitter control"
  Right = mlngRight
End Property

' Purpose    : Sets the distance between the internal top edge of the control
'              and the top edge of VB Splitter control
' Effect     : As specified
' Input      : lngTop (the new Top property value)
Public Property Let Top(ByVal lngTop As Long)
  mlngTop = lngTop
End Property

' Purpose    : Returns the distance between the internal top edge of the control
'              and the top edge of VB Splitter control
' Return     : As specified
Public Property Get Top() As Long
Attribute Top.VB_Description = "Returns the distance between the internal top edge of the control and the top edge of Splitter control"
  Top = mlngTop
End Property

' Purpose    : Sets the width of the control
' Effect     : As specified
' Input      : lngWidth(the new Width property value)
Public Property Let Width(ByVal lngWidth As Long)
  mlngRight = mlngLeft + lngWidth
End Property

' Purpose    : Returns the width of the control
' Return     : As specified
Public Property Get Width() As Long
Attribute Width.VB_Description = "Returns the width of the control"
  Width = mlngRight - mlngLeft
End Property

' Purpose    : Returns the x-coordinate of the control's center
' Return     : As specified
Public Property Get Xc() As Long
Attribute Xc.VB_Description = "Returns the x-coordinate of the control's center"
  Xc = CLng((mlngLeft + mlngRight) / 2)
End Property

' Purpose    : Returns the y-coordinate of the control's center
' Return     : As specified
Public Property Get Yc() As Long
Attribute Yc.VB_Description = "Returns the y-coordinate of the control's center"
  Yc = CLng((mlngTop + mlngBottom) / 2)
End Property
