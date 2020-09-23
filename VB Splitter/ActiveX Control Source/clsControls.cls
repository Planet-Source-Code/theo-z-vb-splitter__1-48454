VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A class module to implement custom collection of class clsControl"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************************************************************
'** File Name  : clsControls.cls                                              **
'** Language   : Visual Basic 6.0                                             **
'** Author     : Theo Zacharias (theo_yz@yahoo.com)                           **
'** Description: A class module to implement custom collection of class       **
'**              clsControl                                                   **
'** Usage      : Instead of using a real control (like TextBox control)       **
'**              properties directly, VB Splitter control saves the necessary **
'**              properties plus several custom properties in a virtual       **
'**              control. Furthermore, VB Splitter control works with this    **
'**              virtual control (next, will be called control) to            **
'**              manipulate its position and size and then write it back to   **
'**              the related real control.                                    **
'** Members    :                                                              **
'**   * Properties: Bottom (r/w), Count (r/o), Height (r/w), Item (def. r/o), **
'**                 Left (r/w), Right (r/w), Top (r/w), Width (r/w)           **
'**   * Methods   : Add, Compact, IsValid, Remove, RemoveHeaps, RemoveHoles,  **
'**                 Stretch                                                   **
'** Note       : * Height properties value are generated from Bottom and Top  **
'**                properties value                                           **
'**              * Width properties value are generated from Left and Right   **
'**                properties value                                           **
'** Last modified on September 11, 2003                                       **
'*******************************************************************************

Option Explicit

'--- Custom Collection Class Variable
Private mcolControls As Collection

'--- Properties Variables

' Developer's Controls Frame Area Position and Size
' Note: These properties below defined the developer's frame of controls
'       collection. It is named developer's frame because the frame area is
'       determined by the developer. After calling Stretch method, the current
'       frame will have the same size of area with the developer's frame's
Private mlngLeft As Long
Private mlngTop As Long
Private mlngRight As Long
Private mlngBottom As Long

' Current Controls Frame Area Position and Size
' Note: These properties below defined the current frame of controls collection.
'       Current frame is the minimum frame area needed to cover all control in
'       controls collection. The frame area is automatically adjusted every time
'       a new control is added in controls collection. These properties are read
'       only.
Private mlngCurLeft As Long
Private mlngCurTop As Long
Private mlngCurRight As Long
Private mlngCurBottom As Long

'----------------------------------
' Class Constructor and Destructor
'----------------------------------

Private Sub Class_Initialize()
  Set mcolControls = New Collection
  
  ' Initialize the left and top to the maximum value and the right and bottom to
  '   the minimum value to make sure that these values will be replaced with
  '   minimum left and top value and maximum right and bottom value
  mlngCurLeft = gconLngInfinite
  mlngCurTop = gconLngInfinite
  mlngCurRight = 0
  mlngCurBottom = 0
End Sub

Private Sub Class_Terminate()
  Set mcolControls = Nothing
End Sub

'-----------------------------------------
' Collection Class Properties and Methods
'-----------------------------------------

' Purpose    : Adds a new item to the collection
' Assumption : Controls with index IdCtl exists in cctl collection
' Effect     : * If control cctl(IdCtl) doesn't have Left, Top, Width or Height
'                property in run-time mode, then this method doesn't have any
'                effect, i.e. no item will be added to the collection
'              * Otherwise, the new item (which is a virtual control of
'                control cctl(IdCtl) has been added to the collection and its
'                position, size and minimum height/width properties have been
'                initialized
' Inputs     : * cctl (controls collection contained in VB Splitter control)
'              * IdCtl (Index of cctl that will be added to the collection)
' Note       : To get the minimum height and width of control cctl(IdCtl), I set
'              its height and width to 0 (if the minimum height/width of the
'              control is larger than 0, control cctl(IdCtl) will automatically
'              adjust its height/value to its minimum value), save control
'              cctl(IdCtl)'s height and width in the control's MinWidth and
'              MinHeight property and then restore control cctl(IdCtl)'s height
'              and width. Note that if control cctl(IdCtl) is visible when you
'              add it to the collection, it may produce a flickering effect.
Public Sub Add(ByVal cctl As ContainedControls, ByVal IdCtl As Long)
Attribute Add.VB_Description = "Adds a new item to the collection"
  Dim lngWidthSave As Long                 'to restore control cctl(IdCtl) width
  Dim lngHeightSave As Long               'to restore control cctl(IdCtl) height
  Dim octlNew As clsControl  'the new item which will be added to the collection
  
  Set octlNew = New clsControl
  octlNew.Id = IdCtl
  
  ' Sets the item's position and size
  ' Note: The error handler is used incase control cctl(IdCtl) doesn't have
  '       Left, Top, Width or Height property in run-time mode.
  On Error GoTo ErrorHandler
  octlNew.Left = cctl(IdCtl).Left
  octlNew.Top = cctl(IdCtl).Top
  octlNew.Width = cctl(IdCtl).Width
  octlNew.Height = cctl(IdCtl).Height
  On Error GoTo 0
  
  ' Determines the control minimum width and height
  ' Note: The error handler is used incase control cctl(IdCtl)'s height or
  '       width is r/o
  On Error Resume Next
  cctl(IdCtl).Visible = False
  lngWidthSave = cctl(IdCtl).Width
  cctl(IdCtl).Width = 0
  octlNew.MinWidth = cctl(IdCtl).Width
  cctl(IdCtl).Width = lngWidthSave
  lngHeightSave = cctl(IdCtl).Height
  cctl(IdCtl).Height = 0
  octlNew.MinHeight = cctl(IdCtl).Height
  cctl(IdCtl).Height = lngHeightSave
  cctl(IdCtl).Visible = True
  On Error GoTo 0
  
  '-- Adds to the collection
  mcolControls.Add Item:=octlNew, Key:=CStr(IdCtl)
  
  '-- Adjusts the current controls frame area
  If octlNew.Left < mlngCurLeft Then mlngCurLeft = octlNew.Left
  If octlNew.Top < mlngCurTop Then mlngCurTop = octlNew.Top
  If octlNew.Right > mlngCurRight Then mlngCurRight = octlNew.Right
  If octlNew.Bottom > mlngCurBottom Then mlngCurBottom = octlNew.Bottom
  
ErrorHandler:
  Set octlNew = Nothing
End Sub

' Purpose    : Returns the number of items in the collection
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of items in the collection"
  Count = mcolControls.Count
End Property

' Purpose    : Returns an item in the collection which has key IdCtl
' Assumption : Key IdCtl exists in the collection
' Input      : IdCtl
' Return     : As specified
Public Property Get Item(ByVal IdCtl As Long) As clsControl
Attribute Item.VB_Description = "Returns an item in the collection which has key IdCtl"
Attribute Item.VB_UserMemId = 0
  Set Item = mcolControls.Item(CStr(IdCtl))
End Property

' Purpose    : Enables For Each ... Next enumeration
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mcolControls.[_NewEnum]
End Property

' Purpose    : Removes an item which has key IdCtl from the collection
' Assumption : Key IdCtl exists in the collection
' Effect     : As specified
' Input      : IdCtl
Public Sub Remove(ByVal IdCtl As Long)
Attribute Remove.VB_Description = "Removes an item which has key IdCtl from the collection"
  mcolControls.Remove CStr(IdCtl)
End Sub

'------------------
' Class Properties
'------------------

' Purpose    : Sets the distance between the bottom edge of the developer's
'              controls frame area and the top edge of VB Splitter control
' Effect     : As specified
' Input      : lngBottom (the new Bottom propety value)
Public Property Let Bottom(ByVal lngBottom As Long)
  mlngBottom = lngBottom
End Property

' Purpose    : Returns the distance between the bottom edge of the developer's
'              controls frame area and the top edge of VB Splitter control
' Return     : As specified
Public Property Get Bottom() As Long
Attribute Bottom.VB_Description = "Returns the distance between the bottom edge of the developer's controls frame area and the top edge of Splitter control"
  Bottom = mlngBottom
End Property

' Purpose    : Sets the height of the developer's controls frame area
' Effect     : As specified
' Input      : lngHeight(the new Height property value)
Public Property Let Height(ByVal lngHeight As Long)
  mlngBottom = mlngTop + lngHeight
End Property

' Purpose    : Returns the height of the developer's controls frame area
' Return     : As specified
Public Property Get Height() As Long
Attribute Height.VB_Description = "Returns the height of the developer's controls frame area"
  Height = mlngBottom - mlngTop
End Property

' Purpose    : Sets the distance between the left edge of the developer's
'              controls frame area and the left edge of VB Splitter control
' Effect     : As specified
' Input      : lngLeft (the new Left property value)
Public Property Let Left(ByVal lngLeft As Long)
  mlngLeft = lngLeft
End Property

' Purpose    : Returns the distance between the left edge of the developer's
'              controls frame area and the left edge of VB Splitter control
' Return     : As specified
Public Property Get Left() As Long
Attribute Left.VB_Description = "Returns the distance between the left edge of the developer's controls frame area and the left edge of Splitter control"
  Left = mlngLeft
End Property

' Purpose    : Sets the distance between the right edge of the developer's
'              controls frame area and the left edge of VB Splitter control
' Effect     : As specified
' Input      : lngRight (the new Right property value)
Public Property Let Right(ByVal lngRight As Long)
  mlngRight = lngRight
End Property

' Purpose    : Returns the distance between the right edge of the developer's
'              controls frame area and the left edge of VB Splitter control
' Return     : As specified
Public Property Get Right() As Long
Attribute Right.VB_Description = "Returns the distance between the right edge of the developer's controls frame area and the left edge of Splitter control"
  Right = mlngRight
End Property

' Purpose    : Sets the distance between the top edge of the developer's
'              controls frame area and the top edge of VB Splitter control
' Effect     : As specified
' Input      : lngTop (the new Top property value)
Public Property Let Top(ByVal lngTop As Long)
  mlngTop = lngTop
End Property

' Purpose    : Returns the distance between the top edge of the developer's
'              controls frame area and the top edge of VB Splitter control
' Return     : As specified
Public Property Get Top() As Long
Attribute Top.VB_Description = "Returns the distance between the top edge of the developer's controls frame area and the top edge of Splitter control"
  Top = mlngTop
End Property

' Purpose    : Sets the width of the developer's controls frame area
' Effect     : As specified
' Input      : lngWidth(the new Width property value)
Public Property Let Width(ByVal lngWidth As Long)
  mlngRight = mlngLeft + lngWidth
End Property

' Purpose    : Returns the width of the developer's controls frame area
' Return     : As specified
Public Property Get Width() As Long
Attribute Width.VB_Description = "Returns the width of the developer's controls frame area"
  Width = mlngRight - mlngLeft
End Property

'---------------
' Class Methods
'---------------

' Purpose    : Compacts each item in the collection so there won't be any space
'              left
' Effect     : As specified
' Note       : Even after calling this method, it may still left special spaces
'              called "holes" that will be removed with RemoveHoles method. See
'              VB Splitter control's documentation or help for the definition of
'              "hole".
Public Sub Compact()
Attribute Compact.VB_Description = "Compacts each item in the collection so there won't be any space left"
  Dim lngXYMeet As Long           'the x- or y-coordinate where the edges of two
                                  '      processed controls will be compacted to
  Dim octl As clsControl            'for enumerating all items in the collection
  
  For Each octl In Me
    With octl
      '-- Compact control octl's top
      ' Sets top-side friend control
      SetFriend .Id
      If .IdCtlFriendTop = gconUninitialized Then
        ' Control octl doesn't have top-side friend, so set the control's top
        '   to current frame area's top
        .Top = mlngCurTop
      Else
        ' Determines lngXYMeet value: where control octl and its top-side friend
        '   will be compacted to
        SetFriend .IdCtlFriendTop
        lngXYMeet = (.Top + Item(.IdCtlFriendTop).Bottom) \ 2
        If Item(.IdCtlFriendTop).IdCtlFriendBottom <> .Id Then
          lngXYMeet = _
            IIf(lngXYMeet < Item(Item(.IdCtlFriendTop).IdCtlFriendBottom).Top, _
                lngXYMeet, Item(Item(.IdCtlFriendTop).IdCtlFriendBottom).Top)
        End If
        ' Compacts control octl's top and its top-side friend's bottom
        .Top = lngXYMeet
        Item(.IdCtlFriendTop).Bottom = lngXYMeet
      End If
      
      '-- Compact control octl's right
      ' Set right-side friend control (must do SetFriend again to compensate
      '   the right-side friend change possibility because of previous compact)
      SetFriend .Id
      If .IdCtlFriendRight = gconUninitialized Then
        ' Control octl doesn't have eight-side friend, so set the control's
        '   right to current frame area's right
        .Right = mlngCurRight
      Else
        ' Determine lngXYMeet value: where control octl and its right-side
        '   friend will be compacted to
        SetFriend .IdCtlFriendRight
        lngXYMeet = (.Right + Item(.IdCtlFriendRight).Left) \ 2
        If Item(.IdCtlFriendRight).IdCtlFriendLeft <> .Id Then
          lngXYMeet = _
            IIf(lngXYMeet > _
                Item(Item(.IdCtlFriendRight).IdCtlFriendLeft).Right, _
                lngXYMeet, Item(Item(.IdCtlFriendRight).IdCtlFriendLeft).Right)
        End If
        ' Compacts control octl's right and its right-side friend's left
        .Right = lngXYMeet
        Item(.IdCtlFriendRight).Left = lngXYMeet
      End If
      
      '-- Compact control octl's bottom
      ' Sets bottom-side friend control (must do SetFriend again to compensate
      '   the bottom-side friend change possibility because of previous compact)
      SetFriend .Id
      If .IdCtlFriendBottom = gconUninitialized Then
        ' Control octl doesn't have bottom-side friend, so set the control's
        '   bottom to current frame area's bottom
        .Bottom = mlngCurBottom
      Else
        ' Determines lngXYMeet value: where control octl and its bottom-side
        '   friend will be compacted to
        SetFriend .IdCtlFriendBottom
        lngXYMeet = (.Bottom + Item(.IdCtlFriendBottom).Top) \ 2
        If Item(.IdCtlFriendBottom).IdCtlFriendTop <> .Id Then
          lngXYMeet = _
            IIf(lngXYMeet > _
                Item(Item(.IdCtlFriendBottom).IdCtlFriendTop).Bottom, _
                lngXYMeet, Item(Item(.IdCtlFriendBottom).IdCtlFriendTop).Bottom)
        End If
        ' Compacts control octl's bottom and its bottom-side friend's top
        .Bottom = lngXYMeet
        Item(.IdCtlFriendBottom).Top = lngXYMeet
      End If
      
      '-- Compact control octl's left
      ' Sets left-side friend control (must do SetFriend again to compensate
      '   the left-side friend change possibility because of previous compact)
      SetFriend .Id
      If .IdCtlFriendLeft = gconUninitialized Then
        ' Control octl doesn't have left-side friend, so set the control's
        '   left to current frame area's left
        .Left = mlngCurLeft
      Else
        ' Determines lngXYMeet value: where control octl and its left-side friend
        '   will be compacted to
        SetFriend .IdCtlFriendLeft
        lngXYMeet = (.Left + Item(.IdCtlFriendLeft).Right) \ 2
        If Item(.IdCtlFriendLeft).IdCtlFriendRight <> .Id Then
          lngXYMeet = _
            IIf(lngXYMeet < _
                Item(Item(.IdCtlFriendLeft).IdCtlFriendRight).Left, _
                lngXYMeet, Item(Item(.IdCtlFriendLeft).IdCtlFriendRight).Left)
        End If
        ' Compacts control octl's left and its left-side friend's right
        .Left = lngXYMeet
        Item(.IdCtlFriendLeft).Right = lngXYMeet
      End If
    End With
  Next
End Sub

' Purpose    : Returns value indicating whether each item's width and height in
'              the collection are not smaller than their minimum value
' Return     : As specified
Public Function IsValid() As Boolean
Attribute IsValid.VB_Description = "Returns value indicating whether each item's width and height in the collection are not smaller than their minimum value"
  Dim blnIsValid As Boolean                                      'returned value
  Dim octl As clsControl            'for enumerating all items in the collection
  
  blnIsValid = True
  For Each octl In Me
    If (octl.Width < octl.MinWidth) Or (octl.Height < octl.MinHeight) Then
      blnIsValid = False
      Exit For
    End If
  Next
  IsValid = blnIsValid
End Function

' Purpose    : Removes all items' heap
' Effect     : As specified
Public Sub RemoveHeaps()
Attribute RemoveHeaps.VB_Description = "Removes all items' heap "
  Dim octl As clsControl            'for enumerating all items in the collection
  
  For Each octl In Me
    RemoveHeap octl.Id
  Next
End Sub

' Purpose    : Removes all "holes" in the collection
' Effect     : As specified
' Note       : Even after calling this method, the collection may still left
'              "holes" for particular complex position of the items in the
'              collection. See VB Splitter control's documentation or help for
'              the definition of "hole" and explanation of complex position
'              which VB Splitter couldn't handle.
Public Sub RemoveHoles()
Attribute RemoveHoles.VB_Description = "Removes all ""holes"" in the collection"
  Dim blnSuccess As Boolean            'indicating whether the remove action for
                                       '          certain direction is successed
  Dim IdTemp As Long                     'to re-arange the "anti-clockwise hole"
  Dim octl As clsControl            'for enumerating all items in the collection
  Dim octlHole As clsControl            'to represent the hole position and size
  
  SetFriends
  For Each octl In Me
    Set octlHole = New clsControl
    With octlHole
      '-- Determines whether item octl is one of four items that construct the
      '   hole
      .IdCtlFriendRight = octl.IdCtlFriendBottom
      If .IdCtlFriendRight <> gconUninitialized Then
        .IdCtlFriendBottom = Item(.IdCtlFriendRight).IdCtlFriendLeft
        If .IdCtlFriendBottom <> gconUninitialized Then
          .IdCtlFriendLeft = Item(.IdCtlFriendBottom).IdCtlFriendTop
          If .IdCtlFriendLeft <> gconUninitialized Then
            .IdCtlFriendTop = Item(.IdCtlFriendLeft).IdCtlFriendRight
          End If
        End If
      End If
      If (.IdCtlFriendTop = octl.Id) Then
        '-- There is a hole and item Octl is one of four items that construct
        '   the hole. The next step is to try to remove the hole.
        '- If the hole is "anti-clockwise" then re-arange the four items' Id that
        '    consruct the hole.
        If Item(.IdCtlFriendTop).Left > Item(.IdCtlFriendRight).Left Then
          IdTemp = .IdCtlFriendTop
          .IdCtlFriendTop = .IdCtlFriendLeft
          .IdCtlFriendLeft = .IdCtlFriendBottom
          .IdCtlFriendBottom = .IdCtlFriendRight
          .IdCtlFriendRight = IdTemp
        End If
        '- If the hole is not "empty", don't remove the hole
        blnSuccess = False
        If IsEmptyHole(octlHole) Then
          '- The hole is not empty, so it must be removed
          ' Try to stretch-down the top-side-hole-item
          If Not blnSuccess Then
            If Item(.IdCtlFriendTop).Right > Item(.IdCtlFriendBottom).Right Then
              If Item(.IdCtlFriendTop).Right = _
                 Item(.IdCtlFriendRight).Right Then
                Item(.IdCtlFriendTop).Bottom = Item(.IdCtlFriendBottom).Top
                Item(.IdCtlFriendRight).Top = Item(.IdCtlFriendBottom).Top
                blnSuccess = True
              End If
            Else
              If Item(.IdCtlFriendTop).Left = Item(.IdCtlFriendLeft).Left Then
                Item(.IdCtlFriendTop).Bottom = Item(.IdCtlFriendBottom).Top
                Item(.IdCtlFriendLeft).Top = Item(.IdCtlFriendBottom).Top
                blnSuccess = True
              End If
            End If
          End If
          ' If it's not successed, try to stretch-left the right-side-hole-item
          If Not blnSuccess Then
            If Item(.IdCtlFriendRight).Bottom > _
               Item(.IdCtlFriendLeft).Bottom Then
              If Item(.IdCtlFriendRight).Bottom = _
                 Item(.IdCtlFriendBottom).Bottom Then
                Item(.IdCtlFriendRight).Left = Item(.IdCtlFriendLeft).Right
                Item(.IdCtlFriendBottom).Right = Item(.IdCtlFriendLeft).Right
                blnSuccess = True
              End If
            Else
              If Item(.IdCtlFriendRight).Top = Item(.IdCtlFriendTop).Top Then
                Item(.IdCtlFriendRight).Left = Item(.IdCtlFriendLeft).Right
                Item(.IdCtlFriendTop).Right = Item(.IdCtlFriendLeft).Right
                blnSuccess = True
              End If
            End If
          End If
          ' If it's not successed, try to stretch-up the bottom-side-hole-item
          If Not blnSuccess Then
            If Item(.IdCtlFriendBottom).Left < Item(.IdCtlFriendTop).Left Then
              If Item(.IdCtlFriendBottom).Left = _
                 Item(.IdCtlFriendLeft).Left Then
                Item(.IdCtlFriendBottom).Top = Item(.IdCtlFriendTop).Bottom
                Item(.IdCtlFriendLeft).Bottom = Item(.IdCtlFriendTop).Bottom
                blnSuccess = True
              End If
            Else
              If Item(.IdCtlFriendBottom).Right = _
                 Item(.IdCtlFriendRight).Right Then
                Item(.IdCtlFriendBottom).Top = Item(.IdCtlFriendTop).Bottom
                Item(.IdCtlFriendRight).Bottom = Item(.IdCtlFriendTop).Bottom
                blnSuccess = True
              End If
            End If
          End If
          ' If it's not successed, try to stretch-right the left-side-hole-item
          If Not blnSuccess Then
            If Item(.IdCtlFriendLeft).Top < Item(.IdCtlFriendRight).Top Then
              If Item(.IdCtlFriendLeft).Top = Item(.IdCtlFriendTop).Top Then
                Item(.IdCtlFriendLeft).Right = Item(.IdCtlFriendRight).Left
                Item(.IdCtlFriendTop).Left = Item(.IdCtlFriendRight).Left
                blnSuccess = True
              End If
            Else
              If Item(.IdCtlFriendLeft).Bottom = _
                 Item(.IdCtlFriendBottom).Bottom Then
                Item(.IdCtlFriendLeft).Right = Item(.IdCtlFriendRight).Left
                Item(.IdCtlFriendBottom).Left = Item(.IdCtlFriendRight).Left
              End If
            End If
          End If
        End If
      End If
    End With
    Set octlHole = Nothing
  Next
End Sub

' Purpose    : Stretchs the items to fill-up the developer's frame area
' Assumption : The current frame area's width and height is greater than 0
' Effects    : * The current frame area's size is equal with the developer's
'                frame area's size
'              * Others, as specified
Public Sub Stretch()
Attribute Stretch.VB_Description = "Stretchs the items to fill-up the developer's frame area"
  Dim lngdXCorner As Long   'the distance between the left edge of current frame
                            '   area and the left edge of developer's frame area
  Dim lngdYCorner As Long    'the distance between the top edge of current frame
                             '   area and the top edge of developer's frame area
  Dim octl As clsControl            'for enumerating all items in the collection
  Dim sngXScale As Single                 'the width scale of current frame area
                                          '   relative to developer's frame area
  Dim sngYScale As Single                'the height scale of current frame area
                                         '    relative to developer's frame area
  
  lngdXCorner = mlngCurLeft - mlngLeft
  lngdYCorner = mlngCurTop - mlngTop
  sngXScale = Width / (mlngCurRight - mlngCurLeft)
  sngYScale = Height / (mlngCurBottom - mlngCurTop)
  For Each octl In Me
    With octl
      .Left = CLng((.Left - lngdXCorner) * sngXScale)
      .Top = CLng((.Top - lngdYCorner) * sngYScale)
      .Right = CLng((.Right - lngdXCorner) * sngXScale)
      .Bottom = CLng((.Bottom - lngdYCorner) * sngYScale)
    End With
  Next
End Sub

'----------------------------------
' Private Functions and Procedures
'----------------------------------

' Purpose    : Returns the value indicating whether the hole octlHole is empty
' Assumption : Hole octlHole's friend controls have been initialized
' Return     : As specified
' Note       : See VB Splitter documentation for the explanation of the hole
'              which is not empty
Private Function IsEmptyHole(ByVal octlHole As clsControl) As Boolean
Attribute IsEmptyHole.VB_Description = "Returns the value indicating whether the hole octlHole is empty"
  Dim blnIsEmptyHole As Boolean                                  'returned value
  Dim octl As clsControl            'for enumerating all items in the collection
  
  If (octlHole.IdCtlFriendTop = gconUninitialized) Or _
     (octlHole.IdCtlFriendRight = gconUninitialized) Or _
     (octlHole.IdCtlFriendBottom = gconUninitialized) Or _
     (octlHole.IdCtlFriendLeft = gconUninitialized) Then
    blnIsEmptyHole = False
  Else
    blnIsEmptyHole = True
    For Each octl In Me
      With octl
        If (.Top >= Item(octlHole.IdCtlFriendTop).Bottom) And _
           (.Right <= Item(octlHole.IdCtlFriendRight).Left) And _
           (.Bottom <= Item(octlHole.IdCtlFriendBottom).Top) And _
           (.Left >= Item(octlHole.IdCtlFriendLeft).Right) Then
          blnIsEmptyHole = False
          Exit For
        End If
      End With
    Next
  End If
  IsEmptyHole = blnIsEmptyHole
End Function

' Purpose    : Removes item IdCtl's heap
' Assumption : Item IdCtl exits in the collection
' Effect     : As specified
' Input      : IdCtl
Private Sub RemoveHeap(ByVal IdCtl As Long)
Attribute RemoveHeap.VB_Description = "Removes item IdCtl's heap"
  Dim lngResizeTop As Long     'the area lost for resizing the top of item IdCtl
  Dim lngResizeRight As Long 'the area lost for resizing the right of item IdCtl
  Dim lngResizeBottom As Long                        'the area lost for resizing
                                                     '  the bottom of item IdCtl
  Dim lngResizeLeft As Long   'the area lost for resizing the left of item IdCtl
  Dim octl1 As clsControl                                 'represents item IdCtl
  Dim octl2 As clsControl           'for enumerating all items in the collection
  Dim octlHeap As clsControl                   'represents the heap area in item
                                                '     octl1 coused by item octl2
   
  Set octlHeap = New clsControl
  Set octl1 = Me(IdCtl)
  For Each octl2 In Me
    If octl2.Id <> octl1.Id Then
      '-- Determines the heap-candidate position and size
      octlHeap.Left = IIf(octl1.Left >= octl2.Left, octl1.Left, octl2.Left)
      octlHeap.Top = IIf(octl1.Top >= octl2.Top, octl1.Top, octl2.Top)
      octlHeap.Right = IIf(octl1.Right <= octl2.Right, octl1.Right, octl2.Right)
      octlHeap.Bottom = IIf(octl1.Bottom <= octl2.Bottom, _
                            octl1.Bottom, octl2.Bottom)
      '-- If octlHeap is truely a heap, remove the heap from item IdCtl
      If ((octlHeap.Left < octlHeap.Right) And _
          (octlHeap.Top < octlHeap.Bottom)) Then
        ' Determines the area lost to resize the item octl1 to certain direction
        lngResizeTop = Abs(octl1.Top - octl2.Bottom)
        lngResizeRight = Abs(octl1.Right - octl2.Left)
        lngResizeBottom = Abs(octl1.Bottom - octl2.Top)
        lngResizeLeft = Abs(octl1.Left - octl2.Right)
        ' Removes the heap using the minimum area lost
        Select Case GetMin(lngResizeTop, lngResizeRight, _
                           lngResizeBottom, lngResizeLeft)
          Case lngResizeTop
            octl1.Top = octlHeap.Yc
            octl2.Bottom = octlHeap.Yc
          Case lngResizeRight
            octl1.Right = octlHeap.Xc
            octl2.Left = octlHeap.Xc
          Case lngResizeBottom
            octl1.Bottom = octlHeap.Yc
            octl2.Top = octlHeap.Yc
          Case lngResizeLeft
            octl1.Left = octlHeap.Xc
            octl2.Right = octlHeap.Xc
        End Select
      End If
    End If
  Next
  Set octl1 = Nothing
  Set octlHeap = Nothing
End Sub

' Purpose    : Sets item IdCtl's friends control
' Assumption : Item IdCtl exits in the collection
' Effect     : As specified
' Input      : IdCtl
Private Sub SetFriend(ByVal IdCtl As Long)
Attribute SetFriend.VB_Description = "Sets item IdCtl's friends control"
  Dim octl1 As clsControl                                 'represents item IdCtl
  Dim octl2 As clsControl           'for enumerating all items in the collection
  
  Set octl1 = Me(IdCtl)
  For Each octl2 In Me
    If octl2.Id <> octl1.Id Then
      With octl1
        ' Sets the top-side friend control
        If (octl2.Bottom <= .Top) And _
           (((.Left < octl2.Left) And (octl2.Left < .Right)) Or _
            ((.Left < octl2.Right) And (octl2.Right < .Right)) Or _
            ((octl2.Left <= .Left) And (.Right <= octl2.Right))) Then
          If .IdCtlFriendTop = gconUninitialized Then
            .IdCtlFriendTop = octl2.Id
          ElseIf (octl2.Bottom > Item(.IdCtlFriendTop).Bottom) Or _
                 ((octl2.Bottom = Item(.IdCtlFriendTop).Bottom) And _
                  (octl2.Right > Item(.IdCtlFriendTop).Right)) Then
            .IdCtlFriendTop = octl2.Id
          End If
        End If
        ' Sets the right-side friend control
        If (.Right <= octl2.Left) And _
           (((.Top < octl2.Top) And (octl2.Top < .Bottom)) Or _
            ((.Top < octl2.Bottom) And (octl2.Bottom < .Bottom)) Or _
            ((octl2.Top <= .Top) And (.Bottom <= octl2.Bottom))) Then
          If .IdCtlFriendRight = gconUninitialized Then
            .IdCtlFriendRight = octl2.Id
          ElseIf (octl2.Left < Item(.IdCtlFriendRight).Left) Or _
                 ((octl2.Left = Item(.IdCtlFriendRight).Left) And _
                  (octl2.Bottom > Item(.IdCtlFriendRight).Bottom)) Then
            .IdCtlFriendRight = octl2.Id
          End If
        End If
        ' Sets the bottom-side friend control
        If (.Bottom <= octl2.Top) And _
           (((.Left < octl2.Left) And (octl2.Left < .Right)) Or _
            ((.Left < octl2.Right) And (octl2.Right < .Right)) Or _
            ((octl2.Left <= .Left) And (.Right <= octl2.Right))) Then
          If .IdCtlFriendBottom = gconUninitialized Then
            .IdCtlFriendBottom = octl2.Id
          ElseIf (octl2.Top < Item(.IdCtlFriendBottom).Top) Or _
                 ((octl2.Top = Item(.IdCtlFriendBottom).Top) And _
                  (octl2.Left < Item(.IdCtlFriendBottom).Left)) Then
            .IdCtlFriendBottom = octl2.Id
          End If
        End If
        ' Sets the left-side friend control
        If (octl2.Right <= .Left) And _
           (((.Top < octl2.Top) And (octl2.Top < .Bottom)) Or _
            ((.Top < octl2.Bottom) And (octl2.Bottom < .Bottom)) Or _
            ((octl2.Top <= .Top) And (.Bottom <= octl2.Bottom))) Then
          If .IdCtlFriendLeft = gconUninitialized Then
            .IdCtlFriendLeft = octl2.Id
          ElseIf (octl2.Right > Item(.IdCtlFriendLeft).Right) Or _
                 ((octl2.Right = Item(.IdCtlFriendLeft).Right) And _
                  (octl2.Top < Item(.IdCtlFriendLeft).Top)) Then
            .IdCtlFriendLeft = octl2.Id
          End If
        End If
      End With
    End If
  Next
  Set octl1 = Nothing
End Sub

' Purpose    : Sets all items' friends control
' Effect     : As specified
Private Sub SetFriends()
Attribute SetFriends.VB_Description = "Sets all items' friends control"
  Dim octl As clsControl            'for enumerating all items in the collection
 
  For Each octl In Me
    SetFriend octl.Id
  Next
End Sub
