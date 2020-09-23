VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A class module to implement collection of class clsId"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************************************************************
'** File Name  : clsIds.cls                                                   **
'** Language   : Visual Basic 6.0                                             **
'** Author     : Theo Zacharias (theo_yz@yahoo.com)                           **
'** Description: A class module to implement custom collection of class clsId **
'** Usage      : Represents list of Ids of class clsControl's and             **
'**              clsSplitter's instances                                      **
'** Members    :                                                              **
'**   * Properties: Count (r/o), Item (def. r/o)                              **
'**   * Methods   : Add, IsExist, IsMarkDeleted, MarkDeleted, Remove,         **
'**                 RemoveDeleted                                             **
'** Last modified on September 11, 2003                                       **
'*******************************************************************************

Option Explicit

'--- Constant Declaration
Private Const conIdDeleted = -2        'represent the Id which is marked deleted

'--- Custom Collection Class Variable
Private mcolIds As Collection

'----------------------------------
' Class Constructor and Destructor
'----------------------------------

Private Sub Class_Initialize()
  Set mcolIds = New Collection
End Sub

Private Sub Class_Terminate()
  Set mcolIds = Nothing
End Sub

'-----------------------------------------------
' Custom Collection Class Properties and Method
'-----------------------------------------------

' Purpose    : Adds a new item to the collection
' Effect     : * If Id = gconUninitialized or Key is already marked deleted
'                then this method doesn't have any effect, i.e. no item will be
'                added to the collection
'              * if Key already exist, the item's value will be replaced with Id
'              * Otherwise, as specified
' Inputs     : * Id
'              * Key (if not passed as parameter, Id will be used as Key)
Public Sub Add(ByVal Id As Long, Optional ByVal Key As Long = gconUninitialized)
Attribute Add.VB_Description = "Adds a new item to the collection"
  Dim oidNew As clsId        'the new item which will be added to the collection
  
  On Error Resume Next
  
  If Id <> gconUninitialized Then
    Set oidNew = New clsId
    oidNew = Id
    If Key = gconUninitialized Then
      mcolIds.Add Item:=oidNew, Key:=CStr(Id)
    ElseIf Not IsMarkDeleted(Key) Then
      mcolIds.Remove CStr(Key)
      mcolIds.Add Item:=oidNew, Key:=CStr(Key)
    End If
    Set oidNew = Nothing
  End If
End Sub

' Purpose    : Returns the number of items in the collection
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of items in the collection"
  Count = mcolIds.Count
End Property

' Purpose    : Returns an item in the collection which has key Key
' Input      : Key
' Return     : * If key Key doesn't exist in the collection, returns item with
'                Id = gconUninitialized
'              * Otherwise, returns as specified
Public Property Get Item(ByVal Key As Long) As clsId
Attribute Item.VB_Description = "Returns an item in the collection which has key Key"
Attribute Item.VB_UserMemId = 0
  Dim oidItem As clsId
  
  On Error GoTo ErrorHandler
  
  Set oidItem = New clsId
  Set oidItem = mcolIds.Item(CStr(Key))
  Set Item = oidItem
  Exit Property

ErrorHandler:
  oidItem = gconUninitialized
  Resume Next
End Property

' Purpose    : Enables For Each ... Next enumeration
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Enables For Each ... Next enumeration"
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mcolIds.[_NewEnum]
End Property

' Purpose    : Removes an item which has key Key from the collection
' Assumption : Key Key exists in the collection
' Effect     : As specified
' Input      : Key
Public Sub Remove(ByVal Key As Long)
Attribute Remove.VB_Description = "Removes an item which has key Key from the collection"
  mcolIds.Remove CStr(Key)
End Sub

'---------------
' Class Methods
'---------------

' Purpose    : Returns a value that determines whether the item with key Key
'              exists in the collection
' Input      : Key
' Note       : Item that is marked deleted is considered not exist
Public Function IsExist(ByVal Key As Long) As Boolean
Attribute IsExist.VB_Description = "Returns a value that determines whether the item with key Key exists in the collection"
  IsExist = (Me(Key).Id <> gconUninitialized) And Not IsMarkDeleted(Key)
End Function

' Purpose    : Returns a value that determines whether the item with key Key
'              is marked deleted
' Input      : Key
Public Function IsMarkDeleted(ByVal Key As Long) As Boolean
Attribute IsMarkDeleted.VB_Description = "Returns a value that determines whether the item with key Key is marked deleted"
  IsMarkDeleted = (Me(Key).Id = conIdDeleted)
End Function

' Purpose    : Marks deleted the item with key Key
' Input      : Key
Public Sub MarkDeleted(ByVal Key As Long)
Attribute MarkDeleted.VB_Description = "Marks deleted the item with key Key"
  Add Id:=conIdDeleted, Key:=CStr(Key)
End Sub

' Purpose    : Removes all items in the collection which is marked deleted
' Input      : lngLastPos (the last position in the collection)
Public Sub RemoveDeleted(ByVal lngLastPos As Long)
Attribute RemoveDeleted.VB_Description = "Removes all items in the collection which is marked deleted"
  Dim i As Long
  
  For i = 0 To lngLastPos
    If IsMarkDeleted(i) Then mcolIds.Remove CStr(i)
  Next
End Sub
