Attribute VB_Name = "mdlGeneral"
Attribute VB_Description = "A module to handle general operations"
'*******************************************************************************
'** File Name  : mdlGeneral.bas                                               **
'** Language   : Visual Basic 6.0                                             **
'** Author     : Theo Zacharias (theo_yz@yahoo.com)                           **
'** Description: A module to handle general operations                        **
'** Last modified on September 12, 2003                                       **
'*******************************************************************************

Option Explicit

'--- Resource File Constants

' Splitter Cursor
Public Const gconCurHSplitter = 101                  'horizontal splitter cursor
Public Const gconCurVSplitter = 102                    'vertical splitter cursor

' Error Message Index
Public Enum genmErrNumber
  errBuildSplitters = 2000
  errSelfContained = 2001
  errMoveSplitter = 2002
End Enum

'--- Other Constants
Public Const gconUninitialized = -1      'represent the Id which is not exist or
                                         '           hasn't been initialized yet
Public Const gconLngInfinite = 2147483647

'--- Variable Declaration
Public gstrControlName As String       'the name of VB Splitter control instance

' Purpose    : Gets minimum value of numbers in array lngValue()
' Assumptions: * Option base is set to 0
'              * Array lngValue() contains only numbers
' Input      : vntValue()
' Return     : * If no parameters passed to vntValue(), returns Empty
'              * Otherwise, returns as specified
Public Function GetMin(ParamArray vntValue() As Variant) As Variant
Attribute GetMin.VB_Description = "Gets minimum value of numbers in array lngValue()"
  Dim i As Long
  Dim vntGetMin As Variant
  
  If Not IsMissing(vntValue) Then
    vntGetMin = vntValue(0)
    For i = 1 To UBound(vntValue)
      If vntValue(i) < vntGetMin Then vntGetMin = vntValue(i)
    Next
    GetMin = vntGetMin
  End If
End Function

' Purpose    : Returns vntrTrue if blnCondition = true, or vntFalse otherwise
' Inputs     : blnCondition, vntTrue, vntFalse
' Return     : As specified
Public Function IIf(ByVal blnCondition As Boolean, _
                    ByVal vntTrue As Variant, _
                    ByVal vntFalse As Variant) As Variant
Attribute IIf.VB_Description = "Returns vntrTrue if blnCondition = true, or vntFalse otherwise"
  If blnCondition Then IIf = vntTrue Else IIf = vntFalse
End Function

' Purpose    : Raises custom error udeErrNumber
' Assumptions: * Error message udeErrNumber exists in the resource file
'              * Global variable gstrControlName has been initialized
' Inputs     : * udeErrNumber
'              * strSource (the location in form ClassName.RoutinesName where
'                the error occur
Public Sub RaiseError(ByVal udeErrNumber As genmErrNumber, _
                      Optional ByVal strSource As String = "")
Attribute RaiseError.VB_Description = "Raises custom error udeErrNumber"
  If strSource <> "." Then strSource = "." & strSource
  Err.Raise Number:=(vbObjectError + udeErrNumber), _
            Source:=gstrControlName & strSource, _
            Description:=LoadResString(udeErrNumber)
End Sub
