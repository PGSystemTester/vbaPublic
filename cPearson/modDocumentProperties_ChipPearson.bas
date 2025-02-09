Attribute VB_Name = "modDocumentProperties"
Option Explicit
Option Compare Text

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modDocumentProperties
' By Chip Pearson, 5-Jan-2008, chip@cpearson.com
' www.cpearson.com
' www.cpearson.com/Excel/DocProp.apsx
'''''''''''''''''''''''''''''''''''''''''''''''''
' This module contains functions for working with the CustomDocumentProperties and
' BuiltInDocumentProperties property sets.
'''''''''''''''''''''''''''''''''''''''''''''''''
' Functions In This Module:
'
'   SetProperty     Sets the value of a BuiltIn or Custom property. It will
'                   create a new Custom property if necesary.
'
'   GetProperty     Retuns the value of a specified property.
'
'   WritePropertiesToRange      List all the BuiltIn and/or Custom properties
'                               on a worksheet.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''
' This values of this enum indicate which
' property set is to be used, either
' BuiltIn properies, Custom properties,
' or both.
''''''''''''''''''''''''''''''''''''''''''
Public Enum PropertyLocation
    PropertyLocationBuiltIn = 1
    PropertyLocationCustom = 2
    PropertyLocationBOth = 3
End Enum

Function SetProperty(PropertyName As String, PropertySet As PropertyLocation, _
    PropertyValue As Variant, Optional ContentLink As Boolean = False, _
    Optional WhatWorkbook As Workbook) As Boolean
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' SetProperty
	' This procedure sets the value of PropertyName to the value of
	' PropertyValue. If PropertyName does not exist, it will be created
	' if Source is either PropertyLocationCustom or PropertyLocationBoth.
	' The parameters are:
	'
	'   PropertyName        The name of the property to update or create.
	'
	'   PropertySet         One of PropertyLocationBuiltIn,
	'                       PropertyLocationCustom, or PropertyLocationBoth.
	'                       This specifies the property set to search.
	'
	'   PropertyValue       The value to assign to the PropertyName property.
	'                       If ContentLink is FALSE, this is the new value
	'                       of the property. If ContentLink is TRUE, this
	'                       is either the name of the range to link to the
	'                       property or a Name object to link to the property.
	'                       The function will fail if PropertyValue is an array
	'                       or any Object other than Excel.Name.
	'
	'   ContentLink         TRUE or FALSE indicating whether the property
	'                       is to be linked. If omitted, FALSE is assumed.
	'                       If TRUE, PropertyValue must be either:
	'                           - A String containing a defined name of the
	'                             linked cell.
	'                           - A Name object.
	'
	'   WhatWorkbook        A reference to the workbook whose properties
	'                       are to be examined. If omitted or Nothing,
	'                       ThisWorkbook is used.
	'
	' The function returns TRUE if successful or FALSE is an error occurred
	' or the property does not exist (BuiltIn properties only).
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
	Dim WB As Workbook
	Dim Prop As Office.DocumentProperty
	Dim BProps As Office.DocumentProperties
	Dim CProps As Office.DocumentProperties
	Dim PropType As Variant
	
	If PropertyName = vbNullString Then
		''''''''''''''''''''''''''''''''''
		' Don't allow blank names.
		''''''''''''''''''''''''''''''''''
		SetProperty = False
		Exit Function
	End If
	If IsArray(PropertyValue) = True Then
		''''''''''''''''''''''''''''''''''''''''''
		' Don't allow arrays for a property value.
		''''''''''''''''''''''''''''''''''''''''''
		SetProperty = False
		Exit Function
	End If
	'''''''''''''''''''''''''''''''''''''''''
	' Set the workbook whose properties we
	' will work with.
	'''''''''''''''''''''''''''''''''''''''''
	If WhatWorkbook Is Nothing Then
		Set WB = ThisWorkbook
	Else
		Set WB = WhatWorkbook
	End If
		
	Set BProps = WB.BuiltinDocumentProperties
	Set CProps = WB.CustomDocumentProperties

	On Error Resume Next
	''''''''''''''''''''''''''''''''''''''''''''
	' If we are working with BuiltIn properties,
	' we can simply assign the value since you
	' can't link content to a BuiltIn property.
	''''''''''''''''''''''''''''''''''''''''''''
	If PropertySet = PropertyLocationBuiltIn Then
		Err.Clear
		Set Prop = BProps(PropertyName)
		If Err.Number = 0 Then
			'''''''''''''''''''''''''''''''''''
			' Property exists. Set the value
			' and get out.
			'''''''''''''''''''''''''''''''''''
			Prop.Value = PropertyValue
			SetProperty = True
			Exit Function
		End If
		'''''''''''''''''''''''''''''''''''
		' Property doesn't exist. Get out.
		'''''''''''''''''''''''''''''''''''
		SetProperty = False
		Exit Function
	End If
	
	If (PropertySet = PropertyLocationCustom) Or _
		(PropertySet = PropertyLocationBOth) Then
		''''''''''''''''''''''''''''''''''''''''''''''''
		' We need to delete the existing CustomProperty
		' and replace it with a new CustomProperty with
		' the same name. This allows us to change a
		' LinkedContent property to an unlinked content
		' property and vice-versa.
		''''''''''''''''''''''''''''''''''''''''''''''''
		Err.Clear
		Set Prop = CProps(PropertyName)
		
		''''''''''''''''''''''''''''''''''''''
		' If the property exists, delete it.
		''''''''''''''''''''''''''''''''''''''
		If Not Prop Is Nothing Then
			Prop.Delete
		End If
		Err.Clear
		If ContentLink = True Then
			''''''''''''''''''''''''''''''''''''''''''''''''
			' If ContentLink is True, then PropertyValue
			' is the defined name to which the property
			' will be linked. In this case, PropertyValue
			' must be a String and the Name must exist.
			''''''''''''''''''''''''''''''''''''''''''''''''
						
			If IsObject(PropertyValue) = True Then
				''''''''''''''''''''''''''''''''''
				' If PropertyValue is an Object,
				' see if it is an Excel.Name.
				''''''''''''''''''''''''''''''''''
				If TypeOf PropertyValue Is Excel.Name Then
					'''''''''''''''''''''''''''''''''''
					' If it is a Name, set the link
					' and get out.
					'''''''''''''''''''''''''''''''''''
					Err.Clear
					
					CProps.Add Name:=PropertyName, LinkToContent:=True, _
						Type:=msoPropertyTypeString, LinkSource:=PropertyValue.Name
					SetProperty = (Err.Number = 0)
					Exit Function
				Else
					''''''''''''''''''''''''''''''''''''''
					' PropertyValue is an object but is
					' not a Name. Get out.
					''''''''''''''''''''''''''''''''''''''
					SetProperty = False
					Exit Function
				End If
			ElseIf VarType(PropertyValue) = vbString Then
				If NameExists(CStr(PropertyValue), WB) = False Then
					''''''''''''''''''''''''''''''''
					' Name doesn't exist. Get out.
					''''''''''''''''''''''''''''''''
					SetProperty = False
					Exit Function
				End If
				'''''''''''''''''''''''''''''''
				' Name exists. Set up the link.
				'''''''''''''''''''''''''''''''
				CProps.Add Name:=PropertyName, Type:=msoPropertyTypeString, _
					LinkSource:=PropertyValue, _
					LinkToContent:=True
					
			Else
				'''''''''''''''''''''''''''''''''
				' PropertyValue is neither a Name
				' nor a String. Get out.
				'''''''''''''''''''''''''''''''''
				SetProperty = False
				Exit Function
			End If
		Else
			''''''''''''''''''''''''''''''''''''''''''
			' We're not linking content. Just create
			' the property, set the value, and get
			' out.
			''''''''''''''''''''''''''''''''''''''''''
			Err.Clear
			PropType = GetPropertyType(V:=PropertyValue)
			If IsNull(PropType) = True Then
				''''''''''''''''''''''''''''''''''
				' Illegal data type.
				''''''''''''''''''''''''''''''''''
				SetProperty = False
				Exit Function
			End If
			Err.Clear
			CProps.Add Name:=PropertyName, LinkToContent:=False, Type:=PropType, Value:=PropertyValue
			SetProperty = (Err.Number = 0)
		End If
	End If
	''''''''''''''''''''''''''''''''
	' If we get this far, success.
	''''''''''''''''''''''''''''''''
	SetProperty = True
End Function

Function GetProperty(PropertyName As String, PropertySet As PropertyLocation, _
    Optional WhatWorkbook As Workbook) As Variant
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' GetProperty
	' This procedure returns the value of a DocumentProperty named in
	' PropertyName. It will examine BuiltinDocumentProperties,
	' or CustomDocumentProperties, or both. The parameters are:
	'
	'   PropertyName        The name of the property to return.
	'
	'   PropertySet         One of PropertyLocationBuiltIn,
	'                       PropertyLocationCustom, or PropertyLocationBoth.
	'                       This specifies the property set to search.
	'
	'   WhatWorkbook        A reference to the workbook whose properties
	'                       are to be examined. If omitted or Nothing,
	'                       ThisWorkbook is used.
	'
	' The function will return:
	'
	'   The value of property named by PropertyName, or
	'
	'   #VALUE if the PropertySet parameter is not valid (test with IsError), or
	'
	'   Null if the property could not be found (test with IsNull)
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Dim WB As Workbook
	Dim Props1 As Office.DocumentProperties
	Dim Props2 As Office.DocumentProperties
	Dim Prop As Office.DocumentProperty
	
	'''''''''''''''''''''''''''''''''''''''''
	' Set the workbook whose properties we
	' will search.
	'''''''''''''''''''''''''''''''''''''''''
	If WhatWorkbook Is Nothing Then
		Set WB = ThisWorkbook
	Else
		Set WB = WhatWorkbook
	End If
	
	'''''''''''''''''''''''''''''''''''''''''
	' Determine what property set we are
	' going to look at.
	'''''''''''''''''''''''''''''''''''''''''
	Select Case PropertySet
		Case PropertyLocationBuiltIn
			Set Props1 = WB.BuiltinDocumentProperties
		Case PropertyLocationCustom
			Set Props1 = WB.CustomDocumentProperties
		Case PropertyLocationBOth
			Set Props1 = WB.BuiltinDocumentProperties
			Set Props2 = WB.CustomDocumentProperties
		Case Else
			GetProperty = CVErr(xlErrValue)
			Exit Function
	End Select
	
	On Error Resume Next
	'''''''''''''''''''''''''''''''''''''''''
	' Search either BuiltIn or Custom.
	'''''''''''''''''''''''''''''''''''''''''
	Set Prop = Props1(PropertyName)
	If Err.Number <> 0 Then
		''''''''''''''''''''''''''''''''''
		' Not found in one set. See if
		' we need to look in the other.
		''''''''''''''''''''''''''''''''''
		If Not Props2 Is Nothing Then
			''''''''''''''''''''''''''''''''''''
			' We'll get here only if both Custom
			' and BuiltIn properties are to be
			' searched.
			''''''''''''''''''''''''''''''''''''
			Err.Clear
			Set Prop = Props2(PropertyName)
			If Err.Number <> 0 Then
				''''''''''''''''''''''''''''''''''''
				' Property not found. Return NULL.
				''''''''''''''''''''''''''''''''''''
				GetProperty = Null
				Exit Function
			End If
		Else
			''''''''''''''''''''''''''''''''''''
			' Property not found. Return NULL.
			''''''''''''''''''''''''''''''''''''
			GetProperty = Null
			Exit Function
		End If
	End If
	
	''''''''''''''''''''''''''''''''''''
	' Property found. Return the value.
	''''''''''''''''''''''''''''''''''''
	GetProperty = Prop.Value
End Function

Function ReadPropertyFromClosedFile(FileName As String, PropertyName As String, _
    PropertySet As PropertyLocation) As Variant
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' ReadPropertyFromClosedFile
	' This uses the DSOFile DLL to read properties from a closed workbook. This DLL is
	' available at http://support.microsoft.com/kb/224351/en-us. This code requires a
	' reference to "DSO OLE Document Properties Reader 2.1". The function returns
	' the value of the property if it exists, or NULL if an error occurs. Be sure to
	' check the return value with IsNull.
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	Dim DSO As DSOFile.OleDocumentProperties
	Dim Prop As Office.DocumentProperty
	Dim V As Variant
	
	If FileName = vbNullString Then
		ReadPropertyFromClosedFile = Null
		Exit Function
	End If
	
	If PropertyName = vbNullString Then
		ReadPropertyFromClosedFile = Null
		Exit Function
	End If
	
	If Dir(FileName, vbNormal + vbSystem + vbHidden) = vbNullString Then
		''''''''''''''''''''''''''''''''
		' File doesn't exist. Get out.
		''''''''''''''''''''''''''''''''
		ReadPropertyFromClosedFile = Null
		Exit Function
	End If
	Select Case PropertySet
		Case PropertyLocationBOth, PropertyLocationBuiltIn, PropertyLocationCustom
			'''''''''''''''''''''''''''''
			' Valid value for PropertySet
			'''''''''''''''''''''''''''''
		Case Else
			'''''''''''''''''''''''''''''
			' Invalid value. Get Out.
			'''''''''''''''''''''''''''''
			ReadPropertyFromClosedFile = Null
			Exit Function
	End Select
	
	On Error Resume Next
	
	Set DSO = New DSOFile.OleDocumentProperties
	'''''''''''''''''''''''''''''''''''''''''''''
	' Open the file.
	'''''''''''''''''''''''''''''''''''''''''''''
	DSO.Open sfilename:=FileName, ReadOnly:=True
	'''''''''''''''''''''''''''''''''''''''''''''
	' If we're working with BuiltIn or Both
	' property sets, try to get the property.
	'''''''''''''''''''''''''''''''''''''''''''''
	If (PropertySet = PropertyLocationBOth) Or (PropertySet = PropertyLocationBuiltIn) Then
		Err.Clear
		''''''''''''''''''''''''''''''''''''''
		' Look first in the BuiltIn (Summary)
		' properties. The SummaryProperties
		' object is not a Collection whose
		' members you can select. Instead,
		' there is a separate property for
		' each of the Summary Properties. Thus,
		' use CallByName to get the values.
		''''''''''''''''''''''''''''''''''''''
		V = CallByName(DSO.SummaryProperties, PropertyName, VbGet)
		If Err.Number <> 0 Then
			If PropertySet = PropertyLocationBOth Then
				'''''''''''''''''''''''''''''''''''''
				' We're looking in both property sets.
				' Not found in BuiltIn. Try Custom.
				'''''''''''''''''''''''''''''''''''''
				Err.Clear
				V = DSO.CustomProperties(PropertyName)
				If Err.Number <> 0 Then
					'''''''''''''''''''''''''''''''''
					' Not found. Return NULL.
					'''''''''''''''''''''''''''''''''
					DSO.Close savebeforeclose:=False
					ReadPropertyFromClosedFile = Null
					Exit Function
				Else
					'''''''''''''''''''''''''''''''''
					' Found. Return value.
					'''''''''''''''''''''''''''''''''
					DSO.Close savebeforeclose:=False
					ReadPropertyFromClosedFile = V
					Exit Function
				End If
			Else
				''''''''''''''''''''''''''''''''''''''
				' Not found in BuiltIn and we're not
				' looking in both sets so return NULL
				' and get out.
				''''''''''''''''''''''''''''''''''''''
				DSO.Close savebeforeclose:=False
				ReadPropertyFromClosedFile = Null
				Exit Function
			End If
		Else
			'''''''''''''''''''''''''''''''''
			' Found. Return value.
			'''''''''''''''''''''''''''''''''
			DSO.Close savebeforeclose:=False
			ReadPropertyFromClosedFile = V
			Exit Function
		End If
	End If
	
	If (PropertySet = PropertyLocationBOth) Or (PropertySet = PropertyLocationCustom) Then
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' We're looking at Custom properties or both. We've already
		' looked in Custom, so don't do it again.
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Err.Clear
		V = DSO.SummaryProperties(PropertyName)
		If Err.Number <> 0 Then
			Err.Clear
			V = DSO.CustomProperties(PropertyName)
			If Err.Number <> 0 Then
				'''''''''''''''''''''''''''''''''
				' Not found. Return NULL.
				'''''''''''''''''''''''''''''''''
				DSO.Close savebeforeclose:=False
				ReadPropertyFromClosedFile = Null
				Exit Function
			Else
				'''''''''''''''''''''''''''''''''
				' Found. Return value.
				'''''''''''''''''''''''''''''''''
				DSO.Close savebeforeclose:=False
				ReadPropertyFromClosedFile = V
				Exit Function
			End If
		Else
			'''''''''''''''''''''''''''''''''
			' Found. Return value.
			'''''''''''''''''''''''''''''''''
			DSO.Close savebeforeclose:=False
			ReadPropertyFromClosedFile = V
			Exit Function
		End If
	End If
	DSO.Close savebeforeclose:=False
End Function

Function WritePropertyToClosedFile(FileName As String, PropertyName As String, _
    PropertyValue As String, PropertySet As PropertyLocation) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' WritePropertyToClosedFile
' This procedure will modify or create either a BuiltIn or Custom property on
' a close file. The parameters are:
'
'       FileName            This is the fully qualified name of the file to
'                           modify.
'
'       PropertyName        This is the name of the property to modify.
'
'       PropertyValue       This is the new value for PropertyName.
'
'       PropertySet         One of PropertyLocationBuiltIn, PropertyLocationCustom,
'                           or PropertyLocationBoth to indicate which property set
'                           to modify.
'
' The function returns TRUE if successful or FALSE if an error occurred or an
' invalid parameter was passed to the function.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim DSO As DSOFile.OleDocumentProperties
Dim V As Variant

If FileName = vbNullString Then
    WritePropertyToClosedFile = False
    Exit Function
End If

If PropertyName = vbNullString Then
    WritePropertyToClosedFile = False
    Exit Function
End If

If Dir(FileName, vbNormal + vbSystem + vbHidden) = vbNullString Then
    WritePropertyToClosedFile = False
    Exit Function
End If
    
On Error Resume Next
Set DSO = New DSOFile.OleDocumentProperties
Err.Clear

DSO.Open sfilename:=FileName, ReadOnly:=False
If Err.Number <> 0 Then
    DSO.Close False
    WritePropertyToClosedFile = False
    Exit Function
End If

If (PropertySet = PropertyLocationBuiltIn) Or _
    (PropertySet = PropertyLocationBOth) Then
    '''''''''''''''''''''''''''''''''''''''''''
    ' We're looking at either BuiltIn or Both.
    '''''''''''''''''''''''''''''''''''''''''''
    Err.Clear
    '''''''''''''''''''''''''''''''''''''''''''
    ' SummaryProperties has a separate procedure
    ' for each property. Thus, we need CallByName
    ' to be the property name.
    '''''''''''''''''''''''''''''''''''''''''''
    V = CallByName(DSO.SummaryProperties, PropertyName, VbGet)
    If Err.Number <> 0 Then
        '''''''''''''''''''''''''''''
        ' Not found. See if we check
        ' Custom properties.
        '''''''''''''''''''''''''''''
        If PropertySet = PropertyLocationBOth Then
            Err.Clear
            V = DSO.CustomProperties(PropertyName)
            If Err.Number <> 0 Then
                ''''''''''''''''''''''''''''''''
                ' Not found in Custom. Get out.
                ''''''''''''''''''''''''''''''''
                DSO.Close False
                WritePropertyToClosedFile = False
                Exit Function
            End If
            Err.Clear
            DSO.CustomProperties(PropertyName).Value = PropertyValue
            DSO.Close savebeforeclose:=True
            WritePropertyToClosedFile = True
            Exit Function
        End If
    Else
        '''''''''''''''''''''''''''''''''''''
        ' Found it in BuiltIn/Summary
        '''''''''''''''''''''''''''''''''''''
        Err.Clear
        CallByName DSO.SummaryProperties, PropertyName, VbLet, PropertyValue
        WritePropertyToClosedFile = (Err.Number = 0)
        DSO.Close savebeforeclose:=True
        Exit Function
    End If
End If

If (PropertySet = PropertyLocationCustom) Or _
    (PropertySet = PropertyLocationBOth) Then
    '''''''''''''''''''''''''''''''''''''''''''
    ' We're looking at either Custom or Both.
    '''''''''''''''''''''''''''''''''''''''''''
    Err.Clear
    V = DSO.CustomProperties(PropertyName).Value
    If Err.Number <> 0 Then
        '''''''''''''''''''''''''''''
        ' Not found in custom. Attempt
        ' to add it.
        '''''''''''''''''''''''''''''
        Err.Clear
        DSO.CustomProperties.Add spropname:=PropertyName, Value:=PropertyValue
        '''''''''''''''''''''''''''''
        ' check the error and get out.
        '''''''''''''''''''''''''''''
        WritePropertyToClosedFile = (Err.Number = 0)
        DSO.Close savebeforeclose:=True
        Exit Function
    Else
        ''''''''''''''''''''''''''''''
        ' Found it. Update it.
        ''''''''''''''''''''''''''''''
        Err.Clear
        DSO.CustomProperties(PropertyName).Value = PropertyValue
        WritePropertyToClosedFile = (Err.Number = 0)
        DSO.Close savebeforeclose:=True
        Exit Function
    End If
End If

End Function




Function PropertyExists(PropertyName As String, PropertySet As PropertyLocation, _
    Optional WhatWorkbook As Workbook) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PropertyExists
' This procedure returns TRUE or FALSE indicating whether PropertyName
' exists in PropertySet in the workbook WhatWorkbook.
' The parameters are:
'
'   PropertyName        The name of the property to be found.
'
'   PropertySet         One of PropertyLocationBuiltIn,
'                       PropertyLocationCustom, or PropertyLocationBoth.
'                       This specifies the property set to search.
'
'   WhatWorkbook        A reference to the workbook whose properties
'                       are to be examined. If omitted or Nothing,
'                       ThisWorkbook is used.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
PropertyExists = Not IsNull(GetProperty(PropertyName, PropertySet, WhatWorkbook))

End Function

Private Function NameExists(NameName As String, Optional WhatWorkbook As Workbook) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' NameExists
' This returns TRUE or FALSE indicating whether the name in NameName exists
' in WhatWorkbook. If WhatWorkbook is omitted or Nothing, the ThisWorkbook
' is used.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim WB As Workbook
If WhatWorkbook Is Nothing Then
    Set WB = ThisWorkbook
Else
    Set WB = WhatWorkbook
End If

On Error Resume Next
NameExists = CBool(Len(WB.Names(NameName).Name))
End Function

Private Function GetPropertyType(V As Variant) As Variant
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' GetPropertyType
	' This tests the data type of V and returns the appropriate Property type.
	' Returns a member of the VbVarType group or NULL if an illegal type (e.g,
	' an Object) is found. Be sure to test the return value with IsNull.
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Select Case VarType(V)
		Case vbArray, vbDataObject, vbEmpty, vbError, vbNull, vbObject, _
			vbUserDefinedType, vbCurrency, vbDecimal
			''''''''''''''''''''''''''''''''''''
			' Illegal types. Return NULL.
			''''''''''''''''''''''''''''''''''''
			GetPropertyType = Null
			Exit Function
		''''''''''''''''''''''''''''''''''
		' All numeric types are rolled up
		' into Floats. Strings and Booleans
		' get their own types.
		''''''''''''''''''''''''''''''''''
		Case vbString
			GetPropertyType = msoPropertyTypeString
		Case vbBoolean
			GetPropertyType = msoPropertyTypeBoolean
		Case Else
			GetPropertyType = msoPropertyTypeFloat
	End Select
End Function

Function WritePropertiesToRange(PropertySet As PropertyLocation, FirstCell As Range, _
    Optional WhatWorkbook As Workbook) As Long
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' WritePropertiesToRange
	' This proceudre writes all the property names and their values to a worksheet,
	' starting with FirstCell. The BuiltIn properties, Custom properties, or both
	' may be specified by the value of PropertySet. The WhatWorkbook parameter
	' specifies the workbook whose properties are to be listed. If WhatWorkbook is
	' omitted or is Nothing, ThisWorkbookis used. The listing is two columns wide.
	' The first (left) column contains the names of the properties, and the second
	' (right) column lists the property values.
	' The function returns the number of properties listed, or -1 if an error
	' occurred.
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Dim WB As Workbook
	Dim R As Range
	Dim S As String
	Dim V As Variant
	Dim N As Long
	Dim BProps As Office.DocumentProperties
	Dim CProps As Office.DocumentProperties
	Dim Prop As Office.DocumentProperty
	If WhatWorkbook Is Nothing Then
		Set WB = ThisWorkbook
	Else
		Set WB = WhatWorkbook
	End If
	Select Case PropertySet
		Case PropertyLocationCustom, PropertyLocationBuiltIn, PropertyLocationBOth
			''''''''''''''''''''''''''''''''''''''''''''''
			' valid value for PropertySet. Do nothing.
			''''''''''''''''''''''''''''''''''''''''''''''
		Case Else
			''''''''''''''''''''''''''''''
			' Invalid value.
			''''''''''''''''''''''''''''''
			WritePropertiesToRange = -1
			Exit Function
	End Select
	If FirstCell Is Nothing Then
		WritePropertiesToRange = -1
		Exit Function
	End If
	Set R = FirstCell
	Set BProps = WB.BuiltinDocumentProperties
	Set CProps = WB.CustomDocumentProperties
	N = 0
	''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Loop through the BuiltIn properties if necessary.
	''''''''''''''''''''''''''''''''''''''''''''''''''''
	On Error Resume Next
	If (PropertySet = PropertyLocationBuiltIn) Or _
		(PropertySet = PropertyLocationBOth) Then
		For Each Prop In BProps
			V = Empty
			Err.Clear
			S = Prop.Name
			V = Prop.Value
			R(1, 1).Value = S
			R(1, 2).Value = V
			N = N + 1
			Set R = R(2, 1)
		Next Prop
	End If
	''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Loop through the Custom properties if necessary.
	''''''''''''''''''''''''''''''''''''''''''''''''''''
	If (PropertySet = PropertyLocationCustom) Or (PropertySet = PropertyLocationBOth) Then
		For Each Prop In CProps
			V = Empty
			Err.Clear
			S = Prop.Name
			V = Prop.Value
			R(1, 1).Value = S
			R(1, 2).Value = V
			N = N + 1
			Set R = R(2, 1)
		Next Prop
	End If
	WritePropertiesToRange = N
End Function
