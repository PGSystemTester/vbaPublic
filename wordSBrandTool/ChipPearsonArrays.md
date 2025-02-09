This is a repost from the late [Chip Pearson's site](http://www.cpearson.com/Excel/PassingAndReturningArrays.htm) on arrays. The formatting from his site is becoming difficult to read, thus this aims to retain his work in markdown.

# Passing And Returning Arrays With Functions
In VBA, you can pass arrays to procedures (Subs, Functions, and Properties), and Functions and Properties (Property Get only) can return arrays as their result. (Note that returning an array as the result of a function or property was added in Office 2000 -- functions in Office 97 cannot return arrays.)  However, you must be aware of the limitations in passing arrays. This page assumes you are familiar with the fundamentals of VBA arrays and the difference between a static and a dynamic array.

As a general coding practice, I always use dynamic arrays and use ReDim to size the array to the necessary dimensions. This makes code more flexible and re-usable. It is quite rare that I will be dealing with a fixed number of entities or objects whose number is known at design time. Using dynamic arrays allows the software to size itself for the task at hand. In fact, I can think of no situation in which a static array would be superior to a dynamic array.

The procedures and code on this page use the array support functions described on the [Functions For VBA Arrays page](http://www.cpearson.com/Excel/VBAArrays.htm). If you are going to use the example code on this page in your VBA code, you should copy the functions on that page into a module in your VBA project, or get the code module here and import this module into your VBA project. We will use the same terminology as described on the [Functions For VBA Arrays page](http://www.cpearson.com/Excel/VBAArrays.htm).

## Passing Arrays To Procedures
A procedure (a Sub, Function or Property) can accept an array as an input parameter. The first thing to understand is that arrays are always passed by reference (`ByRef`). You will receive a compiler error if you attempt to pass an array ByVal. (See the Online VBA Help for the topic Sub Statement for information about ByRef and ByVal.) This means that any modification that the called procedure does to the array parameter is done on the actual array declared in the calling procedure. This is illustrated in the following code:

````vba
Sub AAATest()
	Dim StaticArray(1 To 3) As Long
	Dim N As Long
	StaticArray(1) = 1
	StaticArray(2) = 2
	StaticArray(3) = 3
	PopulatePassedArray Arr:=StaticArray
	For N = LBound(StaticArray) To UBound(StaticArray)
		Debug.Print StaticArray(N)
	Next N
End Sub

Sub PopulatePassedArray(ByRef Arr() As Long)
	''''''''''''''''''''''''''''''''''''
	' PopulatePassedArray
	' This puts some values in Arr.
	''''''''''''''''''''''''''''''''''''
	Dim N As Long
	For N = LBound(Arr) To UBound(Arr)
		Arr(N) = N * 10
	Next N
End Sub
````

In this code, the array `StaticArray` is passed to the `PopulatePassedArray` procedure. The ByRef keyword is not required in the parameter list of the procedure `PopulatePassedArray` (ByRef is the default in VB/VBA), but I tend include `ByRef` in parameter declarations if I am going to modify that variable. It serves as a reminder that a variable in the calling procedure is going to be modified by the called procedure. I don't include the `ByRef` keyword for variables whose content I am not going to modify. You may safely omit ByRef if you prefer. Since the array `StaticArray` is passed by reference, the variable StaticArray in the calling procedure `AAATest` is modified by the code in the called procedure `PopulatePassedArray`.

In a real-word, commercial-quality application, you would first test to ensure that the array Arr has actually been allocated and the the array is single-dimensional. You can use the `IsArrayAllocated` and `NumberOfArrayDimensions` functions, described on the Functions For VBA Arrays page, to test these conditions. For example, you would write the called procedure as:

````vba
Sub PopulatePassedArray(ByRef Arr() As Long)
	''''''''''''''''''''''''''''''''''''
	' PopulatePassedArray
	' This puts some values in Arr.
	''''''''''''''''''''''''''''''''''''
	Dim N As Long
	If IsArrayAllocated(Arr:=Arr) = True Then
		If NumberOfArrayDimensions(Arr:=Arr) = 1 Then
			For N = LBound(Arr) To UBound(Arr)
				Arr(N) = N * 10
			Next N
		Else
			Debug.Print "Array is has multiple dimensions."
		   '''''''''''''''''''''''''''''''''''''''
		   ' Take whatever action is necessary
		   ' for a multi-dimensional array,
		   ' such as resizing the array.
		   '''''''''''''''''''''''''''''''''''''''
		End If
	Else
		Debug.Print "Array Not Allocated."
		'''''''''''''''''''''''''''''''''''''''
		' Take whatever action necessary with
		' an unallocated array, such as ReDim
		' the array.
		'''''''''''''''''''''''''''''''''''''''
   End If
End Sub
````
You may have noticed that the static array StaticArray in `AAATest` has the same data type (`Long`) as the array Arr declared in the parameter list to `PopulatePassedArray`. This is no coincidence. The rule here is that the data type of the array declared in the calling procedure must match the data type declared in the called procedure's parameter list. While you can declare a simple parameter As Variant to accept a parameter to be of any data type, this does not work for arrays. It is a very common misconception that declaring the function parameter `As Variant()` will allow you to accept an array of any type. This is flat wrong. You cannot declare the array in the parameter list of the called procedure As Variant() to accept any data type array. The data types must explicitly match; otherwise, you'll get a "_Type Mismatch: Array or user-defined type expected._" error when you compile and run the code. If you declare the function parameter `As Variant()` then you must pass an array of Variants.

Moreover, if a function parameter is declared as an array, you cannot pass a single Variant as that function parameter, even if the Variant contains an array of the proper data type. For example, the following code will not compile.


````vba
Sub AAATest()
	Dim V As Variant
	Dim L(1 To 3) As Long
	L(1) = 100
	L(2) = 200
	L(3) = 300
	V = L
	BBB L '<<< Works because L is a Long Array.
	BBB V '<<< Compiler error here. V itself is not an array, as expected by BBB.
End Sub

Sub BBB(Arr() As Long)
	Debug.Print Arr(LBound(Arr))
End Sub
````

Since an array passed from the calling procedure to the called procedure is passed by reference, the called procedure may use the ReDim statement to change the size of the passed array and/or number of dimensions (the passed array must declared as a dynamic array in the calling procedure, but it need not be allocated). This is perfectly legal and indeed quite useful. For example, in the procedures below, the array DynArray is declared as a dynamic array in `AAATest`, and it is resized as many times as needed to store the results in the PopulateArrayWithCellValuesGreaterThan10 procedure.

````vba
Sub AAATest()
  Dim DynArray() As Double ' Note that this array is not sized
           ' in the Dim statement. We'll use ReDim
           ' in the called procedure to change the size.
  Dim N As Long
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' Call PopulateArrayWithCellValuesGreaterThan10 to resize
  ' the array and populate its elements with values.
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  PopulateArrayWithCellValuesGreaterThan10 Arr:=DynArray, TestRng:=Range("A1:A10")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' Ensure that PopulateArrayWithCellValuesGreaterThan10
  ' allocated the array.
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''  
  If IsArrayAllocated(Arr:=DynArray) = True Then
    For N = LBound(DynArray) To UBound(DynArray)
      Debug.Print DynArray(N)
    Next N
  End If
End Sub

Sub PopulateArrayWithCellValuesGreaterThan10(ByRef Arr() As Double, TestRng As Range)
	''''''''''''''''''''''''''''''''''''''''''''''''''
	' PopulateArrayWithCellValuesGreaterThan10
	'
	' This resizes Arr and places in it the values
	' in the range TestRng that are greater than 10.
	''''''''''''''''''''''''''''''''''''''''''''''''''
	Dim Rng As Range
	Dim Ndx As Long
	''''''''''''''''''''''''''''''''
	' Ensure TestRng is not Nothing.
	''''''''''''''''''''''''''''''''
	If TestRng Is Nothing Then 
		MsgBox "TestRng Is Nothing"
		Exit Sub
	End If
	''''''''''''''''''''''''''''''''
	' Loop through the range.
	''''''''''''''''''''''''''''''''
	For Each Rng In TestRng.Cells
		If IsNumeric(Rng.Value) = True Then 
			If Rng.Value > 10 Then
				Ndx = Ndx + 1
				ReDim Preserve Arr(1 To Ndx)
				Arr(Ndx) = CDbl(Rng.Value)
			End if
		End If
	Next Rng
End Sub
````

In the `PopulateArrayWithCellValuesGreaterThan10` procedure, the array `Arr` (which is the same array as `DynArray` in `AAATest`) is resized using the `ReDim Preserve` statement each time a cell value greater than 10 is encountered. While this is perfectly legal code, and it illustrates resizing an array parameter, the code is neither safe nor efficient. The code doesn't test whether the array is dynamic and therefore can be resized. If we were passed a static array, we would get a run-time error 10 ("The array is fixed or temporarily locked.") when calling `ReDim Preserve`. Moreover, the code calls the `ReDim Preserve` statement any number of times, as many times as a cell value exceeds 10. `ReDim Preserve` is an expensive operation (especially with large arrays of Strings or Variants) and should be used sparingly. A much better version of the procedure is shown below.

````vba
Sub AAATest()
	Dim DynArray() As Double ' Note that this array is not sized
							 ' in the Dim statement. We'll use ReDim
							 ' to change the size later.
	Dim N As Long
	PopulateArrayWithCellValuesGreaterThan10 Arr:=DynArray, TestRng:=Range("A1:A10")
	If IsArrayAllocated(Arr:=DynArray) = True Then
		For N = LBound(DynArray) To UBound(DynArray)
			Debug.Print DynArray(N)
		Next N
	End If
End Sub

Sub PopulateArrayWithCellValuesGreaterThan10(ByRef Arr() As Double, TestRng As Range)
	''''''''''''''''''''''''''''''''''''''''''''''''''
	' PopulateArrayWithCellValuesGreaterThan10
	'
	' This resizes Arr once to the maximum number
	' of elements we may use, populates the array
	' elements, and then uses ReDim Preserve to
	' resize the array to the number of elements
	' actually used. This avoids calling Redim
	' Preserve for each cell in the range.
	'
	' It calls IsArrayDynamic to ensure that Arr 
	' is a dynamic array that can be resized. 
	''''''''''''''''''''''''''''''''''''''''''''''''''
	Dim Rng As Range
	Dim Ndx As Long
	''''''''''''''''''''''''''''''''
	' Ensure TestRng is not Nothing.
	''''''''''''''''''''''''''''''''
	If TestRng Is Nothing Then 
		MsgBox "TestRng Is Nothing"
		Exit Sub
	End If 
	''''''''''''''''''''''''''''''''
	' Call IsArrayDynamic to ensure
	' that we have a dynamic array.
	''''''''''''''''''''''''''''''''   
	If IsArrayDynamic(Arr:=Arr) = True Then
		'''''''''''''''''''''''''''''''''''
		' ReDim Arr to the number of cells
		' in TestRng. This is the maximum
		' possible entries we might use -- 
		' the number of cells in the TestRng 
		' range. We don't use Preserve 
		' with ReDim here because we don't want
		' to preserve any existing values.
		' Any values currently in the array 
		' will be lost. 
		'''''''''''''''''''''''''''''''''''
		ReDim Arr(1 To TestRng.Cells.Count)
		Ndx = 0
		For Each Rng In TestRng.Cells
		If IsNumeric(Rng.Value) = True Then            
				If Rng.Value > 10 Then
					Ndx = Ndx + 1
					Arr(Ndx) = CDbl(Rng.Value)
				End If
			End If
		Next Rng
		'''''''''''''''''''''''''''''''''''
		' ReDim Preserve to reduce the size 
		' the array to only as many elements 
		' as we used.
		'''''''''''''''''''''''''''''''''''
		ReDim Preserve Arr(1 To Ndx)
	Else
		''''''''''''''''''''''''''''''''''''''''''
		' Code for the case if Arr is not dynamic.
		''''''''''''''''''''''''''''''''''''''''''
		Exit Sub
	End If
End Sub
````

In this code, we first call the function `IsArrayDynamic` (this function is illustrated on the Functions For VBA Arrays page) to ensure that Arr is an dynamic array that we can resize. If this function returns True, the array is dynamic and we can resize it. If the function returns False, the array is static, and we would probably raise an error or just exit the sub. The procedure then calls ReDim once to size the array to the number of cells in the range (the maximum possible size we'll need for the array), and then calls ReDim Preserve at the end to reduce the size to the actual number of elements actually used. (ReDim Preserve  is usually used to increase the size of an array, but it is equally valid to use it to reduce the size of an array). Note that the first call to ReDim doesn't use the Preserve  keyword. This is because we don't want to preserve any values that might be in the array. Calling ReDim without Preserve resizes the array but destroys its existing contents.

Note that when you declare an array in called procedure's parameter list, you do not (and cannot) include its size, even if it is a static array. For example, the following code is illegal and will not compile:

`Public Sub CalledProcedure (Arr(1 to 3) As Long)`

Instead you use code like

`Public Sub CalledProcedure (Arr() As Long)`

The lack of lower and upper bounds within the parentheses in the parameter declaration of `Arr`, however, does not mean that the array somehow has been made a dynamic array. If it was declared static in calling function, it remains static when accessed by the called procedure. The "()" characters after the parameter name in the called procedure's parameter list simply indicate that an array, either static or dynamic, is being passed. It is up to the called procedure to determine whether the passed array is static or dynamic, if necessary. You can use the `IsArrayDynamic` function to test this condition.

There is no way to change a static array into a dynamic array. If it is sized in the Dim statement, it can never be resized. Its size is fixed, and any attempt to resize the array will cause a compiler error ("Array already dimensioned"). You can, of course, create a new dynamic array and load it with the contents of a static array:

````vb
    Dim StaticArray(1 To 3) As Long
	Dim DynArray() As Double
	Dim Ndx As Long
	' Load StaticArray with some data  
	ReDim DynArray(LBound(StaticArray) To UBound(StaticArray))
	For Ndx = LBound(StaticArray) To UBound(StaticArray)
		DynArray(Ndx) = StaticArray(Ndx)
	Next Ndx
````

In this case, the data types need not match. They just must be compatible (e,g, both should be numeric types).

You can pass static arrays to procedures, just as you can dynamic arrays. As with dynamic arrays, static arrays are passed by reference. The only difference is that you cannot resize a static array. For example, the following code passes a static array to a function.

````vba
Sub AAATest()
	Dim StaticArray(1 To 3) As Long
	Dim Result As Long
	StaticArray(1) = 10
	StaticArray(2) = 20
	StaticArray(3) = 30
	Result = SumArray(Arr:=StaticArray)
	Debug.Print Result
End Sub

Function SumArray(Arr() As Long) As Long
	'''''''''''''''''''''''''''''''''''''''''''
	' SumArray
	' This sums the elements of Arr and returns
	' the total.
	''''''''''''''''''''''''''''''''''''''''''' 
	Dim N As Long
	Dim Total As Long
	For N = LBound(Arr) To UBound(Arr)
		Total = Total + Arr(N)
	Next N
	SumArray = Total
End Function
````

The SumArray function just loops through the array, summing the values, and returns the result.

Because the data type of the array in the calling procedure must match the data type in the array parameter declaration in the called procedure, you may wonder how to call a procedure that can handle various other data types, which may not be known until run-time. For example, how would SumArray be written to handle arrays of Integers or Doubles as well as Longs?

To pass an array of any type to a procedure, don't declare the parameter as an array. Instead, declare it as a Variant (not an array of Variants). A single Variant variable may contain an array. This is illustrated in the code below. 

````vba
Sub AAATest()
	Dim StaticArray(1 To 3) As Double
	Dim Result As Double
	StaticArray(1) = 10
	StaticArray(2) = 20
	StaticArray(3) = 30
	Result = SumArray(Arr:=StaticArray)
	Debug.Print Result
End Sub

Function SumArray(Arr As Variant) As Double
	'''''''''''''''''''''''''''''''''''''''''''
	' SumArray
	' This sums the elements of Arr and returns
	' the total.
	'''''''''''''''''''''''''''''''''''''''''''
	Dim N As Long
	Dim Total As Double
	'''''''''''''''''''''''''
	' Ensure Arr is an array.
	'''''''''''''''''''''''''
	If IsArray(Arr) = True Then
		''''''''''''''''''''''''''''''''
		' Ensure the array is allocated.
		''''''''''''''''''''''''''''''''
		If IsArrayAllocated(Arr:=Arr) = True Then
			''''''''''''''''''''''''''''''''
			' Ensure Arr is one-dimensional.
			''''''''''''''''''''''''''''''''
			If NumberOfArrayDimensions(Arr:=Arr) = 1 Then
				'''''''''''''''''''''''''''''''''''''
				' Ensure Arr is a numeric type array.
				'''''''''''''''''''''''''''''''''''''
				If IsNumericDataType(Arr) = True Then
					For N = LBound(Arr) To UBound(Arr)
						'''''''''''''''''''''''''''
						' Ensure Arr(N) is numeric.
						'''''''''''''''''''''''''''
						If IsNumeric(Arr(N)) = True Then
							Total = Total + Arr(N)
						End If
					Next N
				Else
					Debug.Print "Array is not numeric."
					''''''''''''''''''''''''''''''''''
					' Code in case Arr is not numeric.
					''''''''''''''''''''''''''''''''''
					Exit Function
				End If
			Else
				Debug.Print "Array is not one-dimensional."
				''''''''''''''''''''''''''''''''''''''''
				' Code in case Arr is multi-dimensional.
				''''''''''''''''''''''''''''''''''''''''
				Exit Function
			End If
		 Else
			Debug.Print "Array is not allocated."
			 ''''''''''''''''''''''''''''''''''''
			 ' Code in case Arr is not allocated.
			 ''''''''''''''''''''''''''''''''''''
			 Exit Function
		 End If
	Else
		Debug.Print "Input is not an array."
		'''''''''''''''''''''''''''''''''''
		' Code in case Arr is not an array.
		'''''''''''''''''''''''''''''''''''
		Exit Function
	End If
	SumArray = Total
End Function
````
You'll notice that this version of SumArray has much more error checking that the earlier version. This is because in the first version Arr was declared as an array of Longs. This means that we didn't need to test whether it was an array and we didn't need to test whether its elements were numeric. But in this later version, Arr is a Variant that can contain anything, so we need more error checking to ensure everything is valid. With this code, you can pass to an array of any type to  SumArray.

NOTE: You cannot pass an array as an Optional parameter to a procedure. If you need this sort of functionality, declare the parameter As Variant and then use the `IsArray` function to test whether the parameter is in fact an array.

## Using ParamArray
An alternative to passing an array is to use the

### Returning An Array From A Function
Beginning with VBA version 6 (Office 2000 and later), a Function procedure or a Property Get procedure may return an array as its result. (In Office97, the function must store the array in a Variant and return the Variant.) The variable that receives the array result must be a dynamic array and it must have the same data type as the returned array. You cannot declare the receiving array as an array of Variants to accept an array of any type. This will not work. The receiving array must have the same data type as the returned array or it must be a single Variant (not an array of Variants).

For example, the following function will load an array with numbers from Low to High and return the array as its result. Note that the variable Arr in AAATest and the return type of LoadNumbers have the same data type (Long).

````vba
Sub AAATest()
	Dim Arr() As Long
	Dim N As Long
	Arr = LoadNumbers(Low:=101, High:=110)
	If IsArrayAllocated(Arr:=Arr) = True Then
		For N = LBound(Arr) To UBound(Arr)
		   Debug.Print Arr(N)
		Next N
	Else
		''''''''''''''''''''''''''''''''''''
		' Code in case Arr is not allocated.
		''''''''''''''''''''''''''''''''''''
	End If
End Sub

Function LoadNumbers(Low As Long, High As Long) As Long()
	'''''''''''''''''''''''''''''''''''''''
	' Returns an array of Longs, containing
	' the numbers from Low to High. The 
	' number of elements in the returned
	' array will vary depending on the 
	' values of Low and High.
	''''''''''''''''''''''''''''''''''''''''
	
	'''''''''''''''''''''''''''''''''''''''''
	' Declare ResultArray as a dynamic array
	' to be resized based on the values of
	' Low and High.
	'''''''''''''''''''''''''''''''''''''''''
	Dim ResultArray() As Long
	Dim Ndx As Long
	Dim Val As Long
	'''''''''''''''''''''''''''''''''''''''''
	' Ensure Low <= High
	'''''''''''''''''''''''''''''''''''''''''
	If Low > High Then
		Exit Function
	End If
	'''''''''''''''''''''''''''''''''''''''''
	' Resize the array
	'''''''''''''''''''''''''''''''''''''''''
	ReDim ResultArray(1 To (High - Low + 1))
	''''''''''''''''''''''''''''''''''''''''
	' Fill the array with values.
	''''''''''''''''''''''''''''''''''''''''
	Val = Low
	For Ndx = LBound(ResultArray) To UBound(ResultArray)
		ResultArray(Ndx) = Val
		Val = Val + 1
	Next Ndx
	''''''''''''''''''''''''''''''''''''''''
	' Return the array.
	''''''''''''''''''''''''''''''''''''''''
	LoadNumbers = ResultArray()

End Function
````

Note that the array Arr in AAATest has the same data type (Long) as the array returned by LoadNumbers. These data types must match. You cannot declare Arr in AAATest as an array of Variants to receive an array of any data type. If you do, you'll receive a "Can't Assign To Array" compiler error. You will receive the same compiler error if Arr is a static array. The array that is set to the return value of a function must be a dynamic array. It may be allocated, in which case it will be resized automatically to hold the result array, either increasing or decreasing its size. It is not required, though, that the receiving array be allocated. Regardless of whether the receiving array is allocated, it will be automatically sized to match the size of the returned array. 

You can, however, declare the receiving variable as a single Variant. For example, you could use

````vba
Dim Arr As Variant
in place of
Dim Arr() As Long
For example, in the following code, the Arr array will be resized from 100 down to 10 when it receives the result of the LoadNumbers function.

Sub AAATest()
	Dim Arr() As Long
	Dim N As Long
	ReDim Arr(1 To 100)
	Debug.Print "BEFORE LoadNumbers: Number Of Elements in Arr: " & CStr(UBound(Arr) - LBound(Arr) + 1)
	Arr = LoadNumbers(Low:=101, High:=110)
	Debug.Print "AFTER LoadNumbers: Number Of Elements in Arr: " & CStr(UBound(Arr) - LBound(Arr) + 1)
End Sub
Function LoadNumbers(Low As Long, High As Long) As Long()
	'''''''''''''''''''''''''''''''''''''''''''''''''
	' LoadNumbers
	' Returns an array of Longs containing the numbers 
	' between Low and High.
	''''''''''''''''''''''''''''''''''''''''''''''''' 
	Dim ResultArray() As Long
	Dim Ndx As Long
	Dim Val As Long
	If Low > High Then
		Exit Function
	End If
	ReDim ResultArray(1 To (High - Low + 1))
	Val = Low
	For Ndx = LBound(ResultArray) To UBound(ResultArray)
		ResultArray(Ndx) = Val
		Val = Val + 1
	Next Ndx
	LoadNumbers = ResultArray()
End Function

````
If the receiving array has a base index (LBound) that differs from the array it receives, the receiving array will take on a new base value from the returned array. For example,

````vba
    Sub AAATest()
		Dim Arr() As Long
		Dim N As Long
		'''''''''''''''''''''''''''
		' Set the lower and upper
		' bounds of Arr to 0 and 9
		' respectively.
		'''''''''''''''''''''''''''
		ReDim Arr(0 To 9)
		Debug.Print "BEFORE LoadNumbers: LBound: " & CStr(LBound(Arr)) & "  UBound: " & CStr(UBound(Arr))
		' LoadNumbers uses the a lower bound of 1, not 0
		 Arr = LoadNumbers(Low:=101, High:=110)
		' Note that the LBound is now 1 and the UBound is now 10.
		Debug.Print "AFTER LoadNumbers:  LBound: " & CStr(LBound(Arr)) & "  UBound: " & CStr(UBound(Arr))
	End Sub
````

The code above shows that the LBound of Arr was changed from 0 to 1, the LBound of the result array declared and allocated in the  LoadNumbers procedure.

A function can also return a Variant containing an array. Even in this case, the receiving array must have the same data type as the array that is stored in the Variant. For example,

````vba
Sub AAATest()
	Dim Arr() As Long
	Dim N As Long
	Arr = LoadNumbers(Low:=101, High:=110)
	If IsArrayAllocated(Arr:=Arr) = True Then
		For N = LBound(Arr) To UBound(Arr)
			Debug.Print Arr(N)
		Next N
	Else
		''''''''''''''''''''''''''''''''''''
		' Code in case Arr is not allocated.
		''''''''''''''''''''''''''''''''''''
	End If
End Sub

Function LoadNumbers(Low As Long, High As Long) As Variant ' note we return Variant, not Long()
	'''''''''''''''''''''''''''''''''''''''''''''''''
	' LoadNumbers
	' Returns a Variant containing an array containing
	' the numbers between Low and High.
	''''''''''''''''''''''''''''''''''''''''''''''''' 
	Dim ResultArray() As Long
	Dim Ndx As Long
	Dim Val As Long
	If Low > High Then
		Exit Function
	End If
	ReDim ResultArray(1 To (High - Low + 1))
	Val = Low
	For Ndx = LBound(ResultArray) To UBound(ResultArray)
		ResultArray(Ndx) = Val
		Val = Val + 1
	Next Ndx
	LoadNumbers = ResultArray()
End Function
````

If the calling procedure doesn't know what type of data will be in the array returned by a function, it can use a Variant variable to store the result. The Variant will contain the array. Since this version of LoadNumbers returns a Variant, we can make no assumptions about what it might return. The code should test for all contingencies to avoid an unexpected run-time error.

````vba
Sub AAATest()
	Dim Arr As Variant  ' note this is declared As Varaint, not As Long()
	Dim N As Long
	Arr = LoadNumbers(Low:=101, High:=110)
	'''''''''''''''''''''''''
	' Ensure Arr is an array.
	'''''''''''''''''''''''''
	If IsArray(Arr) = True Then
		''''''''''''''''''''''''''''''''
		' Ensure the array is allocated.
		''''''''''''''''''''''''''''''''
		If IsArrayAllocated(Arr:=Arr) = True Then
			''''''''''''''''''''''''''''''''
			' Ensure Arr is one-dimensional.
			''''''''''''''''''''''''''''''''
			If NumberOfArrayDimensions(Arr:=Arr) = 1 Then
				'''''''''''''''''''''''''''''''''
				' Loop through the returned array
				'''''''''''''''''''''''''''''''''
				For N = LBound(Arr) To UBound(Arr)
					''''''''''''''''''''''''''
					' Ensure Arr(N) is numeic.
					''''''''''''''''''''''''''
					If IsNumeric(Arr(N)) = True Then
						Debug.Print Arr(N)
					Else
						Debug.Print "Arr(N) is not numeric."
						'''''''''''''''''''''''''''''''''''''
						' Code in case Arr(N) is not numeric.
						'''''''''''''''''''''''''''''''''''''
					End If
				Next N
			Else
				Debug.Print "Arr is not one-dimensional."
				''''''''''''''''''''''''''''''''''''''''
				' Code in case Arr is multi-dimensional.
				''''''''''''''''''''''''''''''''''''''''
			End If
		Else
			Debug.Print "Arr is not allocated."
			''''''''''''''''''''''''''''''''''''
			' Code in case Arr is not allocated.
			''''''''''''''''''''''''''''''''''''
		End If
	Else
		Debug.Print "Arr is not an array."
		'''''''''''''''''''''''''''''''''''
		' Code in case Arr is not an array.
		'''''''''''''''''''''''''''''''''''
	End If
End Sub

Function LoadNumbers(Low As Long, High As Long) As Variant
   '''''''''''''''''''''''''''''''''''''''''''''''''
	' LoadNumbers
	' Returns a Variant containing an array containing
	' the numbers between Low and High.
	''''''''''''''''''''''''''''''''''''''''''''''''' 
	Dim ResultArray() As Long
	Dim Ndx As Long
	Dim Val As Long
	If Low > High Then
		Exit Function
	End If 
	ReDim ResultArray(1 To (High - Low + 1))
	Val = Low
	For Ndx = LBound(ResultArray) To UBound(ResultArray)
		ResultArray(Ndx) = Val
		Val = Val + 1
	Next Ndx
	LoadNumbers = ResultArray()
End Function
````

As you can see, using Variants to store arrays gives you considerably more flexibility, but it also leaves much more room for error or invalid data. If you are using Variants, your code should contain error checking routines to ensure that you are dealing with the type of data you expect.

### Assigning An Array To An Array
Unfortunately, VBA doesn't let you assign one array to another array, even if the size and data types match. For example, the following code will not work:

````vba
    Dim A(1 To 10) As Long
	Dim B(1 To 10) As Long
	' load B with data
	A = B
You can, however,  assign a Variant containing an array to another Variant. The following code is perfectly legal:

    Dim A As Variant
	Dim B As Variant
	Dim N As Long
	A = Array(11, 22, 33)
	B = A
	Debug.Print "IsArray(B) = " & CStr(IsArray(B))
	For N = LBound(B) To UBound(B)
		Debug.Print B(N)
	Next N
````

If you need to transfer the contents of one array to another, you must loop through the array element-by-element:

````vba
    Dim A(1 To 3) As Long
	Dim B(0 To 5) As Long
	Dim NdxA As Long
	Dim NdxB As Long

	A(1) = 11
	A(2) = 22
	A(3) = 33
	NdxB = LBound(B)
	For NdxA = LBound(A) To UBound(A)
		If NdxB <= UBound(B) Then
			B(NdxB) = A(NdxA)
		Else
			Exit For
		End If
		NdxB = NdxB + 1
	Next NdxA

	For NdxB = LBound(B) To UBound(B)
		Debug.Print B(NdxB)
	Next NdxB
````

The code above will transfer the contents of array A to array B. It does this successfully even if A and B have different LBounds, and will terminate the loop of the UBound of B is exceeded, which would be the case if A contains more elements than B. If A contains fewer elements than B, the unused elements of B will remain intact. If you want to ensure the B is "clean" before transferring the elements of A to it, use the Erase statement and, if B is a dynamic array, ReDim it back to its original size, as shown below:

````vba
    Dim SaveLBound As Long
	Dim SaveUBound As Long
	SaveLBound = LBound(B)
	SaveUBound = UBound(B)
	Erase B
	If IsArrayDynamic(Arr:=B) = True Then
		ReDim B(SaveLBound, SaveUBound)
	End If
````
 
## Multi-Dimensional Arrays

So far, all of the procedures and techniques described have used single-dimensional arrays. So what about multi-dimensional arrays? The short answer is that the same rules and techniques that apply to single-dimensional arrays apply to multi-dimensional arrays. On the Functions For VBA Arrays page, there is a function named `NumberOfArrayDimensions` that will return the number of dimensions of an array. (It returns 0 for unallocated dynamic arrays). You can use this function to determine the number of dimensions of either a static or dynamic array. You can pass a multi-dimensional array to a procedure, as shown in the code below.

````vba
Sub AAATest()
	Dim N As Long
	Dim Sum As Long
	''''''''''''''''''''''''''
	' Declare a dynamic array
	'''''''''''''''''''''''''
	Dim Arr() As Long
	''''''''''''''''''''''''''
	' Size the array for two
	' dimensions.
	''''''''''''''''''''''''''
	ReDim Arr(1 To 2, 1 To 3)
	''''''''''''''''''''''''''
	' Put in some values.
	''''''''''''''''''''''''''
	Arr(1, 1) = 1
	Arr(1, 2) = 2
	Arr(1, 3) = 3
	Arr(2, 1) = 4
	Arr(2, 2) = 5
	Arr(2, 3) = 6
	'''''''''''''''''''''''''''
	' SumMulti will return the
	' sum of the element in a
	' 1 or 2 dimensional array.
	'''''''''''''''''''''''''''
	Sum = SumMulti(Arr:=Arr)
	Debug.Print Sum
End Sub

Function SumMulti(Arr() As Long) As Long
	Dim N As Long
	Dim Ndx1 As Long
	Dim Ndx2 As Long
	Dim NumDims As Long
	Dim Total As Long
	''''''''''''''''''''''''''''''''''''''''''
	' Get the number of array dimensions.
	' NumberOfArrayDimensions will return 0
	' if the array is not allocated.
	'''''''''''''''''''''''''''''''''''''''''
	NumDims = NumberOfArrayDimensions(Arr:=Arr)
	Select Case NumDims
		Case 0
			''''''''''''''''''''''''''''''''
			' unallocated array
			''''''''''''''''''''''''''''''''
			SumMulti = 0
			Exit Function
		Case 1
			''''''''''''''''''''''''''''''''
			' single dimensional array
			''''''''''''''''''''''''''''''''
			For N = LBound(Arr) To UBound(Arr)
				 Total = Total + Arr(N)
			Next N
		Case 2
			'''''''''''''''''''''''''''''''''
			' 2 dimensional array
			'''''''''''''''''''''''''''''''''
			For Ndx1 = LBound(Arr, 1) To UBound(Arr, 1)
				For Ndx2 = LBound(Arr, 2) To UBound(Arr, 2)
					Total = Total + Arr(Ndx1, Ndx2)
				Next Ndx2
			Next Ndx1
		Case Else
			''''''''''''''''''''''''''''''''
			' Too many dimensions.
			''''''''''''''''''''''''''''''''
		   MsgBox "SumMulti works only on 1 or 2 dimensional arrays."
		   Total = 0
	End Select
	''''''''''''''''''
	' return the total
	''''''''''''''''''
	SumMulti = Total
End Function
````

Functions can return multi-dimensional arrays just as they can single-dimensional arrays. The same rules apply: the array receiving the result must be a dynamic array and must have the same data type as the returned array. For example,

````vba
Sub AAATest()
	''''''''''''''''''''''''
	' Dynamic array to hold
	' the result.
	''''''''''''''''''''''''
	Dim ReturnArr() As Long
	Dim Ndx1 As Long
	Dim Ndx2 As Long
	Dim NumDims As Long
	''''''''''''''''''''''''''
	' call the function to get
	' the result array.
	''''''''''''''''''''''''''
	ReturnArr = ReturnMulti()
	NumDims = NumberOfArrayDimensions(Arr:=ReturnArr)
	Select Case NumDims
		Case 0
			'''''''''''''''''''
			' unallocated array
			'''''''''''''''''''
		Case 1
			''''''''''''''''''''''''''
			' single dimensional array
			''''''''''''''''''''''''''
			For Ndx1 = LBound(ReturnArr) To UBound(ReturnArr)
				Debug.Print ReturnArr(Ndx1)
			Next Ndx1
		Case 2
			'''''''''''''''''''''''''''
			' two dimensional array
			'''''''''''''''''''''''''''
			For Ndx1 = LBound(ReturnArr, 1) To UBound(ReturnArr, 1)
				For Ndx2 = LBound(ReturnArr, 2) To UBound(ReturnArr, 2)
					Debug.Print ReturnArr(Ndx1, Ndx2)
				Next Ndx2
			Next Ndx1
		Case Else
			''''''''''''''''''''''
			' too many dimensions
			''''''''''''''''''''''
	End Select
End Sub

Function ReturnMulti() As Long()
  ''''''''''''''''''''''''''''''''''''
  ' Returns a mutli-dimensional array.
  ''''''''''''''''''''''''''''''''''''
  Dim A(1 To 2, 1 To 3) As Long
  '''''''''''''''''''''''''''''
  ' put in some values.
  '''''''''''''''''''''''''''''
  A(1, 1) = 100
  A(1, 2) = 200
  A(1, 3) = 300
  A(2, 1) = 400
  A(2, 2) = 500
  A(2, 3) = 600
  ReturnMulti = A()
End Function
````

### Looping Through Arrays
You may have noticed that when looping through the arrays, all the  code above uses `LBound(Arr)` and `UBound(Arr)` as the lower and upper limits of the 
loop index. This is the safest way to loop through an array. This will properly set the loop index variable bounds regardless of how the array was sized, and 
regardless of the Option Base module setting. Good programming practice dictates that you use `LBound` and `UBound` rather than hard-coding the lower 
and upper values for the loop index. It is a bit more typing, but will ensure that the complete array is looped through and will avoid _Subscript Out Of Range_ run time errors.

Using arrays is a powerful technique in VBA, and passing and returning arrays to and from function only adds to the range of possibilities of your code.
Properly understanding how to pass arrays between procedures is a critical to successfully using array in your applications.
