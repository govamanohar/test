Update VB6 Code
String Operations:
VB.Net 2010 continues to support many VB6 methods. HOWEVER, they are actually
SLOWER than their VB.NET counterparts, so it is very important to go through each
line of code in your project, and replace each VB6 method with it's .NET
counterpart.
VB6 code runs good in VB6, but VB6 code in VB.NET runs bad (very bad). VB.NET
code in VB.NET runs good (VERY good), much faster than VB6 code runs in VB6 (if
that makes sense).
So here are some examples of how to replace VB6 methods with VB.NET
counterparts:
'TIP: VB.Net strings are zero based, in other words, the first position of a string is 0.
In VB6, the first position was 1. This greatly affects how strings are parsed.
CodeDim myString As String = "Go ahead and search for this string"
 Instr - Instead of using the Instr() method to search a string, use the IndexOf()
method.
Code: Instr(myString, "search for this string") --VB
myString.IndexOf("search for this string") -- VB.net
 Mid - Instead of using the Mid() method to get a portion of a string, use the
SubString() method.
Code: Mid(myString, 14) --VB
myString.SubString(13) --VB.net
 Trim - Instead of using the Trim(), LTrim() and RTrim(), use .Trim(), .TrimStart(),
.TrimEnd()
Code: Trim(myString), LTrim(myString), RTrim(myString) --VB
myString.Trim(), myString.TrimStart(), mystring.TrimEnd() --VB.net
 Len - Instead of using the Len() method, use .Length() to get the length of a
string.
Code : Len(myString) --VB
myString.Length() --VB.net
 Replace - Replace the "And" operator with "AndAlso". (Do this in any nonbitwise
comparison).
Code: If 1 = 1 And 2 = 2 And 3 = 3 Then --VB
If 1 = 1 AndAlso 2 = 2 AndAlso 3 = 3 Then --VB.net
 Replace - Replace the "Or" operator with "OrElse". (Do this in any non-bitwise
comparison.)
Code : If 1 = 1 Or 2 = 2 Or 3 = 3 Then --VB
If 1 = 1 OrElse 2 = 2 OrElse 3 = 3 Then --VB.net
 Replace - Replace ALL VB6 File I/O classes with the new .NET File I/O Classes.
They are faster than VB6's so make sure you use them!
Dim myFile As String = "C:\Temp\myfile.txt"
Dim instring As String = String.Empty
///////////////////////////////////////////////////////////////////////////////////////////


///Extension Methods are one of the new features introduced in .Net 3.0. So,
while migrating from VB 6 to VB.Net Application, we make use of extension
methods to reduce our coding effort.
Definition:
Extension methods enable developers to add custom functionality to data
types that are already defined without creating a new derived type.
Extension methods make it possible to write a method that can be called as if
it were an instance method of the existing type.
We created the below extension methods for string class.
i. Left
ii. InStr
iii. Right
iv. Append
<Extension()>
Public Function Left(ByVal str As String, ByVal intLength As Integer) As String
If (String.IsNullOrEmpty(str) Or intLength < 1) Then
Return String.Empty
Else
Return str.Substring(0, Math.Min(intLength, str.Length))
End If
End Function
<Extension()>
Public Function InStr(ByVal strBaseString As String, ByVal stringToCompare As
String, Optional ByVal startIndex As Integer = -1) As Integer
Dim strTempString As String = String.Empty
If startIndex <> -1 Then
strTempString = strBaseString.Substring(startIndex, strBaseString.Length - 1)
Else
strTempString = strBaseString
End If
Return strTempString.IndexOf(stringToCompare) + 1
End Function
<Extension()>
Public Function Right(ByVal str As String, ByVal intLength As Integer) As String
If (String.IsNullOrEmpty(str) Or intLength < 1) Then
Return String.Empty
Else
Return str.Substring(str.Length - intLength)
End If
End Function
<Extension()>
Public Function StartAt(ByVal str As String, ByVal startLength As Integer) As String
If (String.IsNullOrEmpty(str) Or startLength < 1) Then
Return String.Empty
Else
Return str & "@St@" & startLength
End If
Return String.Empty
End Function
<Extension()>
Public Function Append(ByVal str As String, Optional ByVal AppendString As String
= ",") As String
Return str.ToString & AppendString
End Function
By creating these methods, code converted from VB 6 using these methods
need not be replaced with .Net equivalent methods.
There are lots of common methods that can be used similar to VB6 functions
and they are grouped in the below common file which we can included in any
.Net Application migrated from VB6 to reduce the development effort.

////////////////////////////////////////
