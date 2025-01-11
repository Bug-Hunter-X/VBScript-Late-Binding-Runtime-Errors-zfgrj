Late Binding: VBScript's late binding can lead to runtime errors if an object or method doesn't exist.  This is especially problematic when dealing with COM objects or external libraries where the availability of objects might not be guaranteed.

Example:
```vbscript
Dim objExcel
Set objExcel = CreateObject("Excel.Application")
' ... use objExcel ...
Set objExcel = Nothing
```
If Excel is not installed, `CreateObject("Excel.Application")` will fail, causing a runtime error.

Another example:
```vbscript
Dim obj
Set obj = CreateObject("Some.Object.That.Might.Not.Exist")
If Err.Number <> 0 Then
  ' Handle the error
Else
  'Use the object 
End If
```