# Note
* `Environ$("EnvironmentVariable1")` returns the value of EnvironmentVariable1.
* `Worksheet1.PageSetup` throws an error if no printer is available.
  * [You cannot use page setup properties in Excel if no printers were installed @ Microsoft](https://support.microsoft.com/en-us/help/291298/you-cannot-use-page-setup-properties-in-excel-if-no-printers-were-inst)

## Default access levels
Keyword|Access level if omitted
---|---
Const|Private
Procedure|Public
Function|Public

## Type-declaration characters
Type-declaration character|Type
---|---
@|Currency
#|Double
%|Integer
&|Long
!|Single
$|String

# Best practices

## Best practices for performance
* Append a type-declaration character to the end of each function name.
* A rather than B
  * Use `Cells(rowIndex, columnAlphabet)` rather than `Range(columnAlphabet & rowIndex)` when referring to a single cell using variables.
  * Use `Range(singleAddress1, singleAddress2)` rather than `Range(singleAddress1 & ":" & singleAddress2)` when referring to multiple cells using variables.
  * Use `ByVal` rather than `ByRef` 
  * Use `ThisWorkbook.Sheet1` rather than `ThisWorkbook.Worksheets("Sheet1")`.
  * Use `Dictionary` rather than `Collection`.
  * Use `Set` and `With`.
  * Use early binding rather than late binding.
  ```vba
  ' Early binding
  Dim o As New Object1
  
  ' Late binding
  Dim o As Object1
  Set o = New Object1
  ```
  * Minimize the dots. Cache frequently used properties in variables.
  * Use LongLong rather than Long or Integer in 64-bit Office.
    * Use Long in 32-bit Office because LongLong is unavailable in it.
    
## Best practices for maintainability
* Don't bother to specify default access level.
* Don't pass a Worksheet to a function if the function takes a Range because you can get the Worksheet from `Range.Worksheet`.
* Use square brackets to reference cells.

* A rather than B
  * Use `Private` rather than `Public`.
  * Use `vbNewLine` rather than `vbCrLf`.
  * Use `vbNullString` rather than `""`.
  * Use `Application.PathSeparator` rather than `\`.
  * Use `Worksheet` or `chart` ratehr than their parent `Sheet`.
  * In`ThisWorkbook`, use `Me` rather than `ThisWorkbook`.
  * Use `Addins2` rather than `Addins`.
    * `Addins2` = `Addins` + "addins currently open".
  * Use `Range.Value2` rather than `Range.Value`.
    * `Value2` returns `Currency` and `Date` as double (i.e. without formatting).

## References
[Visual Basic Concepts: Optimizing for Speed](https://msdn.microsoft.com/en-us/library/aa263514.aspx)

# References
* [Language reference VBA](https://msdn.microsoft.com/en-us/vba/vba-language-reference)
  * [DataTypes](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/data-types)
  * [Keywords](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/keywords-visual-basic-for-applications)
  * [Keyword Summaries](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/keyword-summaries)
  * [Methods](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/methods-visual-basic-for-applications)
  * [Objects](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/objects-visual-basic-for-applications)
  * [Object Browser](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/object-browser-visual-basic-for-applications)
  * [Statements](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/statements)
    * [Call Statement](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/call-statement)
  * [Error Messages](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/error-messages)
  * [Events](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/events-object-vba-add-in-object-model)
  * [WithEvents](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/withevents-keyword)
