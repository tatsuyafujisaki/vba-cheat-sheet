# Note
* `Environ$("ComputerName")` gets the computer name.

## Access levels if not explicitly specified
Keyword|Access level if not explicitly specified
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
  * Use `ByVal` rather than `ByRef` 
  * Use `ThisWorkbook.Sheet1` rather than `ThisWorkbook.Worksheets("Sheet1")`.

## Best practices for maintainability
* Don't specify default access levels.
* Use square brackets to reference cells.

* A rather than B
  * Use `Private` rather than `Public`.
  * Use `vbNewLine` rather than `vbCrLf`.
  * Use `vbNullString` rather than `""`.
  * Use `Application.PathSeparator` rather than `\`.
  * In`ThisWorkbook`, use `Me` rather than `ThisWorkbook`.
  * Use `Addins2` rather than `Addins`.
    * `Addins2` = `Addins` + "addins currently open".
  * Use `Range.Value2` rather than `Range.Value`.
    * `Value2` returns `Currency` and `Date` as double (i.e. without formatting).

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
