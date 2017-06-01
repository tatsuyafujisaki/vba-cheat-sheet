###### MSDN
[Trappable Errors](https://msdn.microsoft.com/library/aa264975.aspx)

###### Best practices
* Use vbNewLine rather than vbCrLf
* Use Addins2 rather than Addins
  * Addins2 = Addins + "addins currently open"
* Use Range.Value2 rather than Range.Value
  * Value2 returns Currency and Date as double (i.e. without formatting).
