<div align="center">

## If a users attempts to shutdown/restart their computer while your application is still running


</div>

### Description

If a users attempts to shutdown/restart their computer while your application is still running, this simple piece of code will actually allow you to abort the request to shutdown. Can be very useful. Please mail me at mailme@shouvik.tk if you are facing any problems or www.shouvik.tk
 
### More Info
 
www.cisindia.net = Free Software. www.bnetsupport.com and www.vexat.net = Free Tutorials/Ebooks.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[www\.cupidsystems\.com](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/www-cupidsystems-com.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/www-cupidsystems-com-if-a-users-attempts-to-shutdown-restart-their-computer-while-your-app__1-48766/archive/master.zip)

### API Declarations

```
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function AbortSystemShutdown Lib "advapi32.dll" Alias "AbortSystemShutdownA" (ByVal lpMachineName As String) As Long
Function GetName()
 Dim lpBuff As String * 25
 Dim ret As Long, ComputerName As String
 ret = GetComputerName(lpBuff, 25)
 ComputerName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
 GetName = ComputerName
End Function
```


### Source Code

```
Private Sub Command1_Click()
 AbortSystemShutdown (GetName)
End Sub
```

