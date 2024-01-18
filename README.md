# VBScript Cheat Sheet
- Comments begin with '

      'This is a comment
- If you begin a script with **```Option Explicit```**, then you can't use a variable before declaring it.<br/>You must declare all variables using **```Dim```** keyword
- There are 2 ways of calling a function
  - Function_name PARAMS, SEPARATED, BY, COMMAS
  - Call Function_name(PARAMS, SEPARATED, BY, COMMAS)
- Arrays are not dynamic in VBS
- Making API calls

      Dim oWinHTTP
      Set oWinHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
      oWinHTTP.Open "GET", "http://remoteserver/thing.ext", False
      oWinHTTP.SetRequestHeader "User-Agent", "My Agent String"
      oWinHTTP.Send
      response = oWinHTTP.responseText
