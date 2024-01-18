MsgBox "Loading exchnage rates",vbInformation,"Please wait"

'Get you API key from https://exchangeratesapi.io/ 
const apiKey = "your_api_key"
currencies = Array("USD","EUR","INR","JPY","RUB","CNY")
url = "http://api.exchangeratesapi.io/v1/latest?access_key="&apiKey&"&symbols="&Join(currencies,",")

Set req = CreateObject("MSXML2.XMLHTTP")
Call req.open("GET", url, False)
Call req.send()

response = Split(req.responseText,"{")(2)
response = Split(response,"}}")(0)
response = Split(response,",")

index = 0
Dim currencyData(5)
for each data in response
  value = Split(data,":")(1)
  currencyData(index) = CDbl(value)
  index = index + 1
Next

str = Chr(10)
index = 1
for each denom in currencies
  str = str + CStr(index) + ". " + denom + Chr(10)
  index = index + 1
Next

index1 = CInt(InputBox("Choose the index of the currency you wish to convert"&Chr(10)&str,"Exchange Rate calculator by Saptarshi Dey"))
If index1 < 1 or index1 > UBound(currencies) + 1 Then
  MsgBox "Please choose a proper option",vbCritical,"Index error"
  Call WScript.QUIT()
End If

index2 = CInt(InputBox("Choose the index of the currency you wish to convert the previously selected currency to"&Chr(10)&str,"Exchange Rate calculator by Saptarshi Dey"))
If index2 < 1 or index2 > UBound(currencies) + 1 Then
  MsgBox "Please choose a proper option",vbCritical,"Index error"
  Call WScript.QUIT()
End If

value = CDbl(InputBox("Enter the value","Exchange Rate calculator by Saptarshi Dey"))
result = value * currencyData(index2-1) / currencyData(index1-1)
result = Round(result,2)
MsgBox "The value of "&currencies(index1-1)&" "&value&" in "&currencies(index2-1)&" is "&result,vbInformation,"Exchange Rate calculator by Saptarshi Dey"
