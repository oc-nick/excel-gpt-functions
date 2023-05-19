Function GPTgetCC(prompt As String) As String

Dim url As String, apiKey As String
Dim response As Object, json As String

apiKey = ".."
url = "https://api.openai.com/v1/engines/text-davinci-003/completions"

Set response = CreateObject("MSXML2.XMLHTTP")
response.Open "POST", url, False
response.setRequestHeader "Content-Type", "application/json"
response.setRequestHeader "Authorization", "Bearer " + apiKey
response.Send "{""prompt"":""You are an excel function. Return only the answer with no additional text or explanation. Your function is to find credit card details in this string of random text. Return the credit card number with no dashes or spaces [SPACE] expiration in MM/YYYY format [SPACE] CVV: '" & prompt & "'"",""max_tokens"":1024,""temperature"":0}"

json = response.responseText
GPTgetCC = Split(Mid(json, InStr(json, """text"":""") + 8), """")(0)
GPTgetCC = Replace(GPTgetCC, "\n", "")

End Function



Function GPTdelCC(prompt As String) As String

Dim url As String, apiKey As String
Dim response As Object, json As String

apiKey = ".."
url = "https://api.openai.com/v1/engines/text-davinci-003/completions"

Set response = CreateObject("MSXML2.XMLHTTP")
response.Open "POST", url, False
response.setRequestHeader "Content-Type", "application/json"
response.setRequestHeader "Authorization", "Bearer " + apiKey
response.Send "{""prompt"":""You are an excel function. Return only the answer with no additional text or explanation. Your function is to remove all of the credit card-related information from this string and return all of the remaining text. Do not return any credit card information: " & prompt & """,""max_tokens"":1024,""temperature"":0}"

json = response.responseText
GPTdelCC = Split(Mid(json, InStr(json, """text"":""") + 8), """")(0)
GPTdelCC = Replace(GPTdelCC, "\n", "")

End Function