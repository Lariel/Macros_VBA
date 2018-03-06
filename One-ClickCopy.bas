Sub CopyEmail()
Dim oExplorer As Outlook.Explorer
Dim oMail As MailItem
Dim string1 As String, string3 As String

Set oExplorer = Application.ActiveExplorer
Set oMail = oExplorer.Selection.Item(1).Forward

With oMail
.BodyFormat = olFormatPlain

' Texto completo do email
string1 = .Body
End With

' Texto entre "From: "
string2 = Split(string1, "From: ", , vbDatabaseCompare)
oMail.Close olDiscard

' Encontrar posição do "De: " no string2
' Retorna 0 se não encontrar
intPos = InStr(string2(1), "De: ")
If intPos > 0 Then
    ' De: encontrado
    ' string3 recebe todo o conteúdo na esquerda do string2
    string3 = Left(string2(1), intPos - 1)
Else
    ' De: não encontrado
    string3 = string2(1)
End If

CopyTextToClipboard ("From: " & string3)
End Sub

Sub CopyTextToClipboard(ByVal inText As String)
Dim objClipboard As Object
Set objClipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")

objClipboard.SetText inText
objClipboard.PutInClipboard

Set objClipboard = Nothing
End Sub

Function GetTextFromClipboard() As String
Dim objClipboard As Object
Set objClipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")

objClipboard.GetFromClipboard
GetTextFromClipboard = objClipboard.GetText

Set objClipboard = Nothing
End Function
