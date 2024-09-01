' Define constants
Const password = "password"

' Method to perform obfuscation
Function Obfuscate(inputText)
    Dim b64Encoded, rotatedText, i
    b64Encoded = EncodeBase64(inputText)
    
    ' Rotate each character by +1 in ASCII
    rotatedText = ""
    For i = 1 To Len(b64Encoded)
        rotatedText = rotatedText & Chr(Asc(Mid(b64Encoded, i, 1)) + 1)
    Next
    Obfuscate = rotatedText
End Function

' Method to perform deobfuscation
Function Deobfuscate(obfuscatedText)
    Dim b64Decoded, rotatedText, i
    rotatedText = ""
    
    ' Rotate each character by -1 in ASCII
    For i = 1 To Len(obfuscatedText)
        rotatedText = rotatedText & Chr(Asc(Mid(obfuscatedText, i, 1)) - 1)
    Next
    
    ' Decode from Base64
    b64Decoded = DecodeBase64(rotatedText)
    Deobfuscate = b64Decoded
End Function

' Function to encode text to Base64
Function EncodeBase64(inputText)
    Dim objXML, objNode
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.CreateElement("base64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = StreamBinary(inputText)
    EncodeBase64 = objNode.Text
End Function

' Function to decode Base64 text
Function DecodeBase64(base64Text)
    Dim objXML, objNode
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.CreateElement("base64")
    objNode.DataType = "bin.base64"
    objNode.Text = base64Text
    DecodeBase64 = StreamBinaryToText(objNode.nodeTypedValue)
End Function

' Converts string to binary
Function StreamBinary(inputText)
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' Text stream
    stream.Charset = "us-ascii"
    stream.Open
    stream.WriteText inputText
    stream.Position = 0
    stream.Type = 1 ' Binary stream
    StreamBinary = stream.Read
    stream.Close
End Function

' Converts binary to string
Function StreamBinaryToText(binaryData)
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' Binary stream
    stream.Open
    stream.Write binaryData
    stream.Position = 0
    stream.Type = 2 ' Text stream
    stream.Charset = "us-ascii"
    StreamBinaryToText = stream.ReadText
    stream.Close
End Function

' Function to read file content
Function ReadFileContent(filePath)
    Dim fso, file, fileContent
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(filePath, 1)
    fileContent = file.ReadAll
    file.Close
    ReadFileContent = fileContent
End Function

' Function to write file content
Sub WriteFileContent(filePath, content)
    Dim fso, file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(filePath, 2, True)
    file.Write content
    file.Close
End Sub

' Prompt user for password
Dim input, filePath, fileContent
input = InputBox("Enter password")

' Check password
If input = password Then
    ' Simulate file selection (using predefined file path)
    filePath = InputBox("Password correct! Enter file path", "File Prompt", "potato.txt")

    If filePath <> "" Then
        ' Read the file content
        fileContent = ReadFileContent(filePath)
        
        ' Try to deobfuscate the file content
        On Error Resume Next
        Dim deobfuscatedContent
        deobfuscatedContent = Deobfuscate(fileContent)
        
        ' Check for errors during deobfuscation
        If Err.Number <> 0 Then
            MsgBox "Error deobfuscating the file. Obfuscating instead.", 0 + 16, "Error"
            
            ' Clear the error
            Err.Clear
            
            ' Obfuscate the file content and overwrite the file
            Dim obfuscatedContent
            obfuscatedContent = Obfuscate(fileContent)
            WriteFileContent filePath, obfuscatedContent
            
            MsgBox "File has been mashed up more than tuesday's potatoes", 0 + 64, "Success"
        Else
            ' No errors, display the deobfuscated content
            MsgBox "Super secret message: " & vbCrLf & deobfuscatedContent
        End If
        On Error GoTo 0
    Else
        MsgBox "No file selected.", 0 + 16, "Error"
    End If
Else
    MsgBox "Incorrect password. Access denied.", 0 + 16, "Error"
End If
