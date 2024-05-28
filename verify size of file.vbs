' Verify the Size of a File Before Reading It


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile("C:\Windows\Netlogon.log")

If objFile.Size > 0 Then
    Set objReadFile = objFSO.OpenTextFile("E:\VBScript", 1)
    strContents = objReadFile.ReadAll
    Wscript.Echo strContents
    objReadFile.Close
Else
    Wscript.Echo "The file is empty."
End If

