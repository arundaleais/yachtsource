Attribute VB_Name = "modEncrypt"
Option Explicit

'Requires CAPICOM V2.1 Project > Reference including
'Returns unencrypted string

Private CapicomNotFound As Boolean

Public Function Decrypt(kb As String) As String
    Dim Secret As EncryptedData
WriteLog "Decrypting"
    On Error GoTo Decrypt_err
'Err.Raise 429   'test cant create active x
    Set Secret = New EncryptedData
    Secret.Algorithm.Name = CAPICOM_ENCRYPTION_ALGORITHM_AES
    Secret.Algorithm.KeyLength = CAPICOM_ENCRYPTION_KEY_LENGTH_256_BITS
    Secret.SetSecret "My Secret Encryption Key"
'    Secret.Content = "password" ' just so we know that this is being reset by decryption
    Secret.Decrypt kb
    Decrypt = Secret.Content
WriteLog "Decrypted"
Exit Function

Decrypt_err:
    WriteLog "Decrypt Error " & Str(Err.Number) & " " & Err.Description
    Select Case Err.Number
    Case Is = 429   'cant create active x object
        Err.Raise 429, "Decrypt"    'pass error back to DecryptFile
        Exit Function
    End Select
    MsgBox "Decrypt Error " & Str(Err.Number) & " " & Err.Description, vbCritical, "Decrypt"
    Resume Next
'MsgBox Decrypt
End Function

Public Function Encrypt(kb As String) As String
Dim Secret As EncryptedData
    Set Secret = New EncryptedData
    Secret.Algorithm.Name = CAPICOM_ENCRYPTION_ALGORITHM_AES
    Secret.Algorithm.KeyLength = CAPICOM_ENCRYPTION_KEY_LENGTH_256_BITS
    Secret.SetSecret "My Secret Encryption Key"
    Secret.Content = kb ' what we want to encrypt
    Encrypt = Secret.Encrypt
'For Password encryption (AisDecoder)
'we must remove the split lines secret.content includes
'    Encrypt = Replace(Secret.Encrypt, vbCrLf, "")
'    MsgBox Encrypt
End Function

Public Function DecryptFile(EncryptedFileName As String, DecryptedFileName As String)
Dim EncryptedLines As String
Dim DecryptedLines As String
Dim ch As Long
Dim kb As String
Dim FileNotFound As Boolean

WriteLog "Copying " & EncryptedFileName & " to memory"
'MsgBox "Decrypting " & EncryptedFileName & vbCrLf & "to " & DecryptedFileName
    On Error GoTo DecryptFile_err
    ch = FreeFile
    Open EncryptedFileName For Input As #ch
    EncryptedLines = StrConv(InputB(LOF(ch), ch), vbUnicode)
    Close ch
    If EncryptedLines = "" Then Exit Function   'file is empty
    DecryptedLines = Decrypt(EncryptedLines)
    If DecryptedLines <> "" Then    'decrypt failed - could be active x
        Open DecryptedFileName For Output As #ch Len = Len(DecryptedLines)
        Print #ch, DecryptedLines
        Close #ch
    Else    'try for unencrypted file
        FileCopy Replace(DecryptedFileName, ".tmp", ".txt"), DecryptedFileName
        ch = FreeFile
        Open DecryptedFileName For Input As #ch
        EncryptedLines = StrConv(InputB(LOF(ch), ch), vbUnicode)
        Close ch
    End If
Exit Function

DecryptFile_err:
    WriteLog "DecryptFile Error " & Str(Err.Number) & " " & Err.Description & kb
    Select Case Err.Number
    Case Is = 52, 53
        kb = vbCrLf & EncryptedFileName & vbCrLf
        FileNotFound = True
    Case Is = 429   'passed back from Encrypt (cant create active x)
        CapicomNotFound = True
        Resume Next
    End Select
    MsgBox "DecryptFile Error " & Str(Err.Number) & " " & Err.Description & kb, vbCritical, "DecryptFile"
    kb = ""
'Exit subroutine if file not found of any other error
'    If Err <> 53 Then
'        Resume Next
'    End If
'MsgBox Decrypt
End Function

Public Function EncryptFile(DecryptedFileName As String, EncryptedFileName As String)
Dim EncryptedLines As String
Dim DecryptedLines As String
Dim ch As Long
Dim l As Integer

'MsgBox "Encrypting " & DecryptedFileName & vbCrLf & "to " & EncryptedFileName
    ch = FreeFile
    Open DecryptedFileName For Input As #ch
    DecryptedLines = StrConv(InputB(LOF(ch), ch), vbUnicode)
    Close ch
    EncryptedLines = Encrypt(DecryptedLines)
    Open EncryptedFileName For Output As #ch '    Len = Len(EncryptedLines)
    Print #ch, EncryptedLines
    Close #ch
End Function

Public Function EncryptFiles(FilePath As String, DecryptedExt As String, EncryptedExt As String)
Dim DecryptedFileName As String
Dim EncryptedFileName As String

    If CapicomNotFound = True Then Exit Function    'Can occur when RacingSignals profile loaded

'If the user does not have admin they will not be able to access the directory
'This causes the program to fail. Only JNA is expected to require .txt files
'encrypting
    On Error GoTo file_access_err
'MsgBox FilePath
    DecryptedFileName = Dir$(FilePath & "*" & DecryptedExt)
    Do While DecryptedFileName > ""
        If Right$(DecryptedFileName, Len(DecryptedExt)) = DecryptedExt Then
            EncryptedFileName = Replace(DecryptedFileName, DecryptedExt, EncryptedExt)
            Call EncryptFile(FilePath & DecryptedFileName, FilePath & EncryptedFileName)
        End If
        DecryptedFileName = Dir$
    Loop
Exit Function

file_access_err:

End Function
