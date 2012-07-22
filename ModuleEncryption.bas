Attribute VB_Name = "ModuleEncryption"
Option Explicit
'Declare constants
'   KEY is the encryption key that gets XORed with each character of the provided string
'   IV is the initilization vector for the CBC aspect of the encryption. Since a given character needs to be XORed with the key and the previous ciphertext, the IV is used for the first character, because there is no previous ciphertext.
Const KEY As Byte = 165, IV As Byte = 181

'Encrypts or decrypts a string by XORing each character with a key and using the cipher-block chaining (CBC) mode of operation (each block of plaintext is XORed with the previous ciphertext)
'   StartString is the provided string that is to be encrypted or decrypted
'   Encryption determines if the data is to be encrypted (True) or decrypted (False)
Function EncryptDecrypt(StartString As String, Encryption As Boolean) As String
    
    'Declare variables
    '   count is the loop counter for each character in the given string
    '   CurrChar stores the character of the given string in each instance of the loop
    '   EndString stores the characters after they have been encrypted/decrypted
    Dim count As Integer, CurrChar As String, EndString As String
    
    'Loops to encrypt/decrypt each character in the given string
    For count = 1 To Len(StartString)
        
        'Gets the character to be encrypted/decrypted
        CurrChar = Mid(StartString, count, 1)
        
        'The character is XORed with the key
        CurrChar = Chr(Asc(CurrChar) Xor KEY)
        
        'Cipher-block chaining:
        'If it is the first character, the initilization vector is used instead of the previous ciphertext
        If count = 1 Then
            CurrChar = Chr(Asc(CurrChar) Xor IV)
        'If it is anywhere else in the string it is XORed with the previous ciphertext block
        Else
            'If the string is being encrypted, the ciphertext is in EndString
            If Encryption Then
                CurrChar = Chr(Asc(CurrChar) Xor Asc(Mid(EndString, count - 1, 1)))
            'If the string is being decrypted, the ciphertext is in StartString
            Else
                CurrChar = Chr(Asc(CurrChar) Xor Asc(Mid(StartString, count - 1, 1)))
            End If
        End If
        
        'Stores the result
        EndString = EndString & CurrChar
        
    Next count
    
    'Returns the resulting string
    EncryptDecrypt = EndString
    
End Function
