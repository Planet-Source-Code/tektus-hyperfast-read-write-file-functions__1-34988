<div align="center">

## HyperFast\! Read/Write File Functions


</div>

### Description

These two functions are designed to read and write a file as fast as possible in VB. It is faster for some cases than WinApi Read/WriteFile functions because you don't have to convert binary to string in a loop. Thats why, it is very fast and useful for any purpose. I have created it for my Encryption programme and with Windows Crypto functions and this two functions, my encrytion programme works faster than most of the Encrytion software you can download on the net. Hope it will be useful for you. Thanks.

Ozan Yasin Dogan

www.uni-group.org (will be online in 01/06/02)
 
### More Info
 
For ReadFile Function: File Name

For WriteFile Function: File Name and What to Write

Put it in a module.

ReadFile Function: Returns the full content of the file in string format not binary!

WriteFile Function: Returns nothing


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[tektus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tektus.md)
**Level**          |Intermediate
**User Rating**    |4.1 (33 globes from 8 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tektus-hyperfast-read-write-file-functions__1-34988/archive/master.zip)





### Source Code

```
Option Explicit
'------------------- CREDIT --------------------------
'This functions are written by Ozan Yasin Dogan,
'a Turkish student in Istanbul.
'Everybody can copy and use this code without changing
'this credit part. www.uni-group.org
'------------------- CREDIT --------------------------
'HyperFast Read / Write file functions
'How to use:
'Text1.Text = ReadFile("c:\autoexec.bat")
'Notice that textboxes doesn't show after the null characters
'so don't panic, you can check if it is read by using this:
'Text1.Text = Len(ReadFile("c:\autoexec.bat")
'To test both Read and Write file functions, you
'can simply use:
'Call WriteFile("c:\test.bat", ReadFile("c:\autoexec.bat")
'Thank you and please don't forget to vote for me on
'planet source code: www.pscode.com/vb
'This is the buffer lenght, you can change to maximum 32767
'The ReadFile function add to Content variable in the memory
'30000 bytes in each loop
Const Buf As Integer = 30000
'Declarations
Dim FileLen As Long 'To keep file lenght information
Dim Multiply As Long 'It is required to find how many Buf
'bytes exist in the file. For ex: in a 125,000bytes file
'there are 4 multiply. The rest is recorded to Plus variable
Dim Temp As String * Buf 'Temporary string block
'It is necessary for use of Random Access methode.
'If not, you had to open it in Binary mode and convert
'binary data to text, and it is also a loop and slows
'down the process. This is the best methode i think..
Dim Content As String 'Content is the file content,
'the function allocates a space for it first and
'full it with Mid function. It is a very fast methode
'instead of using Content = Content & Something
Dim Plus As Long 'The plus part of the file after dividing
'to Buf variable. It is used when the file lenght is small
'than Buf and to find the rest of the bytes after dividing
'file lenght to Buf
Dim Point As Long 'Point shows on which byte the content is.
Dim FileNo As Byte 'To find a free file number
Dim Counter As Long 'Is required for loops
Public Function ReadFile(FileName As String) As String 'Returns STRING variable!
FileNo = FreeFile 'Find a free file number
Open FileName For Random As #FileNo Len = Buf 'Open the file as Random, each record will have the lenght of Buf
FileLen = LOF(FileNo) 'File lenght
Multiply = Int(FileLen \ Buf) 'How many loops required to read the file
Content = Space(FileLen) 'Allocate a space for file content in the memory
Plus = FileLen - (Multiply * Buf) 'After this loops, there might be also some bytes to read
Point = 1 'Content is in this byte: 1
  If Multiply = 0 Then 'If the file is smaller than Buf (30000 bytes here, you can change it)
    Plus = FileLen: Counter = 1: GoTo Jump1
  End If
  'This loop reads the file as it was defined in a Type,
  'using random access methode and adds each records
  'to the content using Mid function.
  'Because Content = Content & Temp would slow down
  'the loop very much! And as you see, there is no transfer
  'beetween binary to string..
  For Counter = 1 To Multiply
    Get #FileNo, Counter, Temp
      Mid(Content, Point, Buf) = Temp
      Point = Point + Buf
  Next Counter
Jump1:
  'This is for the rest of the file after the loop.
  If Plus > 0 Then
    Get #FileNo, Counter, Temp
      Mid(Content, Point, Plus) = Left(Temp, Plus)
  End If
Close #FileNo
ReadFile = Content
End Function
Public Sub WriteFile(FileName As String, Content As String)
FileNo = FreeFile
Open FileName For Output As #FileNo
Print #FileNo, Content; '; is required for Vb to not write another 2 charachters of new line in the file
Close #FileNo
End Sub
```

