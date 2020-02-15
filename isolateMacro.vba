Dim y As Integer
Dim x As String
Dim K As String


Private Sub CommandButton1_Click()
Dim a As String
Dim b As String
Dim temp As String
Dim named As String
Dim filetype As String
Dim path1 As String
Dim file As String
path1 = Application.ThisWorkbook.Path & "\"
path2 = path1 + "test.bat"





file = Dir(path1 & "*.*")


        
Do While file <> "" ' And file <> ThisWorkbook.Name


'If file <> ThisWorkbook.Name Then
'GoTo lineend
'End If

c = i + 1
b = Str(c)
b = LTrim(b)
y = c
x = "G" + b

Range(x).Select
Selection.Value = file
    
            i = i + 1
            'Ò»¶¨Òª ²»È»ËÀÑ­»·
            file = Dir

            
Loop

Range("H2").Select
Selection.Value = y


End Sub

Private Sub CommandButton2_Click()
Dim i As Integer
Dim y_name As String
Dim x_name As String
path1 = Application.ThisWorkbook.Path & "\"
path2 = path1 + "test.bat"
a = "ren "
K = " "
RRRR = """"



Range("A2").Select
named = Selection.Value
named = RTrim(named)
Range("A4").Select
filetype = Selection.Value
Range("A6").Select
nameda = Selection.Value
nameda = RTrim(nameda)
Range("A8").Select
named3 = Selection.Value
named3 = RTrim(named3)

Range("H2").Select
y = Selection.Value


Range("AB1").Select
yy = Selection.Value



'MsgBox (RRRR)
   
Do While i < y

On Error Resume Next
c = i + 1
b = Str(c)
b = LTrim(b)




x = "G" + b
yyy = "J" + b
If c < 10 Then
b = "0" + b
End If

b = "[" + b + "]"

Range(x).Select
file = Selection.Value

Range(yyy).Select
titley = Selection.Value


x_name = path1 + file
If yy = 0 Then
y_name = path1 + named + b + nameda + "[" + titley + "]" + named3 + filetype
End If

If yy = 1 Then
y_name = path1 + named + "[" + titley + "]" + nameda + b + named3 + filetype
End If

'MsgBox (x_name)
Name x_name As y_name
      i = i + 1
       
'K = K + a + """" + file + """" + " "
'K = K + """" + named + b + filetype + """" + Chr(10) + Chr(10)

   Loop



'MsgBox (K)
'Open path2 For Output As #1
'Print #1, K
'Print #1, "pause"
'Print #1, "del /f %0"
'Close #1

MsgBox ("Éú³ÉÍê³É Çë¼ì²éÎÄ¼þ¼Ð")
End Sub




