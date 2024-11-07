Sub Area(x As Double, y As Double)
   	MsgBox x * y
End Sub


Sub Main() 
    MultiBeep 56 
    Message 
End Sub 

 Sub MultiBeep(numbeeps) 
    For counter = 1 To numbeeps 
        Beep 
    Next counter 
End Sub 
 
Sub Message() 
    MsgBox "Time to take a break!" 
End Sub

 Function findArea(Length As Double, Optional Width As Variant)
   If IsMissing(Width) Then
      findArea = Length * Length
   Else
      findArea = Length * Width
   End If
End Function

Function findArea(Length As Double, Width As Variant)
   area Length, Width ' To Calculate Area 'area' sub proc is called
End Function

Sub Area(x As Double, y As Double)
   MsgBox x * y
End Sub


Private Sub Welcome_Click()
    MsgBox "Welcome"
    wbkname
End Sub


Function wbkname() As String
    myWbk = Application.ThisWorkbook.FullName
    MsgBox myWbk
End Function

Function MessageBox_Demo() 
   'Message Box with just prompt message 
   MsgBox("Welcome")     
   
   'Message Box with title, yes no and cancel Buttons  
   int a = MsgBox("Do you like blue color?",3,"Choose options") 
   ' Assume that you press No Button  
   msgbox ("The Value of a is " & a) 
End Function

Function findArea() 
   Dim Length As Double 
   Dim Width As Double 
   
   Length = InputBox("Enter Length ", "Enter a Number") 
   Width = InputBox("Enter Width", "Enter a Number") 
   findArea = Length * Width 
     End Function



Sub say_helloworld_Click()
   Dim password As String
   password = "Admin#1"

   Dim num As Integer
   num = 1234

   Dim BirthDay As Date
   BirthDay = DateValue("30 / 10 / 2020")

   MsgBox ("Password is " & password & Chr(10) & "Value of num is " & num & Chr(10) & "Value of Birthday is " & BirthDay)


End Sub


Sub if_test()
    Dim x As Integer
    Dim y As Integer
    x = 40
    y = 20
    If x > y Then
        MsgBox "x is greater than y"
    ElseIf y > x Then
        MsgBox "y is greater than x"
    Else
        MsgBox "x and y are equal"
    End If
End Sub

Private Sub nested_If_demo_Click()
Dim x As Integer
x = 30
If x > 0 Then
    MsgBox "a number is a positive number"
    If x = 1 Then
        MsgBox "A number is neither prime nor composite"
    ElseIf x = 2 Then
        MsgBox "A number is the only prime even prime number"
    ElseIf x = 3 Then
        MsgBox "A number is the least odd prime number"
    Else
        MsgBox "The number is not 0, 1, 2, or 3"
    End If
ElseIf x < 0 Then
    MsgBox "A number is a negative number"
Else
    MsgBox "the number is zero"
End If
End Sub

Private Sub switch_demo_Click ()  
Dim MyVar As Integer  
MyVar = 1  
Select Case MyVar  
Case 1  
MsgBox "A number is the least composite number"  
Case 2  
MsgBox "A number is the only even prime number"  
Case 3  
MsgBox "A number is the least odd prime number"  
Case Else   
MsgBox "unknown number"  
End Select  
End Sub   


Function GetGrade(StudentMarks As Integer)
    Dim FinalGrade As String
  
    Select Case StudentMarks
  
    Case Is < 33
        FinalGrade = "F"
  
    Case 33 To 50
        FinalGrade = "E"
  
    Case 51 To 60
        FinalGrade = "D"
  
    Case 61 To 70
        FinalGrade = "C"
      
    Case 71 To 90
        FinalGrade = "B"
  
    Case Else
        FinalGrade = "A"
  
    End Select
    GetGrade = FinalGrade
  
End Function

Sub forloop_demo()
    Dim a As Integer
    a = 5
    
    For i = 0 To a Step 1
        MsgBox "The value of i is : " & i
    Next
End Sub


Sub whileloop_demo()
    Dim a As Integer
    a = 0
    
    Do While a < 5
        a = a + 1
        MsgBox "The value of a is : " & a
    Loop
End Sub

Function Discount(Quantity, Price)

If Quantity > 25 Then
    Discount = Quantity * Price * 0.2
Else: Discount = 0
End If
End Function

Sub ClearContent()
    Answer = MsgBox("Confirm you want to clear?", vbYesNo)
    
    If Answer = vbYes Then
        Rows("6:" & Rows.Count).ClearContents
    Exit Sub
    End If
End Sub

Sub SendEmail()
    Dim OutApp As Object
    Dim OutMail As Object
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    With OutMail
    .To = "hello@gmail.com"
    .Subject = "Excel File"
    .Body = "This is a test email"
    .Attachments.Add ThisWorkbook.FullName
    .Display
    End With
    
    Set OutApp = Nothing
    Set OutMail = Nothing
    
End Sub

