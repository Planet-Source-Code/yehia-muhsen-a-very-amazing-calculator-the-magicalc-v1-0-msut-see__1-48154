Attribute VB_Name = "Module1"
'Programmed By : Yehia Muhsen
'Date          : 8-13-2003
'Copyrights    : Copyright © 2003 Yehia Muhsen, All rights reserved.
'Description   : This program is one of my best programs. Basically it's a calculator,
'                but it's not like any calculator. I called it the MagiCalc 1.0
'                because it's very powerful, and can accurately evaluate any mathematical
'                expression no matter how complicated it is as long as it meets conditions.

'                The operations allowed are *, / , + , - , and ^. You can use parenthesis too.
'                For example:+(2((5+4)/2*6+1.56845)4 + 3^(3-(--5)/2))/(25^.5)+2-3*4/-2^2.2*5

'                In addition to evaluating expressions, this program is supported by a very
'                powerful errors handling. It deals with almost all kind of errors,
'                and it tells you exaclty what and where the error is so that you can
'                fix it. One of the good things about this program is that it uses the
'                memory so effectively. The following expression has almost all kind of
'                errors that this program handles, use that expression to see how Syntax
'                errors get handled. ()+*5.334.2+)453+&76/6(7+

'                A good feature of this calcualtor is that it recolor the expression after the
'                evaluation (Operations in purple, numbers in light blue,invalid inputs in red ,
'                positive and negative signs in dark green, and parenthesis in randome colors).
'                That make the expresssion very clear .

'                Moreover, it's so easy to use. To copy the result to the clipboard, just click
'                on the result. You also can review the old expressoin you entered by using
'                Up and Down keys.

'                If you would like to use the mouse for input, then you can expand the form to
'                view an input keyboard by clicking the dark bar at the bottom. To hide the
'                input keyboard, click on the dark bar again.

'                Also the code has a lot of comments, and I hope it will be easy for you to
'                understand, however if you have any quesions, then please feel free to email
'                me at yehia_sm@hotmail.com
'
'Special Thanks: For Issam Hijazi for inspiring me with the idea.
Option Explicit

Public Function EvaluateExp(ByVal Expression As String, ByRef Result As Double, ByVal OffSet As Integer, ByVal ObjSrc As Object, ByVal ObjDes As Object) As Boolean
On Error GoTo Err
'Variables
Dim Numbers() As Double
Dim Operations() As String
Dim Exp As String, SubExp As String, ExpLen As Integer
Dim TempChar As String, TempChunk As String
Dim I As Integer, N As Integer
Dim X As Byte, sResult As Double
Dim OpenPara As Byte, ClosePara As Byte
Dim NegNum As Boolean, DecPoint As Boolean
Dim Parenthesis As Boolean, DoParenthesis As Boolean
Dim TolerateSigns As Boolean
Dim RndColor As Long

'Get the expression
Exp = Trim(Expression)
ExpLen = Len(Exp)

'Initial Vlues
EvaluateExp = False
I = 0
N = 0
TempChunk = ""
Result = 0
NegNum = False
Parenthesis = False
DecPoint = False
'If there is no sign before or after parenthesis assume the operation is *
TolerateSigns = True

'Check if there is an expressoin or not
If ExpLen = 0 Then
    MsgBox "There is no expression", vbExclamation, "Syntax Error"
    ObjDes = "Error: There is no expression ."
    'Trace Error
    TraceError ObjSrc, OffSet + 1, False
    Exit Function
End If

'Get numbers and operations

For I = 1 To ExpLen

    'Read one character
    TempChar = Mid(Exp, I, 1)
    
    'Check the character
    
    Select Case TempChar
    
        'If it's a number
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "."
            '
            If TempChar = "." Then
                'Check If there more than two decimal points
                If DecPoint Then
                    'Show error message if there is no sign tolerance
                    MsgBox "Extra decimal point .", vbExclamation, "Syntax Error"
                    ObjDes = "Error: Extra decimal point ..."
                    'Trace Error
                    TraceError ObjSrc, OffSet + I, True
                    GoTo Ex
                End If
                DecPoint = True
            End If
            TempChunk = TempChunk + TempChar
            
            'If there is no sign after parenthesis, assume the operatioin is *
            If Parenthesis Then
                If TolerateSigns Then
                    TempChar = "*"
                    GoTo Add_Num_And_Op
                Else
                    'Show error message if there is no sign tolerance
                    MsgBox "There must be an operation after the paranthesis.", vbExclamation, "Syntax Error"
                    ObjDes = "Error: Missing operation ..."
                    'Trace Error
                    TraceError ObjSrc, OffSet + I
                    GoTo Ex
                End If
            End If
    
        
        'If Parenthesis found
        Case "("
            'Check if there is an operation before the Parenthesis
            If I > 1 Then
                If Not IsOperation(Mid(Exp, I - 1, 1)) Then
                    If TolerateSigns Then
                        'If there is no sign before parenthesis, assume the operatioin is *
                        TempChar = "*"
                        'In order to come back and do the parenthesis
                        DoParenthesis = True
                        GoTo Add_Num_And_Op
                    Else
                        'Show error message if there is no sign tolerance
                        MsgBox "There must be an operation before the paranthesis.", vbExclamation, "Syntax Error"
                        ObjDes = "Error: Missing operation ."
                        'Trace Error
                        TraceError ObjSrc, OffSet + I
                        GoTo Ex
                    End If
    
                End If
            End If
            
Do_Parenthesis:
            'Try to find the closed Parenthesis ")"
            OpenPara = 1
            ClosePara = 0
            For X = I + 1 To ExpLen
                'Make sure you're looking for the right closed paranthesis
                If Mid(Exp, X, 1) = "(" Then OpenPara = OpenPara + 1
                If Mid(Exp, X, 1) = ")" Then ClosePara = ClosePara + 1
                If Mid(Exp, X, 1) = ")" And ClosePara = OpenPara Then
                    'Evaluate the expression between Parenthesis
                    SubExp = Mid(Exp, I + 1, X - 1 - I)
                    If Not EvaluateExp(SubExp, sResult, I + OffSet, ObjSrc, ObjDes) Then Exit Function
                    
                    'Set Temp Chunk to "Parenthesis" value, so when reaching
                    'or reaching the end of the expression, insert the sResult
                    'as a number in the Numbers(N) array
                    
                    Parenthesis = True
                    Exit For
                End If
            Next X
            'If the Parenthesis is not closed
            If Not Parenthesis Then
                MsgBox "Missing parenthesis ')' .", vbExclamation, "Syntax Error"
                ObjDes = "Error: Missing ) ."
                'Show the parenthesis in red color
                SetColor ObjSrc, OffSet + I, vbRed, True
                'Trace Error
                TraceError ObjSrc, OffSet + I, True
                GoTo Ex:
            End If
            
            I = X
            
            'Parenthesis are done
            DoParenthesis = False
        
        'If the closed parenthesis found without open parenthesis
        Case ")"
                MsgBox "Missing parenthesis '(' .", vbExclamation, "Syntax Error"
                ObjDes = "Error: Missing ( ."
                'Show the parenthesis in red color
                SetColor ObjSrc, OffSet + I, vbRed, True
                'Trace Error
                TraceError ObjSrc, OffSet + I, True
                GoTo Ex:
        
        'If there is an operation
        Case "*", "/", "+", "-", "^"
            'Check if there is a number before the operation
            If TempChunk = "" And Not Parenthesis Then
                'If the operation is "-" or "+" then consider it a sign (Negative or Positve)
                If TempChar = "-" Or TempChar = "+" Then
                    'Add the negative sign for the next number
                    If TempChar = "-" Then NegNum = Not NegNum
                    'Show the sign's color
                    SetColor ObjSrc, OffSet + I, &HB000&, True
                    GoTo Nxt
                End If
                
                'Show error message
                MsgBox "There is no number before the operation ( " & TempChar & " ).", vbExclamation, "Syntax Error"
                ObjDes = "Error: No number before the operation ..."
                'Trace Error
                TraceError ObjSrc, OffSet + I
                GoTo Ex:
            End If
            
            'If the operation is at the end of the expression
            If I = ExpLen Then
                MsgBox "You have to put a number after the operatoin( " & TempChar & " ).", vbExclamation, "Syntax Error"
                ObjDes = "Error: Operation at the end ..."
                'Trace Error
                TraceError ObjSrc, OffSet + I + 1
                GoTo Ex:
            End If
            
            'Show operaton in dark red color
            SetColor ObjSrc, OffSet + I, &HC000C0, True
            
Add_Num_And_Op:
            'Assign numbers and operations to their arrays
            N = N + 1
            'For memory usage effeciency
            'Reallocate storage space for the two arrays
            ReDim Preserve Numbers(N)
            ReDim Preserve Operations(N)
            
            'This part is important so that power gets evaluated before negative sign
            'Witbout this part  -2^2=-4 gets evaluated as (-2)^2=4 [ Like in Excel ]
            If NegNum And TempChar = "^" Then
                Numbers(N) = "-1"
                Operations(N) = "*"
                NegNum = False
                If N > 1 Then
                    Select Case Operations(N - 1)
                        Case "/"
                           Operations(N) = "/"
                        Case "^"
                            Operations(N) = "^"
                    End Select
                End If
                GoTo Add_Num_And_Op
            End If
            
            'Add the numbers
            If Parenthesis Then
                'Value from the Parenthesis
                Numbers(N) = sResult
                If NegNum Then Numbers(N) = -Numbers(N)
                'Show the parenthesis in random color
                Randomize
                RndColor = RGB(Rnd * 150 + 50, Rnd * 150 + 50, Rnd * 150 + 50)
                SetColor ObjSrc, OffSet + I - Len(SubExp) - 2, RndColor, True
                SetColor ObjSrc, OffSet + I - 1, RndColor, True
                Parenthesis = False
            Else
                'Value from the expression
                Numbers(N) = Val(TempChunk)
                If NegNum Then Numbers(N) = -Numbers(N)
                'Clear TempChunk variable
                TempChunk = ""
                DecPoint = False
            End If
            'Assme the next number is positive
            NegNum = False
            'Save operatoin
            Operations(N) = TempChar
            
            'Do parenthesis
            If DoParenthesis Then GoTo Do_Parenthesis
        
        'If unkown character
        Case Else
            MsgBox "You have entered an invalid input ( " & TempChar & " ).", vbExclamation, "Syntax Error"
            ObjDes = "Error: An invlid input ..."
            'Show the invalid character in red color
            SetColor ObjSrc, OffSet + I, vbRed, True
            'Trace Error
            TraceError ObjSrc, OffSet + I, True
            GoTo Ex
    End Select
Nxt:
Next I

N = N + 1
'For memories usage effeciency
'Reallocate storage space for the array
ReDim Preserve Numbers(N)
'
If Parenthesis Then
    'Value from the Parenthesis
    Numbers(N) = sResult
    If NegNum Then Numbers(N) = -Numbers(N)
    'Show the parenthesis in random color
    Randomize
    RndColor = RGB(Rnd * 150 + 50, Rnd * 150 + 50, Rnd * 150 + 50)
    SetColor ObjSrc, OffSet + I - Len(SubExp) - 2, RndColor, True
    SetColor ObjSrc, OffSet + I - 1, RndColor, True
ElseIf TempChunk <> "" Then
    'Value from the expression
    Numbers(N) = Val(TempChunk)
    If NegNum Then Numbers(N) = -Numbers(N)
ElseIf NegNum Then
    MsgBox "Negative sign without a number.", vbExclamation, "Syntax Error"
    ObjDes = "Error: Negative sign without a number..."
    'Trace Error
    TraceError ObjSrc, OffSet + I - 1, True
    GoTo Ex:
Else
    MsgBox "Unknown Error.", vbExclamation, "Syntax Error"
    ObjDes = "Error: Unknown Error..."
    'Trace Error
    TraceError ObjSrc, OffSet + I - 1, True
    GoTo Ex:
End If

'Evaluate the expression

N = N - 1

'First do the Power from left to right and move result to the right so result can be used
'with multiplicion and division from the right
For I = 1 To N
    If Operations(I) = "^" Then Numbers(I + 1) = Numbers(I) ^ Numbers(I + 1)
Next I

'When there is a power, then move result to the right so it can be used
'with multiplicion and division
For I = N To 1 Step -1
    If Operations(I) = "^" Then Numbers(I) = Numbers(I + 1)
Next I


'Second do the Multiplicatoin and Division from left to right
'Always put the answer to the right so it can be used in other operations
For I = 1 To N
    Select Case Operations(I)
        Case "^"
            'Move result to the left so result can be used with multiplicion
            'or division from the left of the power
            Numbers(I + 1) = Numbers(I)
        Case "*"
            Numbers(I + 1) = Numbers(I) * Numbers(I + 1)
        Case "/"
            Numbers(I + 1) = Numbers(I) / Numbers(I + 1)
    End Select
Next I

'Move results to the right so it can be used
'with addition and subtraction
For I = N To 1 Step -1
    Select Case Operations(I)
        Case "^"
            Numbers(I) = Numbers(I + 1)
        Case "*"
            Numbers(I) = Numbers(I + 1)
        Case "/"
            Numbers(I) = Numbers(I + 1)
    End Select
Next I

'Set the result equal the last number to the right
Result = Numbers(1)

'Finally do the addition and subtraction form right to left

For I = 1 To N
    Select Case Operations(I)
        Case "+"
            Result = Result + Numbers(I + 1)
        Case "-"
            Result = Result - Numbers(I + 1)
    End Select
Next I

'Finished successfully
EvaluateExp = True

'Exit Function
Ex:
    
Exit Function
'
Err:
    MsgBox Err.Description & String(20, " "), vbCritical, "Error!"
    ObjDes = "Error: " & Err.Description
End Function

Private Function IsOperation(ByVal Exp As String) As Boolean
Dim X As Byte
Dim M_Operation(1 To 5) As String

'Fill the mathematical operations array
M_Operation(1) = "*"
M_Operation(2) = "/"
M_Operation(3) = "+"
M_Operation(4) = "-"
M_Operation(5) = "^"

'If the expression is an operation then return true
IsOperation = False
For X = 1 To 5
    If Exp = M_Operation(X) Then
        IsOperation = True
        Exit Function
    End If
Next X
End Function

Private Sub TraceError(ByVal Obj As Object, ByVal Position As Integer, Optional ByVal Sel As Boolean = False)
'
Obj.SelStart = Position - 1
'Select the character or not
If Sel Then Obj.SelLength = 1 Else Obj.SelLength = 0
Obj.SetFocus
End Sub

Private Sub SetColor(ByVal Obj As Object, ByVal Position As Integer, ByVal Color As Long, Optional ByVal Bold As Boolean = False)
'Select the character andn apply the effect on it
Obj.SelStart = Position - 1
Obj.SelLength = 1
Obj.SelColor = Color
Obj.SelBold = Bold
Obj.SelLength = 0
End Sub
