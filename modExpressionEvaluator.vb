' Generic Infix Expression Evaluator
'
' Expressions are converted to postfix (credits: DOEACC)
' Stack and SymbolTable classes are used
'
' Copyright (c) Samuel Gomes (Blade), 2001-2003
' mailto: v_2samg@hotmail.com
'
' This module implements an Algebraic expression evaluator
' for Visual Basic. It support floating point numbers,
' most standard operators, plus or minus unary operators
' and parantheses to override default precedence rules.
' Rudimentary support for user-defined symbols is also
' included.
'
' I do not accept responsibility for any effects,
' adverse or otherwise, that this code may have on you,
' your computer, your sanity, your dog, and anything else
' that you can think of. Use it at your own risk.

Module modExpressionEvaluator

	Private Const STATE_NONE As Short = 0
	Private Const STATE_OPERAND As Short = 1
	Private Const STATE_OPERATOR As Short = 2
	Private Const STATE_UNARYOP As Short = 3

	Private Const UNARY_NEG As String = "(-)"

	Private m_sErrMsg As String

	'Expose symbol table object
	Public EvaluatorSymbolTable As New SymbolTable()

	'Evaluates the expression and returns the result.
	Public Function Evaluate(ByVal sExpression As String) As Double
		Dim sBuffer As String = ""
		Dim nErrPosition As Integer

		'Convert to postfix expression
		nErrPosition = InfixToPostfix(sExpression, sBuffer)
		'Raise trappable error if error in expression
		If nErrPosition <> 0 Then
			Err.Raise(vbObjectError + 1001, , m_sErrMsg & ": Column " & CStr(nErrPosition))
		End If
		'Evaluate postfix expression
		Evaluate = DoEvaluate(sBuffer)
	End Function

	'Converts an infix expression to a postfix expression
	'that contains exactly one space following each token.
	Private Function InfixToPostfix(ByVal sExpression As String, ByRef sBuffer As String) As Integer
		Dim i As Integer
		Dim ch, sTemp As String
		Dim nCurrState As Short
		Dim nParenCount As Integer
		Dim bDecPoint As Boolean
		Dim stkTokens As New Stack()

		nCurrState = STATE_NONE
		nParenCount = 0
		i = 1
		Do Until i > Len(sExpression)
			'Get next character in expression
			ch = Mid(sExpression, i, 1)
			'Respond to character type
			Select Case ch
				Case "("
					'Cannot follow operand
					If nCurrState = STATE_OPERAND Then
						m_sErrMsg = "Operator expected"
						GoTo EvalError
					End If
					'Allow additional unary operators after "("
					If nCurrState = STATE_UNARYOP Then
						nCurrState = STATE_OPERATOR
					End If
					'Push opening parenthesis onto stack
					stkTokens.Push(ch)
					'Keep count of parentheses on stack
					nParenCount += 1
				Case ")"
					'Must follow operand
					If nCurrState <> STATE_OPERAND Then
						m_sErrMsg = "Operand expected"
						GoTo EvalError
					End If
					'Must have matching open parenthesis
					If nParenCount = 0 Then
						m_sErrMsg = "Closing parenthesis without matching open parenthesis"
						GoTo EvalError
					End If
					'Pop all operators until matching "(" found
					sTemp = CStr(stkTokens.Pop)
					Do Until sTemp = "("
						sBuffer = sBuffer & sTemp & " "
						sTemp = CStr(stkTokens.Pop)
					Loop
					'Keep count of parentheses on stack
					nParenCount -= 1
				Case "+", "-", "*", "/", "\", "%", "^"
					'Need a bit of extra code to handle unary operators
					If nCurrState = STATE_OPERAND Then
						'Pop operators with precedence >= operator in ch
						Do While stkTokens.Count > 0
							If GetPrecedence(CStr(stkTokens.Peek())) < GetPrecedence(ch) Then
								Exit Do
							End If
							sBuffer = sBuffer & CStr(stkTokens.Pop) & " "
						Loop
						'Push new operand
						stkTokens.Push(ch)
						nCurrState = STATE_OPERATOR
					ElseIf nCurrState = STATE_UNARYOP Then
						'Don't allow two unary operators in a row
						m_sErrMsg = "Operand expected"
						GoTo EvalError
					Else
						'Test for unary operator
						If ch = "-" Then
							'Push unary minus
							stkTokens.Push(UNARY_NEG)
							nCurrState = STATE_UNARYOP
						ElseIf ch = "+" Then
							'Simply ignore positive unary operator
							nCurrState = STATE_UNARYOP
						Else
							m_sErrMsg = "Operand expected"
							GoTo EvalError
						End If
					End If
				Case "0" To "9", "."
					'Cannot follow other operand
					If nCurrState = STATE_OPERAND Then
						m_sErrMsg = "Operator expected"
						GoTo EvalError
					End If
					sTemp = ""
					bDecPoint = False
					Do While InStr("0123456789.", ch) <> 0
						If ch = "." Then
							If bDecPoint Then
								m_sErrMsg = "Operand contains multiple decimal points"
								GoTo EvalError
							Else
								bDecPoint = True
							End If
						End If
						sTemp &= ch
						i += 1
						If i > Len(sExpression) Then Exit Do
						ch = Mid(sExpression, i, 1)
					Loop
					'i will be incremented at end of loop
					i -= 1
					'Error if number contains decimal point only
					If sTemp = "." Then
						m_sErrMsg = "Invalid operand"
						GoTo EvalError
					End If
					sBuffer = sBuffer & sTemp & " "
					nCurrState = STATE_OPERAND
				Case Is <= " "              'Ignore spaces, tabs, etc.
				Case Else
					'Symbol name cannot follow other operand
					If nCurrState = STATE_OPERAND Then
						m_sErrMsg = "Operator expected"
						GoTo EvalError
					End If
					If IsSymbolCharFirst(ch) Then
						sTemp = ch
						i += 1
						If i <= Len(sExpression) Then
							ch = Mid(sExpression, i, 1)
							Do While IsSymbolChar(ch)
								sTemp &= ch
								i += 1
								If i > Len(sExpression) Then Exit Do
								ch = Mid(sExpression, i, 1)
							Loop
						End If
					Else
						'Unexpected character
						m_sErrMsg = "Unexpected character encountered"
						GoTo EvalError
					End If
					'See if symbol is defined
					If EvaluatorSymbolTable.IsDefined(sTemp) Then
						sBuffer = sBuffer & CStr(EvaluatorSymbolTable.Value(sTemp)) & " "
						nCurrState = STATE_OPERAND
						'i will be incremented at end of loop
						i -= 1
					Else
						m_sErrMsg = "Undefined symbol: '" & sTemp & "'"
						'Reset error position to start of symbol
						i -= Len(sTemp)
						GoTo EvalError
					End If
			End Select
			i += 1
		Loop
		'Expression cannot end with operator
		If nCurrState = STATE_OPERATOR Or nCurrState = STATE_UNARYOP Then
			m_sErrMsg = "Operand expected"
			GoTo EvalError
		End If
		'Check for balanced parentheses
		If nParenCount > 0 Then
			m_sErrMsg = "Closing parenthesis expected"
			GoTo EvalError
		End If
		'Retrieve remaining operators from stack
		Do Until stkTokens.Count = 0
			sBuffer = sBuffer & CStr(stkTokens.Pop()) & " "
		Loop
		'Indicate no error
		InfixToPostfix = 0
		Exit Function
EvalError:
		'Report error postion
		InfixToPostfix = i
		Exit Function
	End Function

	'Returns a number that indicates the relative precedence of an operator.
	Private Function GetPrecedence(ByVal ch As String) As Short
		Select Case ch
			Case "+", "-"
				GetPrecedence = 1
			Case "*", "/", "\", "%"
				GetPrecedence = 2
			Case "^"
				GetPrecedence = 3
			Case UNARY_NEG
				GetPrecedence = 10
			Case Else
				GetPrecedence = 0
		End Select
	End Function

	'Evaluates the given expression and returns the result.
	'It is assumed that the expression has been converted to
	'a postix expression and that a space follows each token.
	Private Function DoEvaluate(ByVal sExpression As String) As Double
		Dim i, j As Integer
		Dim stkTokens As New Stack()
		Dim sTemp As String
		Dim Op1, Op2 As Object

		'Locate first token
		i = 1
		j = InStr(sExpression, " ")
		Do Until j = 0
			'Extract token from expression
			sTemp = Mid(sExpression, i, j - i)
			If IsNumeric(sTemp) Then
				'If operand, push onto stack
				stkTokens.Push(CDbl(sTemp))
			Else
				'If operator, perform calculations
				Select Case sTemp
					Case "+"
						stkTokens.Push(CDbl(stkTokens.Pop) + CDbl(stkTokens.Pop))
					Case "-"
						Op1 = stkTokens.Pop
						Op2 = stkTokens.Pop
						stkTokens.Push(CDbl(Op2) - CDbl(Op1))
					Case "*"
						stkTokens.Push(CDbl(stkTokens.Pop) * CDbl(stkTokens.Pop))
					Case "/"
						Op1 = stkTokens.Pop
						Op2 = stkTokens.Pop
						stkTokens.Push(CDbl(Op2) / CDbl(Op1))
					Case "\"
						Op1 = stkTokens.Pop
						Op2 = stkTokens.Pop
						stkTokens.Push(CLng(CDec(Op2) / CDec(Op1)))
					Case "%"
						Op1 = stkTokens.Pop
						Op2 = stkTokens.Pop
						stkTokens.Push(CDec(Op2) Mod CDec(Op1))
					Case "^"
						Op1 = stkTokens.Pop
						Op2 = stkTokens.Pop
						stkTokens.Push(CDbl(Op2) ^ CDbl(Op1))
					Case UNARY_NEG
						stkTokens.Push(-CDbl(stkTokens.Pop))
					Case Else
						'This should never happen (bad tokens caught in InfixToPostfix)
						Err.Raise(vbObjectError + 1002, , "Bad token in Evaluate: " & sTemp)
				End Select
			End If
			'Find next token
			i = j + 1
			j = InStr(i, sExpression, " ")
		Loop
		'Remaining item on stack contains result
		If stkTokens.Count > 0 Then
			DoEvaluate = CDbl(stkTokens.Pop())
		Else
			'Null expression; return 0
			DoEvaluate = 0
		End If
	End Function

	'Returns a boolean value that indicates if sChar is a valid
	'character to be used as the first character in symbols names
	Private Function IsSymbolCharFirst(ByVal sChar As String) As Boolean
		Dim c As String

		c = UCase(Left(sChar, 1))
		IsSymbolCharFirst = (c >= "A" And c <= "Z") Or (InStr("_", c) <> 0)
	End Function

	'Returns a boolean value that indicates if sChar is a valid
	'character to be used in symbols names
	Private Function IsSymbolChar(ByVal sChar As String) As Boolean
		Dim c As String

		c = UCase(Left(sChar, 1))
		IsSymbolChar = (c >= "A" And c <= "Z") Or (InStr("0123456789_", c) <> 0)
	End Function
End Module
