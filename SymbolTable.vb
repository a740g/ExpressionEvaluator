' Generic Symbol Table class using Collections
' Copyright (c) Samuel Gomes (Blade), 2001-2020
' mailto: v_2samg@hotmail.com

Friend Class SymbolTable

	' Private collection to hold symbol values
	Private ReadOnly m_Symbols As New Collection()

	'Add a symbol to the symbol table
	Public Sub Add(ByVal sName As String, Optional ByVal nValue As Object = 0)
		If Not IsDefined(sName) Then
			m_Symbols.Add(nValue, sName)
		Else
			Err.Raise(vbObjectError + 1021, , "Symbol already defined")
		End If
	End Sub

	'Delete the specified symbol from the symbol table
	'If the symbol is not defined, the call is ignored
	Public Sub Delete(ByVal sName As String)
		If IsDefined(sName) Then
			m_Symbols.Remove(sName)
		Else
			Err.Raise(vbObjectError + 1022, , "Symbol not defined")
		End If
	End Sub

	'Indicates if the specified symbol name is currently defined
	Public Function IsDefined(ByVal sName As String) As Boolean
		Dim nValue As Object

		On Error Resume Next
		nValue = m_Symbols.Item(sName)

		Return Err.Number = 0
	End Function

	'Sets the value of the specified symbol
	'Raises a run-time error if the symbol is not currently defined or
	'if the value is non-numeric
	'Returns the value for the specified symbol
	'Raises a run-time error if the symbol is not currently defined
	Public Property Value(ByVal sName As String) As Object
		Get
			If IsDefined(sName) Then
				Value = m_Symbols.Item(sName)
			Else
				Err.Raise(vbObjectError + 1022, , "Symbol not defined")
				Value = 0
			End If
		End Get
		Set(ByVal Value As Object)
			If IsDefined(sName) Then
				Delete(sName)
				Add(sName, Value)
			Else
				Err.Raise(vbObjectError + 1022, , "Symbol not defined")
			End If
		End Set
	End Property

	'Returns the number of symbols in table
	Public ReadOnly Property Count() As Integer
		Get
			Return m_Symbols.Count()
		End Get
	End Property
End Class