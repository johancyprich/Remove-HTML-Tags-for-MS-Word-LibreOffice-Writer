Sub RemoveHTMLTags()
	' Macro for Microsoft Word.
	' Author: https://excelwordvbamacros.blogspot.ca/2012/10/vba-string-manipulations-remove-html.html.

	Dim MyRange As Range
	Dim pos As Long

	Set MyRange = ActiveDocument.Range
		With MyRange.Find
			Do While .Execute(findText:="(\<*\>)", _
				MatchWildcards:=True, _
				Wrap:=wdFindStop, Forward:=True) = True
				MyRange.Delete
			Loop
		End With

	Set MyRange = Nothing
	MsgBox "end macro"
End Sub
