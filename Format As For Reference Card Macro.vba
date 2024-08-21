' Format As For Reference Card Macro v1.0.0 - Stable Edition
' Created on 2024-08-17.
' https://github.com/KSXia/Verbatim-Format-As-For-Reference-Card-Macro---Stable-Edition
' Thanks to Truf for creating and providing his "ForReference" macro, which this macro is partly based upon! You can find Truf's macros on his website at https://debate-decoded.ghost.io/leveling-up-verbatim/
Sub FormatAsForReferenceCard()
	' Check if any text is selected
	If Selection.Type = wdSelectionIP Then
		MsgBox "You have not selected any text." & vbNewLine & "Please select the text you want" & vbNewLine & "to format as a ""For Reference"" card.", Title:="Error in Formatting as" & vbNewLine & "a ""For Reference"" Card"
        Exit Sub
	End If
	
	Dim SelectedText As Range
	Set SelectedText = Selection.Range
	
	Dim Character As Range
	
	' Loop through each character in the selected text
	For Each Character In SelectedText.Characters
		' Check if the character is highlighted
		If Character.HighlightColorIndex <> wdNoHighlight Then
			' If the character is highlighted, change the highlight color to light gray
			Character.HighlightColorIndex = wdGray25
		End If
	Next Character
End Sub