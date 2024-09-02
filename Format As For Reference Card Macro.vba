' ---Format As For Reference Card Macro - Stable Edition - v1.0.2---
' Updated on 2024-09-01.
' https://github.com/KSXia/Verbatim-Format-As-For-Reference-Card-Macro---Stable-Edition
' Thanks to Truf for creating and providing his "ForReference" macro, which this macro is partly based upon! You can find Truf's macros on his website at https://debate-decoded.ghost.io/leveling-up-verbatim/
Sub FormatAsForReferenceCard()
	' Check if any text is selected.
	If Selection.Type = wdSelectionIP Then
		MsgBox "You have not selected any text." & vbNewLine & "Please select the text you want" & vbNewLine & "to format as a ""For Reference"" card.", Title:="Error in Formatting as" & vbNewLine & "a ""For Reference"" Card"
        Exit Sub
	End If
	
	Dim SelectionRange As Range
	Set SelectionRange = Selection.Range
	
	If Len(SelectionRange.Text) > 1 Then
		' Record the user's default highlight color.
		Dim UserDefaultHighlightColor As Long
		UserDefaultHighlightColor = Options.DefaultHighlightColorIndex
		
		' Set the default highlight color to the "For Reference" highlight color.
		Options.DefaultHighlightColorIndex = wdGray25
		
		' Find all highlighted characters and replace their highlight color with the default highlight color, which should be set to the "For Reference" highlight color.
		With SelectionRange.Find
			' Specify find criteria.
			.ClearFormatting
			.MatchWildcards = True
			.Text = "*"
			.Highlight = True
			
			' Ensure other find options are set to their defaults.
			.MatchCase = False
			.MatchWholeWord = False
			.MatchSoundsLike = False
			.MatchAllWordForms = False
			.MatchPrefix = False
			.MatchSuffix = False
			.MatchPhrase = False
			
			' Specify replacement criteria.
			.Replacement.ClearFormatting
			.Replacement.Text = ""
			.Replacement.Highlight = True
			
			' Set execution properties.
			.Format = True
			.Forward = True
			.Wrap = wdFindStop
			.Execute Replace:=wdReplaceAll
		End With
		
		' Reset the default highlight color back to the user's default highlight color.
		Options.DefaultHighlightColorIndex = UserDefaultHighlightColor
	ElseIf Len(SelectionRange.Text) = 1 Then
		If SelectionRange.HighlightColorIndex <> wdNoHighlight Then
			SelectionRange.HighlightColorIndex = wdGray25
		End If
	End If
End Sub