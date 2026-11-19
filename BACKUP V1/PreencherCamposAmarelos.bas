Attribute VB_Name = "PreencherCamposAmarelos"
Option Explicit

Public Sub PreencherCamposAmarelos()
    Dim replacements As Object
    Dim story As Range
    Dim currentStory As Range
    Dim searchRange As Range
    Dim matchRange As Range
    Dim storyStart As Long
    Dim placeholder As String
    Dim key As Variant
    Dim answer As Variant
    Dim replacedCount As Long

    Set replacements = CreateObject("Scripting.Dictionary")

    For Each story In ActiveDocument.StoryRanges
        Set currentStory = story

        Do While Not currentStory Is Nothing
            storyStart = currentStory.Start
            Set searchRange = currentStory.Duplicate

            With searchRange.Find
                .ClearFormatting
                .Text = vbNullString
                .Format = True
                .Highlight = True
                .Forward = False
                .Wrap = wdFindStop
            End With

            Do While searchRange.Find.Execute
                Set matchRange = searchRange.Duplicate

                If matchRange.HighlightColorIndex = wdYellow Then
                    placeholder = matchRange.Text
                    If Not replacements.Exists(placeholder) Then
                        replacements.Add placeholder, Empty
                    End If
                End If

                If matchRange.Start <= storyStart Then Exit Do
                searchRange.SetRange Start:=storyStart, End:=matchRange.Start
            Loop

            Set currentStory = currentStory.NextStoryRange
        Loop
    Next story

    If replacements.Count = 0 Then
        MsgBox "Nao foram encontrados campos com realce amarelo neste documento.", vbInformation
        Exit Sub
    End If

    For Each key In replacements.Keys
        answer = Application.InputBox( _
            Prompt:="Informe o texto que substituira:" & vbCrLf & vbCrLf & CStr(key), _
            Title:="Preencher campos amarelos", _
            Type:=2)

        If VarType(answer) = vbBoolean Then Exit Sub
        replacements(key) = CStr(answer)
    Next key

    For Each key In replacements.Keys
        replacedCount = replacedCount + ReplaceHighlightedText(CStr(key), CStr(replacements(key)))
    Next key

    MsgBox CStr(replacedCount) & " campo(s) preenchido(s).", vbInformation

    If MsgBox("Pretende gravar o documento preenchido noutra localizacao?", vbQuestion Or vbYesNo) = vbYes Then
        With Application.FileDialog(msoFileDialogSaveAs)
            .Title = "Gravar documento preenchido noutra localizacao"
            .InitialFileName = Left(ActiveDocument.FullName, Len(ActiveDocument.FullName) - Len(ActiveDocument.Name)) _
                & Left(ActiveDocument.Name, InStrRev(ActiveDocument.Name, ".") - 1) & " - preenchido.docx"

            If .Show = -1 Then
                On Error GoTo SaveCopyError
                ActiveDocument.SaveCopyAs .SelectedItems(1)
                MsgBox "Documento gravado em:" & vbCrLf & .SelectedItems(1), vbInformation
                Exit Sub
            End If
        End With
    End If

    Exit Sub

SaveCopyError:
    MsgBox "Nao foi possivel gravar o documento:" & vbCrLf & Err.Description, vbExclamation
End Sub

Private Function ReplaceHighlightedText(ByVal placeholder As String, ByVal replacement As String) As Long
    Dim story As Range
    Dim currentStory As Range
    Dim searchRange As Range
    Dim matchRange As Range
    Dim storyStart As Long

    For Each story In ActiveDocument.StoryRanges
        Set currentStory = story

        Do While Not currentStory Is Nothing
            storyStart = currentStory.Start
            Set searchRange = currentStory.Duplicate

            With searchRange.Find
                .ClearFormatting
                .Text = placeholder
                .Format = True
                .Highlight = True
                .MatchCase = True
                .Forward = False
                .Wrap = wdFindStop
            End With

            Do While searchRange.Find.Execute
                Set matchRange = searchRange.Duplicate

                If matchRange.HighlightColorIndex = wdYellow Then
                    matchRange.Text = replacement
                    matchRange.HighlightColorIndex = wdNoHighlight
                    ReplaceHighlightedText = ReplaceHighlightedText + 1
                End If

                If matchRange.Start <= storyStart Then Exit Do
                searchRange.SetRange Start:=storyStart, End:=matchRange.Start
            Loop

            Set currentStory = currentStory.NextStoryRange
        Loop
    Next story
End Function
