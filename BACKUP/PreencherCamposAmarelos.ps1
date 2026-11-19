Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Show-TextPrompt {
    param(
        [string]$Placeholder,
        [int]$Index,
        [int]$Total,
        [string]$CurrentValue
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Preencher campos amarelos"
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(560, 290)
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false
    $form.FormBorderStyle = "FixedDialog"
    $form.Tag = "Cancel"

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Campo $Index de $Total`r`nInforme o texto que substituira:"
    $label.Location = New-Object System.Drawing.Point(12, 12)
    $label.Size = New-Object System.Drawing.Size(520, 40)
    $form.Controls.Add($label)

    $placeholderBox = New-Object System.Windows.Forms.RichTextBox
    $placeholderBox.Location = New-Object System.Drawing.Point(12, 54)
    $placeholderBox.Size = New-Object System.Drawing.Size(520, 24)
    $placeholderBox.Multiline = $false
    $placeholderBox.ReadOnly = $true
    $placeholderBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $placeholderBox.BackColor = $form.BackColor
    $placeholderBox.Font = $label.Font
    $placeholderBox.Text = $Placeholder
    $form.Controls.Add($placeholderBox)

    $boldFont = $null
    foreach ($bracketMatch in [regex]::Matches($Placeholder, '\[(?<content>[^\[\]]+)\]')) {
        $content = $bracketMatch.Groups['content']
        if ($null -eq $boldFont) {
            $boldFont = New-Object System.Drawing.Font($placeholderBox.Font, [System.Drawing.FontStyle]::Bold)
        }
        $placeholderBox.Select($content.Index, $content.Length)
        $placeholderBox.SelectionFont = $boldFont
    }
    $placeholderBox.Select(0, 0)

    $noteLabel = New-Object System.Windows.Forms.Label
    $noteLabel.Text = "Atencao: Caso nao saibas o que colocar neste campo preenche com XXXX, poderas alterar directamente no WORD posteriormente"
    $noteLabel.Location = New-Object System.Drawing.Point(12, 80)
    $noteLabel.Size = New-Object System.Drawing.Size(520, 36)
    $form.Controls.Add($noteLabel)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(12, 120)
    $textBox.Size = New-Object System.Drawing.Size(520, 70)
    $textBox.Multiline = $true
    $textBox.ScrollBars = "Vertical"
    $textBox.Text = $CurrentValue
    $textBox.SelectionStart = $textBox.Text.Length
    $textBox.SelectionLength = 0
    $form.Controls.Add($textBox)
    $form.Add_Shown({
        $textBox.Focus()
    })

    $previousButton = New-Object System.Windows.Forms.Button
    $previousButton.Text = "Anterior"
    $previousButton.Location = New-Object System.Drawing.Point(12, 205)
    $previousButton.Size = New-Object System.Drawing.Size(90, 30)
    $previousButton.Enabled = ($Index -gt 1)
    $previousButton.Add_Click({
        $form.Tag = "Previous"
        $form.Close()
    })
    $form.Controls.Add($previousButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancelar"
    $cancelButton.Location = New-Object System.Drawing.Point(432, 205)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
    $cancelButton.Add_Click({
        $form.Tag = "Cancel"
        $form.Close()
    })
    $form.Controls.Add($cancelButton)

    $nextButton = New-Object System.Windows.Forms.Button
    if ($Index -eq $Total) {
        $nextButton.Text = "Concluir"
    }
    else {
        $nextButton.Text = "Seguinte"
    }
    $nextButton.Location = New-Object System.Drawing.Point(326, 205)
    $nextButton.Size = New-Object System.Drawing.Size(100, 30)
    if ($Index -eq $Total) {
        $nextButton.Add_Click({
            $form.Tag = "Complete"
            $form.Close()
        })
    }
    else {
        $nextButton.Add_Click({
            $form.Tag = "Next"
            $form.Close()
        })
    }
    $form.Controls.Add($nextButton)

    $form.AcceptButton = $nextButton
    $form.CancelButton = $cancelButton
    $form.ShowDialog() | Out-Null
    $answer = [pscustomobject]@{
        Action = [string]$form.Tag
        Value = $textBox.Text
    }
    $form.Dispose()
    if ($null -ne $boldFont) {
        $boldFont.Dispose()
    }
    return $answer
}

function Show-ReviewSummary {
    param([string]$Summary)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Conferencia das alteracoes"
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(720, 520)
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false
    $form.FormBorderStyle = "Sizable"

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Os textos alterados estao a verde; sequencias de XXXX estao realcadas a vermelho:"
    $label.Location = New-Object System.Drawing.Point(12, 12)
    $label.Size = New-Object System.Drawing.Size(680, 32)
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(12, 48)
    $textBox.Size = New-Object System.Drawing.Size(680, 385)
    $textBox.Multiline = $true
    $textBox.ReadOnly = $true
    $textBox.ScrollBars = "Both"
    $textBox.WordWrap = $false
    $textBox.Text = $Summary
    $form.Controls.Add($textBox)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Fechar"
    $closeButton.Location = New-Object System.Drawing.Point(617, 445)
    $closeButton.Size = New-Object System.Drawing.Size(75, 28)
    $closeButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($closeButton)
    $form.AcceptButton = $closeButton

    $form.ShowDialog() | Out-Null
    $form.Dispose()
}

$fileDialog = New-Object System.Windows.Forms.OpenFileDialog
$fileDialog.Title = "Selecione o documento do Word"
$fileDialog.Filter = "Documentos do Word (*.doc;*.docx;*.docm;*.rtf)|*.doc;*.docx;*.docm;*.rtf|Todos os arquivos (*.*)|*.*"
$fileDialog.Multiselect = $false

if ($fileDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    $fileDialog.Dispose()
    exit 0
}
$documentPath = $fileDialog.FileName
$fileDialog.Dispose()

$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $document = $word.Documents.Open($documentPath)

    $replacements = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::Ordinal)
    $occurrences = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([System.StringComparer]::Ordinal)
    $wdYellow = 7
    $wdBrightGreen = 4
    $wdRed = 6
    $wdFindStop = 0

    foreach ($story in $document.StoryRanges) {
        $currentStory = $story
        while ($null -ne $currentStory) {
            $storyStart = $currentStory.Start
            $searchRange = $currentStory.Duplicate
            $find = $searchRange.Find
            $find.ClearFormatting()
            $find.Text = ""
            $find.Format = $true
            $find.Highlight = $true
            $find.Forward = $false
            $find.Wrap = $wdFindStop

            while ($find.Execute()) {
                $matchRange = $searchRange.Duplicate
                if ($matchRange.HighlightColorIndex -eq $wdYellow) {
                    $placeholder = [string]$matchRange.Text
                    if (-not $replacements.ContainsKey($placeholder)) {
                        $replacements.Add($placeholder, "")
                    }
                }

                if ($matchRange.Start -le $storyStart) {
                    break
                }
                $searchRange.SetRange($storyStart, $matchRange.Start)
            }

            $currentStory = $currentStory.NextStoryRange
        }
    }

    if ($replacements.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Nao foram encontrados campos com realce amarelo neste documento.",
            "Preencher campos amarelos",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        exit 0
    }

    $placeholders = @($replacements.Keys)
    $promptIndex = 0
    $inputComplete = $false
    while (-not $inputComplete) {
        $placeholder = $placeholders[$promptIndex]
        $answer = Show-TextPrompt `
            -Placeholder $placeholder `
            -Index ($promptIndex + 1) `
            -Total $placeholders.Count `
            -CurrentValue $replacements[$placeholder]
        $replacements[$placeholder] = $answer.Value

        switch ($answer.Action) {
            "Previous" { $promptIndex-- }
            "Next" { $promptIndex++ }
            "Complete" { $inputComplete = $true }
            default { exit 0 }
        }
    }

    $replacedCount = 0
    foreach ($placeholder in @($replacements.Keys)) {
        foreach ($story in $document.StoryRanges) {
            $currentStory = $story
            while ($null -ne $currentStory) {
                $storyStart = $currentStory.Start
                $searchRange = $currentStory.Duplicate
                $find = $searchRange.Find
                $find.ClearFormatting()
                $find.Text = $placeholder
                $find.Format = $true
                $find.Highlight = $true
                $find.MatchCase = $true
                $find.Forward = $false
                $find.Wrap = $wdFindStop

                while ($find.Execute()) {
                    $matchRange = $searchRange.Duplicate
                    if ($matchRange.HighlightColorIndex -eq $wdYellow) {
                        $replacementText = $replacements[$placeholder]
                        $replacementStart = $matchRange.Start
                        $matchRange.Text = $replacementText
                        $matchRange.HighlightColorIndex = $wdBrightGreen

                        foreach ($xMatch in [regex]::Matches($replacementText, 'X{4,}', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
                            $reviewRange = $matchRange.Duplicate
                            $reviewRange.SetRange(
                                $replacementStart + $xMatch.Index,
                                $replacementStart + $xMatch.Index + $xMatch.Length
                            )
                            $reviewRange.HighlightColorIndex = $wdRed
                        }

                        $replacedCount++
                        if ($occurrences.ContainsKey($placeholder)) {
                            $occurrences[$placeholder]++
                        }
                        else {
                            $occurrences.Add($placeholder, 1)
                        }
                    }

                    if ($matchRange.Start -le $storyStart) {
                        break
                    }
                    $searchRange.SetRange($storyStart, $matchRange.Start)
                }

                $currentStory = $currentStory.NextStoryRange
            }
        }
    }

    $summaryLines = New-Object 'System.Collections.Generic.List[string]'
    $summaryLines.Add("$replacedCount campo(s) preenchido(s).")
    $summaryLines.Add("")
    foreach ($placeholder in @($replacements.Keys)) {
        $summaryLines.Add("Original: $placeholder")
        $summaryLines.Add("Substituido por: $($replacements[$placeholder])")
        $summaryLines.Add("Ocorrencias alteradas: $($occurrences[$placeholder])")
        $summaryLines.Add("")
    }

    Show-ReviewSummary -Summary ($summaryLines -join [Environment]::NewLine)
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Nao foi possivel preencher o documento:`r`n$($_.Exception.Message)",
        "Erro",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}
