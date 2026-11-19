Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Show-TextPrompt {
    param(
        [string]$Placeholder,
        [int]$Index,
        [int]$Total,
        [string]$CurrentValue,
        [IntPtr]$OwnerHandle
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Preencher campos amarelos"
    $form.StartPosition = "CenterScreen"
    $form.ClientSize = New-Object System.Drawing.Size(720, 490)
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false
    $form.FormBorderStyle = "FixedDialog"
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $baseFont = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Font = $baseFont
    $form.Tag = "Cancel"

    $header = New-Object System.Windows.Forms.Panel
    $header.Location = New-Object System.Drawing.Point(0, 0)
    $header.Size = New-Object System.Drawing.Size(720, 80)
    $header.BackColor = [System.Drawing.Color]::FromArgb(31, 55, 85)
    $form.Controls.Add($header)

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Preencher campos amarelos"
    $titleLabel.Location = New-Object System.Drawing.Point(24, 16)
    $titleLabel.Size = New-Object System.Drawing.Size(470, 32)
    $titleLabel.ForeColor = [System.Drawing.Color]::White
    $titleFont = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $titleLabel.Font = $titleFont
    $header.Controls.Add($titleLabel)

    $stepLabel = New-Object System.Windows.Forms.Label
    $stepLabel.Text = "Campo $Index de $Total"
    $stepLabel.Location = New-Object System.Drawing.Point(550, 27)
    $stepLabel.Size = New-Object System.Drawing.Size(145, 24)
    $stepLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $stepLabel.ForeColor = [System.Drawing.Color]::White
    $header.Controls.Add($stepLabel)

    $placeholderLabel = New-Object System.Windows.Forms.Label
    $placeholderLabel.Text = "Campo a preencher"
    $placeholderLabel.Location = New-Object System.Drawing.Point(24, 99)
    $placeholderLabel.Size = New-Object System.Drawing.Size(672, 22)
    $placeholderLabel.ForeColor = [System.Drawing.Color]::FromArgb(55, 65, 81)
    $form.Controls.Add($placeholderLabel)

    $placeholderBox = New-Object System.Windows.Forms.RichTextBox
    $placeholderBox.Location = New-Object System.Drawing.Point(24, 124)
    $placeholderBox.Size = New-Object System.Drawing.Size(672, 42)
    $placeholderBox.Multiline = $false
    $placeholderBox.ReadOnly = $true
    $placeholderBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $placeholderBox.BackColor = [System.Drawing.Color]::White
    $placeholderBox.ForeColor = [System.Drawing.Color]::FromArgb(31, 41, 55)
    $placeholderBox.Font = New-Object System.Drawing.Font("Segoe UI", 11)
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

    $notePanel = New-Object System.Windows.Forms.Panel
    $notePanel.Location = New-Object System.Drawing.Point(24, 182)
    $notePanel.Size = New-Object System.Drawing.Size(672, 62)
    $notePanel.BackColor = [System.Drawing.Color]::FromArgb(255, 247, 220)
    $form.Controls.Add($notePanel)

    $noteLabel = New-Object System.Windows.Forms.Label
    $noteLabel.Text = "Atencao: Caso nao saibas o que colocar neste campo preenche com XXXX, poderas alterar directamente no WORD posteriormente"
    $noteLabel.Location = New-Object System.Drawing.Point(14, 8)
    $noteLabel.Size = New-Object System.Drawing.Size(644, 46)
    $noteLabel.ForeColor = [System.Drawing.Color]::FromArgb(110, 75, 15)
    $notePanel.Controls.Add($noteLabel)

    $inputLabel = New-Object System.Windows.Forms.Label
    $inputLabel.Text = "Texto de substituicao"
    $inputLabel.Location = New-Object System.Drawing.Point(24, 261)
    $inputLabel.Size = New-Object System.Drawing.Size(672, 22)
    $inputLabel.ForeColor = [System.Drawing.Color]::FromArgb(55, 65, 81)
    $form.Controls.Add($inputLabel)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(24, 287)
    $textBox.Size = New-Object System.Drawing.Size(672, 112)
    $textBox.Multiline = $true
    $textBox.ScrollBars = "Vertical"
    $textBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $textBox.Font = New-Object System.Drawing.Font("Segoe UI", 11)
    $textBox.Text = $CurrentValue
    $textBox.SelectionStart = $textBox.Text.Length
    $textBox.SelectionLength = 0
    $form.Controls.Add($textBox)
    $form.Add_Shown({
        $form.Activate()
        $textBox.Focus()
    })

    $previousButton = New-Object System.Windows.Forms.Button
    $previousButton.Text = "Anterior"
    $previousButton.Location = New-Object System.Drawing.Point(24, 430)
    $previousButton.Size = New-Object System.Drawing.Size(104, 38)
    $previousButton.Enabled = ($Index -gt 1)
    $previousButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $previousButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(203, 213, 225)
    $previousButton.BackColor = [System.Drawing.Color]::White
    $previousButton.ForeColor = [System.Drawing.Color]::FromArgb(31, 41, 55)
    $previousButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $previousButton.Add_Click({
        $form.Tag = "Previous"
        $form.Close()
    })
    $form.Controls.Add($previousButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancelar"
    $cancelButton.Location = New-Object System.Drawing.Point(592, 430)
    $cancelButton.Size = New-Object System.Drawing.Size(104, 38)
    $cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $cancelButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(203, 213, 225)
    $cancelButton.BackColor = [System.Drawing.Color]::White
    $cancelButton.ForeColor = [System.Drawing.Color]::FromArgb(31, 41, 55)
    $cancelButton.Cursor = [System.Windows.Forms.Cursors]::Hand
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
    $nextButton.Location = New-Object System.Drawing.Point(476, 430)
    $nextButton.Size = New-Object System.Drawing.Size(104, 38)
    $nextButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $nextButton.FlatAppearance.BorderSize = 0
    $nextButton.BackColor = [System.Drawing.Color]::FromArgb(37, 99, 235)
    $nextButton.ForeColor = [System.Drawing.Color]::White
    $nextButton.Cursor = [System.Windows.Forms.Cursors]::Hand
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
    $ownerWindow = New-Object System.Windows.Forms.NativeWindow
    $ownerWindow.AssignHandle($OwnerHandle)
    try {
        $form.ShowDialog($ownerWindow) | Out-Null
    }
    finally {
        $ownerWindow.ReleaseHandle()
    }
    $answer = [pscustomobject]@{
        Action = [string]$form.Tag
        Value = $textBox.Text
    }
    $form.Dispose()
    if ($null -ne $boldFont) {
        $boldFont.Dispose()
    }
    $titleFont.Dispose()
    $baseFont.Dispose()
    return $answer
}

function Show-ReviewSummary {
    param(
        [string]$Summary,
        [object]$Document,
        [string]$OriginalPath,
        [IntPtr]$OwnerHandle
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Conferencia das alteracoes"
    $form.StartPosition = "CenterScreen"
    $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
    $form.TopMost = $true
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

    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Gravar noutra localizacao"
    $saveButton.Location = New-Object System.Drawing.Point(380, 445)
    $saveButton.Size = New-Object System.Drawing.Size(225, 28)
    $saveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $saveButton.FlatAppearance.BorderSize = 0
    $saveButton.BackColor = [System.Drawing.Color]::FromArgb(37, 99, 235)
    $saveButton.ForeColor = [System.Drawing.Color]::White
    $saveButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $saveButton.Add_Click({
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Title = "Gravar documento preenchido noutra localizacao"
        $saveDialog.Filter = "Documento do Word (*.docx)|*.docx|Documento do Word 97-2003 (*.doc)|*.doc|Documento RTF (*.rtf)|*.rtf"

        $originalDirectory = [System.IO.Path]::GetDirectoryName($OriginalPath)
        $originalName = [System.IO.Path]::GetFileNameWithoutExtension($OriginalPath)
        $saveDialog.InitialDirectory = $originalDirectory
        $saveDialog.FileName = "$originalName - preenchido.docx"
        $saveDialog.AddExtension = $true
        $saveDialog.DefaultExt = "docx"
        $saveDialog.OverwritePrompt = $true

        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $extension = [System.IO.Path]::GetExtension($saveDialog.FileName).ToLowerInvariant()
                $saveFormat = switch ($extension) {
                    ".doc" { 0 }
                    ".rtf" { 6 }
                    default { 16 }
                }
                $Document.SaveAs2($saveDialog.FileName, $saveFormat)
                [System.Windows.Forms.MessageBox]::Show(
                    "Documento gravado em:`r`n$($saveDialog.FileName)",
                    "Documento gravado",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                ) | Out-Null
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "Nao foi possivel gravar o documento:`r`n$($_.Exception.Message)",
                    "Erro",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                ) | Out-Null
            }
        }
        $saveDialog.Dispose()
    })
    $form.Controls.Add($saveButton)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Fechar"
    $closeButton.Location = New-Object System.Drawing.Point(617, 445)
    $closeButton.Size = New-Object System.Drawing.Size(75, 28)
    $closeButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($closeButton)
    $form.AcceptButton = $closeButton

    $form.Add_Shown({
        $form.Activate()
    })
    $ownerWindow = New-Object System.Windows.Forms.NativeWindow
    $ownerWindow.AssignHandle($OwnerHandle)
    try {
        $form.ShowDialog($ownerWindow) | Out-Null
    }
    finally {
        $ownerWindow.ReleaseHandle()
    }
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
    $wordWindowHandle = $document.ActiveWindow.Hwnd
    if ($null -eq $wordWindowHandle -or [long]$wordWindowHandle -eq 0) {
        throw "Nao foi possivel obter o identificador da janela ativa do Word."
    }

    while (-not $inputComplete) {
        $placeholder = $placeholders[$promptIndex]
        $word.Activate()
        $answer = Show-TextPrompt `
            -Placeholder $placeholder `
            -Index ($promptIndex + 1) `
            -Total $placeholders.Count `
            -CurrentValue $replacements[$placeholder] `
            -OwnerHandle ([IntPtr]([long]$wordWindowHandle))
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

    $word.Activate()
    Show-ReviewSummary `
        -Summary ($summaryLines -join [Environment]::NewLine) `
        -Document $document `
        -OriginalPath $documentPath `
        -OwnerHandle ([IntPtr]([long]$wordWindowHandle))
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
