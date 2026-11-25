# Manga Chapter Compressor - localized PowerShell utility.
# Combines chapter archives or image folders into one CBZ or separate volume files.
# Requirements: Windows 10/11, Windows PowerShell 5.1 or PowerShell 7+.

$ErrorActionPreference = 'Stop'

$ImageExtensions = @('.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp', '.avif', '.tif', '.tiff')
$ArchiveExtensions = @('.cbz', '.zip')
$ExportFolderName = 'CBZ_Export'

# Detect the Windows display language. Russian is supported; English is the fallback.
$script:IsRussian = (Get-UICulture).TwoLetterISOLanguageName -eq 'ru'
$script:Text = if ($script:IsRussian) {
    @{
        Header = ' Manga Compressor | Создание CBZ-архивов для читалки'
        FolderDialog = 'Выберите папку одной манги (папку с главами или архивами CBZ/ZIP).'
        ArchiveReadError = '[!] Не удалось прочитать архив {0}: {1}'
        EmptyGroup = 'В выбранной группе нет изображений.'
        MissingPage = 'Не найдена страница {0} внутри {1}'
        ProgressActivity = 'Создание CBZ: {0}'
        Cancelled = 'Папка не выбрана. Работа отменена.'
        SelectedFolder = 'Выбрана папка: {0}'
        WhatToDo = 'Что сделать с главами?'
        MergeChoice = 'Собрать всё в один CBZ'
        SplitChoice = 'Разделить по томам (по обозначениям в именах архивов/папок)'
        ModePrompt = 'Ваш выбор [1/2]'
        ModeInvalid = 'Введите 1 или 2.'
        MangaPrompt = 'Название манги'
        MangaNameEmpty = 'Название манги не может быть пустым.'
        Scanning = 'Ищу главы и архивы...'
        NoImages = 'Изображения не найдены. Выберите папку одной манги, внутри которой находятся папки глав или архивы CBZ/ZIP.'
        FoundChapters = 'Найдено глав: {0}'
        FoundPages = 'Найдено страниц: {0}'
        VolumeWarning = 'Внимание: у {0} глав не найден номер тома. Они попадут в tom-1.'
        NoVolumeLabels = 'Обозначения томов не найдены. Все главы будут собраны в tom-1.'
        VolumeLabel = 'Том {0}'
        ResultSaved = 'Результат будет сохранён в: {0}'
        Creating = 'Создаю {0}...'
        Created = 'Готово: {0}'
        AllDone = 'Все архивы успешно созданы!'
        OutputFolder = 'Папка: {0}'
        Error = 'Ошибка: {0}'
        ClosePrompt = 'Нажмите Enter, чтобы закрыть окно...'
    }
} else {
    @{
        Header = ' Manga Compressor | Create CBZ archives for e-readers'
        FolderDialog = 'Select the folder containing one manga title (chapter folders or CBZ/ZIP archives).'
        ArchiveReadError = '[!] Could not read archive {0}: {1}'
        EmptyGroup = 'The selected group contains no images.'
        MissingPage = 'Page {0} was not found inside {1}'
        ProgressActivity = 'Creating CBZ: {0}'
        Cancelled = 'No folder selected. Operation cancelled.'
        SelectedFolder = 'Selected folder: {0}'
        WhatToDo = 'What would you like to do with the chapters?'
        MergeChoice = 'Merge everything into one CBZ'
        SplitChoice = 'Split into volumes (using labels in archive/folder names)'
        ModePrompt = 'Choose [1/2]'
        ModeInvalid = 'Enter 1 or 2.'
        MangaPrompt = 'Manga title'
        MangaNameEmpty = 'Manga title cannot be empty.'
        Scanning = 'Scanning for chapters and archives...'
        NoImages = 'No images found. Select one manga folder containing chapter folders or CBZ/ZIP archives.'
        FoundChapters = 'Detected chapters: {0}'
        FoundPages = 'Detected pages: {0}'
        VolumeWarning = 'Warning: {0} chapters have no volume number. They will be added to tom-1.'
        NoVolumeLabels = 'No volume labels found. All chapters will be placed in tom-1.'
        VolumeLabel = 'Volume {0}'
        ResultSaved = 'Output will be saved to: {0}'
        Creating = 'Creating {0}...'
        Created = 'Created: {0}'
        AllDone = 'All archives were created successfully!'
        OutputFolder = 'Folder: {0}'
        Error = 'Error: {0}'
        ClosePrompt = 'Press Enter to close this window...'
    }
}

function Get-NaturalKey {
    param([Parameter(Mandatory = $true)][string]$Value)

    $lower = $Value.ToLowerInvariant()
    return [regex]::Replace($lower, '\d+', {
        param($match)
        try {
            return ([int64]$match.Value).ToString('D20')
        }
        catch {
            return $match.Value.PadLeft(20, '0')
        }
    })
}

function Get-VolumeNumber {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }

    $patterns = @(
        '(?i)(?<![A-Za-zА-Яа-я])(?:том|tom|volume|vol|v)[\s._-]*0*(\d+)(?!\d)',
        '(?i)(?<!\d)0*(\d+)[\s._-]*(?:том|tom|volume|vol)(?![A-Za-zА-Яа-я])'
    )

    foreach ($pattern in $patterns) {
        $match = [regex]::Match($Value, $pattern)
        if ($match.Success) {
            try { return [int]$match.Groups[1].Value } catch { return $null }
        }
    }

    return $null
}

function ConvertTo-SafeFileName {
    param([string]$Value)

    $safe = [regex]::Replace($Value, '[<>:"/\\|?*\x00-\x1F]', '_')
    $safe = [regex]::Replace($safe, '\s+', ' ').Trim()
    $safe = $safe.TrimEnd([char[]]@('.', ' '))
    if ([string]::IsNullOrWhiteSpace($safe)) { return 'Manga' }
    return $safe
}

function Get-UniqueOutputPath {
    param(
        [string]$Folder,
        [string]$FileName
    )

    $candidate = Join-Path $Folder $FileName
    if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $extension = [System.IO.Path]::GetExtension($FileName)
    $number = 2
    do {
        $candidate = Join-Path $Folder ("{0}_{1}{2}" -f $baseName, $number, $extension)
        $number++
    } while (Test-Path -LiteralPath $candidate)

    return $candidate
}

function Select-MangaFolder {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $script:Text.FolderDialog
    $dialog.ShowNewFolderButton = $false

    try {
        $result = $dialog.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK -and -not [string]::IsNullOrWhiteSpace($dialog.SelectedPath)) {
            return $dialog.SelectedPath
        }
        return $null
    }
    finally {
        $dialog.Dispose()
    }
}

function Find-Chapters {
    param([Parameter(Mandatory = $true)][string]$Root)

    $chapters = New-Object System.Collections.Generic.List[object]
    $allFiles = @(Get-ChildItem -LiteralPath $Root -File -Recurse -Force -ErrorAction SilentlyContinue |
        Where-Object {
            $_.FullName -notmatch '(?i)(^|[\\/])CBZ_Export([\\/]|$)' -and
            $_.FullName.Substring($Root.Length) -notmatch '(^|[\\/])\.[^\\/]+'
        })

    # 1) Каждый CBZ/ZIP является отдельной главой.
    # Раньше архивы из одной папки объединялись в одну группу, из-за чего
    # определялся только первый найденный номер тома.
    foreach ($file in @($allFiles | Where-Object {
        $ArchiveExtensions -contains $_.Extension.ToLowerInvariant()
    })) {
        $pageRefs = New-Object System.Collections.Generic.List[object]
        $inputZip = $null
        try {
            $inputZip = [System.IO.Compression.ZipFile]::OpenRead($file.FullName)
            $members = @($inputZip.Entries | Where-Object {
                -not [string]::IsNullOrEmpty($_.Name) -and
                $ImageExtensions -contains [System.IO.Path]::GetExtension($_.FullName).ToLowerInvariant()
            } | Sort-Object { Get-NaturalKey $_.FullName })

            foreach ($member in $members) {
                $pageRefs.Add([pscustomobject]@{
                    Path = $file.FullName
                    EntryName = $member.FullName
                    Name = $member.Name
                    IsArchived = $true
                    SortKey = Get-NaturalKey $member.FullName
                })
            }
        }
        catch {
            Write-Host ($script:Text.ArchiveReadError -f $file.Name, $_.Exception.Message) -ForegroundColor Yellow
        }
        finally {
            if ($null -ne $inputZip) { $inputZip.Dispose() }
        }

        if ($pageRefs.Count -eq 0) { continue }

        $relativeDir = $file.DirectoryName.Substring($Root.Length).TrimStart([char[]]@('\', '/'))
        $relativeName = if ([string]::IsNullOrWhiteSpace($relativeDir)) {
            $file.BaseName
        } else {
            ($relativeDir -replace '\\', '/') + '/' + $file.BaseName
        }

        # Номер тома ищем сначала в имени самого архива, затем в родительских папках.
        $volume = Get-VolumeNumber $file.BaseName
        if ($null -eq $volume -and -not [string]::IsNullOrWhiteSpace($relativeDir)) {
            $parts = $relativeDir -split '[\\/]'
            for ($i = $parts.Count - 1; $i -ge 0; $i--) {
                $volume = Get-VolumeNumber $parts[$i]
                if ($null -ne $volume) { break }
            }
        }
        if ($null -eq $volume) { $volume = Get-VolumeNumber (Split-Path -Leaf $Root) }

        $chapters.Add([pscustomobject]@{
            RelativeName = $relativeName
            Pages = @($pageRefs.ToArray())
            Volume = $volume
        })
    }

    # 2) Изображения, лежащие непосредственно в папках, считаются одной главой на папку.
    $imageFiles = @($allFiles | Where-Object {
        $ImageExtensions -contains $_.Extension.ToLowerInvariant()
    })
    foreach ($group in @($imageFiles | Group-Object -Property DirectoryName)) {
        $directory = $group.Name
        $pageRefs = New-Object System.Collections.Generic.List[object]

        foreach ($file in $group.Group) {
            $pageRefs.Add([pscustomobject]@{
                Path = $file.FullName
                EntryName = $null
                Name = $file.Name
                IsArchived = $false
                SortKey = Get-NaturalKey $file.Name
            })
        }

        if ($pageRefs.Count -eq 0) { continue }

        $relativeDir = $directory.Substring($Root.Length).TrimStart([char[]]@('\', '/'))
        if ([string]::IsNullOrWhiteSpace($relativeDir)) {
            $relativeName = Split-Path -Leaf $Root
        } else {
            $relativeName = $relativeDir -replace '\\', '/'
        }

        $volume = $null
        if (-not [string]::IsNullOrWhiteSpace($relativeDir)) {
            $parts = $relativeDir -split '[\\/]'
            for ($i = $parts.Count - 1; $i -ge 0; $i--) {
                $volume = Get-VolumeNumber $parts[$i]
                if ($null -ne $volume) { break }
            }
        }
        if ($null -eq $volume) { $volume = Get-VolumeNumber (Split-Path -Leaf $Root) }

        $chapters.Add([pscustomobject]@{
            RelativeName = $relativeName
            Pages = @($pageRefs.ToArray())
            Volume = $volume
        })
    }

    return @($chapters | Sort-Object { Get-NaturalKey $_.RelativeName })
}

function New-CbzArchive {
    param(
        [Parameter(Mandatory = $true)][object[]]$Chapters,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][string]$Label
    )

    $total = 0
    foreach ($chapter in $Chapters) { $total += $chapter.Pages.Count }
    if ($total -le 0) { throw $script:Text.EmptyGroup }

    $completed = 0
    $startTime = Get-Date
    $outputStream = $null
    $outputZip = $null
    $inputZip = $null
    $currentInputPath = $null

    try {
        $outputStream = [System.IO.File]::Open(
            $OutputPath,
            [System.IO.FileMode]::CreateNew,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None
        )
        $outputZip = [System.IO.Compression.ZipArchive]::new(
            $outputStream,
            [System.IO.Compression.ZipArchiveMode]::Create,
            $false
        )

        foreach ($chapter in $Chapters) {
            foreach ($page in $chapter.Pages) {
                $inputStream = $null
                $pageOutputStream = $null
                try {
                    if ($page.IsArchived) {
                        if ($currentInputPath -ne $page.Path) {
                            if ($null -ne $inputZip) { $inputZip.Dispose() }
                            $inputZip = [System.IO.Compression.ZipFile]::OpenRead($page.Path)
                            $currentInputPath = $page.Path
                        }
                        $inputEntry = $inputZip.GetEntry($page.EntryName)
                        if ($null -eq $inputEntry) {
                            throw ($script:Text.MissingPage -f $page.EntryName, $page.Path)
                        }
                        $inputStream = $inputEntry.Open()
                    }
                    else {
                        $inputStream = [System.IO.File]::OpenRead($page.Path)
                    }

                    $extension = [System.IO.Path]::GetExtension($page.Name).ToLowerInvariant()
                    if ($ImageExtensions -notcontains $extension) { $extension = '.jpg' }
                    $entryName = '{0:D7}{1}' -f ($completed + 1), $extension
                    $outputEntry = $outputZip.CreateEntry($entryName, [System.IO.Compression.CompressionLevel]::NoCompression)
                    $pageOutputStream = $outputEntry.Open()
                    $inputStream.CopyTo($pageOutputStream)
                }
                finally {
                    if ($null -ne $pageOutputStream) { $pageOutputStream.Dispose() }
                    if ($null -ne $inputStream) { $inputStream.Dispose() }
                }

                $completed++
                $percent = [math]::Floor(($completed / $total) * 100)
                $elapsed = (Get-Date) - $startTime
                $remaining = 0
                if ($completed -gt 0) {
                    $remaining = [int][math]::Max(0, ($elapsed.TotalSeconds / $completed) * ($total - $completed))
                }
                $currentName = [System.IO.Path]::GetFileName($page.Name)
                Write-Progress -Activity ($script:Text.ProgressActivity -f $Label) `
                    -Status ("{0}% | {1}/{2} | {3}" -f $percent, $completed, $total, $currentName) `
                    -PercentComplete $percent -SecondsRemaining $remaining
            }
        }
    }
    finally {
        Write-Progress -Activity ($script:Text.ProgressActivity -f $Label) -Completed
        if ($null -ne $inputZip) { $inputZip.Dispose() }
        if ($null -ne $outputZip) { $outputZip.Dispose() }
        if ($null -ne $outputStream) { $outputStream.Dispose() }
    }
}

try {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    Clear-Host
    Write-Host '============================================================' -ForegroundColor DarkCyan
    Write-Host $script:Text.Header -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor DarkCyan
    Write-Host ''

    $root = Select-MangaFolder
    if ([string]::IsNullOrWhiteSpace($root)) {
        Write-Host $script:Text.Cancelled -ForegroundColor Yellow
        exit 0
    }

    Write-Host ($script:Text.SelectedFolder -f $root) -ForegroundColor Gray
    Write-Host ''
    Write-Host $script:Text.WhatToDo -ForegroundColor Cyan
    Write-Host ("  1 - {0}" -f $script:Text.MergeChoice)
    Write-Host ("  2 - {0}" -f $script:Text.SplitChoice)

    do {
        $mode = Read-Host $script:Text.ModePrompt
        if ($mode -notin @('1', '2')) { Write-Host $script:Text.ModeInvalid -ForegroundColor Yellow }
    } while ($mode -notin @('1', '2'))

    $mangaNameInput = Read-Host $script:Text.MangaPrompt
    if ([string]::IsNullOrWhiteSpace($mangaNameInput)) {
        throw $script:Text.MangaNameEmpty
    }
    $mangaName = ConvertTo-SafeFileName $mangaNameInput

    Write-Host ''
    Write-Host $script:Text.Scanning -ForegroundColor Cyan
    $chapters = @(Find-Chapters -Root $root)
    if ($chapters.Count -eq 0) {
        throw $script:Text.NoImages
    }

    $totalPages = 0
    foreach ($chapter in $chapters) { $totalPages += $chapter.Pages.Count }
    Write-Host ($script:Text.FoundChapters -f $chapters.Count) -ForegroundColor Green
    Write-Host ($script:Text.FoundPages -f $totalPages) -ForegroundColor Green

    $outputFolder = Join-Path $root $ExportFolderName
    New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null

    $jobs = New-Object System.Collections.Generic.List[object]
    if ($mode -eq '1') {
        $outputPath = Get-UniqueOutputPath -Folder $outputFolder -FileName ("{0}.cbz" -f $mangaName)
        $jobs.Add([pscustomobject]@{ Chapters = $chapters; OutputPath = $outputPath; Label = $mangaName })
    }
    else {
        $volumeGroups = @{}
        foreach ($chapter in $chapters) {
            $volume = 1
            if ($null -ne $chapter.Volume) { $volume = [int]$chapter.Volume }
            if (-not $volumeGroups.ContainsKey($volume)) {
                $volumeGroups[$volume] = New-Object System.Collections.Generic.List[object]
            }
            $volumeGroups[$volume].Add($chapter)
        }

        $chaptersWithoutVolume = @($chapters | Where-Object { $null -eq $_.Volume }).Count
        if ($chaptersWithoutVolume -gt 0) {
            Write-Host ($script:Text.VolumeWarning -f $chaptersWithoutVolume) -ForegroundColor Yellow
        }
        if (@($chapters | Where-Object { $null -ne $_.Volume }).Count -eq 0) {
            Write-Host $script:Text.NoVolumeLabels -ForegroundColor Yellow
        }

        foreach ($volume in @($volumeGroups.Keys | Sort-Object { [int]$_ })) {
            $fileName = "{0}-tom-{1}.cbz" -f $mangaName, $volume
            $outputPath = Get-UniqueOutputPath -Folder $outputFolder -FileName $fileName
            $jobs.Add([pscustomobject]@{
                Chapters = @($volumeGroups[$volume].ToArray())
                OutputPath = $outputPath
                Label = ($script:Text.VolumeLabel -f $volume)
            })
        }
    }

    Write-Host ''
    Write-Host ($script:Text.ResultSaved -f $outputFolder) -ForegroundColor Gray

    foreach ($job in $jobs) {
        Write-Host ''
        Write-Host ($script:Text.Creating -f [System.IO.Path]::GetFileName($job.OutputPath)) -ForegroundColor Cyan
        try {
            New-CbzArchive -Chapters $job.Chapters -OutputPath $job.OutputPath -Label $job.Label
            Write-Host ($script:Text.Created -f $job.OutputPath) -ForegroundColor Green
        }
        catch {
            if (Test-Path -LiteralPath $job.OutputPath) {
                Remove-Item -LiteralPath $job.OutputPath -Force -ErrorAction SilentlyContinue
            }
            throw
        }
    }

    Write-Host ''
    Write-Host $script:Text.AllDone -ForegroundColor Green
    Write-Host ($script:Text.OutputFolder -f $outputFolder) -ForegroundColor Gray
}
catch {
    Write-Host ''
    Write-Host ($script:Text.Error -f $_.Exception.Message) -ForegroundColor Red
}

Write-Host ''
Write-Host $script:Text.ClosePrompt
[void](Read-Host)
