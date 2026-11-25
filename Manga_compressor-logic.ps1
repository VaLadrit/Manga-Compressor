# Mihon_to_CBZ.ps1
# Собирает изображения из папок глав и CBZ/ZIP-архивов в один CBZ или по томам.
# Требования: Windows 10/11, встроенный Windows PowerShell 5.1 или PowerShell 7+.

$ErrorActionPreference = 'Stop'

$ImageExtensions = @('.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp', '.avif', '.tif', '.tiff')
$ArchiveExtensions = @('.cbz', '.zip')
$ExportFolderName = 'CBZ_Export'

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
    $dialog.Description = 'Выберите папку одной манги из Mihon (папку с главами или архивами)'
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
            Write-Host ("[!] Не удалось прочитать архив {0}: {1}" -f $file.Name, $_.Exception.Message) -ForegroundColor Yellow
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
    if ($total -le 0) { throw 'В выбранной группе нет изображений.' }

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
                            throw ("Не найдена страница {0} внутри {1}" -f $page.EntryName, $page.Path)
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
                Write-Progress -Activity ("Создание CBZ: {0}" -f $Label) `
                    -Status ("{0}% | {1}/{2} | {3}" -f $percent, $completed, $total, $currentName) `
                    -PercentComplete $percent -SecondsRemaining $remaining
            }
        }
    }
    finally {
        Write-Progress -Activity ("Создание CBZ: {0}" -f $Label) -Completed
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
    Write-Host ' Mihon -> CBZ | Сборка манги для электронной книги' -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor DarkCyan
    Write-Host ''

    $root = Select-MangaFolder
    if ([string]::IsNullOrWhiteSpace($root)) {
        Write-Host 'Папка не выбрана. Работа отменена.' -ForegroundColor Yellow
        exit 0
    }

    Write-Host ("Выбрана папка: {0}" -f $root) -ForegroundColor Gray
    Write-Host ''
    Write-Host 'Что сделать с главами?' -ForegroundColor Cyan
    Write-Host '  1 - Собрать всё в один CBZ'
    Write-Host '  2 - Разделить по томам (по номеру тома в имени каждого архива/папки)'

    do {
        $mode = Read-Host 'Ваш выбор [1/2]'
        if ($mode -notin @('1', '2')) { Write-Host 'Введите 1 или 2.' -ForegroundColor Yellow }
    } while ($mode -notin @('1', '2'))

    $mangaNameInput = Read-Host 'Название манги'
    if ([string]::IsNullOrWhiteSpace($mangaNameInput)) {
        throw 'Название манги не может быть пустым.'
    }
    $mangaName = ConvertTo-SafeFileName $mangaNameInput

    Write-Host ''
    Write-Host 'Ищу папки глав и архивы...' -ForegroundColor Cyan
    $chapters = @(Find-Chapters -Root $root)
    if ($chapters.Count -eq 0) {
        throw 'Изображения не найдены. Выберите папку одной манги, внутри которой находятся папки глав или CBZ/ZIP.'
    }

    $totalPages = 0
    foreach ($chapter in $chapters) { $totalPages += $chapter.Pages.Count }
    Write-Host ("Найдено папок с главами: {0}" -f $chapters.Count) -ForegroundColor Green
    Write-Host ("Найдено страниц: {0}" -f $totalPages) -ForegroundColor Green

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
            Write-Host ("Внимание: у {0} папок не найден номер тома. Они попадут в tom-1." -f $chaptersWithoutVolume) -ForegroundColor Yellow
        }
        if (@($chapters | Where-Object { $null -ne $_.Volume }).Count -eq 0) {
            Write-Host 'Обозначения томов не найдены. Все главы будут собраны в tom-1.' -ForegroundColor Yellow
        }

        foreach ($volume in @($volumeGroups.Keys | Sort-Object { [int]$_ })) {
            $fileName = "{0}-tom-{1}.cbz" -f $mangaName, $volume
            $outputPath = Get-UniqueOutputPath -Folder $outputFolder -FileName $fileName
            $jobs.Add([pscustomobject]@{
                Chapters = @($volumeGroups[$volume].ToArray())
                OutputPath = $outputPath
                Label = "Том $volume"
            })
        }
    }

    Write-Host ''
    Write-Host ("Результат будет сохранён в: {0}" -f $outputFolder) -ForegroundColor Gray

    foreach ($job in $jobs) {
        Write-Host ''
        Write-Host ("Создаю {0}..." -f [System.IO.Path]::GetFileName($job.OutputPath)) -ForegroundColor Cyan
        try {
            New-CbzArchive -Chapters $job.Chapters -OutputPath $job.OutputPath -Label $job.Label
            Write-Host ("Готово: {0}" -f $job.OutputPath) -ForegroundColor Green
        }
        catch {
            if (Test-Path -LiteralPath $job.OutputPath) {
                Remove-Item -LiteralPath $job.OutputPath -Force -ErrorAction SilentlyContinue
            }
            throw
        }
    }

    Write-Host ''
    Write-Host 'Все архивы успешно созданы!' -ForegroundColor Green
    Write-Host ("Папка: {0}" -f $outputFolder) -ForegroundColor Gray
}
catch {
    Write-Host ''
    Write-Host ("Ошибка: {0}" -f $_.Exception.Message) -ForegroundColor Red
}

Write-Host ''
Write-Host 'Нажмите Enter, чтобы закрыть окно...'
[void](Read-Host)
