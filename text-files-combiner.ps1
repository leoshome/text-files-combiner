Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 語言定義
$languages = @{
    "zh-TW" = @{
        FormTitle = "文字檔合併工具"
        FileExplorer = "檔案總管"
        SelectedFiles = "已選擇的檔案"
        CurrentPath = "目前路徑:"
        Browse = "瀏覽..."
        AddFiles = "加入 →"
        RemoveFiles = "← 移除"
        ClearFiles = "清空"
        MergeFiles = "合併檔案"
        FileName = "檔案名稱"
        FileType = "類型"
        FileSize = "大小"
        ModifiedDate = "修改日期"
        SelectFolder = "選擇資料夾"
        PleaseSelectFiles = "請先選擇要合併的檔案！"
        SelectSaveLocation = "選擇合併後檔案的儲存位置"
        MergeComplete = "所有檔案已成功合併至:"
        MergeError = "合併過程發生錯誤:"
        PathNotExists = "路徑不存在或無效："  
        Hint = "提示"
        Error = "錯誤"
        Complete = "完成"
        Language = "語言/Language"
        File = "檔案"
        ChineseTraditional = "繁體中文"
        English = "English"
        CannotReadFile = "【無法讀取此檔案內容】"
        TextFiles = "文字檔 (*.txt)|*.txt"
        AllFiles = "所有檔案 (*.*)|*.*"
        SelectedCount = "已選擇 {0} 個檔案"
        PleaseSelectToMerge = "請選擇要合併的檔案"
    }
    "en-US" = @{
        FormTitle = "Text File Merger Tool"
        FileExplorer = "File Explorer"
        SelectedFiles = "Selected Files"
        CurrentPath = "Current Path:"
        Browse = "Browse..."
        AddFiles = "Add →"
        RemoveFiles = "← Remove"
        ClearFiles = "Clear"
        MergeFiles = "Merge Files"
        FileName = "File Name"
        FileType = "Type"
        FileSize = "Size"
        ModifiedDate = "Modified Date"
        SelectFolder = "Select Folder"
        PleaseSelectFiles = "Please select files to merge first!"
        SelectSaveLocation = "Select save location for merged file"
        MergeComplete = "All files merged successfully to:"
        MergeError = "Error occurred during merge:"
        PathNotExists = "Path does not exist or invalid:"
        Hint = "Hint"
        Error = "Error"
        Complete = "Complete"
        Language = "語言/Language"
        File = "File"
        ChineseTraditional = "繁體中文"
        English = "English"
        CannotReadFile = "【Cannot read file content】"
        TextFiles = "Text Files (*.txt)|*.txt"
        AllFiles = "All Files (*.*)|*.*"
        SelectedCount = "Selected {0} files"
        PleaseSelectToMerge = "Please select files to merge"
    }
}

# 當前語言設定 - 預設為繁體中文
$currentLanguage = "zh-TW"

# 取得當前語言文字的函數
function Get-Text($key) {
    if ($languages[$currentLanguage][$key]) {
        return $languages[$currentLanguage][$key]
    }
    return $key  # 如果找不到翻譯，返回原始鍵值
}

# 更新界面語言的函數
function Update-Language {
    $form.Text = Get-Text "FormTitle"
    $leftLabel.Text = Get-Text "FileExplorer"
    $rightLabel.Text = Get-Text "SelectedFiles"
    $pathLabel.Text = Get-Text "CurrentPath"
    $browseButton.Text = Get-Text "Browse"
    $addButton.Text = Get-Text "AddFiles"
    $removeButton.Text = Get-Text "RemoveFiles"
    $clearButton.Text = Get-Text "ClearFiles"
    $mergeButton.Text = Get-Text "MergeFiles"
    
    # 更新語言 GroupBox 和 Radio Button 文字
    $languageGroupBox.Text = Get-Text "Language"
    $zhTWRadio.Text = Get-Text "ChineseTraditional"
    $enUSRadio.Text = Get-Text "English"
    
    # 更新 ListView 欄位標題
    $fileListView.Columns[0].Text = Get-Text "FileName"
    $fileListView.Columns[1].Text = Get-Text "FileType"
    $fileListView.Columns[2].Text = Get-Text "FileSize"
    $fileListView.Columns[3].Text = Get-Text "ModifiedDate"
    
    # 更新狀態標籤
    Update-StatusLabel
}

# 全域字體設定
$globalFontSize = 12  # 可以修改這個值來調整全域字體大小
$globalFont = New-Object System.Drawing.Font("微軟正黑體", $globalFontSize)

# 創建主表單
$form = New-Object System.Windows.Forms.Form
$form.Text = Get-Text "FormTitle"
$form.Size = New-Object System.Drawing.Size(1180, 700)
$form.StartPosition = "CenterScreen"
$form.Font = $globalFont

# 主要內容區域
$mainPanel = New-Object System.Windows.Forms.Panel
$mainPanel.Location = New-Object System.Drawing.Point(10, 10)
$mainPanel.Size = New-Object System.Drawing.Size(1100, 520)
$mainPanel.Font = $globalFont
$form.Controls.Add($mainPanel)

# 創建 TableLayoutPanel
$tableLayout = New-Object System.Windows.Forms.TableLayoutPanel
$tableLayout.Dock = "Fill"
$tableLayout.ColumnCount = 3
$tableLayout.RowCount = 2
$tableLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 65)))
$tableLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 110)))
$tableLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 30)))
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tableLayout.Font = $globalFont
$mainPanel.Controls.Add($tableLayout)

$leftLabel = New-Object System.Windows.Forms.Label
$leftLabel.Text = Get-Text "FileExplorer"
$leftLabel.Dock = "Fill"
$leftLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
$leftLabel.Font = $globalFont
$tableLayout.Controls.Add($leftLabel, 0, 0)

$rightLabel = New-Object System.Windows.Forms.Label
$rightLabel.Text = Get-Text "SelectedFiles"
$rightLabel.Dock = "Fill"
$rightLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
$rightLabel.Font = $globalFont
$tableLayout.Controls.Add($rightLabel, 2, 0)

# 左側檔案總管區域
$leftPanel = New-Object System.Windows.Forms.Panel
$leftPanel.Dock = "Fill"
$leftPanel.BorderStyle = "FixedSingle"
$leftPanel.Font = $globalFont
$tableLayout.Controls.Add($leftPanel, 0, 1)

# 創建分割面板（樹狀結構 + 檔案清單）
$splitContainer = New-Object System.Windows.Forms.SplitContainer
$splitContainer.Dock = "Fill"
$splitContainer.SplitterDistance = 50
$splitContainer.Orientation = "Vertical"
$splitContainer.Font = $globalFont
$leftPanel.Controls.Add($splitContainer)

# 樹狀結構（資料夾）
$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Dock = "Fill"
$treeView.Font = $globalFont
$splitContainer.Panel1.Controls.Add($treeView)

# 檔案清單（只顯示檔案，不顯示資料夾）
$fileListView = New-Object System.Windows.Forms.ListView
$fileListView.Dock = "Fill"
$fileListView.View = "Details"
$fileListView.FullRowSelect = $true
$fileListView.GridLines = $true
$fileListView.MultiSelect = $true
$fileListView.Sorting = "None"
$fileListView.Font = $globalFont
$fileListView.Columns.Add((Get-Text "FileName"), 350) | Out-Null
$fileListView.Columns.Add((Get-Text "FileType"), 80) | Out-Null
$fileListView.Columns.Add((Get-Text "FileSize"), 100) | Out-Null
$fileListView.Columns.Add((Get-Text "ModifiedDate"), 130) | Out-Null
$splitContainer.Panel2.Controls.Add($fileListView)

# 中間按鈕區域
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Dock = "Fill"
$buttonPanel.Font = $globalFont
$tableLayout.Controls.Add($buttonPanel, 1, 1)

$addButton = New-Object System.Windows.Forms.Button
$addButton.Text = Get-Text "AddFiles"
$addButton.Location = New-Object System.Drawing.Point(0, 100)
$addButton.Size = New-Object System.Drawing.Size(100, 40)
$addButton.Font = $globalFont
$buttonPanel.Controls.Add($addButton)

$removeButton = New-Object System.Windows.Forms.Button
$removeButton.Text = Get-Text "RemoveFiles"
$removeButton.Location = New-Object System.Drawing.Point(0, 160)
$removeButton.Size = New-Object System.Drawing.Size(100, 40)
$removeButton.Font = $globalFont
$buttonPanel.Controls.Add($removeButton)

$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Text = Get-Text "ClearFiles"
$clearButton.Location = New-Object System.Drawing.Point(0, 210)
$clearButton.Size = New-Object System.Drawing.Size(100, 40)
$clearButton.Font = $globalFont
$buttonPanel.Controls.Add($clearButton)

$mergeButton = New-Object System.Windows.Forms.Button
$mergeButton.Text = Get-Text "MergeFiles"
$mergeButton.Location = New-Object System.Drawing.Point(0, 340)
$mergeButton.Size = New-Object System.Drawing.Size(100, 40)
$mergeButton.Font = $globalFont
$buttonPanel.Controls.Add($mergeButton)

# 右側已選檔案區域
$rightPanel = New-Object System.Windows.Forms.Panel
$rightPanel.Dock = "Fill"
$rightPanel.BorderStyle = "FixedSingle"
$rightPanel.Font = $globalFont
$tableLayout.Controls.Add($rightPanel, 2, 1)

$rightListBox = New-Object System.Windows.Forms.ListBox
$rightListBox.Dock = "Fill"
$rightListBox.SelectionMode = "MultiExtended"
$rightListBox.Font = $globalFont
$rightPanel.Controls.Add($rightListBox)

# 底部按鈕區域
$bottomPanel = New-Object System.Windows.Forms.Panel
$bottomPanel.Location = New-Object System.Drawing.Point(10, 540)
$bottomPanel.Size = New-Object System.Drawing.Size(1100, 120)  # 從100改為120，增加高度
$bottomPanel.Font = $globalFont
$form.Controls.Add($bottomPanel)

# 路徑輸入區域
$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Text = Get-Text "CurrentPath"
$pathLabel.Location = New-Object System.Drawing.Point(10, 10)
$pathLabel.Size = New-Object System.Drawing.Size(80, 25)
$pathLabel.Font = $globalFont
$bottomPanel.Controls.Add($pathLabel)

# ComboBox
$pathComboBox = New-Object System.Windows.Forms.ComboBox
$pathComboBox.Location = New-Object System.Drawing.Point(95, 8)
$pathComboBox.Size = New-Object System.Drawing.Size(400, 25)
$pathComboBox.Font = $globalFont
$pathComboBox.DropDownStyle = "DropDown"
$pathComboBox.AutoCompleteMode = "SuggestAppend"
$pathComboBox.AutoCompleteSource = "FileSystemDirectories"
$bottomPanel.Controls.Add($pathComboBox)

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = Get-Text "Browse"
$browseButton.Location = New-Object System.Drawing.Point(505, 7)
$browseButton.Size = New-Object System.Drawing.Size(80, 27)
$browseButton.Font = $globalFont
$bottomPanel.Controls.Add($browseButton)

# 語言選擇 GroupBox
$languageGroupBox = New-Object System.Windows.Forms.GroupBox
$languageGroupBox.Text = Get-Text "Language"
$languageGroupBox.Location = New-Object System.Drawing.Point(10, 45)
$languageGroupBox.Size = New-Object System.Drawing.Size(250, 50)  # 從200改為250，增加寬度
$languageGroupBox.Font = $globalFont
$bottomPanel.Controls.Add($languageGroupBox)

# 繁體中文 Radio Button
$zhTWRadio = New-Object System.Windows.Forms.RadioButton
$zhTWRadio.Text = Get-Text "ChineseTraditional"
$zhTWRadio.Location = New-Object System.Drawing.Point(10, 20)
$zhTWRadio.Size = New-Object System.Drawing.Size(100, 20)  # 從80改為100，增加寬度
$zhTWRadio.Font = $globalFont
$zhTWRadio.Checked = $true  # 預設選中繁體中文
$languageGroupBox.Controls.Add($zhTWRadio)

# 英文 Radio Button
$enUSRadio = New-Object System.Windows.Forms.RadioButton
$enUSRadio.Text = Get-Text "English"
$enUSRadio.Location = New-Object System.Drawing.Point(120, 20)  # 從100改為120，向右移動
$enUSRadio.Size = New-Object System.Drawing.Size(80, 25)
$enUSRadio.Font = $globalFont
$languageGroupBox.Controls.Add($enUSRadio)

# Radio Button 事件處理
$zhTWRadio.Add_CheckedChanged({
    if ($zhTWRadio.Checked) {
        $script:currentLanguage = "zh-TW"
        Update-Language
    }
})

$enUSRadio.Add_CheckedChanged({
    if ($enUSRadio.Checked) {
        $script:currentLanguage = "en-US"
        Update-Language
    }
})
# 排序相關變數
$sortColumn = -1
$sortOrder = "Ascending"

# 儲存已選檔案的完整路徑資訊
$selectedFiles = @{}

# 路徑歷史記錄變數
$pathHistory = @()

# 載入磁碟機到樹狀結構
function Load-Drives {
    $treeView.Nodes.Clear()
    Get-WmiObject -Class Win32_LogicalDisk | ForEach-Object {
        $node = $treeView.Nodes.Add($_.DeviceID, "$($_.DeviceID) ($($_.VolumeName))")
        $node.Tag = $_.DeviceID + "\"
        $node.Nodes.Add("") | Out-Null
    }
}

# 載入子資料夾
function Load-SubFolders($node, $path) {
    try {
        $node.Nodes.Clear()
        Get-ChildItem -Path $path -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $subNode = $node.Nodes.Add($_.Name, $_.Name)
            $subNode.Tag = $_.FullName
            if (Get-ChildItem -Path $_.FullName -Directory -ErrorAction SilentlyContinue) {
                $subNode.Nodes.Add("") | Out-Null
            }
        }
    }
    catch {}
}

# 載入檔案（只載入檔案，不載入資料夾）
function Load-FilesOnly($path) {
    try {
        $fileListView.Items.Clear()
        
        $files = Get-ChildItem -Path $path -File -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            $item = $fileListView.Items.Add($file.Name)
            $extension = if ($file.Extension) { $file.Extension } else { Get-Text "File" }
            $item.SubItems.Add($extension) | Out-Null
            
            if ($file.Length -gt 0) {
                $sizeKB = [math]::Round($file.Length / 1KB, 2)
                $item.SubItems.Add("$sizeKB KB") | Out-Null
            } else {
                $item.SubItems.Add("0 KB") | Out-Null
            }
            
            $item.SubItems.Add($file.LastWriteTime.ToString("yyyy/MM/dd HH:mm")) | Out-Null
            $item.SubItems.Add($file.Length.ToString()) | Out-Null
            $item.SubItems.Add($file.LastWriteTime.Ticks.ToString()) | Out-Null
            
            $item.Tag = $file.FullName
        }
    }
    catch {
        # 忽略錯誤
    }
}

# 排序函數
function Sort-ListView($column) {
    if ($script:sortColumn -eq $column) {
        $script:sortOrder = if ($script:sortOrder -eq "Ascending") { "Descending" } else { "Ascending" }
    } else {
        $script:sortColumn = $column
        $script:sortOrder = "Ascending"
    }
    
    $items = @()
    for ($i = 0; $i -lt $fileListView.Items.Count; $i++) {
        $items += $fileListView.Items[$i]
    }
    
    $fileListView.Items.Clear()
    
    switch ($column) {
        0 {
            if ($script:sortOrder -eq "Ascending") {
                $sortedItems = $items | Sort-Object { $_.Text }
            } else {
                $sortedItems = $items | Sort-Object { $_.Text } -Descending
            }
        }
        1 {
            if ($script:sortOrder -eq "Ascending") {
                $sortedItems = $items | Sort-Object { $_.SubItems[1].Text }
            } else {
                $sortedItems = $items | Sort-Object { $_.SubItems[1].Text } -Descending
            }
        }
        2 {
            if ($script:sortOrder -eq "Ascending") {
                $sortedItems = $items | Sort-Object { [long]$_.SubItems[4].Text }
            } else {
                $sortedItems = $items | Sort-Object { [long]$_.SubItems[4].Text } -Descending
            }
        }
        3 {
            if ($script:sortOrder -eq "Ascending") {
                $sortedItems = $items | Sort-Object { [long]$_.SubItems[5].Text }
            } else {
                $sortedItems = $items | Sort-Object { [long]$_.SubItems[5].Text } -Descending
            }
        }
    }
    
    foreach ($item in $sortedItems) {
        $fileListView.Items.Add($item)
    }
}

# 更新路徑ComboBox的函數
function Update-PathComboBox($path) {
    $pathComboBox.Text = $path
    
    if ($pathHistory -notcontains $path) {
        $pathHistory += $path
        $pathComboBox.Items.Add($path)
        
        if ($pathHistory.Count -gt 20) {
            $oldestPath = $pathHistory[0]
            $pathHistory = $pathHistory[1..($pathHistory.Count - 1)]
            $pathComboBox.Items.Remove($oldestPath)
        }
    }
}

# 初始化常用路徑
function Initialize-CommonPaths {
    $commonPaths = @(
        [Environment]::GetFolderPath('Desktop'),
        [Environment]::GetFolderPath('MyDocuments'),
        [Environment]::GetFolderPath('MyPictures'),
        [Environment]::GetFolderPath('MyMusic'),
        [Environment]::GetFolderPath('MyVideos'),
        "C:\",
        $PSScriptRoot
    )
    
    foreach ($path in $commonPaths) {
        if ($path -and (Test-Path $path)) {
            $pathHistory += $path
            $pathComboBox.Items.Add($path)
        }
    }
}

# 更新已選檔案數量顯示（狀態標籤功能）
function Update-StatusLabel {
    $count = $rightListBox.Items.Count
    if ($count -eq 0) {
        # 可以在這裡顯示狀態，目前沒有狀態標籤
    } else {
        # 可以顯示選中的檔案數量
    }
}

# 導航並選擇腳本所在的路徑
function Select-NodeByPath($path) {
    if (-not $path -or -not (Test-Path $path -PathType Container)) {
        return
    }
    
    $targetPath = $path.TrimEnd('\')
    $pathParts = $targetPath.Split([System.IO.Path]::DirectorySeparatorChar) | Where-Object { $_ }

    $currentNode = $null
    
    foreach ($node in $treeView.Nodes) {
        $nodeDrive = $node.Tag.TrimEnd('\')
        if ($targetPath.StartsWith($nodeDrive, [System.StringComparison]::OrdinalIgnoreCase)) {
            $currentNode = $node
            $currentNode.Expand()
            break
        }
    }

    if ($currentNode) {
        if ($targetPath.Equals($currentNode.Tag.TrimEnd('\'), [System.StringComparison]::OrdinalIgnoreCase)) {
            $treeView.SelectedNode = $currentNode
            return
        }
        
        $relativePath = $targetPath.Substring($currentNode.Tag.Length)
        $relativeParts = $relativePath.Split([System.IO.Path]::DirectorySeparatorChar) | Where-Object { $_ }
        
        foreach ($part in $relativeParts) {
            if ($currentNode.Nodes.Count -eq 1 -and $currentNode.Nodes[0].Text -eq "") {
                Load-SubFolders $currentNode $currentNode.Tag
            }
            
            $foundChild = $false
            foreach ($childNode in $currentNode.Nodes) {
                if ($childNode.Text.Equals($part, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $currentNode = $childNode
                    $currentNode.Expand()
                    $foundChild = $true
                    break
                }
            }
            
            if (-not $foundChild) {
                break
            }
        }
        
        $treeView.SelectedNode = $currentNode
    }
}

# 事件處理
$treeView.add_BeforeExpand({
    param($sender, $e)
    if ($e.Node.Nodes.Count -eq 1 -and $e.Node.Nodes[0].Text -eq "") {
        Load-SubFolders $e.Node $e.Node.Tag
    }
})

$treeView.add_AfterSelect({
    param($sender, $e)
    if ($e.Node.Tag) {
        Load-FilesOnly $e.Node.Tag
        Update-PathComboBox $e.Node.Tag
        Update-StatusLabel
    }
})

$fileListView.add_ColumnClick({
    param($sender, $e)
    Sort-ListView $e.Column
})

$pathComboBox.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq "Enter") {
        $inputPath = $pathComboBox.Text.Trim()
        if ($inputPath -and (Test-Path $inputPath -PathType Container)) {
            Select-NodeByPath $inputPath
            Load-FilesOnly $inputPath
            Update-PathComboBox $inputPath
            Update-StatusLabel
        } else {
            [System.Windows.Forms.MessageBox]::Show((Get-Text "PathNotExists") + $inputPath, (Get-Text "Error"))
            if ($treeView.SelectedNode -and $treeView.SelectedNode.Tag) {
                $pathComboBox.Text = $treeView.SelectedNode.Tag
            }
        }
    }
})

$pathComboBox.Add_SelectedIndexChanged({
    $selectedPath = $pathComboBox.SelectedItem
    if ($selectedPath -and (Test-Path $selectedPath -PathType Container)) {
        Select-NodeByPath $selectedPath
        Load-FilesOnly $selectedPath
        Update-StatusLabel
    }
})

$pathComboBox.Add_Leave({
    $inputPath = $pathComboBox.Text.Trim()
    if ($inputPath -and (Test-Path $inputPath -PathType Container)) {
        if ($treeView.SelectedNode -eq $null -or $treeView.SelectedNode.Tag -ne $inputPath) {
            Select-NodeByPath $inputPath
            Load-FilesOnly $inputPath
            Update-PathComboBox $inputPath
            Update-StatusLabel
        }
    } else {
        if ($treeView.SelectedNode -and $treeView.SelectedNode.Tag) {
            $pathComboBox.Text = $treeView.SelectedNode.Tag
        }
    }
})

$browseButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = Get-Text "SelectFolder"
    if ($pathComboBox.Text -and (Test-Path $pathComboBox.Text)) {
        $folderBrowser.SelectedPath = $pathComboBox.Text
    }
    
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $selectedPath = $folderBrowser.SelectedPath
        Select-NodeByPath $selectedPath
        Load-FilesOnly $selectedPath
        Update-PathComboBox $selectedPath
        Update-StatusLabel
    }
})

$addButton.Add_Click({
    foreach ($item in $fileListView.SelectedItems) {
        $fileName = $item.Text
        $fullPath = $item.Tag
        
        if (-not $selectedFiles.ContainsKey($fullPath)) {
            $rightListBox.Items.Add($fileName)
            $selectedFiles[$fullPath] = $fileName
        }
    }
    Update-StatusLabel
})

$removeButton.Add_Click({
    $selectedIndices = @($rightListBox.SelectedIndices)
    for ($i = $selectedIndices.Count - 1; $i -ge 0; $i--) {
        $index = $selectedIndices[$i]
        $fileName = $rightListBox.Items[$index]
        
        $pathToRemove = $null
        foreach ($path in $selectedFiles.Keys) {
            if ($selectedFiles[$path] -eq $fileName) {
                $pathToRemove = $path
                break
            }
        }
        if ($pathToRemove) {
            $selectedFiles.Remove($pathToRemove)
        }
        
        $rightListBox.Items.RemoveAt($index)
    }
    Update-StatusLabel
})

$clearButton.Add_Click({
    $rightListBox.Items.Clear()
    $selectedFiles.Clear()
    Update-StatusLabel
})

$mergeButton.Add_Click({
    if ($rightListBox.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show((Get-Text "PleaseSelectFiles"), (Get-Text "Hint"))
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = (Get-Text "TextFiles") + "|" + (Get-Text "AllFiles")
    $saveFileDialog.Title = Get-Text "SelectSaveLocation"
    $saveFileDialog.FileName = "merged.txt"
    
    if ($saveFileDialog.ShowDialog() -eq "OK") {
        $outputPath = $saveFileDialog.FileName
        
        try {
            Clear-Content -Path $outputPath -ErrorAction SilentlyContinue

            foreach ($fullPath in $selectedFiles.Keys) {
                if (Test-Path $fullPath) {
                    $form.Refresh()
                    try {
                        Get-Content $fullPath -Encoding Default | Add-Content $outputPath -Encoding Default
                    }
                    catch {
                        (Get-Text "CannotReadFile") | Add-Content $outputPath -Encoding UTF8
                    }
                }
            }
            
            [System.Windows.Forms.MessageBox]::Show((Get-Text "MergeComplete") + " $outputPath", (Get-Text "Complete"))
            Update-StatusLabel
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show((Get-Text "MergeError") + " $($_.Exception.Message)", (Get-Text "Error"))
        }
    }
})

# 初始化
Load-Drives
Initialize-CommonPaths
Select-NodeByPath $PSScriptRoot
if ($treeView.SelectedNode -and $treeView.SelectedNode.Tag) {
    Update-PathComboBox $treeView.SelectedNode.Tag
}
Update-StatusLabel

# 初始化語言設定
Update-Language

# 顯示表單
$form.ShowDialog()