Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# XAML 定義
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="OfflineList DAT CSV Converter" Height="520" Width="720"
        WindowStartupLocation="CenterScreen"
        AllowDrop="True">
    <Grid Margin="8">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>   <!-- 0: 入力ファイル -->
            <RowDefinition Height="Auto"/>   <!-- 1: 出力フォルダ -->
            <RowDefinition Height="Auto"/>   <!-- 2: configuration -->
            <RowDefinition Height="Auto"/>   <!-- 3: newDat -->
            <RowDefinition Height="*"/>      <!-- 4: ログ -->
            <RowDefinition Height="Auto"/>   <!-- 5: 出力形式 + 実行ボタン -->
        </Grid.RowDefinitions>

        <!-- 入力ファイル -->
        <GroupBox Header="入力ファイル" Grid.Row="0" Margin="0,0,0,6">
            <Grid Margin="6">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Text="ファイル:" VerticalAlignment="Center" Margin="0,0,4,0"/>
                <TextBox x:Name="TxtInputFile" Grid.Column="1" Margin="0,0,4,0"/>
                <Button x:Name="BtnSelectFile" Grid.Column="2" Content="参照..." Width="80"/>
            </Grid>
        </GroupBox>

        <!-- 出力フォルダ -->
        <GroupBox Header="出力フォルダ" Grid.Row="1" Margin="0,0,0,6">
            <Grid Margin="6">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox x:Name="TxtOutputFolder" Grid.Column="0" Margin="0,0,4,0"/>
                <Button x:Name="BtnSelectOutputFolder" Grid.Column="1" Content="参照..." Width="80"/>
            </Grid>
        </GroupBox>

        <!-- configuration（screenshots を 1 行目に移動） -->
        <GroupBox Header="configuration" Grid.Row="2" Margin="0,0,0,6">
            <StackPanel Margin="6">

                <!-- 行1：screenshots + チェックボックス2つ -->
                <Grid Margin="0,0,0,4">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <StackPanel Grid.Column="0" Orientation="Horizontal">
                        <TextBlock Text="Screenshots (W×H):" VerticalAlignment="Center" Margin="0,0,4,0"/>
                        <TextBox x:Name="TxtScreenshotsWidth" Width="60" Margin="0,0,4,0"/>
                        <TextBlock Text="×" VerticalAlignment="Center" Margin="0,0,4,0"/>
                        <TextBox x:Name="TxtScreenshotsHeight" Width="60"/>
                    </StackPanel>

                    <CheckBox x:Name="ChkVersionFile"
                              Grid.Column="1"
                              Margin="12,0,4,0"
                              Content="バージョンファイル出力"/>

                    <CheckBox x:Name="ChkUseDatFileName"
                              Grid.Column="2"
                              Content="出力ファイル名を DATファイル に合わせる"/>
                </Grid>

                <!-- 行2：datName / romTitle -->
                <Grid Margin="0,0,0,4">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <StackPanel Grid.Column="0" Orientation="Horizontal">
                        <TextBlock Text="datName:" VerticalAlignment="Center" Margin="0,0,4,0"/>
                        <TextBox x:Name="TxtDatName" Width="180"/>
                    </StackPanel>

                    <StackPanel Grid.Column="1" Orientation="Horizontal">
                        <TextBlock Text="romTitle:" VerticalAlignment="Center" Margin="12,0,4,0"/>
                        <TextBox x:Name="TxtRomTitle" Width="180" Margin="0,0,4,0" Text="%n"/>
                        <Button x:Name="BtnEditRomTitle" Content="編集..." Width="60"/>
                    </StackPanel>
                </Grid>

                <!-- 行3：system / imFolder / ROM拡張子 -->
                <Grid Margin="0,0,0,4">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <StackPanel Grid.Column="0" Orientation="Horizontal">
                        <TextBlock Text="system:" VerticalAlignment="Center" Margin="0,0,4,0"/>
                        <TextBox x:Name="TxtSystem" Width="140"/>
                    </StackPanel>

                    <StackPanel Grid.Column="1" Orientation="Horizontal">
                        <TextBlock Text="imFolder:" VerticalAlignment="Center" Margin="12,0,4,0"/>
                        <TextBox x:Name="TxtImFolder" Width="140" IsReadOnly="True" Background="#EEE"/>
                    </StackPanel>

                    <StackPanel Grid.Column="2" Orientation="Horizontal">
                        <TextBlock Text="ROM拡張子:" VerticalAlignment="Center" Margin="12,0,4,0"/>
                        <TextBox x:Name="TxtCanOpenExt" Width="140"/>
                    </StackPanel>
                </Grid>

            </StackPanel>
        </GroupBox>

        <!-- newDat（Header に baseURL を配置） -->
        <GroupBox Grid.Row="3" Margin="0,0,0,6">
            <GroupBox.Header>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="newDat" Margin="0,0,12,0"/>
                    <TextBlock Text="baseURL:" VerticalAlignment="Center" Margin="0,0,4,0"/>
                    <TextBox x:Name="TxtBaseUrl" Width="260" Margin="0,0,4,0"/>
                    <TextBlock Text="/" VerticalAlignment="Center"/>
                </StackPanel>
            </GroupBox.Header>

            <Grid Margin="6">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0" Orientation="Horizontal">
                    <TextBlock Text="バージョンファイル:" VerticalAlignment="Center" Margin="0,0,4,0"/>
                    <TextBox x:Name="TxtDatVersionUrl" Width="120" IsReadOnly="True" Background="#EEE"/>
                </StackPanel>

                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <TextBlock Text="DATファイル:" VerticalAlignment="Center" Margin="12,0,4,0"/>
                    <TextBox x:Name="TxtDatUrlFileName" Width="120" IsReadOnly="True" Background="#EEE"/>
                </StackPanel>

                <StackPanel Grid.Column="2" Orientation="Horizontal">
                    <TextBlock Text="サムネイルフォルダ:" VerticalAlignment="Center" Margin="12,0,4,0"/>
                    <TextBox x:Name="TxtImUrl" Width="120" IsReadOnly="True" Background="#EEE"/>
                </StackPanel>
            </Grid>
        </GroupBox>

        <!-- ログ -->
        <GroupBox Header="ログ" Grid.Row="4" Margin="0,0,0,6">
            <Grid Margin="6">
                <TextBox x:Name="TxtLog" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto"/>
            </Grid>
        </GroupBox>

        <!-- 出力形式 + 実行ボタン（右詰め） -->
        <Grid Grid.Row="5">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>

            <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center" Margin="0,0,8,0">
                <TextBlock Text="出力形式:" VerticalAlignment="Center" Margin="0,0,4,0"/>
                <ComboBox x:Name="CmbOutputKindBottom" Width="140" SelectedIndex="0">
                    <ComboBoxItem Content="XML" Tag="Xml"/>
                    <ComboBoxItem Content="ZIP" Tag="Zip"/>
                </ComboBox>
            </StackPanel>

            <StackPanel Grid.Column="2" Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="BtnConvert" Content="変換実行" Width="100" Margin="0,0,8,0"/>
                <Button x:Name="BtnClose" Content="閉じる" Width="80"/>
            </StackPanel>
        </Grid>

    </Grid>
</Window>
"@

function Show-ErrorMessage($msg) {
    [System.Windows.MessageBox]::Show($msg, "エラー", "OK", "Error") | Out-Null
}

function Write-Log([string]$msg) {
    $time = (Get-Date).ToString("HH:mm:ss")
    $TxtLog.AppendText("[$time] $msg`r`n")
    $TxtLog.ScrollToEnd()
}

# XAML パース
$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# コントロール取得
$BtnSelectFile      = $window.FindName("BtnSelectFile")
$TxtInputFile       = $window.FindName("TxtInputFile")
$CmbOutputKindBottom      = $window.FindName("CmbOutputKindBottom")
$ChkVersionFile     = $window.FindName("ChkVersionFile")
$ChkUseDatFileName  = $window.FindName("ChkUseDatFileName")
$TxtDatName         = $window.FindName("TxtDatName")
$TxtSystem          = $window.FindName("TxtSystem")
$TxtCanOpenExt      = $window.FindName("TxtCanOpenExt")
$TxtBaseUrl         = $window.FindName("TxtBaseUrl")
$TxtScreenshotsWidth  = $window.FindName("TxtScreenshotsWidth")
$TxtScreenshotsHeight = $window.FindName("TxtScreenshotsHeight")
$TxtRomTitle        = $window.FindName("TxtRomTitle")
$BtnEditRomTitle    = $window.FindName("BtnEditRomTitle")
$TxtDatVersionUrl   = $window.FindName("TxtDatVersionUrl")
$TxtDatUrlFileName  = $window.FindName("TxtDatUrlFileName")
$TxtDatUrl          = $window.FindName("TxtDatUrl")
$TxtImUrl           = $window.FindName("TxtImUrl")
$TxtImFolder       = $window.FindName("TxtImFolder")
$TxtLog             = $window.FindName("TxtLog")
$BtnConvert         = $window.FindName("BtnConvert")
$BtnClose           = $window.FindName("BtnClose")

# --- system 入力制御（半角英数字 . _ - のみ） ---
$regexSystem = '^[A-Za-z0-9._-]+$'

# キー入力制御
$TxtSystem.Add_PreviewTextInput({
    if ($_.Text -notmatch $regexSystem) { $_.Handled = $true }
})

# 貼り付けイベント
$pasteEvent = [Windows.DataObject]::PastingEvent

$onPasteSystem = {
    $text = [Windows.Clipboard]::GetText()
    if ($text -notmatch $regexSystem) {
        $_.Handled = $true
    }
}

# --- baseURL 入力制御（ASCII 全記号許可） ---
$regexBaseUrl = '^[\x21-\x7E]+$'   # 全角禁止、ASCII のみ

# キー入力制御
$TxtBaseUrl.Add_PreviewTextInput({
    if ($_.Text -notmatch $regexBaseUrl) { $_.Handled = $true }
})

# 貼り付けイベント
$onPasteBaseUrl = {
    $text = [Windows.Clipboard]::GetText()
    if ($text -notmatch $regexBaseUrl) {
        $_.Handled = $true
    }
}

$TxtBaseUrl.AddHandler(
    $pasteEvent,
    [Windows.DataObjectPastingEventHandler]$onPasteBaseUrl
)

$TxtSystem.AddHandler(
    $pasteEvent,
    [Windows.DataObjectPastingEventHandler]$onPasteSystem
)

# 7z 検出
$Global:SevenZipPath = $null
if (Get-Command "7za.exe" -ErrorAction SilentlyContinue) {
    $Global:SevenZipPath = "7za.exe"
    Write-Log "7za.exe を使用します（ZIP 出力および 7z 解凍に対応）"
}
elseif (Get-Command "7z.exe" -ErrorAction SilentlyContinue) {
    $Global:SevenZipPath = "7z.exe"
    Write-Log "7z.exe を使用します（ZIP 出力および 7z 解凍に対応）"
}
else {
    Write-Log "7za.exe / 7z.exe が見つかりません。ZIP 出力は無効になります。"
}

# 初期値
$TxtScreenshotsWidth.Text  = "320"
$TxtScreenshotsHeight.Text = "224"

function Get-BaseNameFromPath($path) {
    if (-not $path) { return "" }
    return [System.IO.Path]::GetFileNameWithoutExtension($path)
}

function Normalize-BaseUrl($url) {
    if ([string]::IsNullOrWhiteSpace($url)) { return "" }
    if ($url.EndsWith("/")) { return $url }
    return $url + "/"
}

function Read-XmlWithDeclaredEncoding([string]$path) {
    # 先頭数行だけ生バイトで読む
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $textProbe = [System.Text.Encoding]::ASCII.GetString($bytes, 0, [Math]::Min($bytes.Length, 512))

    $encName = "UTF-8"  # デフォルト
    if ($textProbe -match 'encoding\s*=\s*"(.*?)"') {
        $encName = $matches[1]
    }

    try {
        $enc = [System.Text.Encoding]::GetEncoding($encName)
    } catch {
        $enc = [System.Text.Encoding]::UTF8
    }

    $sr = New-Object System.IO.StreamReader($path, $enc, $true)
    $xmlText = $sr.ReadToEnd()
    $sr.Close()

    return [xml]$xmlText
}

function Update-NewDatFields {

    $system  = $TxtSystem.Text
    $baseUrl = $TxtBaseUrl.Text

    # baseURL の末尾に / を補完（UI には反映しない）
    if ($baseUrl -and -not $baseUrl.EndsWith("/")) {
        $baseUrl = $baseUrl + "/"
    }

    if ($system) {

        # --- configuration 側の imFolder（system + "img"） ---
        $TxtImFolder.Text = "${system}img"

        # --- newDat 側のバージョンファイル（ファイル名のみ） ---
        $TxtDatVersionUrl.Text = "$system.txt"

        # --- newDat 側の DATファイル（ファイル名のみ） ---
        $TxtDatUrlFileName.Text = "$system.zip"

        # --- newDat 側の imFolder（configuration とリンク） ---
        $TxtImUrl.Text = $TxtImFolder.Text
    }
    else {
        $TxtImFolder.Text      = ""
        $TxtDatVersionUrl.Text = ""
        $TxtDatUrlFileName.Text = ""
        $TxtImUrl.Text         = ""
    }
}

# system / baseURL 変更時に newDat 更新
$TxtSystem.Add_TextChanged({ Update-NewDatFields })
$TxtBaseUrl.Add_TextChanged({ Update-NewDatFields })

function Set-From-InputFile($path) {
    $TxtInputFile.Text = $path
    $baseName = Get-BaseNameFromPath $path

    if ($baseName) {
        # datName は全角OK
        $TxtDatName.Text = $baseName

        # system は半角英数字 . _ - のみ許可
        if ($baseName -match '^[A-Za-z0-9._-]+$') {
            $TxtSystem.Text = $baseName
        } else {
            $TxtSystem.Text = ""
        }
    }

    Update-NewDatFields
}

# ファイル選択
$BtnSelectFile.Add_Click({
    $ofd = New-Object Microsoft.Win32.OpenFileDialog
    if ($Global:SevenZipPath) {
        $ofd.Filter = "対応ファイル|*.xml;*.csv;*.zip"
    } else {
        $ofd.Filter = "対応ファイル|*.xml;*.csv"
    }
    if ($ofd.ShowDialog()) {
        Set-From-InputFile $ofd.FileName
    }
})

# D&D
$window.Add_DragOver({
    if ($_.Data.GetDataPresent([Windows.DataFormats]::FileDrop)) {
        $_.Effects = [Windows.DragDropEffects]::Copy
    } else {
        $_.Effects = [Windows.DragDropEffects]::None
    }
    $_.Handled = $true
})
$window.Add_Drop({
    $files = $_.Data.GetData([Windows.DataFormats]::FileDrop)
    if ($files -and $files.Count -gt 0) {
        $file = $files[0]
        $ext = [System.IO.Path]::GetExtension($file).ToLowerInvariant()
        $allowed = @(".xml",".csv")
        if ($Global:SevenZipPath) {
            $allowed += ".zip",".7z"
        }
        if ($ext -in $allowed) {
            Set-From-InputFile $file
        } else {
            Show-ErrorMessage "対応していない拡張子です: $ext"
        }
    }
})

$BtnSelectOutputFolder = $window.FindName("BtnSelectOutputFolder")
$TxtOutputFolder       = $window.FindName("TxtOutputFolder")

Add-Type -AssemblyName System.Windows.Forms

$BtnSelectOutputFolder.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "出力フォルダを選択してください"

    if ($dialog.ShowDialog() -eq "OK") {
        $TxtOutputFolder.Text = $dialog.SelectedPath
    }
})

# romTitle 編集ウィンドウ
$BtnEditRomTitle.Add_Click({
    $rtWindow = New-Object Windows.Window
    $rtWindow.Title = "romTitle 編集"
    $rtWindow.Width = 360
    $rtWindow.Height = 260
    $rtWindow.WindowStartupLocation = "CenterOwner"
    $rtWindow.Owner = $window

    # × を押したらキャンセル扱いにする
    $rtWindow.WindowStyle = 'SingleBorderWindow'
    $rtWindow.ShowInTaskbar = $false
    $rtWindow.ResizeMode = 'NoResize'
    $rtWindow.Add_Closing({
        if (-not $script:rtConfirmed) {
            $script:rtCanceled = $true
        }
    })

    $grid = New-Object Windows.Controls.Grid
    $grid.Margin = '8'
    $grid.RowDefinitions.Add((New-Object Windows.Controls.RowDefinition -Property @{Height="*"}))
    $grid.RowDefinitions.Add((New-Object Windows.Controls.RowDefinition -Property @{Height="Auto"}))
    $grid.RowDefinitions.Add((New-Object Windows.Controls.RowDefinition -Property @{Height="Auto"}))

    # 編集テキストボックス
    $txt = New-Object Windows.Controls.TextBox
    $txt.AcceptsReturn = $false
    $txt.Text = $TxtRomTitle.Text
    $grid.Children.Add($txt)

    # コード挿入ボタン群
    $panel = New-Object Windows.Controls.WrapPanel
    $panel.Margin = '0,8,0,0'
    [Windows.Controls.Grid]::SetRow($panel,1)
    $grid.Children.Add($panel)

    $buttons = @(
        @{Label="Title"; Code="%n"},
        @{Label="Publisher"; Code="%p"},
        @{Label="Source"; Code="%g"},
        @{Label="Comment"; Code="%e"},
        @{Label="Language"; Code="%a"},
        @{Label="Size"; Code="%i"},
        @{Label="Save"; Code="%s"},
        @{Label="Location"; Code="%o"},
        @{Label="Rom number"; Code="%u"},
        @{Label="CRC"; Code="%c"}
    )

foreach ($b in $buttons) {

    $label = $b.Label
    $code  = $b.Code

    $btn = New-Object Windows.Controls.Button
    $btn.Content = $label
    $btn.Margin = '2'

    # ★ スクリプトブロックを「その場で生成」して code を固定する
    $btn.Add_Click([scriptblock]::Create("
        `$txt.Text += '$code'
        `$txt.CaretIndex = `$txt.Text.Length
    "))

    $panel.Children.Add($btn)
}

    # OK / Cancel ボタン
    $btnPanel = New-Object Windows.Controls.StackPanel
    $btnPanel.Orientation = 'Horizontal'
    $btnPanel.HorizontalAlignment = 'Right'
    $btnPanel.Margin = '0,8,0,0'
    [Windows.Controls.Grid]::SetRow($btnPanel,2)
    $grid.Children.Add($btnPanel)

    $btnOK = New-Object Windows.Controls.Button
    $btnOK.Content = "OK"
    $btnOK.Width = 80
    $btnOK.Margin = "0,0,8,0"
    $btnOK.Add_Click({
        $script:rtConfirmed = $true
        $rtWindow.Close()
    })
    $btnPanel.Children.Add($btnOK)

    $btnCancel = New-Object Windows.Controls.Button
    $btnCancel.Content = "キャンセル"
    $btnCancel.Width = 80
    $btnCancel.Add_Click({
        $script:rtCanceled = $true
        $rtWindow.Close()
    })
    $btnPanel.Children.Add($btnCancel)

    $rtWindow.Content = $grid

    # 初期状態
    $script:rtConfirmed = $false
    $script:rtCanceled = $false

    $rtWindow.ShowDialog() | Out-Null

    # OK のときだけ反映
    if ($script:rtConfirmed) {
        $TxtRomTitle.Text = $txt.Text.Trim()
    }
})

function Ensure-CanOpenExt {
    $ext = $TxtCanOpenExt.Text.Trim()
    if ($ext -and -not $ext.StartsWith(".")) {
        $TxtCanOpenExt.Text = "." + $ext
    }
}

function Build-ConfigurationXml([string]$datVersion) {
    Ensure-CanOpenExt
    $datName = $TxtDatName.Text
    $system  = $TxtSystem.Text
    $canExt  = $TxtCanOpenExt.Text
    $baseUrl = Normalize-BaseUrl $TxtBaseUrl.Text
    $romTitle = $TxtRomTitle.Text
    $w = $TxtScreenshotsWidth.Text
    $h = $TxtScreenshotsHeight.Text

    # エスケープ
    $datNameEsc = [System.Security.SecurityElement]::Escape($datName)
    $systemEsc  = [System.Security.SecurityElement]::Escape($system)
    $canExtEsc  = [System.Security.SecurityElement]::Escape($canExt)
    $romTitleEsc = [System.Security.SecurityElement]::Escape($romTitle)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("  <configuration>")
    [void]$sb.AppendLine("    <datName>$datNameEsc</datName>")
    [void]$sb.AppendLine("    <datVersion>$datVersion</datVersion>")
    [void]$sb.AppendLine("    <system>$systemEsc</system>")
    # configuration の imFolder を出力
    $imFolderConf = $TxtImFolder.Text
    if ($imFolderConf) {
        $imConfEsc = [System.Security.SecurityElement]::Escape($imFolderConf)
        [void]$sb.AppendLine("    <imFolder>$imConfEsc</imFolder>")
    } else {
        [void]$sb.AppendLine("    <imFolder/>")
    }
    # screenshotsWidth / screenshotsHeight 出力
    $w = $TxtScreenshotsWidth.Text
    $h = $TxtScreenshotsHeight.Text

    if ($w) {
        [void]$sb.AppendLine("    <screenshotsWidth>$w</screenshotsWidth>")
    } else {
        [void]$sb.AppendLine("    <screenshotsWidth/>")
    }

    if ($h) {
        [void]$sb.AppendLine("    <screenshotsHeight>$h</screenshotsHeight>")
    } else {
        [void]$sb.AppendLine("    <screenshotsHeight/>")
    }

    [void]$sb.AppendLine("    <infos>")
    $infos = @(
        "releaseNumber",
        "imageNumber",
        "im1CRC",
        "im2CRC",
        "title",
        "publisher",
        "sourceRom",
        "location",
        "languages",
        "comment",
        "romSize",
        "romCRC",
        "saveType"
    )
    foreach ($name in $infos) {
        [void]$sb.AppendLine("      <${name} visible=""true"" inNamingOption=""true"" default=""true"" />")
    }
    [void]$sb.AppendLine("    </infos>")
    [void]$sb.AppendLine("    <canOpen>")
    [void]$sb.AppendLine("      <extension>$canExtEsc</extension>")
    [void]$sb.AppendLine("    </canOpen>")
    [void]$sb.AppendLine("    <newDat>")

    $systemRaw = $TxtSystem.Text
    $fileNameAttr = $TxtDatUrlFileName.Text
    if ([string]::IsNullOrWhiteSpace($fileNameAttr)) {
        $fileNameAttr = "$systemRaw.zip"
    }

    # baseURL（Normalize-BaseUrl で末尾 / 保証済み）
    $baseUrl = Normalize-BaseUrl $TxtBaseUrl.Text
    
    # newDat の値（UI にはファイル名だけ入っている）
    $datVersionFile = $TxtDatVersionUrl.Text
    $datFile        = $TxtDatUrlFileName.Text
    $imFolder       = $TxtImFolder.Text

    # XML に書き込む完全 URL を生成
    $datVersionUrl = if ($datVersionFile) { $baseUrl + $datVersionFile } else { "" }
    $datUrl        = if ($datFile)        { $baseUrl + $datFile }        else { "" }
    $imUrl         = if ($imFolder)       { $baseUrl + $imFolder }       else { "" }

    if ($datVersionUrl) {
        $dvEsc = [System.Security.SecurityElement]::Escape($datVersionUrl)
        [void]$sb.AppendLine("      <datVersionURL>$dvEsc</datVersionURL>")
    } else {
        [void]$sb.AppendLine("      <datVersionURL/>")
    }

    $fnEsc = [System.Security.SecurityElement]::Escape($datFile)
    if ($datUrl) {
        $duEsc = [System.Security.SecurityElement]::Escape($datUrl)
        [void]$sb.AppendLine("      <datURL fileName=""$fnEsc"">$duEsc</datURL>")
    } else {
        [void]$sb.AppendLine("      <datURL fileName=""$fnEsc""/>")
    }

    if ($imUrl) {
        $imEsc = [System.Security.SecurityElement]::Escape($imUrl)
        [void]$sb.AppendLine("      <imURL>$imEsc</imURL>")
    } else {
        [void]$sb.AppendLine("      <imURL/>")
    }

    [void]$sb.AppendLine("    </newDat>")
    [void]$sb.AppendLine("    <search>")
    $searchTargets = @("publisher","sourceRom","saveType","location","languages")
    foreach ($v in $searchTargets) {
        [void]$sb.AppendLine("      <to value=""$v"" default=""true"" auto=""true"" />")
    }
    [void]$sb.AppendLine("    </search>")
    [void]$sb.AppendLine("    <romTitle>$romTitleEsc</romTitle>")
    [void]$sb.AppendLine("  </configuration>")
    return $sb.ToString()
}

function Csv-ToXml([string]$csvPath, [string]$outKind) {
    Write-Log "CSV → XML 変換開始: $csvPath"

    $rows = Import-Csv -Path $csvPath -Delimiter ','

    if (-not $rows -or $rows.Count -eq 0) {
        throw "CSV にデータがありません。"
    }

    $headers = $rows[0].PSObject.Properties.Name
    if (-not $headers -or $headers.Count -eq 0) {
        throw "CSV の 1 行目がヘッダーとして認識できません。"
    }

    $datVersion = (Get-Date).ToString("yyyyMMddHHmm")

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="no"?>')
    [void]$sb.AppendLine('<dat xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="datas.xsd">')
    [void]$sb.Append((Build-ConfigurationXml $datVersion))
    [void]$sb.AppendLine("  <games>")

    foreach ($row in $rows) {
        [void]$sb.AppendLine("    <game>")

        # releaseNumber（任意）
        $releaseNumber = $row.releaseNumber
        if ($releaseNumber) {
            $relEsc = [System.Security.SecurityElement]::Escape($releaseNumber)
            [void]$sb.AppendLine("      <releaseNumber>$relEsc</releaseNumber>")
        }

        # imageNumber（必須・数字、無ければ 0）
        $imageNumber = $row.imageNumber
        if (-not $imageNumber) { $imageNumber = "0" }
        [void]$sb.AppendLine("      <imageNumber>$imageNumber</imageNumber>")

        # im1CRC / im2CRC / title / publisher / sourceRom / location / language / comment / saveType / romSize
        $fields = @{
            im1CRC    = $row.im1CRC
            im2CRC    = $row.im2CRC
            title     = $row.title
            publisher = $row.publisher
            sourceRom = $row.sourceRom
            location  = $row.location
            language  = $row.language
            comment   = $row.comment
            saveType  = $row.saveType
            romSize   = $row.romSize
        }

        foreach ($k in $fields.Keys) {
            $v = $fields[$k]
            if (-not $v) { $v = "" }  # 必須でないものも空要素OK
            $esc = [System.Security.SecurityElement]::Escape($v)
            [void]$sb.AppendLine("      <$k>$esc</$k>")
        }

        # files / romCRC（必須）
        [void]$sb.AppendLine("      <files>")
        $ext = $row.extension
        $crc = $row.romCRC
        if (-not $crc) { $crc = "00000000" }
        if (-not $ext) { $ext = "" }
        $extEsc = [System.Security.SecurityElement]::Escape($ext)
        $crcEsc = [System.Security.SecurityElement]::Escape($crc)
        [void]$sb.AppendLine("        <romCRC extension=""$extEsc"">$crcEsc</romCRC>")
        [void]$sb.AppendLine("      </files>")

        # duplicateID（任意）
        $dup = $row.duplicateID
        if ($dup -ne $null -and $dup -ne "") {
            $dupEsc = [System.Security.SecurityElement]::Escape($dup)
            [void]$sb.AppendLine("      <duplicateID>$dupEsc</duplicateID>")
        }

        [void]$sb.AppendLine("    </game>")
    }

    [void]$sb.AppendLine("  </games>")
    [void]$sb.AppendLine("</dat>")

    $xmlContent = $sb.ToString()

    $system = $TxtSystem.Text
    if (-not $system) {
        $system = (Get-BaseNameFromPath $csvPath)
    }

    # 出力ファイル名の決定
    $outDir = [System.IO.Path]::GetDirectoryName($csvPath)

    if ($ChkUseDatFileName.IsChecked) {
        # チェック ON → GUI の DATファイル名を使用
        $fileName = $TxtDatUrlFileName.Text
        if (-not $fileName) {
            # 念のため system.zip から system を抽出
            $fileName = ($TxtSystem.Text + ".zip")
        }
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    }
    else {
        # チェック OFF → 元の CSV のファイル名を使用
        $baseName = (Get-BaseNameFromPath $csvPath)
    }

    $outXmlPath = Join-Path $outDir ($baseName + ".xml")

    # UTF-8 (BOMなし) で出力
    $sw = New-Object System.IO.StreamWriter($outXmlPath, $false, (New-Object System.Text.UTF8Encoding($false)))
    $sw.Write($xmlContent)
    $sw.Close()
    Write-Log "XML 出力: $outXmlPath"

    if ($ChkVersionFile.IsChecked) {
        $verPath = Join-Path $outDir ($system + ".txt")
        $datVersion | Out-File -FilePath $verPath -Encoding ASCII
        Write-Log "バージョンファイル出力: $verPath"
    }

    if ($outKind -eq "Zip") {
        if (-not $Global:SevenZipPath) {
            throw "7za.exe が無いため ZIP 出力はできません。"
        }

        $archiveExt = ".zip"
        $archivePath = Join-Path $outDir ($baseName + $archiveExt)

        if (Test-Path $archivePath) { Remove-Item $archivePath -Force }

        $args = @("a","-y",$archivePath,$outXmlPath)
        & $Global:SevenZipPath @args | Out-Null
        Write-Log "アーカイブ出力: $archivePath"
    }
}

function Xml-ToCsv([string]$xmlPath, [string]$outDir) {
    Write-Log "XML → CSV 変換開始: $xmlPath"

    $xml = Read-XmlWithDeclaredEncoding $xmlPath
    if (-not $xml.dat) {
        throw "dat ルート要素が見つかりません。"
    }

    $games = $xml.dat.games.game
    if (-not $games) {
        throw "games/game 要素がありません。"
    }

    $rows = @()
    foreach ($g in $games) {
        $obj = [PSCustomObject]@{
            releaseNumber = $g.releaseNumber
            imageNumber   = $g.imageNumber
            im1CRC        = $g.im1CRC
            im2CRC        = $g.im2CRC
            title         = $g.title
            publisher     = $g.publisher
            sourceRom     = $g.sourceRom
            location      = $g.location
            language      = $g.language
            comment       = $g.comment
            saveType      = $g.saveType
            romSize       = $g.romSize
            duplicateID   = $g.duplicateID
            extension     = $null
            romCRC        = $null
        }

        $romNode = $g.files.romCRC | Select-Object -First 1
        if ($romNode) {
            $obj.extension = $romNode.extension
            $obj.romCRC    = $romNode.'#text'
        }

        $rows += $obj
    }

    # ★ outDir が渡されていない場合は xml と同じフォルダ
    if (-not $outDir) {
        $outDir = [System.IO.Path]::GetDirectoryName($xmlPath)
    }

    $base   = [System.IO.Path]::GetFileNameWithoutExtension($xmlPath)
    $outCsv = Join-Path $outDir ($base + ".csv")

    # UTF-8 (BOMなし) で出力
    $sw = New-Object System.IO.StreamWriter($outCsv, $false, (New-Object System.Text.UTF8Encoding($false)))
    $csv = $rows | ConvertTo-Csv -NoTypeInformation
    foreach ($line in $csv) { $sw.WriteLine($line) }
    $sw.Close()

    Write-Log "CSV 出力: $outCsv"
}

function Extract-XmlFromArchive([string]$archivePath) {
    if (-not $Global:SevenZipPath) {
        throw "7za.exe / 7z.exe が無いためアーカイブ入力はできません。"
    }

    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("OLCSV_" + [System.Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $tempDir | Out-Null

    $args = @("e","-y","-o$tempDir",$archivePath,"*.xml")
    & $Global:SevenZipPath @args | Out-Null

    $xmlFiles = Get-ChildItem -Path $tempDir -Filter *.xml
    if ($xmlFiles.Count -ne 1) {
        throw "アーカイブ内の XML が単一ではありません。"
    }

    # ★ 元の zip のフォルダも返す
    return @{
        XmlPath = $xmlFiles[0].FullName
        OutDir  = [System.IO.Path]::GetDirectoryName($archivePath)
    }
}

$BtnConvert.Add_Click({
    try {
        # baseURL の URL 妥当性チェック
        $baseUrl = $TxtBaseUrl.Text
        $uriRef = $null
        if ($baseUrl -and -not [System.Uri]::TryCreate($baseUrl, [System.UriKind]::Absolute, [ref]$uriRef)) {
            Show-ErrorMessage "baseURL が URL として正しくありません。処理は続行しますが、OfflineList で正しく動作しない可能性があります。"
        }

        # 入力ファイルチェック
        $path = $TxtInputFile.Text
        if (-not (Test-Path $path)) {
            Show-ErrorMessage "入力ファイルが指定されていません。"
            return
        }

        # ★ 出力フォルダの決定（ここを追加）
        if ($TxtOutputFolder.Text) {
            # ユーザーが指定したフォルダ
            $outputDir = $TxtOutputFolder.Text
        } else {
            # 入力ファイルと同じフォルダ
            $outputDir = Split-Path $path
        }

        # フォルダが存在しなければ作成
        if (-not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir | Out-Null
        }

        # 拡張子判定
        $ext = [System.IO.Path]::GetExtension($path).ToLowerInvariant()

        switch ($ext) {
        
            ".xml" {
                Xml-ToCsv -xmlPath $path -outDir $outputDir
            }
        
            ".csv" {
                $selectedItem = $CmbOutputKindBottom.SelectedItem
                $kind = $selectedItem.Tag
                Csv-ToXml -csvPath $path -outKind $kind -outDir $outputDir
            }
        
            ".zip" {
                if (-not $Global:SevenZipPath) {
                    throw "7za.exe / 7z.exe が無いため zip 入力はできません。"
                }
                $result = Extract-XmlFromArchive -archivePath $path
                Xml-ToCsv -xmlPath $result.XmlPath -outDir $outputDir
            }
        
            ".7z" {
                if (-not $Global:SevenZipPath) {
                    throw "7za.exe / 7z.exe が無いため 7z 入力はできません。"
                }
                $result = Extract-XmlFromArchive -archivePath $path
                Xml-ToCsv -xmlPath $result.XmlPath -outDir $outputDir
            }
        
            default {
                throw "対応していない拡張子です: $ext"
            }
        }

        Write-Log "変換完了。出力先: $outputDir"

    } catch {
        Write-Log "エラー: $($_.Exception.Message)"
        Show-ErrorMessage $_.Exception.Message
    }
})

$BtnClose.Add_Click({ $window.Close() })

$window.ShowDialog() | Out-Null
