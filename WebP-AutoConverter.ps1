# WebP自動変換スクリプト
# 機能: PNG/BMP/JPEGファイルを自動的にWebPに変換
# 作成者: Claude Code
# バージョン: 1.0

param(
    [string]$ConfigPath = "",
    [switch]$OneTime,
    [switch]$Help
)

# ヘルプ表示
if ($Help) {
    Write-Host "WebP自動変換スクリプト" -ForegroundColor Green
    Write-Host "使用法:" -ForegroundColor Yellow
    Write-Host "  .\WebP-AutoConverter.ps1              # 常駐監視モード"
    Write-Host "  .\WebP-AutoConverter.ps1 -OneTime      # 一回限りの変換"
    Write-Host "  .\WebP-AutoConverter.ps1 -ConfigPath <path> # カスタム設定ファイル"
    Write-Host "  .\WebP-AutoConverter.ps1 -Help         # このヘルプ"
    exit
}

# 必要なモジュールの読み込み
Add-Type -AssemblyName System.IO

# グローバル変数
$script:Config = $null
$script:Logger = $null
$script:FileWatcher = $null
$script:ProcessedFiles = @{}
$script:Mutex = $null

# 設定クラス
class WebPConfig {
    [string[]]$WatchPaths = @()
    [string[]]$SupportedExtensions = @(".png", ".jpg", ".jpeg", ".bmp")
    [int]$JpegQuality = 95
    [bool]$LosslessPngBmp = $true
    [bool]$PreserveMetadata = $true
    [bool]$DeleteOriginal = $true
    [string]$OriginalBackupFolder = "Originals"
    [bool]$EnableNotifications = $true
    [string]$LogLevel = "INFO"
    [int]$ProcessDelayMs = 1000
    [bool]$PreventDuplicateProcessing = $true
    [string[]]$ExcludePatterns = @()

    # デフォルト設定を読み込み
    static [WebPConfig] LoadDefault() {
        $config = [WebPConfig]::new()
        
        # デフォルトの監視フォルダ
        $downloadsPath = [Environment]::GetFolderPath("MyDocuments") + "\Downloads"
        if (-not (Test-Path $downloadsPath)) {
            $downloadsPath = $env:USERPROFILE + "\Downloads"
        }
        
        $screenshotsPath = [Environment]::GetFolderPath("MyPictures") + "\Screenshots"
        if (-not (Test-Path $screenshotsPath)) {
            $screenshotsPath = $env:USERPROFILE + "\Pictures\Screenshots"
        }
        
        $config.WatchPaths = @($downloadsPath, $screenshotsPath) | Where-Object { Test-Path $_ }
        
        return $config
    }
    
    # 設定をJSONファイルから読み込み
    static [WebPConfig] LoadFromFile([string]$path) {
        try {
            if (Test-Path $path) {
                $json = Get-Content -Path $path -Encoding UTF8 | ConvertFrom-Json
                $config = [WebPConfig]::new()
                
                # プロパティを手動でコピー
                $config.WatchPaths = $json.WatchPaths
                $config.SupportedExtensions = $json.SupportedExtensions
                $config.JpegQuality = $json.JpegQuality
                $config.LosslessPngBmp = $json.LosslessPngBmp
                $config.PreserveMetadata = $json.PreserveMetadata
                $config.DeleteOriginal = $json.DeleteOriginal
                $config.OriginalBackupFolder = $json.OriginalBackupFolder
                $config.EnableNotifications = $json.EnableNotifications
                $config.LogLevel = $json.LogLevel
                $config.ProcessDelayMs = $json.ProcessDelayMs
                $config.PreventDuplicateProcessing = $json.PreventDuplicateProcessing
                $config.ExcludePatterns = $json.ExcludePatterns
                
                return $config
            }
        }
        catch {
            Write-Warning "設定ファイルの読み込みに失敗しました: $($_.Exception.Message)"
        }
        
        return [WebPConfig]::LoadDefault()
    }
    
    # 設定をJSONファイルに保存
    [void] SaveToFile([string]$path) {
        try {
            $dir = Split-Path -Parent $path
            if (-not (Test-Path $dir)) {
                New-Item -ItemType Directory -Path $dir -Force | Out-Null
            }
            
            $this | ConvertTo-Json -Depth 10 | Out-File -FilePath $path -Encoding UTF8
        }
        catch {
            Write-Warning "設定ファイルの保存に失敗しました: $($_.Exception.Message)"
        }
    }
}

# ログ機能クラス
class WebPLogger {
    [string]$LogPath
    [string]$LogLevel
    
    WebPLogger([string]$logPath, [string]$logLevel) {
        $this.LogPath = $logPath
        $this.LogLevel = $logLevel
        
        # ログディレクトリの作成
        $logDir = Split-Path -Parent $this.LogPath
        if (-not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }
    }
    
    [void] WriteLog([string]$level, [string]$message) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "[$timestamp] [$level] $message"
        
        # コンソールに出力
        switch ($level) {
            "ERROR" { Write-Host $logEntry -ForegroundColor Red }
            "WARN"  { Write-Host $logEntry -ForegroundColor Yellow }
            "INFO"  { Write-Host $logEntry -ForegroundColor Green }
            "DEBUG" { if ($this.LogLevel -eq "DEBUG") { Write-Host $logEntry -ForegroundColor Gray } }
        }
        
        # ファイルに出力
        try {
            Add-Content -Path $this.LogPath -Value $logEntry -Encoding UTF8
        }
        catch {
            Write-Warning "ログファイルへの書き込みに失敗しました: $($_.Exception.Message)"
        }
    }
    
    [void] Info([string]$message) { $this.WriteLog("INFO", $message) }
    [void] Warn([string]$message) { $this.WriteLog("WARN", $message) }
    [void] Error([string]$message) { $this.WriteLog("ERROR", $message) }
    [void] Debug([string]$message) { $this.WriteLog("DEBUG", $message) }
}

# 通知機能
function Send-ToastNotification {
    param(
        [string]$Title,
        [string]$Message,
        [string]$Type = "Info"
    )
    
    if (-not $script:Config.EnableNotifications) {
        return
    }
    
    try {
        # BurntToastモジュールが利用可能かチェック
        if (Get-Module -ListAvailable -Name BurntToast) {
            Import-Module BurntToast -ErrorAction SilentlyContinue
            
            $toastParams = @{
                Text = @($Title, $Message)
                AppLogo = $null
            }
            
            New-BurntToastNotification @toastParams
            $script:Logger.Debug("BurntToastで通知を送信しました")
        }
        else {
            # フォールバック: シンプルなバルーン通知
            Add-Type -AssemblyName System.Windows.Forms
            $balloon = New-Object System.Windows.Forms.NotifyIcon
            $balloon.Icon = [System.Drawing.SystemIcons]::Information
            $balloon.BalloonTipTitle = $Title
            $balloon.BalloonTipText = $Message
            $balloon.BalloonTipIcon = $Type
            $balloon.Visible = $true
            $balloon.ShowBalloonTip(5000)
            
            # クリーンアップ
            Start-Sleep -Milliseconds 100
            $balloon.Dispose()
            $script:Logger.Debug("標準バルーン通知を送信しました")
        }
    }
    catch {
        $script:Logger.Warn("通知の送信に失敗しました: $($_.Exception.Message)")
    }
}

# ImageMagick/cwebpツールのチェック
function Test-ConversionTools {
    $tools = @()
    
    # ImageMagickのチェック
    try {
        $magickResult = & magick -version 2>$null
        if ($LASTEXITCODE -eq 0) {
            $tools += "ImageMagick"
            $script:Logger.Info("ImageMagickが検出されました")
        }
    }
    catch {
        $script:Logger.Debug("ImageMagickが見つかりませんでした")
    }
    
    # cwebpのチェック
    try {
        $cwebpResult = & cwebp -version 2>$null
        if ($LASTEXITCODE -eq 0) {
            $tools += "cwebp"
            $script:Logger.Info("cwebpが検出されました")
        }
    }
    catch {
        $script:Logger.Debug("cwebpが見つかりませんでした")
    }
    
    if ($tools.Count -eq 0) {
        $script:Logger.Error("ImageMagickまたはcwebpが必要です。どちらも検出されませんでした")
        throw "変換ツールが見つかりません"
    }
    
    return $tools
}

# ファイルをWebPに変換
function Convert-ToWebP {
    param(
        [string]$InputPath,
        [string]$OutputPath,
        [string[]]$AvailableTools
    )
    
    $script:Logger.Info("変換開始: $InputPath -> $OutputPath")
    
    try {
        # 長いパスの処理
        if ($InputPath.Length -gt 260) {
            $InputPath = "\\?\$InputPath"
        }
        if ($OutputPath.Length -gt 260) {
            $OutputPath = "\\?\$OutputPath"
        }
        
        # ファイル拡張子を取得
        $extension = [System.IO.Path]::GetExtension($InputPath).ToLower()
        
        # ImageMagickを優先して使用
        if ($AvailableTools -contains "ImageMagick") {
            $magickArgs = @($InputPath)
            
            # JPEG以外は無劣化、JPEGは指定品質
            if ($extension -eq ".jpg" -or $extension -eq ".jpeg") {
                $magickArgs += "-quality", $script:Config.JpegQuality
            }
            elseif ($script:Config.LosslessPngBmp) {
                $magickArgs += "-define", "webp:lossless=true"
            }
            
            # メタデータの処理
            if (-not $script:Config.PreserveMetadata) {
                $magickArgs += "-strip"
            }
            
            $magickArgs += $OutputPath
            
            $result = & magick @magickArgs 2>&1
            
            if ($LASTEXITCODE -ne 0) {
                throw "ImageMagickでの変換に失敗しました: $result"
            }
        }
        elseif ($AvailableTools -contains "cwebp") {
            $cwebpArgs = @()
            
            # JPEG以外は無劣化、JPEGは指定品質
            if ($extension -eq ".jpg" -or $extension -eq ".jpeg") {
                $cwebpArgs += "-q", $script:Config.JpegQuality
            }
            elseif ($script:Config.LosslessPngBmp) {
                $cwebpArgs += "-lossless"
            }
            
            # メタデータの処理
            if ($script:Config.PreserveMetadata) {
                $cwebpArgs += "-metadata", "all"
            }
            
            $cwebpArgs += $InputPath, "-o", $OutputPath
            
            $result = & cwebp @cwebpArgs 2>&1
            
            if ($LASTEXITCODE -ne 0) {
                throw "cwebpでの変換に失敗しました: $result"
            }
        }
        
        # 変換後ファイルのチェック
        if (-not (Test-Path $OutputPath)) {
            throw "変換後ファイルが作成されませんでした"
        }
        
        $originalSize = (Get-Item $InputPath).Length
        $webpSize = (Get-Item $OutputPath).Length
        $compressionRatio = [Math]::Round((1 - ($webpSize / $originalSize)) * 100, 1)
        
        $script:Logger.Info("変換完了: $([Math]::Round($originalSize/1KB, 1))KB -> $([Math]::Round($webpSize/1KB, 1))KB ($compressionRatio% 削減)")
        
        return $true
    }
    catch {
        $script:Logger.Error("変換エラー: $($_.Exception.Message)")
        
        # 失敗時の通知
        Send-ToastNotification -Title "WebP変換エラー" -Message "ファイル: $(Split-Path -Leaf $InputPath)`nエラー: $($_.Exception.Message)" -Type "Error"
        
        return $false
    }
}

# 元ファイルの処理
function Handle-OriginalFile {
    param(
        [string]$FilePath,
        [string]$WebPPath
    )
    
    try {
        if ($script:Config.DeleteOriginal) {
            Remove-Item -Path $FilePath -Force
            $script:Logger.Info("元ファイルを削除しました: $FilePath")
        }
        else {
            # Originalsフォルダに移動
            $parentDir = Split-Path -Parent $FilePath
            $backupDir = Join-Path $parentDir $script:Config.OriginalBackupFolder
            
            if (-not (Test-Path $backupDir)) {
                New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
            }
            
            $backupPath = Join-Path $backupDir (Split-Path -Leaf $FilePath)
            
            # 同名ファイルがある場合は連番を付与
            $counter = 1
            while (Test-Path $backupPath) {
                $name = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
                $ext = [System.IO.Path]::GetExtension($FilePath)
                $backupPath = Join-Path $backupDir "$name($counter)$ext"
                $counter++
            }
            
            Move-Item -Path $FilePath -Destination $backupPath -Force
            $script:Logger.Info("元ファイルをバックアップしました: $FilePath -> $backupPath")
        }
    }
    catch {
        $script:Logger.Error("元ファイルの処理に失敗しました: $($_.Exception.Message)")
    }
}

# ファイル処理のメイン関数
function Process-ImageFile {
    param(
        [string]$FilePath,
        [string[]]$AvailableTools
    )
    
    try {
        # ファイルが存在するかチェック
        if (-not (Test-Path $FilePath)) {
            $script:Logger.Debug("ファイルが存在しません: $FilePath")
            return
        }
        
        # 除外パターンのチェック
        foreach ($pattern in $script:Config.ExcludePatterns) {
            if ($FilePath -like $pattern) {
                $script:Logger.Debug("除外パターンにマッチしました: $FilePath")
                return
            }
        }
        
        # 対応拡張子かチェック
        $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
        if ($script:Config.SupportedExtensions -notcontains $extension) {
            return
        }
        
        # 重複処理の防止
        if ($script:Config.PreventDuplicateProcessing) {
            $fileInfo = Get-Item $FilePath
            $fileKey = "$($fileInfo.FullName):$($fileInfo.LastWriteTime)"
            
            if ($script:ProcessedFiles.ContainsKey($fileKey)) {
                $script:Logger.Debug("重複処理をスキップしました: $FilePath")
                return
            }
            
            $script:ProcessedFiles[$fileKey] = $true
        }
        
        # ファイルがロックされていないかチェック（最大5回リトライ）
        $retryCount = 0
        do {
            try {
                $fileStream = [System.IO.File]::Open($FilePath, 'Open', 'Read', 'None')
                $fileStream.Close()
                break
            }
            catch {
                $retryCount++
                if ($retryCount -ge 5) {
                    $script:Logger.Warn("ファイルがロックされているため処理をスキップしました: $FilePath")
                    return
                }
                Start-Sleep -Milliseconds 500
            }
        } while ($retryCount -lt 5)
        
        # 出力パスの生成
        $directory = Split-Path -Parent $FilePath
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
        $webpPath = Join-Path $directory "$baseName.webp"
        
        # 同名ファイルが存在する場合は連番を付与
        $counter = 1
        while (Test-Path $webpPath) {
            $webpPath = Join-Path $directory "$baseName($counter).webp"
            $counter++
        }
        
        # 変換実行
        $success = Convert-ToWebP -InputPath $FilePath -OutputPath $webpPath -AvailableTools $AvailableTools
        
        if ($success) {
            # 元ファイルの処理
            Handle-OriginalFile -FilePath $FilePath -WebPPath $webpPath
        }
    }
    catch {
        $script:Logger.Error("ファイル処理中にエラーが発生しました: $FilePath - $($_.Exception.Message)")
    }
}

# ファイル監視イベントハンドラー
function On-FileCreated {
    param($sender, $e)
    
    $script:Logger.Debug("ファイル作成イベント: $($e.FullPath)")
    
    # 少し待ってからファイル処理（ファイル書き込み完了待ち）
    Start-Sleep -Milliseconds $script:Config.ProcessDelayMs
    
    $tools = Test-ConversionTools
    Process-ImageFile -FilePath $e.FullPath -AvailableTools $tools
}

# 一回限りの処理モード
function Invoke-OneTimeConversion {
    $script:Logger.Info("一回限りの変換モードを開始しました")
    
    $tools = Test-ConversionTools
    $processedCount = 0
    
    foreach ($watchPath in $script:Config.WatchPaths) {
        if (-not (Test-Path $watchPath)) {
            $script:Logger.Warn("監視パスが存在しません: $watchPath")
            continue
        }
        
        $script:Logger.Info("処理開始: $watchPath")
        
        foreach ($extension in $script:Config.SupportedExtensions) {
            $pattern = "*$extension"
            $files = Get-ChildItem -Path $watchPath -Filter $pattern -File -Recurse
            
            foreach ($file in $files) {
                Process-ImageFile -FilePath $file.FullName -AvailableTools $tools
                $processedCount++
            }
        }
    }
    
    $script:Logger.Info("一回限りの変換が完了しました。処理ファイル数: $processedCount")
}

# ファイル監視モード
function Start-FileWatcher {
    $script:Logger.Info("ファイル監視モードを開始しました")
    
    $tools = Test-ConversionTools
    $script:FileWatcher = @()
    
    foreach ($watchPath in $script:Config.WatchPaths) {
        if (-not (Test-Path $watchPath)) {
            $script:Logger.Warn("監視パスが存在しません: $watchPath")
            continue
        }
        
        $script:Logger.Info("監視開始: $watchPath")
        
        $watcher = New-Object System.IO.FileSystemWatcher
        $watcher.Path = $watchPath
        $watcher.Filter = "*.*"
        $watcher.EnableRaisingEvents = $true
        $watcher.IncludeSubdirectories = $true
        
        # イベント登録
        Register-ObjectEvent -InputObject $watcher -EventName "Created" -Action {
            param($sender, $e)
            
            try {
                $extension = [System.IO.Path]::GetExtension($e.FullPath).ToLower()
                if ($script:Config.SupportedExtensions -contains $extension) {
                    Start-Sleep -Milliseconds $script:Config.ProcessDelayMs
                    $tools = Test-ConversionTools
                    Process-ImageFile -FilePath $e.FullPath -AvailableTools $tools
                }
            }
            catch {
                $script:Logger.Error("ファイル監視イベント処理エラー: $($_.Exception.Message)")
            }
        } | Out-Null
        
        $script:FileWatcher += $watcher
    }
    
    $script:Logger.Info("ファイル監視を開始しました。Ctrl+Cで終了してください。")
    
    # 監視継続
    try {
        while ($true) {
            Start-Sleep -Seconds 1
        }
    }
    finally {
        # クリーンアップ
        Stop-FileWatcher
    }
}

# ファイル監視停止
function Stop-FileWatcher {
    if ($script:FileWatcher) {
        foreach ($watcher in $script:FileWatcher) {
            $watcher.Dispose()
        }
        $script:FileWatcher = $null
        $script:Logger.Info("ファイル監視を停止しました")
    }
}

# ミューテックス管理
function Initialize-Mutex {
    try {
        $script:Mutex = New-Object System.Threading.Mutex($false, "WebPAutoConverter")
        if (-not $script:Mutex.WaitOne(0)) {
            Write-Warning "WebP自動変換スクリプトは既に実行中です。"
            exit 1
        }
        return $true
    }
    catch {
        Write-Warning "ミューテックスの初期化に失敗しました: $($_.Exception.Message)"
        return $false
    }
}

function Cleanup-Mutex {
    if ($script:Mutex) {
        $script:Mutex.ReleaseMutex()
        $script:Mutex.Dispose()
        $script:Mutex = $null
    }
}

# メイン処理
function Main {
    try {
        # 多重起動防止
        if (-not (Initialize-Mutex)) {
            exit 1
        }
        
        # 設定読み込み
        if ($ConfigPath) {
            $script:Config = [WebPConfig]::LoadFromFile($ConfigPath)
        }
        else {
            $configDir = "$env:LOCALAPPDATA\WebP-Auto"
            $defaultConfigPath = Join-Path $configDir "config.json"
            
            if (Test-Path $defaultConfigPath) {
                $script:Config = [WebPConfig]::LoadFromFile($defaultConfigPath)
            }
            else {
                $script:Config = [WebPConfig]::LoadDefault()
                # デフォルト設定を保存
                $script:Config.SaveToFile($defaultConfigPath)
                Write-Host "デフォルト設定ファイルを作成しました: $defaultConfigPath" -ForegroundColor Green
            }
        }
        
        # ログ初期化
        $logDir = "$env:LOCALAPPDATA\WebP-Auto\logs"
        $logFile = Join-Path $logDir "$(Get-Date -Format 'yyyy-MM-dd').log"
        $script:Logger = [WebPLogger]::new($logFile, $script:Config.LogLevel)
        
        $script:Logger.Info("WebP自動変換スクリプトを開始しました")
        $script:Logger.Info("監視パス: $($script:Config.WatchPaths -join ', ')")
        
        # 変換ツールの確認
        Test-ConversionTools | Out-Null
        
        # モード別実行
        if ($OneTime) {
            Invoke-OneTimeConversion
        }
        else {
            Start-FileWatcher
        }
    }
    catch {
        if ($script:Logger) {
            $script:Logger.Error("スクリプト実行エラー: $($_.Exception.Message)")
        }
        else {
            Write-Error "スクリプト実行エラー: $($_.Exception.Message)"
        }
        exit 1
    }
    finally {
        # クリーンアップ
        Stop-FileWatcher
        Cleanup-Mutex
        if ($script:Logger) {
            $script:Logger.Info("WebP自動変換スクリプトを終了しました")
        }
    }
}

# Ctrl+Cハンドリング
Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
    Stop-FileWatcher
    Cleanup-Mutex
} | Out-Null

# スクリプト実行
Main