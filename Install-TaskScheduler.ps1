# WebP自動変換 タスクスケジューラ登録スクリプト
# 機能: WebP-AutoConverter.ps1をWindowsのタスクスケジューラに登録
# 作成者: Claude Code

param(
    [switch]$Uninstall,
    [switch]$Help,
    [string]$ScriptPath = "",
    [string]$TaskName = "WebP Auto Converter"
)

# ヘルプ表示
if ($Help) {
    Write-Host "WebP自動変換 タスクスケジューラ登録スクリプト" -ForegroundColor Green
    Write-Host ""
    Write-Host "使用法:" -ForegroundColor Yellow
    Write-Host "  .\Install-TaskScheduler.ps1                  # タスクを登録"
    Write-Host "  .\Install-TaskScheduler.ps1 -Uninstall       # タスクを削除"
    Write-Host "  .\Install-TaskScheduler.ps1 -ScriptPath <path> # カスタムスクリプトパス"
    Write-Host "  .\Install-TaskScheduler.ps1 -TaskName <name>  # カスタムタスク名"
    Write-Host "  .\Install-TaskScheduler.ps1 -Help             # このヘルプ"
    Write-Host ""
    Write-Host "注意:" -ForegroundColor Red
    Write-Host "  - 管理者権限で実行してください"
    Write-Host "  - デフォルトでは同一フォルダのWebP-AutoConverter.ps1を登録します"
    exit
}

# 管理者権限チェック
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# スクリプトパスの解決
function Get-ScriptPath {
    if ($ScriptPath) {
        if (Test-Path $ScriptPath) {
            return $ScriptPath
        }
        else {
            throw "指定されたスクリプトパスが見つかりません: $ScriptPath"
        }
    }
    
    # デフォルト: 同一フォルダのWebP-AutoConverter.ps1
    $defaultPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "WebP-AutoConverter.ps1"
    
    if (Test-Path $defaultPath) {
        return $defaultPath
    }
    else {
        throw "WebP-AutoConverter.ps1が見つかりません。-ScriptPathパラメータで指定してください。"
    }
}

# タスク削除
function Remove-ScheduledTask {
    try {
        Write-Host "タスクスケジューラからタスクを削除しています..." -ForegroundColor Yellow
        
        # タスクが存在するかチェック
        $existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        
        if ($existingTask) {
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
            Write-Host "タスク '$TaskName' を削除しました。" -ForegroundColor Green
        }
        else {
            Write-Host "タスク '$TaskName' は存在しませんでした。" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Error "タスクの削除に失敗しました: $($_.Exception.Message)"
        exit 1
    }
}

# タスク登録
function Register-ScheduledTask {
    try {
        Write-Host "タスクスケジューラにタスクを登録しています..." -ForegroundColor Yellow
        
        # スクリプトパスを取得
        $scriptFullPath = Get-ScriptPath
        Write-Host "スクリプトパス: $scriptFullPath" -ForegroundColor Green
        
        # 既存タスクを削除（存在する場合）
        $existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        if ($existingTask) {
            Write-Host "既存のタスクを削除しています..." -ForegroundColor Yellow
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        }
        
        # タスクアクション設定
        $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptFullPath`""
        
        # トリガー設定（ログオン時に開始）
        $trigger = New-ScheduledTaskTrigger -AtLogOn
        
        # 設定
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable:$false
        
        # プリンシパル設定（現在のユーザーで実行）
        $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive
        
        # タスク登録
        Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description "WebP自動変換スクリプト - PNG/BMP/JPEGファイルを自動的にWebPに変換します"
        
        Write-Host ""
        Write-Host "タスクの登録が完了しました！" -ForegroundColor Green
        Write-Host "タスク名: $TaskName" -ForegroundColor Cyan
        Write-Host "実行ユーザー: $currentUser" -ForegroundColor Cyan
        Write-Host "開始条件: ログオン時" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "確認方法:" -ForegroundColor Yellow
        Write-Host "  1. タスクスケジューラを開く (taskschd.msc)"
        Write-Host "  2. タスクスケジューラライブラリで '$TaskName' を探す"
        Write-Host "  3. 手動で実行してテストすることも可能"
        Write-Host ""
        Write-Host "注意:" -ForegroundColor Red
        Write-Host "  - 次回ログオン時から自動的に開始されます"
        Write-Host "  - スクリプトは非表示で実行されます"
        Write-Host "  - ログは %LOCALAPPDATA%\\WebP-Auto\\logs に出力されます"
    }
    catch {
        Write-Error "タスクの登録に失敗しました: $($_.Exception.Message)"
        Write-Host ""
        Write-Host "トラブルシューティング:" -ForegroundColor Yellow
        Write-Host "  1. PowerShellを管理者として実行していますか？"
        Write-Host "  2. WebP-AutoConverter.ps1が存在しますか？"
        Write-Host "  3. 実行ポリシーが適切に設定されていますか？"
        Write-Host "     Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser"
        exit 1
    }
}

# タスク状態確認
function Show-TaskStatus {
    try {
        $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        
        if ($task) {
            Write-Host ""
            Write-Host "=== タスク情報 ===" -ForegroundColor Cyan
            Write-Host "タスク名: $($task.TaskName)"
            Write-Host "状態: $($task.State)"
            Write-Host "作成者: $($task.Author)"
            Write-Host "説明: $($task.Description)"
            
            # 最後の実行結果
            $taskInfo = Get-ScheduledTaskInfo -TaskName $TaskName -ErrorAction SilentlyContinue
            if ($taskInfo) {
                Write-Host "最後の実行: $($taskInfo.LastRunTime)"
                Write-Host "最後の結果: $($taskInfo.LastTaskResult)"
                Write-Host "次の実行: $($taskInfo.NextRunTime)"
            }
            
            Write-Host ""
            Write-Host "手動でタスクを開始するには:" -ForegroundColor Yellow
            Write-Host "  Start-ScheduledTask -TaskName '$TaskName'"
            
            Write-Host ""
            Write-Host "タスクを停止するには:" -ForegroundColor Yellow
            Write-Host "  Stop-ScheduledTask -TaskName '$TaskName'"
        }
        else {
            Write-Host "タスク '$TaskName' は登録されていません。" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Warning "タスク情報の取得に失敗しました: $($_.Exception.Message)"
    }
}

# メイン処理
function Main {
    Write-Host "WebP自動変換 タスクスケジューラ管理" -ForegroundColor Green
    Write-Host "======================================" -ForegroundColor Green
    
    # 管理者権限チェック
    if (-not (Test-Administrator)) {
        Write-Error "このスクリプトは管理者権限で実行する必要があります。"
        Write-Host ""
        Write-Host "管理者権限でPowerShellを開く方法:" -ForegroundColor Yellow
        Write-Host "  1. スタートメニューで 'PowerShell' を検索"
        Write-Host "  2. 'Windows PowerShell' を右クリック"
        Write-Host "  3. '管理者として実行' を選択"
        exit 1
    }
    
    if ($Uninstall) {
        Remove-ScheduledTask
    }
    else {
        Register-ScheduledTask
    }
    
    # タスク状態表示
    Show-TaskStatus
    
    Write-Host ""
    Write-Host "完了しました！" -ForegroundColor Green
}

# 実行
try {
    Main
}
catch {
    Write-Error "予期しないエラーが発生しました: $($_.Exception.Message)"
    exit 1
}