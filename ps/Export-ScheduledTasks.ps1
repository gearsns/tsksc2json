function Get-TaskInfo {
    param ($task, $path)

    $definition = $task.Definition
    # アクション情報の取得
    $actions = @()
    foreach ($action in $definition.Actions) {
        if ($action.Type -eq 0) {  # TASK_ACTION_EXEC
            $a = @{}
            if ($action.Path) { $a.Command = $action.Path }
            if ($action.Arguments) { $a.Arguments = $action.Arguments }
            if ($action.WorkingDirectory) { $a.WorkingDirectory = $action.WorkingDirectory }
            if ($a.Count -ge 0) { $actions += $a }
        }
    }
    # トリガー情報の取得
    $triggers = @()
    foreach ($trigger in $definition.Triggers) {
        $t = @{ Enabled = $trigger.Enabled }
        if ($trigger.StartBoundary) { $t.StartBoundary = $trigger.StartBoundary }

        switch ($trigger.Type) {
            2 {  # Daily
                $a = @{}
                if ($trigger.DaysInterval) { $a.DaysInterval = $trigger.DaysInterval }
                if ($a.Count -gt 0){ $t.ScheduleByDay = $a }
            }
            3 {  # Weekly
                $a = @{}
                if ($trigger.WeeksInterval) { $a.WeeksInterval = $trigger.WeeksInterval }
                # DaysOfWeekは0になる可能性があるのでnullチェックのみ
                if ($null -ne $trigger.DaysOfWeek.value__) { $a.DaysOfWeek = $trigger.DaysOfWeek.value__ }
                if ($a.Count -gt 0){ $t.ScheduleByWeek = $a }
            }
            4 {  # Monthly
                $a = @{}
                if ($trigger.DaysOfMonth) { $a.Days = $trigger.DaysOfMonth }
                if ($trigger.MonthsOfYear) { $a.Months = $trigger.MonthsOfYear }
                if ($trigger.RunOnLastDayOfMonth) { $a.RunOnLastDayOfMonth = $trigger.RunOnLastDayOfMonth }
                if ($a.Count -gt 0){ $t.ScheduleByMonth = $a }
            }
        }

        if ($trigger.Repetition) {
            $rep = $trigger.Repetition
            $a = @{}
            if ($rep.Interval) { $a.Interval = $rep.Interval }
            if ($rep.Duration) { $a.Duration = $rep.Duration }
            if ($rep.StopAtDurationEnd) { $a.StopAtDurationEnd = $rep.StopAtDurationEnd }
            if ($a.Count -gt 0) { $t.Repetition = $a }
        }
        $triggers += $t
    }
    # タスク基本情報の構築
    $taskInfo =  @{
        TaskPath = $path
        TaskName = $task.Name
        State = [int]$task.State
        LastResult = $task.LastTaskResult
        Description = $definition.RegistrationInfo.Description
    }
    # 日付プロパティは0でない場合のみ追加
    if ($task.LastRunTime -ne 0) {
        $taskInfo.LastRunTime = $task.LastRunTime.ToString("yyyy/MM/dd HH:mm:ss")
    }
    if ($task.NextRunTime -ne 0) {
        $taskInfo.NextRunTime = $task.NextRunTime.ToString("yyyy/MM/dd HH:mm:ss")
    }
    # アクションとトリガーが存在する場合のみ追加
    if ($actions.Count -gt 0) { $taskInfo.Actions = $actions }
    if ($triggers.Count -gt 0) { $taskInfo.Triggers = $triggers }

    $taskInfo
}

function Get-Tasks {
    param ($folder, $path)

    $allTasks = @()

    try {
        # 現在のフォルダのタスクを取得
        $allTasks += $folder.GetTasks(1) | ForEach-Object { Get-TaskInfo -task $_ -path $path }
        # サブフォルダを再帰的に処理
        foreach ($subFolder in $folder.GetFolders(0)) {
            $allTasks += Get-Tasks -folder $subFolder -path $subFolder.Path
        }
    }
    catch {
        Write-Warning "Failed to get tasks or subfolders from $($folder.Path): $($_.Exception.Message)"
    }

    $allTasks
}

# --- メイン処理 ---
try {
    $rootPath = if ($args.Length -gt 0) { $args[0] } else { "\" }
    $service = New-Object -ComObject "Schedule.Service"
    $service.Connect()

    $result = @()

    if ($rootPath -match "^(.+)\\(.+[^\\])$") { # ディレクトリとファイル名に分割
        $path = $matches[1]
        $name = $matches[2]
        $folder = $service.GetFolder($path)
        $task = $folder.GetTask($name)
        $result = @(Get-TaskInfo -task $task -path $path)
    } elseif ($inputPath.EndsWith('\')) {
        # パスが'\'で終わっている場合は削除（ルートパスの場合は維持）
        $rootPath = $rootPath -replace "^(.*[^\\])\\$", "$1"
        $folder = $service.GetFolder($rootPath)
        $result = Get-Tasks -folder $folder -path $rootPath
    }

    $result | ConvertTo-Json -Depth 10
}
catch [System.UnauthorizedAccessException] {
    Write-Error "アクセス拒否されました。管理者権限で実行してください。`n$($_.Exception.Message)"
}
catch [System.Management.Automation.RuntimeException] {
    Write-Error "スクリプト実行中にエラーが発生しました: $($_.Exception.Message)"
}
catch {
    Write-Error "予期せぬエラーが発生しました: $($_.Exception.Message)"
}