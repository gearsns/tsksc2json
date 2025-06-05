# tsksc2json
Windowsのタスクスケジュールをjson形式で出力

## 概要
COMを使用してタスクスケジューラに登録されているタスクをjson形式で出力します。

## 背景
Windowsのタスクスケジューラ情報を取得には、標準のschtasksコマンドを使用する方法がありますが、
CSV形式で出力した場合、内容が正しく整形されていないことがあり後処理が煩雑になります。
代わりにXML形式で出力することも可能ですが、この形式ではトリガー情報が含まれないという制約があります。

また、PowerShellでは Get-ScheduledTask、Get-ScheduledTaskInfo、Export-ScheduledTask などの
コマンドを組み合わせて同様の情報を取得できますが、
こちらも出力されるオブジェクトに不要な情報が多く、用途に応じた整形が必要になります。

これらを踏まえ、タスクスケジューラ情報をJSON形式で出力できるアプリケーションを作成しました。

## 実行
全てのタスクを出力
```bash
tskc2json
```

特定のフォルダ以下を出力
※末尾を\で終了
```bash
tskc2json \タスフォルダ\
```

特定のタスクのみを出力
```bash
tskc2json \タスフォルダ\タスク名
```

## PowerShell版
全てのタスクを出力
```bash
.\Export-ScheduledTasks.ps1
```

特定のフォルダ以下を出力
※末尾を\で終了
```bash
.\Export-ScheduledTasks.ps1 \タスフォルダ\
```

特定のタスクのみを出力
```bash
.\Export-ScheduledTasks.ps1 \タスフォルダ\タスク名
```
