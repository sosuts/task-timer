# task timer

## アプリの目的

- GitLab, vscode,を使ってコードレビューしている時間を計測
- vscode, visual studioを使って実装をしている時間を計測
- WordやExcelで仕様書を作成している時間を計測

## 機能

- ネットワーク通信はしない。スタンドアロン
- タスクごとの時間を計測、ラベル付けし、csvで出力する
- バックグラウンドで常時動作
- 手動でタスクの開始、終了を記録することも可能
- 見た目がおしゃれで、使いやすいUI
- zorderを常に最前面にする機能あり
- タスクの開始、終了を自動で検知する機能あり
- マウスやキーボードが{一定時間}動いていないときは、タスクを一時停止する機能あり
- GitLabは社内オンプレドメインに対応

## 技術スタック

- .NET 8 (WPF)
- CommunityToolkit.Mvvm (MVVM パターン)
- Hardcodet.NotifyIcon.Wpf (システムトレイ)
- Win32 API (アイドル検知・プロセス監視)

## プロジェクト構成

```
src/TaskTimer/
├── App.xaml / App.xaml.cs               # アプリケーションエントリ・タスクトレイ
├── MainWindow.xaml / MainWindow.xaml.cs  # メインウィンドウUI
├── TaskTimer.csproj                     # プロジェクトファイル
├── Models/
│   ├── AppSettings.cs                   # アプリ設定（JSON保存）
│   ├── ProcessMapping.cs                # プロセス→カテゴリのマッピング
│   ├── TaskCategory.cs                  # タスクカテゴリ列挙型
│   ├── TaskRecord.cs                    # タスク記録モデル
│   └── TaskState.cs                     # タスク状態列挙型
├── ViewModels/
│   └── MainViewModel.cs                 # メインViewModel
├── Services/
│   ├── CsvExportService.cs              # CSV出力
│   ├── IdleDetectionService.cs          # アイドル検知（Win32 API）
│   └── ProcessMonitorService.cs         # プロセス監視・自動検知
├── Converters/
│   └── Converters.cs                    # 値コンバーター
└── Resources/
    └── Styles.xaml                      # カスタムスタイル・テーマ
```

## ビルド・実行方法

**前提条件**: Windows 10/11 に .NET 8 SDK がインストールされていること

```bash
# ビルド
dotnet build src/TaskTimer/TaskTimer.csproj -c Release

# 実行
dotnet run --project src/TaskTimer/TaskTimer.csproj

# 単体実行ファイルとして発行
dotnet publish src/TaskTimer/TaskTimer.csproj -c Release -r win-x64 --self-contained -p:PublishSingleFile=true -o ./publish
```

## 設定ファイル

設定は `%APPDATA%/TaskTimer/settings.json` に自動保存されます。

| 設定項目 | 説明 | デフォルト値 |
|---------|------|------------|
| IdleThresholdSeconds | アイドル判定までの秒数 | 300 (5分) |
| ProcessCheckIntervalSeconds | プロセス監視間隔 | 5秒 |
| AlwaysOnTop | 常に最前面表示 | false |
| GitLabDomain | GitLabオンプレドメイン | gitlab.example.com |
| CsvOutputDirectory | CSV出力先 | マイドキュメント/TaskTimer |
