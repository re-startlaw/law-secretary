# システム管理業務ルール

事務所のMac・NAS・自動化スクリプト・LaunchAgent等の運用管理を扱う業務。

## 基本方針
- **既存稼働中のシステムは原則触らない。** 設定変更・ログ確認・復元手順の提示を中心に対応する。
- **システム領域（`~/`配下、`~/Library/LaunchAgents/`、`/etc/`等）への変更は必ず米谷尚起の承認を得てから実行する。**
- 破壊的操作（`launchctl unload`、`rm`、plist書き換え等）の前に現状のバックアップ有無を確認する。
- トラブル時はまずログを確認し、ユーザー指示なしに設定変更しない。

## 運用中システム一覧

### Pocket3自動バックアップ（Mac）

#### 概要
OsmoPocket3のSDカードをUSBリーダー経由でMacに挿すと、自動でUGREEN NASにコピーされ、SDカードから削除されるシステム。

#### 環境
- **NAS**: UGREEN DXP2800 (192.168.0.209)
- **SMB保存先**: `/Volumes/personal_folder/Photos/osmo pocket/YYYYMMDD/`
- **SDカードマウント名**: `/Volumes/SD_Card`
- **SDカード内パス**: `/Volumes/SD_Card/DCIM/DJI_001`
- **UGOSユーザー名**: `komecome0720@gmail.com`

#### 構成ファイル
- `~/pocket3_backup.sh` — バックアップスクリプト本体（zsh）
- `~/Library/LaunchAgents/com.pocket3.backup.plist` — `/Volumes`監視のLaunchAgent
- `~/pocket3_backup.log` — 実行ログ
- `~/.pocket3_backup.lock` — 多重起動防止ロック
- `~/.pocket3_sd_processed` — 処理済みSDカードUUID保存
- `~/.pocket3_cooldown` — クールダウンタイムスタンプ

#### 仕組み
1. `/Volumes`の変化をLaunchAgentが検知
2. 5秒待機してSDカードのマウント完了を待つ
3. NAS書き込みテスト（成功時のみ続行）
4. 対象ファイル（MP4/MOV/JPG/DNG/LRF）をコピー
5. 元サイズと宛先サイズを比較検証
6. サイズ一致のファイルのみSDカードから削除
7. 処理済みSDカードUUIDを記録
8. 通知を表示

#### 安全機構
- NAS書き込みテスト（失敗時は削除しない）
- サイズ検証による削除前確認
- 60秒クールダウン（無限ループ防止）
- 10分のロックファイルによる多重起動防止
- SDカードUUIDベースの処理済みフラグ

#### 必要な権限
システム設定→プライバシーとセキュリティ→フルディスクアクセス:
- ターミナル
- bash（/bin/bash）

#### 復元手順（新Mac用）
1. NAS上の`/personal_folder/Photos/scripts/pocket3_setup.sh`を実行
2. フルディスクアクセス権限を付与
3. SMB接続で`smb://192.168.0.209`にアクセスしてパスを有効化

#### トラブルシューティング
- 通知が来ない → `~/pocket3_backup.log`を確認
- フルディスクアクセス権限の再付与が必要な場合あり
- LaunchAgent再読み込み: `launchctl unload/load ~/Library/LaunchAgents/com.pocket3.backup.plist`
- SDが処理済み扱いで止まる → `rm ~/.pocket3_sd_processed` で解除
- クールダウンで止まる → `rm ~/.pocket3_cooldown` で解除

## システム変更時のチェックリスト
- [ ] 現状の設定ファイル・スクリプト内容を確認（Read）
- [ ] 現状の稼働状態を確認（`launchctl list`、ログ末尾）
- [ ] 変更内容を米谷尚起に提示して承認を得る
- [ ] 変更後は動作確認（テスト実行またはログ監視）
- [ ] 変更内容を本ファイルに追記
