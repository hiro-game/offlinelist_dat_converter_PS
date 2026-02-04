# OfflineList Dat Converter

---

## このアプリについて
**OfflineList Dat Converter** は  
**Microsoft Copilot と PowerShell/WPF によって構築された Copilot 製アプリ**です。

OfflineList DAT（XML / XML(zip)）を CSV に変換し、  
その CSV を編集したうえで **再び CSV → XML / XML(zip)** に戻して  
OfflineList で利用できる形に整えるためのツールです。

Windows 11 + PowerShell 7.5.4 で動作確認済み。  
PowerShell 5.1 でも動作します。

---

## スクリーンショット
![OfflineList Dat Converter]( "アプリウィンドウ")

---

## 特徴

- **OfflineList DAT（XML / XML(zip)）→ CSV → XML / XML(zip)** の往復変換に対応
- **WPF GUI による直感的な操作**
- **OfflineList DAT の仕様に準拠した XML を生成**
  - `<system>`
  - `<imFolder>`
  - `<screenshotsWidth>` / `<screenshotsHeight>`
  - `<newDat>`（`datVersionURL` / `datURL` / `imURL`）
- **出力形式は XML / XML(zip) の 2 種類**
- **ZIP 出力は 7za.exe / 7z.exe がある場合のみ対応**
- **ZIP の入力は PowerShell 標準で対応**
- **7Z の入力は 7za.exe / 7z.exe がある場合のみ対応**

---

## 使用方法

### 1. 基本的な流れ（全体像）

1. **OfflineList の DAT（XML / XML(zip)）を用意する**
2. 本アプリで **XML / XML(zip) → CSV に変換**
3. CSV をエディタ等で編集する
4. 編集済み CSV を本アプリで **CSV → XML / XML(zip)** に再変換
5. 生成された DAT（XML / XML(zip)）を OfflineList で使用する

---

### 2. XML / XML(zip) → CSV 変換

1. アプリ下部の **「入力ファイル」欄** に  
   - DAT の XML ファイル  
   - または XML を含む ZIP（XML(zip)）  
   を **ドラッグ＆ドロップ** するか、ファイル選択ボタンから指定します。
2. **出力先フォルダ** を指定します。  
   - 指定しない場合は、**入力ファイルと同じフォルダ**に CSV が出力されます。
3. 「変換」ボタンを押すと、  
   **`<system>.csv` または 元ファイル名ベースの CSV** が生成されます。
4. 生成された CSV をエディタ等で編集します。

---

### 3. CSV → XML / XML(zip) 変換

1. アプリ下部の **「入力ファイル」欄** に  
   編集済みの **CSV ファイル**を  
   **ドラッグ＆ドロップ** するか、ファイル選択ボタンから指定します。
2. 上部の設定欄で必要項目を入力します：
   - **System**  
     - OfflineList の `<system>` 要素に入る値  
     - 例：`md`、`snes` など
   - **DAT 名 / バージョン**  
     - `<datName>` / `<datVersion>` に反映されます
   - **baseURL**  
     - `datVersionURL` / `datURL` / `imURL` の共通プレフィックス  
     - 例：`https://example.com/dat/`
   - **imFolder**  
     - `<imFolder>` に出力されるフォルダ名  
     - 例：`mdimg`
   - **screenshotsWidth / screenshotsHeight**  
     - `<screenshotsWidth>` / `<screenshotsHeight>` に出力される数値  
     - 例：`320` / `224`
   - **newDat の URL 関連**  
     - `datVersionURL`：DAT バージョン情報ファイル（例：`md_ver.txt`）  
     - `datURL`：DAT 本体（例：`md.zip`）  
     - `imURL`：画像フォルダ（例：`mdimg`）  
     - これらは `baseURL + 各ファイル名/フォルダ名` として XML に出力されます。
3. **出力形式** を選択します：
   - **XML**：プレーンな XML ファイルを出力  
   - **XML (zip)**：XML を単一ファイル ZIP にまとめて出力  
     - ZIP 出力には 7za.exe / 7z.exe が必要です
4. 「変換」ボタンを押すと、  
   - XML の場合：`<元 CSV 名>.xml`  
   - XML(zip) の場合：`<元 CSV 名>.zip`  
   が出力されます。

---

### 4. 出力ファイル名とチェックボックスの挙動

- **出力ファイル名の基本ルール**
  - チェック **未使用時**：  
    → **元の CSV ファイル名** をベースに `<名前>.xml` / `<名前>.zip` を生成
  - チェック **使用時（「DAT ファイル名に合わせる」等）**：  
    → GUI 上の **DAT ファイル名（例：`md.zip`）** をベースに  
      `md.xml` / `md.zip` を生成

- **system と出力ファイル名の関係**
  - `<system>` 要素の値（例：`md`）は  
    **XML 内の `<system>` にのみ影響**し、  
    出力ファイル名とは切り離されています。

---

### 5. 注意事項

- **ZIP 出力には 7za.exe / 7z.exe が必須**です。  
  見つからない場合、XML(zip) 出力はエラーになります。
- **ZIP の入力（解凍）は PowerShell 標準機能で対応**しています。
- **7Z の入力（解凍）は 7za.exe / 7z.exe が必要**です。
- OfflineList の仕様に依存するため、  
  不正な値や空欄が多い場合、OfflineList 側で正しく認識されない可能性があります。
- `baseURL` と `datVersionURL` / `datURL` / `imURL` の組み合わせは  
  実際にアクセス可能な URL になるように設定してください。

---

## 動作要件

- Windows 10 / 11  
- PowerShell 5.1 または PowerShell 7.x  
- ZIP 出力を行う場合は **7za.exe または 7z.exe が必要**  
- 7Z の入力（解凍）にも **7za.exe / 7z.exe が必要**

---

## インストール

特別なインストールは不要です。  
スクリプトを任意のフォルダに配置して実行するだけで動作します。

### ショートカットで使用する場合

```powershell
pwsh -WindowStyle Hidden -ExecutionPolicy Bypass -File .\OLDATCONV.ps1
ライセンス
本プロジェクトは MIT License のもとで公開されています。

注意事項
本アプリは OfflineList DAT の仕様に基づいて動作します

ZIP 出力には 7za.exe  / 7z.exe  が必要です

7Z の入力（解凍）にも 7za.exe  / 7z.exe  が必要です

本アプリは Microsoft Copilot の支援により作成されています