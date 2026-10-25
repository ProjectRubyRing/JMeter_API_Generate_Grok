# JMeter APIテスト 利用方法

> 対象: Apache JMeter 5.6.3 ／ RHEL 9.8 ／ Windows 11 ／ ツール: jmeter_api_tool v1.0.0（2026-10-24 出力）

本書は、REST API の定義（URL・HTTPメソッド・クエリパラメータ・ヘッダ・リクエストボディ）を Excel または CSV に記入し、JMeter のテスト計画ファイル（JMX）を自動生成して実行し、結果を確認するまでの手順をまとめたものです。

## 1.概要
<!-- summary: ツールでできること・ファイル構成・作業の流れ -->

### このツールでできること
- Excel / CSV に記入した API 定義から、JMeter 5.6.3 用の JMX ファイルを自動生成します。
- 生成される JMX には、API テストでよく使う要素（待機タイマ、連番・乱数、CSV / Excel データ読込、クッキー、前処理・後処理、正規表現 / JSON 抽出、応答の検証、結果出力）があらかじめ組み込まれています。
- 既定は「1スレッド × 1回」の実行です。実行時のプロパティ指定（例: -JTHREADS=10 -JLOOPS=100）だけで、大量件数の負荷試験に切り替えられます。
- 実行結果は JTL（CSV）、HTML レポート、jmeter.log、結果ログ（api_result.log）、エラー詳細（error_detail.xml）に出力されます。
- 既存の JMX を -t で指定すると、新規の JMX は作らず、Excel / CSV / OpenAPI の定義との差分だけをその JMX へ更新します。更新箇所は画面と *_update.txt に表示します。手で追加した要素は残します。--diff（PowerShell は -Diff）を付けると、JMX は書き換えずに差分だけを表示します。
- cli コマンドは、既存の JMX から Linux 用 CLI 手順（*_cli.txt）を作り直します。手順書には環境変数、nohup によるバックグラウンド実行、jmeter -g によるレポート再生成も含みます。
- 生成される要素の名前には [必須] / [任意] を付けています。[任意] の要素は JMeter GUI で右クリック →「削除」するだけで、他の要素に影響を与えずに取り除けます。

### ファイル構成
<!-- cols: 36,74 -->
| パス | 内容 |
|---|---|
| jmeter_api_tool.sh | RHEL 9.8 用スクリプト（bash + python3 標準ライブラリのみ使用） |
| jmeter_api_tool.ps1 | Windows 11 用スクリプト（Windows PowerShell 5.1 / PowerShell 7 で動作） |
| jmeter_api_tool.bat | Windows 用の起動ファイル（実行ポリシーを一時的に回避して .ps1 を起動） |
| input/api_definition.xlsx | 入力テンプレート（Excel）：「API一覧」「設定」「記入方法」シート |
| input/api_definition.csv | 入力テンプレート（CSV）：API 一覧 |
| input/api_definition_settings.csv | 入力テンプレート（CSV）：設定 |
| input/openapi_config.csv | OpenAPI 定義（.json）から生成する場合の設定（接続先・HTTP ヘッダ・アクセストークン・定義の変数 ${public_base_url} / ${name_prefix} の値）。generate の -c / -OpenApiConfig で指定 |
| output/api_test.jmx | 生成された JMX（サンプル） |
| output/data/testdata.csv | CSV データ読込のサンプルデータ |
| output/data/testdata.xlsx | Excel データ読込のサンプルデータ |
| docs/JMeter_Usage_Guide.xlsx / .md | 本書（manual コマンドで出力） |
| result/JMX名_日時/ | run コマンドの実行結果（JTL・ログ・HTML レポート） |

### 作業の流れ
1. テンプレートを作成する（init コマンド）
2. API 定義と設定を記入する（Excel または CSV）
3. JMX を生成する（generate コマンド）
4. JMeter GUI で内容を確認し、不要な [任意] 要素を削除・調整する
5. CLI モードで実行する（run コマンド）
6. 結果（HTML レポート・JTL・ログ）を確認する

> ファイル名・フォルダ名はすべて英数字にしています。Windows と Linux の間でファイルを転送したときの文字化けを防ぐためです。

## 2.環境構築
<!-- summary: Java・JMeter 5.6.3 のインストールと動作確認（RHEL 9.8 / Windows 11） -->

### 必要なソフトウェア
<!-- cols: 18,30,30,40 -->
| ソフトウェア | RHEL 9.8 | Windows 11 | 備考 |
|---|---|---|---|
| Java | java-17-openjdk（dnf） | JDK 17（Eclipse Temurin 等） | JMeter 5.6.3 は Java 8 以上で動作。17 を推奨 |
| Apache JMeter | 5.6.3（.tgz） | 5.6.3（.zip） | https://archive.apache.org/dist/jmeter/binaries/ |
| スクリプト実行環境 | bash・python3（OS 標準） | Windows PowerShell 5.1（OS 標準） | 追加のモジュール・ライブラリは不要 |
| Excel | 不要 | 不要（記入・閲覧には推奨） | xlsx の読み書きはスクリプト自身が行います |

### RHEL 9.8 へのインストール
```bash
# Java 17（GUI を使わないサーバーなら java-17-openjdk-headless でも可）
sudo dnf install -y java-17-openjdk
java -version

# JMeter 5.6.3（ダウンロード・チェックサム確認・展開）
cd /opt
sudo curl -LO https://archive.apache.org/dist/jmeter/binaries/apache-jmeter-5.6.3.tgz
sudo curl -LO https://archive.apache.org/dist/jmeter/binaries/apache-jmeter-5.6.3.tgz.sha512
sha512sum -c apache-jmeter-5.6.3.tgz.sha512
sudo tar xzf apache-jmeter-5.6.3.tgz

# 環境変数（~/.bashrc に追記）
echo 'export JMETER_HOME=/opt/apache-jmeter-5.6.3' >> ~/.bashrc
echo 'export PATH=$JMETER_HOME/bin:$PATH' >> ~/.bashrc
source ~/.bashrc
jmeter -v
```
- python3 は RHEL 9 のパッケージ管理（dnf）が使用しているため、通常は導入済みです（確認: `python3 --version`）。無い場合は `sudo dnf install -y python3` で導入します。
- ツールのファイルを Windows から転送して改行コードが CRLF になった場合は、`sed -i 's/\r$//' jmeter_api_tool.sh` で LF に戻してください。
- GUI を使う場合に日本語が □ で表示されるときは、日本語フォントを導入します（例: `sudo dnf install -y google-noto-sans-cjk-ttc-fonts`）。

### Windows 11 へのインストール
1. JDK 17 をインストールします（例: Eclipse Temurin 17 の .msi。インストール時に「JAVA_HOME 変数の設定」と「PATH への追加」を有効にします）。
2. https://archive.apache.org/dist/jmeter/binaries/apache-jmeter-5.6.3.zip をダウンロードし、ファイルのプロパティで「許可する（ブロック解除）」にチェックしてから、C:\apache-jmeter-5.6.3 などスペースや日本語を含まないフォルダに展開します。
3. 環境変数 JMETER_HOME に展開先のフォルダを設定します（下記のコマンド、または「システムの詳細設定」→「環境変数」）。
4. PowerShell を開き直して動作を確認します。

```powershell
[Environment]::SetEnvironmentVariable('JMETER_HOME', 'C:\apache-jmeter-5.6.3', 'User')
# PowerShell を開き直してから確認
java -version
& "$env:JMETER_HOME\bin\jmeter.bat" -v
```

### ツールの配置と動作確認
ツール一式を任意のフォルダ（例: RHEL は ~/JMeter_API_Generate、Windows は C:\work\JMeter_API_Generate）に配置し、ヘルプが表示されることを確認します。

```bash
# RHEL 9.8
cd ~/JMeter_API_Generate
chmod +x jmeter_api_tool.sh
./jmeter_api_tool.sh help
```

```powershell
# Windows 11（PowerShell またはコマンドプロンプト）
cd C:\work\JMeter_API_Generate
.\jmeter_api_tool.bat help
```
- jmeter_api_tool.bat は PowerShell の実行ポリシー（既定では .ps1 の実行が禁止）を、その実行に限って回避して起動します。PowerShell から直接実行する場合は `powershell -ExecutionPolicy Bypass -File .\jmeter_api_tool.ps1 help` としてください。
- インターネットから取得した ZIP を展開した場合は、事前に `Get-ChildItem -Recurse | Unblock-File` でブロックを解除しておくと確実です。

## 3.JMX生成
<!-- summary: テンプレート作成から JMX 生成・GUI での確認までの手順 -->

### 手順1：テンプレートを作成する
```bash
./jmeter_api_tool.sh init          # RHEL 9.8
```
```powershell
.\jmeter_api_tool.bat init         # Windows 11
```
input フォルダに次の 4 ファイルが作成されます（既存ファイルは上書きしません。上書きする場合は -f / -Force を付けます）。
- api_definition.xlsx（Excel 形式のテンプレート）
- api_definition.csv（CSV 形式：API 一覧）
- api_definition_settings.csv（CSV 形式：設定）
- openapi_config.csv（CSV 形式：OpenAPI 定義から生成する場合の接続先・HTTP ヘッダ・アクセストークン・定義の変数 ${public_base_url} / ${name_prefix} の値。generate の -c / -OpenApiConfig で指定）

### curl コマンドから API 定義へ 1 行追加する
curl コマンドをファイルに書いて `curl` コマンドを実行すると、既存の `api_definition.xlsx` の最初の空白行と、同じ名前の `api_definition.csv` の末尾に 1 行追加します。curl 1 件が 1 行です。ファイルには複数コマンドを書けます。行継続（`\`、`^`、`` ` ``）と引用符に対応します。既存の API 行と「設定」シートは変更しません。

```bash
./jmeter_api_tool.sh curl -i requests.curl -o input/api_definition.xlsx
./jmeter_api_tool.sh generate -i input/api_definition.xlsx
```
```powershell
.\jmeter_api_tool.bat curl -i requests.curl -o input\api_definition.xlsx
.\jmeter_api_tool.bat generate -i input\api_definition.xlsx
```
- URL の `?` 以降は「クエリパラメータ」列に分けます。`-G` と `--data-urlencode` もクエリになります。
- `-d` / `--json` / `-F` は本文とボディ形式にします。`-d` で Content-Type を書いていない場合は、curl と同じ `application/x-www-form-urlencoded` として取り込みます（JSON に見える文字列でも `application/json` にはしません）。
- もう一度実行すると、その続きの空白行（CSV は末尾）にさらに 1 行追加します。
- ファイルが無い場合は、テンプレートの記入例の次の空白行から追加します。
- 期待ステータスは空欄です。必要なら記入してから `generate` してください。

### 手順2：API 定義を記入する
- Excel 形式：「API一覧」シートに 1行 = 1API で記入し、「設定」シートの「設定値」列で全体設定を変更します。
- CSV 形式：api_definition.csv に API を、api_definition_settings.csv に設定を記入します。
- テンプレートには記入例として、postman-echo.com（送信した内容をそのまま返す公開サービス）を呼び出す 4 件の API が入っています。自分の API に書き換えてください。
- 入力項目の詳細は「4.入力項目_設定」「5.入力項目_API」シートを参照してください。

> CSV を Excel で保存する場合は「CSV UTF-8（コンマ区切り）」を選んでください。Shift_JIS の「CSV（コンマ区切り）」で保存したファイルも自動判別して読み込めます。

### 手順3：JMX を生成する
```bash
# RHEL 9.8
./jmeter_api_tool.sh generate -i input/api_definition.xlsx
./jmeter_api_tool.sh generate -i input/api_definition.csv -s input/api_definition_settings.csv
```
```powershell
# Windows 11
.\jmeter_api_tool.bat generate -i input\api_definition.xlsx
.\jmeter_api_tool.bat generate -i input\api_definition.csv -s input\api_definition_settings.csv
```

<!-- cols: 22,30,66 -->
| sh のオプション | PowerShell のオプション | 説明 |
|---|---|---|
| -i ファイル | -InputFile（-i） | 入力ファイル（.xlsx / .csv / .json）。必須 |
| -s ファイル | -SettingsFile（-s） | 設定 CSV。省略時は「入力ファイル名_settings.csv」を探し、無ければ既定値を使用（Excel 入力時は「設定」シートを使用） |
| -o フォルダ | -OutDir（-o） | 出力先フォルダ（既定: output） |
| -n ファイル名 | -JmxName（-n） | 出力する JMX のファイル名（例: order_test.jmx）。設定「出力JMXファイル名」（output_file）より優先します |
| --split | -Split | API ごとに別々の JMX を出力する |
| -f | -Force（-f） | サンプルデータ（data/testdata.csv / .xlsx）が既にあっても上書きする |
| -t ファイル | -Jmx（-t） | 既存の JMX。指定すると新規作成せず、Excel / CSV / OpenAPI の定義との差分だけをそのファイルへ更新する。更新箇所は画面と *_update.txt に表示する。--split とは同時に使えない |
| --diff | -Diff | 既存 JMX との差分箇所だけを表示し、JMX は書き換えない。-t を省略したときは、出力先にある同名の JMX を対象にする |

### 生成されるファイル
<!-- cols: 34,84 -->
| ファイル | 説明 |
|---|---|
| output/出力JMXファイル名 | 生成された JMX。同名のファイルがある場合は「.日時.bak」を付けて退避してから作成します。-t で既存の JMX を指定したときは、退避せず差分だけを更新します |
| output/出力JMXファイル名_cli.txt | Linux サーバで JMeter CLI を実行するコマンドと、各パラメータの説明。JMX と同じフォルダに出力する |
| output/出力JMXファイル名_update.txt | -t で既存の JMX を更新したときの更新箇所。差分が無いときも出力する |
| output/data/testdata.csv | CSV データ読込用のサンプル（ファイルが無い場合のみ作成） |
| output/data/testdata.xlsx | Excel データ読込用のサンプル（ファイルが無い場合のみ作成） |

既存の JMX を更新する場合は、-t にそのファイルを指定します。Excel、CSV、OpenAPI のどれからでも同じです。新しい JMX は作らず、定義と違う箇所だけを書き換えます。更新・追加・削除した箇所は画面に表示し、JMX と同じフォルダの *_update.txt にも書きます。名前が [必須] / [任意] で始まらない要素、および APIxx ではないサンプラは、手で追加したものとして残します。差分が無いときは JMX を書き換えません。

```bash
./jmeter_api_tool.sh generate -i input/api_definition.xlsx -t output/api_test.jmx
./jmeter_api_tool.sh generate -i openapi.json -t output/api_test.jmx
```
```powershell
.\jmeter_api_tool.bat generate -i input\api_definition.xlsx -t output\api_test.jmx
.\jmeter_api_tool.bat generate -i openapi.json -t output\api_test.jmx
```

### 手順4：JMeter GUI で確認・調整する
- RHEL（デスクトップ環境あり）：`jmeter -t output/api_test.jmx`
- Windows：`%JMETER_HOME%\bin\jmeter.bat` をダブルクリックし、「File」→「Open」で JMX を選択
- 不要な [任意] 要素は右クリック →「Remove」で削除、一時的に外す場合は右クリック →「Disable」で無効化できます。
- GUI モードは動作確認・デバッグ用です。負荷試験は必ず CLI モード（run コマンド）で実行してください。

## 4.入力項目_設定
<!-- summary: 「設定」シート（設定 CSV）の全項目と既定値 -->
「設定」シート（CSV の場合は api_definition_settings.csv）の項目です。設定値が空欄の項目は既定値が使われます。

- ON / OFF / DISABLE の項目は、ON = 生成して有効、OFF = 生成しない、DISABLE = 生成するが無効化（GUI で右クリック →「Enable」ですぐ使える）を表します。
- 説明欄の「-J名前」は、実行時に -J名前=値（run コマンドでは -J 名前=値 / -Prop 名前=値）で値を変更できることを表します。

<!-- cols: 12,28,22,16,62 -->
| 区分 | 設定キー | 設定項目 | 既定値 | 説明 |
|---|---|---|---|---|
| 基本 | `test_plan_name` | テストプラン名 | APIテスト | JMeter 上に表示されるテストプラン名 |
| 基本 | `output_file` | 出力JMXファイル名 | api_test.jmx | 生成する JMX のファイル名（英数字・_・-・. 以外の文字は _ に置き換えます） |
| 基本 | `base_url` | ベースURL | （空欄） | API一覧の URL が「/」で始まる場合に先頭に付ける URL（例: https://api.example.com/v1） |
| 基本 | `content_encoding` | 文字コード | UTF-8 | リクエストの文字コード |
| 負荷 | `threads` | スレッド数 | 1 | 同時に実行する仮想ユーザー数。-JTHREADS |
| 負荷 | `ramp_up` | ランプアップ(秒) | 1 | 全スレッドが起動し終わるまでの秒数。-JRAMP_UP |
| 負荷 | `loops` | ループ回数 | 1 | 各スレッドが API 一覧を繰り返す回数。-1 = 無限（実行時間と併用）。-JLOOPS |
| 負荷 | `duration` | 実行時間(秒) | （空欄） | 指定するとスケジューラを有効にし、この秒数で終了（ループ回数は -1 を推奨）。-JDURATION |
| 負荷 | `on_sample_error` | エラー時の動作 | continue | continue = 続行 / startnextloop = 次のループへ / stopthread = スレッド停止 / stoptest = テスト停止 / stoptestnow = 即時停止 |
| 負荷 | `throughput_timer` | スループット制御 | DISABLE | 1 分あたりのリクエスト数を一定に保つ Constant Throughput Timer |
| 負荷 | `throughput_per_min` | 目標スループット(件/分) | 600 | スループット制御の目標値（メイン処理全体の合計）。-JTHROUGHPUT |
| HTTP | `http_defaults` | HTTPリクエスト初期値 | ON | 文字コード・タイムアウトの共通設定（HTTP Request Defaults） |
| HTTP | `connect_timeout` | 接続タイムアウト(ms) | 10000 | -JCONNECT_TIMEOUT |
| HTTP | `response_timeout` | 応答タイムアウト(ms) | 30000 | -JRESPONSE_TIMEOUT |
| HTTP | `common_headers` | 共通ヘッダ | （空欄） | 全 API に付けるヘッダ。「名前: 値」を改行区切り（テンプレートの記入例: Accept: application/json / User-Agent: JMeter-API-Test/1.0 / X-Client-Id: ${CLIENT_ID}） |
| HTTP | `access_token_dir` | アクセストークンのフォルダ | （空欄） | 指定すると、このフォルダ配下のファイルからアクセストークンを読み込み、共通ヘッダに Authorization: Bearer ${ACCESS_TOKEN} を付けます（共通ヘッダに Authorization がある場合はその値を使用）。空欄 = 使わない。JMX からの相対パスまたは絶対パス（区切りは \ と / のどちらでも可。; で区切って複数書くと、存在する最初のフォルダを使用）。-JACCESS_TOKEN_DIR |
| HTTP | `access_token_file` | アクセストークンのファイル名 | access_token.txt | アクセストークンを書いたファイルの名前（フォルダからの相対パス）。*.token のように * を含めると、一致するファイルのうち更新日時が最も新しいファイルを使います。前後の空白・改行は除き、実行中にファイルが更新されると読み直します。-JACCESS_TOKEN_FILE |
| HTTP | `follow_redirects` | リダイレクト追従 | true | true = 3xx 応答の Location を自動で追従 |
| HTTP | `keep_alive` | KeepAlive | true | true = 接続を再利用する |
| HTTP | `default_wait_ms` | 既定の待機時間(ms) | 0 | API一覧の「待機時間」が空欄の API に使う値（0 = 待機なし） |
| クッキー | `cookie_manager` | クッキー管理 | ON | HTTP Cookie Manager（Set-Cookie を保存して以降のリクエストで送信） |
| クッキー | `cookie_clear_each_iteration` | 反復ごとにクリア | true | true = ループのたびにクッキーを消去（毎回新しいセッション） |
| クッキー | `cookies` | 事前設定クッキー | （空欄） | 最初から送信するクッキー。「名前=値」を改行区切り（テンプレートの記入例: LANG=ja） |
| クッキー | `cookie_domain` | クッキーのドメイン | （空欄） | 事前設定クッキーのドメイン。空欄 = 1 つ目の接続先（${HOST}） |
| 変数 | `user_variables` | 共通変数 | （空欄） | 「名前=値」を改行区切り。${名前} で参照でき、実行時に -J名前=値 で変更可（テンプレートの記入例: ENV=staging / CLIENT_ID=jmeter-client） |
| 変数 | `counter` | 連番 | ON | 反復ごとに連番を設定する Counter |
| 変数 | `counter_name` | 連番の変数名 | SEQ | 連番を入れる変数名 |
| 変数 | `counter_start` | 連番の開始値 | 1 |  |
| 変数 | `counter_increment` | 連番の増分 | 1 |  |
| 変数 | `counter_format` | 連番の書式 | 000000 | 0 埋めの桁数など（Java の DecimalFormat 形式）。空欄 = 書式なし |
| 変数 | `counter_per_thread` | スレッドごとの連番 | false | true = スレッドごとに開始値から採番 / false = 全スレッドで一意 |
| 変数 | `random_variable` | 乱数 | ON | 反復ごとに乱数を設定する Random Variable |
| 変数 | `random_name` | 乱数の変数名 | RANDOM_NO |  |
| 変数 | `random_min` | 乱数の最小値 | 1 |  |
| 変数 | `random_max` | 乱数の最大値 | 99999 |  |
| CSVデータ | `csv_data` | CSVデータ読込 | ON | CSV Data Set Config（1 反復につき 1 行を読み込み、列を変数に設定） |
| CSVデータ | `csv_file` | CSVファイル | data/testdata.csv | JMX からの相対パスまたは絶対パス。-JCSV_FILE |
| CSVデータ | `csv_variables` | CSVの変数名 | （空欄） | カンマ区切り。空欄 = 生成時に CSV ファイルの 1 行目を読み取って設定 |
| CSVデータ | `csv_encoding` | CSVの文字コード | UTF-8 |  |
| CSVデータ | `csv_delimiter` | CSVの区切り文字 | , |  |
| CSVデータ | `csv_recycle` | 終端で先頭に戻る | true | true = 最終行の次は先頭行に戻る |
| CSVデータ | `csv_stop_thread` | 終端でスレッド停止 | false | true = 最終行まで読んだらスレッドを停止（csv_recycle=false と併用） |
| CSVデータ | `csv_share_mode` | CSVの共有範囲 | all | all = 全スレッドで共有 / group = スレッドグループ単位 / thread = スレッドごとに先頭から |
| Excelデータ | `excel_data` | Excelデータ読込 | ON | Excel（.xlsx）の読み込み。Apache POI などの追加ライブラリは不要 |
| Excelデータ | `excel_file` | Excelファイル | data/testdata.xlsx | JMX からの相対パスまたは絶対パス。-JEXCEL_FILE |
| Excelデータ | `excel_sheet` | シート名 | （空欄） | 空欄 = 先頭のシート。-JEXCEL_SHEET |
| 前後処理 | `setup_teardown` | 開始・終了処理 | ON | テスト開始・終了時のログ出力と後片付け（setUp / tearDown Thread Group） |
| 前後処理 | `common_preprocessor` | 共通前処理 | ON | 全 API の直前に実行する JSR223 PreProcessor（REQUEST_ID・TIMESTAMP 等を生成） |
| 前後処理 | `common_postprocessor` | 共通後処理 | ON | 全 API の直後に実行する JSR223 PostProcessor（LAST_STATUS 等を設定） |
| 結果出力 | `result_dir` | 結果出力フォルダ | ~/result | JTL・結果ログの出力先。~/ = JMX のフォルダ。-JRESULT_DIR（run コマンドは自動設定） |
| 結果出力 | `view_results_tree` | 結果ツリー | ON | View Results Tree（GUI での確認用） |
| 結果出力 | `summary_report` | 集計レポート | ON | Summary Report（GUI での確認用） |
| 結果出力 | `aggregate_report` | 統計レポート | ON | Aggregate Report（GUI での確認用。パーセンタイル） |
| 結果出力 | `jtl_writer` | JTL出力(全件) | ON | 全リクエストの結果を result_all.jtl（CSV）に出力 |
| 結果出力 | `error_writer` | エラー詳細出力 | ON | 失敗したリクエストのみ error_detail.xml（リクエスト・レスポンス全文）に出力 |
| 結果出力 | `result_log` | 結果ログ出力 | ON | 1 リクエスト 1 行の結果ログ api_result.log を出力（JSR223 Listener） |
| 結果出力 | `log_response` | 結果ログの本文出力 | error | none = 出力しない / error = 失敗時のみ / all = 全件。-JLOG_RESPONSE |
| 結果出力 | `log_max_chars` | 本文の最大文字数 | 1000 | 結果ログに出力する本文の最大文字数。-JLOG_MAX_CHARS |
| 結果出力 | `debug_sampler` | 変数確認(デバッグ) | DISABLE | Debug Sampler（全変数を結果ツリーに表示） |

## 5.入力項目_API
<!-- summary: 「API一覧」シート（API CSV）の列と記入ルール -->
「API一覧」シート（CSV の場合は api_definition.csv）の列です。1行 = 1API で、上から順に実行されます。

<!-- cols: 18,6,62,36 -->
| 列名 | 必須 | 説明 | 記入例 |
|---|---|---|---|
| No |  | 実行順の番号（ラベル APIxx に使用）。空欄は自動採番。# で始まる行は読み飛ばす | `1` |
| 有効 |  | Y = 有効、N = 無効化した状態で生成。空欄 = Y | `Y` |
| API名 |  | HTTP リクエストの名前（レポートのラベル）。空欄はメソッドとパス | `ログイン` |
| メソッド | ○ | HTTP メソッド（GET / POST。PUT / PATCH / DELETE も可） | `POST` |
| URL | ○ | 完全な URL、またはベースURL 設定時は「/」で始まるパス。${変数} 可。? 以降のクエリも可 | `https://api.example.com/v1/users` |
| クエリパラメータ |  | URL のクエリパラメータ | `page=1`<br>`size=20` |
| ヘッダ |  | この API だけに付けるヘッダ（全 API 共通のヘッダは設定 common_headers） | `Authorization: Bearer ${AUTH_TOKEN}` |
| ボディ形式 |  | リクエストボディの形式。空欄 = auto（自動判定） | `json` |
| リクエストボディ |  | POST 等で送る本文。form 形式は「名前=値」を改行 / & 区切り | `{"userId":"${USER_ID}"}` |
| 待機時間(ms) |  | API 呼び出し前に待つ時間（Constant Timer）。空欄 = 設定 default_wait_ms、0 = 待機なし | `1000` |
| 期待ステータス |  | 期待する HTTP ステータス（カンマ区切りで複数可）。4xx / 5xx を書くと異常系テストとして扱う | `200` |
| 応答検証文字列 |  | 応答本文に含まれるべき文字列。複数書くとすべて含む場合に成功 | `"result":"OK"` |
| 応答時間上限(ms) |  | この時間を超えたら失敗（Duration Assertion） | `3000` |
| 正規表現抽出 |  | 応答本文から ( ) 内を抽出して変数に設定。見つからない場合は NOT_FOUND | `TOKEN="token"\s*:\s*"([^"]+)"` |
| JSON抽出 |  | 応答 JSON から値を抽出して変数に設定。見つからない場合は NOT_FOUND | `USER_ID=$.data.id` |
| 前処理スクリプト |  | API 呼び出しの直前に実行する Groovy（JSR223 PreProcessor） | `vars.put('X', '1')` |
| 後処理スクリプト |  | API 呼び出しの直後に実行する Groovy（JSR223 PostProcessor）。抽出の後に実行 | `log.info(prev.getResponseCode())` |
| 備考 |  | JMeter のコメント欄に設定 |  |

### 記入のルール
- 複数の値を書く列（クエリパラメータ、ヘッダ、正規表現抽出など）は、セル内改行（Excel では Alt + Enter）で区切ります。クエリパラメータと form 形式のボディは & 区切りでも記入できます。
- 先頭の列（No）が # で始まる行はコメントとして読み飛ばします（テンプレートの 4 行目の記入説明など）。
- URL・パラメータ・ヘッダ・ボディには ${変数名} で JMeter の変数・関数を埋め込めます（例: `${USER_ID}`、`${SEQ}`、`${__UUID()}`）。
- 「有効」を N にした API は、生成はされますが無効化された状態になります。
- GET / DELETE / HEAD / OPTIONS のクエリパラメータは HTTP リクエストの「Parameters」タブに設定されます。POST / PUT / PATCH の場合は URL のパスに付与されます（値は UTF-8 で URL エンコード。${...} の部分はそのまま）。
- 既に URL エンコード済みの値（%E3%83%86 など）は、二重にエンコードしないよう自動で判定します。
- GET / DELETE ではリクエストボディは使用できません（指定した場合は警告を表示して無視します）。

### ボディ形式（auto の場合の自動判定）
<!-- cols: 26,16,76 -->
| 本文の内容 | 判定 | 送信方法 |
|---|---|---|
| { または [ で始まる | json | 本文をそのまま送信。Content-Type が未指定なら application/json を付与 |
| < で始まる | raw（XML） | 本文をそのまま送信。Content-Type が未指定なら application/xml を付与 |
| 名前=値 の形式（改行 / & 区切り） | form | パラメータとして送信（application/x-www-form-urlencoded） |
| 上記以外 | raw | 本文をそのまま送信（Content-Type はヘッダ列で指定） |

## 6.JMX構成
<!-- summary: 生成される JMX の要素一覧と、削除してよいかの判断基準 -->
生成される JMX のツリー構成です（記入例を生成した場合）。名前が [必須] の要素は削除しないでください。[任意] の要素は、他の要素に影響を与えずに削除できます。

```text
テストプラン
├─ [必須] 共通変数 (User Defined Variables)
├─ [任意] HTTPリクエスト初期値 (HTTP Request Defaults)
├─ [任意] 共通ヘッダ (HTTP Header Manager)
├─ [任意] クッキー管理 (HTTP Cookie Manager)
├─ [任意] 事前処理 (setUp Thread Group)
│   ├─ [任意] テスト開始処理 (JSR223 Sampler)
│   └─ [任意] Excelデータ読込 (JSR223 Sampler)
├─ [必須] メイン処理 (Thread Group)
│   ├─ [任意] 連番 SEQ (Counter)
│   ├─ [任意] 乱数 RANDOM_NO (Random Variable)
│   ├─ [任意] CSVデータ読込 (CSV Data Set Config)
│   ├─ [任意] Excelデータ行取得 (JSR223 PreProcessor)
│   ├─ [任意] 共通前処理 (JSR223 PreProcessor)
│   ├─ [任意] 共通後処理 (JSR223 PostProcessor)
│   ├─ [任意] スループット制御 (Constant Throughput Timer) ※既定は無効
│   ├─ API01_ログイン(POST JSON)  ← HTTP リクエスト（1API = 1つ）
│   │   ├─ [任意] API01 ヘッダ (HTTP Header Manager)
│   │   ├─ [任意] API01 待機 500ms (Constant Timer)
│   │   ├─ [任意] API01 前処理 (JSR223 PreProcessor)
│   │   ├─ [任意] API01 正規表現抽出 ECHO_URL (Regular Expression Extractor)
│   │   ├─ [任意] API01 JSON抽出 AUTH_TOKEN (JSON Extractor)
│   │   ├─ [任意] API01 後処理 (JSR223 PostProcessor)
│   │   ├─ [任意] API01 ステータス検証 200 (Response Assertion)
│   │   ├─ [任意] API01 応答本文検証 (Response Assertion)
│   │   └─ [任意] API01 応答時間検証 5000ms (Duration Assertion)
│   ├─ API02_商品検索(GET) / API03_注文登録(POST form) / API04_クッキー確認(GET)
│   └─ [任意] 変数確認 (Debug Sampler) ※既定は無効
├─ [任意] 事後処理 (tearDown Thread Group)
│   └─ [任意] テスト終了処理 (JSR223 Sampler)
├─ [任意] 結果ツリー (View Results Tree)
├─ [任意] 集計レポート (Summary Report)
├─ [任意] 統計レポート (Aggregate Report)
├─ [任意] JTL出力 全件 (Simple Data Writer)
├─ [任意] エラー詳細出力 XML (Simple Data Writer)
└─ [任意] 結果ログ出力 (JSR223 Listener)
```

### 要素の役割と削除時の影響
<!-- cols: 30,28,52,44 -->
| 要素 | JMeter の要素 | 役割 | 削除したときの影響 |
|---|---|---|---|
| [必須] 共通変数 | User Defined Variables | 接続先（PROTOCOL / HOST / PORT）と共通変数を定義。値は `${__P(名前,既定値)}` 形式で、実行時に -J で変更可能 | 削除不可（各 API が ${HOST} 等を参照） |
| [任意] HTTPリクエスト初期値 | HTTP Request Defaults | 文字コードと接続 / 応答タイムアウト（CONNECT_TIMEOUT / RESPONSE_TIMEOUT） | タイムアウトが無制限になる |
| [任意] 共通ヘッダ | HTTP Header Manager | 全 API に付与するヘッダ（設定 common_headers） | 共通ヘッダが送信されない |
| [任意] クッキー管理 | HTTP Cookie Manager | 応答の Set-Cookie を保存し、以降のリクエストで送信。事前設定クッキーも送信 | クッキーが送信されない（セッションを使う API は失敗） |
| [任意] 事前処理 | setUp Thread Group | テスト開始前に 1 回だけ実行する処理の入れ物 | 配下の処理（Excel 読込など）も削除される |
| [任意] テスト開始処理 | JSR223 Sampler | 実行条件をログに出力し、集計値をリセット | 開始ログが出ない（動作には影響なし） |
| [任意] Excelデータ読込 | JSR223 Sampler | Excel（.xlsx）をメモリに読み込む（Apache POI 不要） | Excel の値が変数に設定されない |
| [必須] メイン処理 | Thread Group | API を実行する。THREADS / RAMP_UP / LOOPS（/ DURATION）で制御 | 削除不可 |
| [任意] 連番 | Counter | 反復ごとに連番を ${SEQ} に設定（全スレッドで一意） | ${SEQ} が未定義になる |
| [任意] 乱数 | Random Variable | 反復ごとに乱数を ${RANDOM_NO} に設定 | ${RANDOM_NO} が未定義になる |
| [任意] CSVデータ読込 | CSV Data Set Config | CSV を 1 反復につき 1 行読み込み、列を変数に設定 | CSV の変数が未定義になる |
| [任意] Excelデータ行取得 | JSR223 PreProcessor | Excel データから反復ごとに 1 行を取り出して変数に設定 | Excel の変数が未定義になる |
| [任意] 共通前処理 | JSR223 PreProcessor | 全 API の直前に実行。${REQUEST_ID} ${TIMESTAMP} ${EPOCH_MS} を生成 | 左記の変数が未定義になる |
| [任意] 共通後処理 | JSR223 PostProcessor | 全 API の直後に実行。${LAST_STATUS} ${LAST_ELAPSED_MS} を設定 | 左記の変数が未定義になる |
| [任意] スループット制御 | Constant Throughput Timer | 1 分あたりの目標リクエスト数（THROUGHPUT）に調整。既定は無効 | なし |
| APIxx_名前 | HTTP Request | API 定義 1 行分のリクエスト | その API が実行されない（配下の要素も削除される） |
| [任意] APIxx ヘッダ | HTTP Header Manager | その API だけに付けるヘッダ | そのヘッダが送信されない |
| [任意] APIxx 待機 | Constant Timer | API 呼び出し前の待機時間 | 待機しなくなる |
| [任意] APIxx 前処理 / 後処理 | JSR223 PreProcessor / PostProcessor | API ごとの前処理・後処理（Groovy） | その処理が行われない |
| [任意] APIxx 正規表現抽出 / JSON抽出 | Regular Expression Extractor / JSON Extractor | 応答から値を取り出して変数に設定 | 後続の API でその変数が未定義になる |
| [任意] APIxx ステータス検証 / 応答本文検証 / 応答時間検証 | Response Assertion / Duration Assertion | 応答の検証。失敗はエラーとして集計 | 検証されなくなる |
| [任意] 変数確認 | Debug Sampler | 変数の一覧を結果ツリーに表示（デバッグ用）。既定は無効 | なし |
| [任意] 事後処理 / テスト終了処理 | tearDown Thread Group / JSR223 Sampler | 合計・成功・失敗件数をログに出力し、結果ログを閉じる | 終了時のサマリが出ない |
| [任意] 結果ツリー / 集計レポート / 統計レポート | View Results Tree / Summary Report / Aggregate Report | GUI での結果確認用（ファイル出力なし） | GUI で結果が見られない（CLI 実行には影響なし） |
| [任意] JTL出力 全件 | Simple Data Writer | 全リクエストの結果を result_all.jtl（CSV）に出力 | JTL が出力されない（run コマンドは別途 result.jtl を出力） |
| [任意] エラー詳細出力 XML | Simple Data Writer | 失敗したリクエストのみ、リクエスト・レスポンス全文を error_detail.xml に出力 | エラー詳細が出力されない |
| [任意] 結果ログ出力 | JSR223 Listener | 1 リクエスト 1 行の結果ログ api_result.log を出力 | 結果ログが出力されない |

### 記入例で使っている変数の流れ
<!-- cols: 30,34,54 -->
| 変数 | 設定する要素 | 使っている場所（記入例） |
|---|---|---|
| PROTOCOL / HOST / PORT | [必須] 共通変数 | 全 API の接続先 |
| SEQ | [任意] 連番 | API01 のボディ、API03 の orderNo、API04 のクエリ |
| RANDOM_NO | [任意] 乱数 | API02 のクエリ rand |
| USER_ID / PASSWORD / KEYWORD | [任意] CSVデータ読込 | API01 のボディ、API02 のクエリ |
| ITEM_CODE / ITEM_NAME / QUANTITY | [任意] Excelデータ行取得 | API02 のクエリ、API03 のボディ |
| REQUEST_ID / TIMESTAMP | [任意] 共通前処理 | API01 のヘッダ・ボディ |
| SIGNATURE | API01 前処理 | API01 のヘッダ X-Signature |
| AUTH_TOKEN | API01 JSON抽出 | API02 / API03 の Authorization ヘッダ |
| ORDER_NO | API03 正規表現抽出 | API03 後処理のログ出力 |
| CLIENT_ID | 共通変数（設定 user_variables） | 共通ヘッダ X-Client-Id |

### 要素を削除・無効化するときのポイント
- 一時的に外す場合は、削除より「Disable」（右クリック → Disable、または Ctrl + T）が安全です。
- 変数を設定している要素を削除すると、その変数を使っている箇所では ${変数名} という文字列がそのまま送信されます。上の「変数の流れ」を確認してから削除してください。
- CSV データ読込は、ファイルが見つからないとスレッドがエラーで停止します。CSV を使わない場合は削除してください。
- 生成時に設定を OFF にすると、その要素は最初から生成されません（「4.入力項目_設定」参照）。

## 7.実行方法
<!-- summary: GUI / CLI での実行方法、run コマンド、実行時に変更できるプロパティ -->

### 実行モードの使い分け
<!-- cols: 16,46,56 -->
| モード | 用途 | コマンド例 |
|---|---|---|
| GUI モード | シナリオの作成・動作確認・デバッグ（少量の実行） | `jmeter -t output/api_test.jmx` |
| CLI モード | 負荷試験・大量件数の実行・定期実行 | `jmeter -n -t output/api_test.jmx -l result.jtl -e -o report` |

### Linux サーバで JMeter CLI を実行する
`generate` は JMX と同じフォルダに `*_cli.txt` を出力します。JMX とデータファイルを Linux サーバへコピーし、このテキストの「実行コマンド」を JMX と同じフォルダで実行します。サーバにこのツールが無くても、Java と JMeter 5.6.3 があれば実行できます。コマンドには、その JMX の接続先・スレッド数・タイムアウトなどの `-J` プロパティを、既定値つきで書いてあります。各パラメータの意味は同じファイルの「パラメータの説明」にあります。

### run コマンド（推奨）
run コマンドは CLI モードで JMeter を実行し、結果フォルダ（result/JMX名_日時/）に JTL・ログ・HTML レポートをまとめて出力し、最後に集計結果を表示します。

```bash
# RHEL 9.8
# 既定（1スレッド × 1回）
./jmeter_api_tool.sh run -t output/api_test.jmx
# 10 スレッド × 100 回（1 API あたり 1,000 件）、ランプアップ 10 秒
./jmeter_api_tool.sh run -t output/api_test.jmx --threads 10 --loops 100 --rampup 10
# 接続先を切り替え、ヒープを 2GB にして実行
./jmeter_api_tool.sh run -t output/api_test.jmx --threads 50 --loops 200 -J HOST=stg.example.com --heap 2g
```

```powershell
# Windows 11
.\jmeter_api_tool.bat run -t output\api_test.jmx
.\jmeter_api_tool.bat run -t output\api_test.jmx -Threads 10 -Loops 100 -RampUp 10
.\jmeter_api_tool.bat run -t output\api_test.jmx -Threads 50 -Loops 200 -Prop HOST=stg.example.com -Heap 2g
```

<!-- cols: 24,30,64 -->
| sh のオプション | PowerShell のオプション | 説明 |
|---|---|---|
| -t ファイル | -Jmx（-t） | 実行する JMX。必須 |
| --threads 数 | -Threads | スレッド数（THREADS） |
| --rampup 秒 | -RampUp | ランプアップ秒数（RAMP_UP） |
| --loops 数 | -Loops | ループ回数（LOOPS）。-1 で無限 |
| --duration 秒 | -Duration | 実行時間（DURATION）。設定 duration を指定して生成した JMX のみ有効 |
| -J 名前=値 | -Prop 名前=値,名前=値 | 任意のプロパティ。sh は複数回指定可、PowerShell はカンマ区切り |
| -o フォルダ | -OutDir（-o） | 結果の出力先（既定: result） |
| --heap サイズ | -Heap | JVM のヒープサイズ（例: 2g）。既定 1g |
| --jmeter-home フォルダ | -JMeterHome | JMeter のフォルダ。省略時は環境変数 JMETER_HOME → PATH → 標準的な場所の順に探索 |
| --fail-on-error | -FailOnError | 1 件でもエラーがあれば終了コード 2 を返す（CI 用） |
| -- 以降の引数 | -ExtraArgs "引数" | JMeter にそのまま渡す追加オプション（例: プロキシ -H proxy.example.com -P 8080） |

### 実行時に変更できるプロパティ
<!-- cols: 34,32,52 -->
| プロパティ | 既定値 | 説明 |
|---|---|---|
| THREADS | 設定 threads | 同時に実行するスレッド数（仮想ユーザー数） |
| RAMP_UP | 設定 ramp_up | 全スレッドが起動するまでの秒数 |
| LOOPS | 設定 loops | 各スレッドの繰り返し回数（-1 = 無限） |
| DURATION | 設定 duration | 実行時間（秒）。duration を指定して生成した JMX のみ |
| THROUGHPUT | 設定 throughput_per_min | スループット制御を有効にした場合の目標（件 / 分） |
| PROTOCOL / HOST / PORT | URL から自動設定 | 接続先。2 つ目以降のホストは HOST_2、HOST_3 … |
| CONNECT_TIMEOUT / RESPONSE_TIMEOUT | 10000 / 30000 | 接続 / 応答タイムアウト（ミリ秒） |
| CSV_FILE | data/testdata.csv | CSV データファイル（相対パスは JMX のフォルダ基準） |
| EXCEL_FILE / EXCEL_SHEET | data/testdata.xlsx / 先頭シート | Excel データファイルとシート名 |
| RESULT_DIR | ~/result | JTL・結果ログの出力先（~/ は JMX のフォルダ。run コマンドは自動設定） |
| LOG_RESPONSE | error | 結果ログへの本文出力（none / error / all） |
| LOG_MAX_CHARS | 1000 | 結果ログに出力する本文の最大文字数 |
| 共通変数の名前（例: ENV） | 設定 user_variables | 設定で定義した共通変数 |

### JMeter を直接実行する場合
```bash
# RHEL 9.8（-f は既存の結果ファイル・レポートフォルダを削除してから実行）
jmeter -n -t output/api_test.jmx -l result/result.jtl -j result/jmeter.log \
  -e -o result/report -f -JTHREADS=10 -JLOOPS=100 -JRESULT_DIR=result
```
```powershell
# Windows 11
& "$env:JMETER_HOME\bin\jmeter.bat" -n -t output\api_test.jmx -l result\result.jtl -j result\jmeter.log -e -o result\report -f -JTHREADS=10 -JLOOPS=100 -JRESULT_DIR=result
```
- RESULT_DIR を相対パスで指定した場合は、JMeter を起動したフォルダ（カレントディレクトリ）が基準になります。
- HTML レポートの出力先（-o）は、存在しないか空のフォルダである必要があります（-f を付けると自動で削除されます）。

### GUI モードで実行する場合
- メニューの「Run」→「Start」（Ctrl + R）で実行し、「結果ツリー」「集計レポート」で結果を確認します。
- GUI で実行した場合も、JTL・結果ログは RESULT_DIR（既定: JMX と同じフォルダの result）に出力されます。
- GUI での大量実行は避けてください。画面の描画で JMeter 自体が遅くなり、正しく測定できません。

## 8.大量実行
<!-- summary: 件数の考え方、負荷試験の推奨設定、JVM・OS の調整 -->

### リクエスト件数の考え方
総リクエスト件数 = THREADS（スレッド数）× LOOPS（ループ回数）× 有効な API の数

<!-- cols: 22,50,46 -->
| 目的 | 指定例 | 件数（API 4 本の場合） |
|---|---|---|
| 疎通確認（既定） | THREADS=1, LOOPS=1 | 4 件 |
| 少量の連続実行 | THREADS=1, LOOPS=100 | 400 件 |
| 並列実行 | THREADS=10, LOOPS=100 | 4,000 件 |
| 時間指定の負荷試験 | 設定 duration=600 で生成し、THREADS=50, LOOPS=-1 | 600 秒間継続 |
| CSV の行数だけ実行 | LOOPS=-1、設定 csv_recycle=false・csv_stop_thread=true | CSV 行数 × API 数 |

### 負荷試験の推奨設定
- CLI モード（run コマンド、または jmeter -n）で実行します。
- [任意] 結果ログ出力は 1 件ごとにファイルへ書き込むため、毎秒数百件を超える場合は無効化するか、LOG_RESPONSE=none にします。
- [任意] JTL出力 全件 は run コマンドの result.jtl と同じ内容を別ファイルに書き込むため、大量実行時は無効化してかまいません。
- [任意] 結果ツリー / 集計レポート / 統計レポートは CLI では負荷になりにくいですが、不要なら削除または無効化します。
- 待機時間（Constant Timer）で実際の操作間隔を再現し、スループット制御で目標の件数 / 分に合わせます。
- ランプアップは「スレッド数 ÷ 1 秒あたりに増やしたいスレッド数」を目安にし、急激に負荷をかけないようにします。
- Excel データは全行をメモリに読み込むため、数万行を超えるデータは CSV を使ってください。

### JVM のヒープサイズ
<!-- cols: 22,30,66 -->
| 規模の目安 | ヒープ | 指定方法 |
|---|---|---|
| 〜100 スレッド | 1g（既定） | 指定不要 |
| 〜500 スレッド | 2g 〜 4g | run コマンドの --heap 4g / -Heap 4g |
| 500 スレッド超 | 4g 以上、または分散実行 | 同上。JMeter を直接実行する場合は環境変数 HEAP（例: -Xms4g -Xmx4g -XX:MaxMetaspaceSize=256m） |

### OS の調整（RHEL 9.8）
```bash
# ファイルディスクリプタの上限（スレッド数が多い場合）
ulimit -n                 # 現在値の確認
ulimit -n 65535           # 一時的に変更（恒久的には /etc/security/limits.conf）
# エフェメラルポートの範囲
sysctl net.ipv4.ip_local_port_range
```

### OS の調整（Windows 11）
- 動的ポートの範囲は `netsh int ipv4 show dynamicport tcp` で確認できます。
- 長時間の試験では、スリープ・省電力の設定を無効にしてから実行します。
- ウイルス対策ソフトのリアルタイムスキャンが結果ファイルの書き込みを遅くする場合があります（除外設定は社内のルールに従ってください）。

### 分散実行（参考）
1 台で必要な負荷をかけられない場合は、JMeter の分散実行（コントローラ 1 台 + 負荷生成サーバー複数台）を利用します。各サーバーで `jmeter-server` を起動し、`jmeter -n -t output/api_test.jmx -R サーバー1,サーバー2` で実行します。CSV / Excel のデータファイルは各サーバーの同じパスに配置が必要です。詳細は JMeter 公式マニュアルの「Remote Testing」を参照してください。

## 9.結果確認
<!-- summary: 出力ファイルの場所と見方（HTML レポート・JTL・ログ） -->

### run コマンドの出力先
```text
result/api_test_20261011_103000/
├─ result.jtl          全リクエストの結果（CSV 形式）
├─ jmeter.log          JMeter の実行ログ
├─ report/index.html   HTML ダッシュボードレポート
├─ api_result.log      結果ログ（[任意] 結果ログ出力）
├─ result_all.jtl      JTL（[任意] JTL出力 全件）
└─ error_detail.xml    失敗したリクエストの詳細（[任意] エラー詳細出力。失敗がある場合のみ）
```

### HTML レポート（report/index.html）
<!-- cols: 28,90 -->
| 画面 | 確認するポイント |
|---|---|
| Dashboard | APDEX（応答の満足度）、成功 / 失敗の割合、Statistics（API ごとの件数・エラー率・平均・90/95/99 パーセンタイル・スループット）、Errors（エラーの種類） |
| Charts → Over Time | 応答時間・スループットの時間推移（性能劣化やエラー急増のタイミング） |
| Charts → Throughput | 1 秒あたりのリクエスト数、応答コード別の件数 |
| Charts → Response Times | 応答時間の分布、パーセンタイル |

JTL から HTML レポートを作り直す場合は `jmeter -g result.jtl -o report2` を実行します。

### JTL（result.jtl）の主な列
<!-- cols: 22,96 -->
| 列 | 内容 |
|---|---|
| timeStamp | 開始時刻（エポックミリ秒） |
| elapsed | 応答時間（ミリ秒） |
| label | API 名（HTTP リクエストの名前） |
| responseCode / responseMessage | HTTP ステータスとメッセージ |
| threadName | スレッド名 |
| success | 成否（true / false） |
| failureMessage | 検証（Assertion）エラーの内容 |
| bytes / sentBytes | 受信 / 送信バイト数 |
| grpThreads / allThreads | 実行中のスレッド数 |
| URL | リクエストの URL |
| Latency / Connect / IdleTime | 最初の応答までの時間 / 接続時間 / 待機時間（ミリ秒） |

### 結果ログ（api_result.log）
タブ区切りのテキストで、1 リクエスト 1 行です。列は「日時・結果（OK / NG）・ステータス・応答時間ms・受信バイト・ラベル・リクエスト（メソッドと URL）・スレッド・メッセージ」です。失敗時（LOG_RESPONSE=all の場合は全件）は次の行に [REQUEST] / [RESPONSE] として送信内容と応答本文（LOG_MAX_CHARS 文字まで）を出力します。最終行には合計・成功・失敗件数が出力されます。

### エラー詳細（error_detail.xml）
失敗したリクエストだけを、リクエストヘッダ・本文とレスポンスヘッダ・本文付きで出力します。JMeter GUI の「結果ツリー（View Results Tree）」の「Filename」欄でこのファイルを開くと、内容を画面で確認できます。

### jmeter.log
- 「==== テスト開始」「==== テスト終了」の行に、実行条件と合計・成功・失敗件数が出力されます。
- 失敗したリクエストは「[API NG]」で始まる警告として出力されます。
- スクリプトのエラーは ERROR / WARN として出力されます。うまく動かない場合はまずこのファイルを確認してください。

## 10.変数と関数
<!-- summary: 生成される変数、よく使う関数、JSR223（Groovy）・抽出の書き方 -->

### 生成される変数
<!-- cols: 34,26,58 -->
| 変数 | 設定する要素 | 内容 |
|---|---|---|
| `${PROTOCOL}` `${HOST}` `${PORT}` | 共通変数 | 接続先 |
| `${SEQ}` | 連番 | 000001 から始まる連番（全スレッドで一意） |
| `${RANDOM_NO}` | 乱数 | 1 〜 99999 の乱数 |
| `${REQUEST_ID}` | 共通前処理 | リクエストごとの UUID |
| `${TIMESTAMP}` / `${EPOCH_MS}` | 共通前処理 | 現在時刻（ISO 8601 形式 / エポックミリ秒） |
| `${LAST_STATUS}` / `${LAST_ELAPSED_MS}` | 共通後処理 | 直前の API のステータスと応答時間 |
| `${EXCEL_ROW_NO}` | Excelデータ行取得 | 使用中の Excel の行番号 |
| CSV / Excel の列名 | CSVデータ読込 / Excelデータ行取得 | 1 行目の列名がそのまま変数名になる |

### よく使う関数
<!-- cols: 18,46,54 -->
| 関数 | 記入例 | 結果 |
|---|---|---|
| `__P` | `${__P(THREADS,1)}` | プロパティの値（未指定なら既定値） |
| `__UUID` | `${__UUID()}` | ランダムな UUID |
| `__time` | `${__time(yyyyMMddHHmmss,)}` | 現在時刻（書式指定） |
| `__Random` | `${__Random(1,100,)}` | 1 〜 100 の乱数 |
| `__RandomString` | `${__RandomString(8,abcdef0123456789,)}` | 指定した文字からなるランダムな文字列 |
| `__counter` | `${__counter(FALSE,)}` | 全スレッド共通の連番（TRUE にするとスレッドごと） |
| `__threadNum` | `${__threadNum}` | スレッド番号 |
| `__urlencode` | `${__urlencode(${KEYWORD})}` | URL エンコード |
| `__digest` | `${__digest(SHA-256,${USER_ID},,,)}` | ハッシュ値 |
| `__timeShift` | `${__timeShift(yyyy-MM-dd,,P1D,,)}` | 日付の計算（例: 翌日） |
| `__groovy` | `${__groovy(vars.get('SEQ').toInteger() * 2,)}` | Groovy の式の結果 |
| `__V` | `${__V(USER_${SEQ})}` | 変数名を組み立てて参照 |

### JSR223（Groovy）の書き方
```groovy
// 変数の取得・設定（値は文字列）
String userId = vars.get('USER_ID')
vars.put('USER_ID_UPPER', userId.toUpperCase())

// プロパティ（全スレッド共通）の取得・設定
String host = props.getProperty('HOST', 'localhost')
props.put('SHARED_TOKEN', vars.get('AUTH_TOKEN'))

// 直前の応答（後処理・リスナーで使用）
String code = prev.getResponseCode()
def json = new groovy.json.JsonSlurper().parseText(prev.getResponseDataAsString())
vars.put('ORDER_ID', String.valueOf(json.data.id))

// ログ出力（jmeter.log）
log.info('order id = ' + vars.get('ORDER_ID'))
```
- JSR223 のスクリプト内では `${変数名}` を使わず、`vars.get('変数名')` を使います（${} を使うとスクリプトのキャッシュが効かず、値によっては構文エラーになります）。
- 言語は groovy、「Cache compiled script if available」はオンにします（生成した JMX は設定済み）。
- vars はスレッドごと、props は全スレッド共通です。スレッドをまたいで値を共有する場合は props を使います。

### 正規表現抽出の書き方
<!-- cols: 40,78 -->
| 抽出したいもの | 正規表現の例 |
|---|---|
| JSON の "token":"abc" の abc | `"token"\s*:\s*"([^"]+)"` |
| 数値の項目 "id": 123 の 123 | `"id"\s*:\s*(\d+)` |
| HTML の hidden 項目の値 | `name="csrf" value="([^"]+)"` |
| 応答ヘッダの値（GUI で対象を Response Headers に変更） | `Location: (.+)` |
- ( ) で囲んだ部分が変数に入ります（テンプレート `$1$`）。見つからない場合は NOT_FOUND が入ります。
- 一致したものをすべて取得するには、GUI で「Match No.」を -1 にします。`変数名_1`、`変数名_2` … と、件数を表す `変数名_matchNr` が作られます。

### JSON 抽出（JSONPath）の書き方
<!-- cols: 40,78 -->
| 抽出したいもの | JSONPath |
|---|---|
| ルートの id | `$.id` |
| 入れ子の値 | `$.data.user.name` |
| 配列の先頭要素 | `$.items[0].id` |
| 条件に一致する要素 | `$.items[?(@.status == 'active')].id` |
| すべての id（Match No. を -1 にする） | `$..id` |

## 11.トラブル対応
<!-- summary: よくあるエラーと対処方法 -->
<!-- cols: 36,36,48 -->
| 症状 | 主な原因 | 対処 |
|---|---|---|
| `python3: command not found`（RHEL） | python3 が未導入 | `sudo dnf install -y python3` |
| `$'\r': command not found`（RHEL） | 改行コードが CRLF になっている | `sed -i 's/\r$//' jmeter_api_tool.sh` |
| このシステムではスクリプトの実行が無効になっている（Windows） | PowerShell の実行ポリシー | jmeter_api_tool.bat から実行する |
| 日本語が文字化けする（コンソール） | 端末の文字コード | RHEL は `export LANG=ja_JP.UTF-8` など UTF-8 の端末を使用 |
| JMeter が見つからない | JMETER_HOME が未設定 | 環境変数 JMETER_HOME を設定するか、--jmeter-home / -JMeterHome で指定 |
| Java が見つからない・バージョンが古い | Java が未導入、PATH が未設定 | Java 17 をインストールし、JAVA_HOME と PATH を設定 |
| `${USER_ID}` などがそのまま送信される | 変数を設定する要素が削除・無効、または変数名の誤り | 結果ツリーでリクエストを確認し、変数名と設定元の要素を見直す |
| CSV の値が `<EOF>` になる | CSV の最終行に達した（csv_recycle=false） | csv_recycle=true にするか、ループ回数を見直す |
| `File … must exist and be readable` | CSV ファイルが見つからない | CSV_FILE を JMX からの相対パス、または絶対パスで指定 |
| 応答の日本語が文字化けする | 応答に文字コードの指定がない | JMeter 5.6.3 は既定で UTF-8。異なる場合は `-Jsampleresult.default.encoding=文字コード` |
| Connection refused / Connect timed out | 接続先・ポートの誤り、ファイアウォール、プロキシ | URL を確認。プロキシ経由の場合は -H プロキシ -P ポート を追加（run は -- / -ExtraArgs で指定） |
| SSLHandshakeException | TLS のバージョン・証明書の問題 | jmeter.log で詳細を確認。クライアント証明書が必要な場合は system.properties で keystore を設定 |
| `java.lang.OutOfMemoryError` | ヒープ不足 | --heap / -Heap で増やす。CLI で実行し、結果ツリーを無効化 |
| `BindException: Address already in use` | クライアント側のポート枯渇 | KeepAlive を有効（既定）にする、スレッド数を減らす、OS のポート範囲を広げる |
| `Too many open files`（RHEL） | ファイルディスクリプタの上限 | `ulimit -n` で上限を引き上げる |
| HTML レポートが作成されない | JTL が空、出力先フォルダが空でない | jmeter.log を確認。`jmeter -g result.jtl -o 新しいフォルダ` で再作成 |
| Excel のデータが読み込まれない | ファイルパス・シート名の誤り、.xls 形式 | jmeter.log の [Excel] の行を確認。.xlsx 形式で保存する（.xls は非対応） |

## 12.コマンド一覧
<!-- summary: ツールと JMeter のコマンドラインオプション -->

### ツールのコマンド
<!-- cols: 12,46,46,30 -->
| コマンド | RHEL 9.8（sh） | Windows 11（bat / ps1） | 説明 |
|---|---|---|---|
| init | `./jmeter_api_tool.sh init [-o フォルダ] [-f]` | `jmeter_api_tool.bat init [-OutDir フォルダ] [-Force]` | 入力テンプレートを作成（既定: input） |
| generate | `./jmeter_api_tool.sh generate -i ファイル [-s 設定CSV] [-o フォルダ] [-n ファイル名] [--split] [-f]` | `jmeter_api_tool.bat generate -i ファイル [-s 設定CSV] [-OutDir フォルダ] [-JmxName ファイル名] [-Split] [-Force]` | JMX を生成（既定: output） |
| manual | `./jmeter_api_tool.sh manual [-o フォルダ]` | `jmeter_api_tool.bat manual [-OutDir フォルダ]` | 本書（xlsx / md）を出力（既定: docs） |
| run | `./jmeter_api_tool.sh run -t JMX [オプション]` | `jmeter_api_tool.bat run -t JMX [オプション]` | CLI モードで実行（既定: result） |
| help | `./jmeter_api_tool.sh help` | `jmeter_api_tool.bat help` | ヘルプを表示 |

### JMeter の主なオプション
<!-- cols: 24,94 -->
| オプション | 説明 |
|---|---|
| -n | CLI（非 GUI）モードで実行 |
| -t ファイル | 実行する JMX |
| -l ファイル | 結果（JTL）の出力先 |
| -j ファイル | ログ（jmeter.log）の出力先 |
| -e | テスト終了後に HTML レポートを作成 |
| -o フォルダ | HTML レポートの出力先（存在しないか空のフォルダ） |
| -f | 既存の結果ファイル・レポートフォルダを削除してから実行 |
| -g ファイル | 既存の JTL から HTML レポートだけを作成（-o と併用） |
| -J名前=値 | JMeter プロパティを設定（${__P(名前,既定値)} で参照） |
| -G名前=値 | 分散実行時に全サーバーへプロパティを設定 |
| -D名前=値 | Java のシステムプロパティを設定 |
| -q ファイル | 追加のプロパティファイルを読み込む |
| -H ホスト / -P ポート | プロキシサーバーのホスト / ポート |
| -R ホスト1,ホスト2 | 分散実行（指定したサーバーで実行） |
| -L レベル | ログレベル（例: -LDEBUG） |
| -v | バージョンを表示 |
