# コードレビュー (2026-11-03)

対象: CSV2XL v1.1.5 (`app/`, `services/`, `utils/`, `main.py`, `build.py`)
観点: 可読性・メンテナンス性・KISS

---

## 総評

処理の流れ(CSV 検索 → 読み込み → 加工 → Excel 書き込み → バックアップ → Excel で並べ替え)は
`services/csv_excel_transfer.py` に集まっていて追いやすく、モジュール分割の方向性も妥当です。

一方で、次の 4 点がメンテナンス性を下げています。

1. **`ConfigManager` の冗長さ**: 同じ形の getter/setter が 20 個近くあり、既定値が二重に書かれ、使われていないメソッドも残っている。
2. **エラー処理の重複と混在**: `try/except → print → raise` がほぼ全関数にあり、`QMessageBox` が services 層と UI 層の両方から出ている。
3. **マジックナンバー**: CSV の列番号 (6, 9, 8, 10, 3, 5) や Excel の列番号が説明なしに並んでいる。
4. **防御的すぎるコード**: `clear_all_filters` や `write_data_to_excel` で `try/except: pass` や不要な `None` チェックが多く、本筋が読みにくい。

以下、優先度順に記載します。

---

## 🔴 優先度: 高(不具合・不具合の恐れ)

### H1. exe 実行時に設定変更が保存されない

`utils/config_manager.py:9-10`

```python
if getattr(sys, 'frozen', False):
    base_path = Path(sys._MEIPASS)
```

PyInstaller の `--onefile` でなくても、`_MEIPASS` は配布物の内部フォルダ(onefile なら一時フォルダ)です。
onefile の場合は終了時に削除されるため、ダイアログで変更した設定が次回起動時に消えます。
`build.py` は `--onefile` を付けていないので現状は onedir 構成ですが、`_internal` 配下の ini を書き換える運用になっており、再ビルドのたびに設定が上書きされます。

**提案**: exe 実行時は `Path(sys.executable).parent / 'config.ini'` を読み書きし、無ければ同梱版をコピーする。(仕様として許容しているなら CLAUDE.md に明記)

### H2. `apply_cell_formats` で垂直中央配置が上書きされて消える

`services/excel_processor.py:49-55`

```python
cell.alignment = Alignment(vertical='center')
if col in [1, 2, 5, 6]:
    cell.alignment = Alignment(horizontal='center')        # vertical が消える
elif col in [3, 4]:
    cell.alignment = Alignment(horizontal='left', shrink_to_fit=True)  # vertical が消える
```

A〜F 列はどちらかの分岐に必ず入るため、`vertical='center'` は一度も効いていません。

```diff
-            cell.alignment = Alignment(vertical='center')
-
-            if col in [1, 2, 5, 6]:  # A,B列とE,F列
-                cell.alignment = Alignment(horizontal='center')
-
-            elif col in [3, 4]:  # C列とD列
-                cell.alignment = Alignment(horizontal='left', shrink_to_fit=True)
+            if col in [3, 4]:  # C列とD列
+                cell.alignment = Alignment(vertical='center', horizontal='left', shrink_to_fit=True)
+            else:
+                cell.alignment = Alignment(vertical='center', horizontal='center')
```

### H3. `AppearanceDialog.accept` が不正値で落ちる/範囲外の値を保存する

`app/dialogs.py:238-243`

- 入力欄が空のまま OK を押すと `int("")` で `ValueError` になりアプリが落ちる。
- `QIntValidator(200, 1000)` は入力途中の値(例: `"5"`)を `Intermediate` として許すため、幅 5 がそのまま保存され、次回起動時に極小ウィンドウになる。

**提案**: `hasAcceptableInput()` で検証し、不可なら警告を出して `return`。

```python
def accept(self) -> None:
    inputs = [self.font_size_input, self.window_width_input, self.window_height_input]
    if not all(field.hasAcceptableInput() for field in inputs):
        QMessageBox.warning(self, "警告", "入力値が範囲外です。")
        return
    ...
```

### H4. Excel ファイルが見つからないとき、ユーザーに何も表示されない

`services/excel_processor.py:119-121`

`write_data_to_excel` はファイルが無い/拡張子が `.xlsm` でない場合に `print` して `False` を返すだけです。
`--windowed` の exe では `print` は見えないため、ボタンを押しても何も起きないように見えます。
`.xlsm` 以外の場合も「見つかりません」と表示しており、メッセージも不正確です。

### H5. `open_and_sort_excel` の `finally` が常に `Win+↓` を送る

`services/excel_processor.py:343-344`

エラー時や `return` 時(ワークブックを開けなかった場合)も実行され、その時点で前面にある別のウィンドウを最小化してしまいます。
最小化は正常終了時のみ、try ブロックの最後で行うのが自然です。

---

## 🟡 優先度: 中(可読性・保守性)

### M1. `ConfigManager` の getter/setter を整理する

`utils/config_manager.py` 全体

- `configparser.get(section, key, fallback=...)` はセクションが無くても `fallback` を返すため、各メソッド冒頭の `if 'X' not in self.config: return 既定値` はすべて不要。既定値が 2 回書かれ、片方だけ変更する事故の元になる。
- setter の「セクションが無ければ作る」処理が 9 回繰り返されているのに、そのための `_ensure_section`(169行目)は一度も使われていない。
- 次のメソッドはアプリ・テストのどこからも呼ばれていない: `set_processed_path`, `set_share_button_position`, `set_share_button_wait_time`, `set_backup_retention_days`, `_ensure_section`。
- `save_config` の例外メッセージが `"Failed to load config"` になっている(save なのに load)。
- 既定パス `C:\Shinseikai\CSV2XL\...` が 3 箇所に散らばっている。

**提案**: 既定値の二重定義を消すだけでも約 40 行減ります。

```diff
     def get_backup_path(self) -> str:
-        if 'Paths' not in self.config:
-            return r"C:\Shinseikai\CSV2XL\backup"
         return self.config.get('Paths', 'backup_path', fallback=r"C:\Shinseikai\CSV2XL\backup")
```

setter は共通化します。

```python
def _set(self, section: str, key: str, value: str) -> None:
    if section not in self.config:
        self.config[section] = {}
    self.config[section][key] = value
    self.save_config()

def set_backup_path(self, path: str) -> None:
    self._set('Paths', 'backup_path', path)
```

`get_exclude_docs` / `get_exclude_doctors` も同じ処理なので `_get_list(section)` に寄せられます。
未使用メソッドは削除を推奨します(必要になったら追加すればよい)。

### M2. `ExcludeItemDialog` が `ConfigManager` の内部 (`config.config`) を直接触っている

`app/dialogs.py:52-80`

読み込みは `get_exclude_docs()` と同じ分割処理を自前で再実装し、保存もセクション作成から自前で行っています。
`ConfigManager` に `get_list(section)` / `set_list(section, items)` を用意すれば、ダイアログ側は数行になります。

また `ExcludeDocsDialog` / `ExcludeDoctorsDialog` は `title` と `item_label` に同じ文字列を渡しているので、引数を 1 つにまとめられます。

### M3. エラー処理の重複と UI 依存

- `MainWindow.import_csv`(`app/main_window.py:88-92`)と `transfer_csv_to_excel`(`services/csv_excel_transfer.py:46-47`)が同じ `except Exception` と同じメッセージを持っており、内側が必ず先に捕まえるので外側は事実上不要。
- services 層の 3 ファイルが `QMessageBox` を直接呼んでいるため、GUI なしでテストできず、メッセージの出し場所も散らばっている。

**提案 (KISS 版)**: services は例外を投げる(または `str` のメッセージを返す)だけにし、`QMessageBox` は `MainWindow.import_csv` の 1 箇所にまとめる。

```python
# services 側
raise FileNotFoundError("ダウンロードフォルダにCSVファイルが見つかりません。")

# app/main_window.py
def import_csv(self) -> None:
    try:
        transfer_csv_to_excel()
    except Exception as e:
        QMessageBox.critical(self, "エラー", f"CSVファイルの取り込み中にエラーが発生しました:\n{e}")
```

### M4. `except → print → raise` の定型句を削除する

次の関数は例外をログ出力して再送出しているだけです。`--windowed` の exe では `print` が見えないので、実質的に何もしていません。

- `process_csv_data`(csv_processor.py:84-86)
- `process_completed_csv`(csv_processor.py:117-119)
- `sort_excel_data`(excel_processor.py:87-89)
- `backup_excel_file`(file_manager.py:28-30)
- `ensure_directories_exist`(file_manager.py:87-89)

try/except を外せば各関数が 3〜4 行短くなり、本筋が見えやすくなります。
ログが必要なら `print` ではなく `logging` に統一してファイル出力するのがよいです。

### M5. CSV の列番号に名前を付ける

`services/csv_processor.py:57-80`

`df.columns[6]`, `[9]`, `[8, 10]`, `[3:]`, `df.columns[3]`, `df.columns[5]` と、列を削除する前後で同じ列が違う番号になっており、コメントを読まないと追えません。

**提案**: 元 CSV の列位置を定数にし、削除前にフィルタまで済ませる。

```python
DOC_NAME_COL = 6     # G列: 文書名
DOCTOR_NAME_COL = 9  # J列: 医師名
DROP_COLS = {0, 1, 2, 8, 10}  # A〜C列, I列, K列
```

フィルタも `for` で 1 件ずつ `filter` せず、まとめて書けます。

```python
doc_col = pl.col(df.columns[DOC_NAME_COL])
for doc in config.get_exclude_docs():
    df = df.filter(~doc_col.str.contains(doc, literal=True))
```

(`if exclude_docs:` のガードは空リストなら for が回らないだけなので不要)

列名の一意化(48-55行目)も 1 行で書けます。

```python
df.columns = [f"col_{i}_{name}" for i, name in enumerate(df.columns)]
```

### M6. `write_data_to_excel` が長すぎる(約 125 行)

`services/excel_processor.py:107-232`

ガイドラインの 50 行を大きく超えています。次の 3 つに分けると読みやすくなります。

1. `read_existing_keys(ws, last_row) -> set[tuple[str, ...]]`
2. `filter_new_rows(rows, existing_keys) -> list[list[str]]`
3. `write_rows(ws, start_row, rows) -> None`

あわせて以下も簡略化できます。

- `cell1 = ws.cell(...)` 〜 `cell6` と `if cellN else ''` の繰り返し(140-160行目): `ws.cell()` は `None` を返さないためチェック不要。`ws.iter_rows(min_row=2, max_row=last_row, max_col=6, values_only=True)` を使えばループ 1 つで書ける。
- `if 'wb' in locals():`(230行目): `wb` は必ず定義済みなので不要。
- `to_numpy().tolist()`(167行目): `temp_df.rows()` で同じ結果が得られ、`numpy` 依存を外せる。
- 日付比較で「Excel 側は `YYYYMMDD`、CSV 側は `YYYY-MM-DD` を `YYYYMMDD` に変換」している。どちらも `YYYY-MM-DD`(`date.isoformat()`)にそろえれば CSV 側の変換が不要になる。

### M7. `clear_all_filters` の防御的コードを減らす

`services/excel_processor.py:235-285`

`try/except` が 6 重に入れ子になり、`print` が 7 回出てきます。
`ShowAllData()` でフィルタ条件はすべて解除されるため、「念のため各列のフィルタを個別にクリア」(257-272行目)は不要と考えられます。
実機で必要性を確認したうえで、次程度に縮めることを推奨します。

```python
def clear_all_filters(worksheet: Any, workbook: Any) -> None:
    """フィルタ条件をクリアし、共有ブックでなければオートフィルタも解除"""
    if not worksheet.AutoFilterMode:
        return
    if worksheet.FilterMode:
        worksheet.ShowAllData()
    if not workbook.MultiUserEditing:
        worksheet.AutoFilterMode = False
```

### M8. 最終行の取得が重複している

`excel_processor.py:68` と `:328` で同じ `worksheet.Cells(worksheet.Rows.Count, "A").End(-4162).Row` を書いています。
`sort_excel_data` は最終行を返しているのに、呼び出し側で捨てて再計算しています。

```diff
-        sort_excel_data(worksheet)
-
-        # 最終行にカーソルを移動
-        last_row = worksheet.Cells(worksheet.Rows.Count, "A").End(-4162).Row
+        last_row = sort_excel_data(worksheet)
         worksheet.Cells(last_row, 1).Select()
```

`-4162` は `XL_UP = -4162` のように定数化すると意味が伝わります(`-4137` も同様に `XL_MAXIMIZED`)。

### M9. `find_latest_csv` の条件式が読みにくい

`services/csv_processor.py:133-136`

- `split('_')` を 1 ファイルにつき 3 回呼んでいる。
- `0 <= len(...)` は常に真。
- docstring の「職員ID_YYYYMMDDHHmmss」と CLAUDE.md の `XXXX_YYYYMMDDhhmmss` を正規表現にすればそのまま読める。

```python
CSV_NAME_PATTERN = re.compile(r"^[^_]{0,5}_\d{14}\.csv$")

csv_files = [f for f in Path(downloads_path).glob('*.csv') if CSV_NAME_PATTERN.match(f.name)]
```

### M10. `ConfigManager()` を何度も生成している

1 回の取り込みで `ConfigManager()` が 6 回以上生成され、そのたびに ini を読み直しています
(`transfer_csv_to_excel`, `ensure_directories_exist`, `cleanup_old_csv_files`, `process_csv_data`, `process_completed_csv`, `backup_excel_file`, `open_and_sort_excel`)。
`FolderPathDialog` は `__init__` の中で 2 回生成しています(100行目と134行目)。

**提案**: `transfer_csv_to_excel` で 1 回だけ生成し、必要な値(パスや除外リスト)を引数で渡す。
関数が「設定ファイル」ではなく「値」に依存するため、テストでのモックも減ります。

### M11. フォルダ作成が 3 箇所に分散している

`ensure_directories_exist` で processed/backup フォルダを作成済みなのに、
`process_completed_csv`(112行目)と `backup_excel_file`(17-18行目)でも再度作成しています。
また `ensure_directories_exist` はダウンロードフォルダまで作成しますが、存在しなければ CSV も無いので作成は不要です。

さらに `if not directory.exists(): directory.mkdir(exist_ok=True)` は `exist_ok=True` があるので `if` が不要です。

### M12. `cleanup_old_*` の重複

`file_manager.py:33-68` の 2 関数は「フォルダ・パターン・保持日数」以外が同じです。

```python
def delete_old_files(directory: Path, pattern: str, retention_days: int) -> None:
    threshold = datetime.datetime.now() - datetime.timedelta(days=retention_days)
    for file in directory.glob(pattern):
        if datetime.datetime.fromtimestamp(file.stat().st_mtime) <= threshold:
            file.unlink(missing_ok=True)
```

なお、処理済み CSV の削除にも `get_backup_retention_days()` を使っているので、設定名と実際の用途がずれています。

### M13. バックアップのファイル名がハードコード

`file_manager.py:22`, `:61`

`医療文書担当一覧_{timestamp}.xlsm` が固定のため、設定で Excel ファイル名を変えるとバックアップ名と一致しません。
`Path(excel_path).stem` を使えば設定と連動します。

---

## 🟢 優先度: 低(体裁・規約)

| # | 場所 | 内容 |
|---|------|------|
| L1 | `app/main_window.py`, `app/dialogs.py`, `build.py` | 型ヒントが無い(規約では必須)。`-> None` と引数型を付ける |
| L2 | `app/main_window.py:6-11` | import 順が規約(アルファベット順)と違う。`from app import __version__` が最後にある |
| L3 | `app/dialogs.py:12` | クラス定義前の空行が 1 行(PEP8 は 2 行)。`config_manager.py:6, 16` も同様 |
| L4 | `services/coordinate_tracker.py` | QMainWindow を継承した UI 部品なので `app/` に置くほうが構成と合う |
| L5 | `app/dialogs.py:167-187` | `browse_folder('downloads')` のような文字列分岐より、フォルダ用・ファイル用の 2 メソッドに分けたほうが単純 |
| L6 | `app/dialogs.py` `FolderPathDialog` | `processed_path` は設定項目にあるがダイアログで編集できない(意図的なら問題なし) |
| L7 | `services/csv_processor.py:12` | CP932 は Shift-JIS の上位互換なので `'shift-jis'` を先に試す意味は薄い。`['cp932', 'utf-8']` で足りる(CLAUDE.md の記載も要更新) |
| L8 | `services/csv_processor.py:30-33` | 正常時の `print` 4 行はデバッグ用の名残と思われる |
| L9 | `services/csv_processor.py:107-108` | 直前に見つけたファイルなので `exists()` チェックは不要 |
| L10 | `services/csv_processor.py:3`, `excel_processor.py` | Python 3.13 なので `Optional[X]` / `List[X]` より `X \| None` / `list[X]` が簡潔 |
| L11 | `build.py:15` | プレースホルダの無い f-string。`subprocess.run` に `check=True` が無く、失敗しても「成功」と表示される |
| L12 | `pyproject.toml` | `psutil` はどこからも import されていない。`numpy` も M6 の対応後は不要になる |
| L13 | `main_window.py:40-55` | 長いスタイルシートは定数に出すと `__init__` が短くなる(`__init__` は現在約 70 行) |

---

## 推奨する対応順

1. **H1〜H5**(不具合): 影響が大きく、修正範囲も小さい。
2. **M1・M2**(ConfigManager 整理): 行数が最も減り、ダイアログ側も簡単になる。
3. **M3・M4**(エラー処理の一本化): services からの UI 依存が消え、テストが書きやすくなる。
4. **M5・M6**(列番号の定数化と `write_data_to_excel` の分割): 仕様変更(列の追加など)への耐性が上がる。
5. それ以外は、関連箇所を触るついでに対応。

各ステップの前後で `uv run pytest tests/ -v --tb=short` と `uv run pyright` が通ることを確認してください。
