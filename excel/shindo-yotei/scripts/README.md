# scripts/

VBA を含む `.xlsm` を安全に編集するための Python ツール集。

## 必要なライブラリ

```bash
pip install oletools olefile
```

Web 版 Claude Code で使う場合は Setup script に書いておくと毎セッション自動で入ります。

## ファイル

| ファイル | 役割 |
|---|---|
| `ms_ovba_compress.py` | MS-OVBA 圧縮アルゴリズム（VBA ソース埋め込み用） |
| `extract_vba.py` | `.xlsm` → `<ソースディレクトリ>/*.bas` (.cls) にソースを取り出す |
| `build_vba.py` | `<ソースディレクトリ>/*.bas` (.cls) → `.xlsm` に書き戻す |
| `clear_data.py` | 記入済みの `.xlsm` から実データを消して配布できる状態に戻す |

## 2 つのブックを扱う

このディレクトリにはブックが 2 つあり、それぞれ別の VBA ソースディレクトリを持ちます。

| ブック | ソースディレクトリ |
|---|---|
| `進度予定表システム_鳴海.xlsm` | `vba/` |
| `進度予定表システム_共通.xlsm` | `vba-共通/` |

```bash
python3 scripts/extract_vba.py <xlsm> <ソースディレクトリ>   # 既定: 鳴海版 → vba/
python3 scripts/build_vba.py <ソースディレクトリ>            # 既定: vba/
```

`build_vba.py` の書き戻し先は `<ソースディレクトリ>/_meta.json` の `xlsm_path` で決まるので、
ディレクトリさえ合っていれば別のブックを上書きすることはありません。

ソースディレクトリを作り直すときは、いったん消してから抽出してください。
シート数が減った場合に古い `SheetN.cls` が残ってしまうためです。

```bash
rm -rf vba && python3 scripts/extract_vba.py 進度予定表システム_鳴海.xlsm vba
```

## 典型的なワークフロー

```bash
# 1. 最新の VBA を取り出す
python3 scripts/extract_vba.py 進度予定表システム_鳴海.xlsm vba

# 2. vba/Module1.bas などをエディタで編集

# 3. .xlsm に反映
python3 scripts/build_vba.py vba

# 4. 往復が壊れていないか確認（差分ゼロになるはず）
python3 scripts/extract_vba.py 進度予定表システム_鳴海.xlsm vba
git diff --stat vba/

# 5. コミット
git add 進度予定表システム_鳴海.xlsm vba/
git commit -m "..."
git push
```

学校共通版は `進度予定表システム_共通.xlsm` と `vba-共通` に読み替えてください。

## clear_data.py — 記入データを消す

実際に使ったブックをリポジトリに置く前に、中身を消すためのツールです。

```bash
python3 scripts/clear_data.py 進度予定表システム_鳴海.xlsm 出力先.xlsm
```

やること:

1. 第01〜16週シートの中身を「週案_テンプレート」と同じ状態に戻す
2. 基本設定・時間割マスターの入力欄を消す（共通版 VBA の `全データ消去` と同じ範囲）
3. **数式のキャッシュ値を全シートから消す** — セルを空にしても、数式セルには
   消す前の計算結果（クラス名など）が `<v>` に残っているため
4. **参照されなくなった共有文字列を捨てる** — `xl/sharedStrings.xml` には
   授業内容の文字列が丸ごと残るため、使われているものだけに作り直す
5. `docProps/core.xml` の作成者名・最終更新者・最終印刷日時を消す
6. 陳腐化した `xl/calcChain.xml` を捨て、`fullCalcOnLoad` を立てて開いたとき再計算させる

**シートは消しません。** VBA のドキュメントモジュール（`Sheet4`〜`Sheet19`）と
ワークシートが一対一で対応しており、シートだけ消すと VBA プロジェクトとの対応が
崩れるためです。`xl/vbaProject.bin` はバイト単位で無変更のまま残ります。

## 仕組み

### extract_vba.py
1. `.xlsm` (実体は ZIP) を読み、`xl/vbaProject.bin` を取り出す
2. vbaProject.bin は OLE 複合ドキュメント。`olefile` でストリームを読む
3. `VBA/dir` を MS-OVBA 解凍し、各モジュールの `textoffset`（圧縮ソース開始位置）と MODULETYPE（`0x21`=標準モジュール→`.bas` / `0x22`=ドキュメント・クラスモジュール→`.cls`）を取得
4. 各モジュールストリームの `[textoffset:]` を解凍 → cp932 でテキスト化。ワークシートを増やすとシートモジュールも増えるので、モジュール名の決め打ちはせず MODULETYPE から拡張子を決める
5. 末尾のオートパディング（旧版で挿入したランダムコメント、現版の MS-OVBA 空チャンクで挿入される末尾空白）をヒューリスティクで除去
6. `<ソースディレクトリ>/<Name>.<ext>` に保存。`<ソースディレクトリ>/_meta.json` に書き戻し先の `xlsm_path` と、各モジュールの `textoffset` / `target_compressed_size` / `stream_size` を記録

### build_vba.py
1. `<ソースディレクトリ>/_meta.json` から書き戻し先と各モジュールのメタ情報を読む
2. `<ソースディレクトリ>/<Name>.<ext>` を読んで MS-OVBA で圧縮
3. **サイズが target に満たない場合のパディング**
   - 元の OLE ストリームと同じバイト数で書き戻す必要があるため、`compressed_size` を `target_compressed_size` に **完全一致** させる必要がある
   - 戦略 (2段階):
     - **ソース側**: ランダムなコメント行を末尾に追加。圧縮後サイズが target にできるだけ近く、かつ gap が 0 または 3 以上になるよう探索
     - **圧縮側**: 残った 0/3+バイトの隙間を MS-OVBA の「空チャンク」（3 バイトで 0 バイト解凍）と微小チャンク（4-5 バイトで 1-2 文字解凍）で埋める
4. `olefile.write_stream()` で vbaProject.bin に書き戻し
5. **キャッシュ無効化**: 各モジュールの PerformanceCache、`_VBA_PROJECT` のバージョン欄、`__SRP_*` ストリームを全部ゼロクリア。Excel が次に開いたときソースから再コンパイルする
6. xl/* 全体を再 ZIP して `.xlsm` に上書き

## サイズ上限に当たったら

`build_vba.py` は `compressed_size > target_compressed_size` でエラーになります。
このとき選択肢は:

1. **編集量を減らす**: 不要なコメント・空白を削る、ロジックを圧縮する
2. **別のモジュールから領域を分けてもらう**: 大きい修正は Module1 以外に分散
3. **ベースラインから作り直す**: `.xlsm.backup` をコピーしてから再構築（履歴がリセットされる）

target サイズは元の `.xlsm` を Excel が保存した際に決まったストリーム長です。Excel で `.xlsm` を一度開いて保存し直すと、現状のソースサイズに合わせて target が広げ直されます（手元 PC で Excel が使える場合の回避策）。
