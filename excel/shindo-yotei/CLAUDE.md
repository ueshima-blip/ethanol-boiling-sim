# 進度予定表システム — Claude 向け作業ガイド

このディレクトリは中学校教諭用の進度予定表（Excel マクロ有効ブック）です。VBA を含むバイナリ `.xlsm` を Python スクリプト経由で安全に編集します。

**ブックは 2 つあり、それぞれ別の VBA ソースを持ちます。**片方を直しても、もう片方には反映されません。

## ファイル構成

```
進度予定表システム_鳴海.xlsm            # 鳴海中用（編集対象・記入データは消去済み）
進度予定表システム_共通.xlsm            # 学校共通用（編集対象・週シートなしの配布用テンプレート）
進度予定表システム_鳴海.xlsm.backup     # オリジナル（触らない）
進度予定表システム_鳴海.xlsm.before_*   # 過去スナップショット（触らない）
vba/                                # 鳴海用 VBA ソース（Source of truth）
  ThisWorkbook.cls
  Module1.bas                       # 実質ここだけが中身のあるモジュール
  Sheet1.cls 〜 Sheet19.cls          # 各ワークシートのクラスモジュール（中身は空）
  _meta.json                        # 書き戻し先の xlsm / textoffset / target_compressed_size
vba-共通/                            # 学校共通用 VBA ソース（Sheet1〜3 のみ）
scripts/
  ms_ovba_compress.py               # MS-OVBA 圧縮ライブラリ
  extract_vba.py                    # xlsm → <ソースディレクトリ>/*.bas (.cls)
  build_vba.py                      # <ソースディレクトリ>/ → xlsm（VBA だけを書き戻し）
```

`_meta.json` の `xlsm_path` が書き戻し先のブックを決めます。`vba/` は鳴海版、`vba-共通/` は
共通版を指しているので、ディレクトリさえ間違えなければ別のブックを上書きすることはありません。

## 標準ワークフロー（VBA を直す場合）

引数を省略すると鳴海版（`vba/`）が対象になります。共通版を直すときは必ず引数を付けてください。

1. **最新版を取り込む**
   ```bash
   python3 scripts/extract_vba.py 進度予定表システム_鳴海.xlsm vba
   python3 scripts/extract_vba.py 進度予定表システム_共通.xlsm vba-共通
   ```

2. **`Module1.bas` や `ThisWorkbook.cls` を編集**
   - ふつうのテキスト編集でよい（cp932 / Shift-JIS で書く）
   - VBA の構文・属性（`Attribute VB_Name =` 等）は触らない

3. **ビルドして `.xlsm` を更新**
   ```bash
   python3 scripts/build_vba.py vba
   python3 scripts/build_vba.py vba-共通
   ```

4. **往復が壊れていないか確認**（extract → build → extract で差分ゼロになるはず）
   ```bash
   python3 scripts/build_vba.py vba && python3 scripts/extract_vba.py 進度予定表システム_鳴海.xlsm vba && git diff --stat vba/
   ```

5. **コミット & プッシュ**
   ```bash
   git add 進度予定表システム_鳴海.xlsm vba/
   git commit -m "fix: ..."
   git push
   ```

## VBA を編集する際の注意

- **サイズ上限**: `build_vba.py` は元の OLE ストリームと同じバイト数で書き戻す制約があるので、`_meta.json` の `target_compressed_size` を超えるとビルドが失敗する
  - 現状 Module1 の圧縮後サイズは上限（鳴海 25531 / 共通 25635 バイト）ぴったりまで使っている。**行を増やすなら、同じくらい別の場所を削る必要がある**
  - どうしても入らない場合は Excel で一度開いて保存し直すと target が広がる（手元 PC に Excel がある場合）
- **VBA キャッシュ無効化は自動**: build 時に `_VBA_PROJECT` のバージョン欄と `__SRP_*`、各モジュールの PerformanceCache をゼロクリアする。Excel が次回開いたときソースから再コンパイルする
- **AUTO_PADDING**: ビルド時に末尾にランダムなコメントが付くが、extract で自動的に剥がされる。気にしなくてよい
- **2 つのブックの差分**: 共通版だけが `学活方式を準備` / `学活方式チェンジ処理` / `全データ消去` を持ち、`集計補正` が基本設定 B4 を見る。共通してほしい修正は両方に手で移植する
- **空のシートモジュール**: `Sheet2.cls` 以降は `Attribute` 行だけの空モジュール。`.xlsm` のシート数と一対一で対応するので、消さずにそのまま置いておく
- **記入データを入れない**: リポジトリのブックは中身を空にしてある。実際の授業記録が入ったファイルをコミットしないこと。`.xlsm` は ZIP なので、セルを消しても `xl/sharedStrings.xml`・数式のキャッシュ値・`docProps/core.xml` の作成者名にデータが残る。消すときはこの 3 か所も必ず確認する

## ワークシートの内容（時間割や日付）を編集する場合

`xl/worksheets/sheetN.xml` を直接編集する。
- `sheet1.xml` = 基本設定
- `sheet2.xml` = 時間割マスター
- `sheet3.xml` = 週案_テンプレート
- `sheet4.xml`〜 = 第01週、第02週、…

`.xlsm` は単なる ZIP なので、`unzip` → 編集 → `zip` で再パッケージ。
`scripts/build_vba.py` の `os.walk(work_dir)` を参考にできる。

## このシステムで実装済みの機能

1. **設定を完了** — 基本設定から時間割マスター・週案_テンプレート・集計表を一括生成し、シートを保護
2. **基本設定のボタン 6 つ** — 新しい週を追加 / 1・2・3 学期 PDF 出力 / 次年度ファイルを作成 / 別ファイルからインポート。`Auto_Open` がブックを開くたびに貼り直す
3. **授業の自動転記** — 週シートの時数欄に `月3` と入れると、時間割マスターから科目名を隣のセルへ転記（`Workbook_SheetChange`）
4. **総合のカウント** — 集計欄が col=15 までループ
5. **累計を更新の範囲制限** — 現在シートの週番号以下のシートだけを合算
6. **テスト機能**
   - 基本設定 B14 に「テスト」が登録されており、時間割セルのドロップダウンに表示される
   - 「テスト」を選ぶと `Workbook_SheetChange` が `HandleTestEntry` を呼び出し、対象クラスを聞く InputBox を表示
   - セル値を `テスト[,2-1,3-5,]` 形式で記録（前後のカンマで一意マッチを担保）し、表示は `;;;"テスト"` 書式で隠す
   - 集計式（PLAN 行）が `COUNTIF(範囲, "*,"&クラス名&",*")` で対象クラスにのみ +1
7. **旧ファイルから移行** — 別の `.xlsm` を選び、時間割マスター `B3:F9` と第NN週シートを取り込む。存在しない週はテンプレートから作る
8. **長期休業の自動スキップ** — 基本設定 `A37:B43` に夏・冬・春休みの開始／終了を `7/18` 形式で入力。`新しい週を追加` が `休業スキップ調整` を通し、休業期間内の週を飛ばす。設定セルは `休業設定を準備` が `Auto_Open` 時に自動で作る
9. **【共通版のみ】学活カウント方式** — 基本設定 B4 で「特活に含める / 学活で独立」を選び、`集計補正` が切り替える。B4 変更時は `学活方式チェンジ処理` が集計表を組み直す
10. **【共通版のみ】全データ消去** — 週シートを全削除し、基本設定・時間割マスターを空に戻す（配布テンプレートのリセット用）

## Web 版 Claude Code（claude.ai/code）での作業時

リポジトリをクローンしたら最初に下記が必要：

```bash
pip install oletools olefile
```

Setup script（環境設定で自動実行）に書いておくと毎回不要。

## 困ったとき

- `build_vba.py` が `target_compressed_size` 超過で失敗 → 編集量を減らす or 別のモジュールから削る
- `extract_vba.py` で AUTO_PADDING が混じる → `extract_vba.py` の `strip_padding()` のヒューリスティクを調整
- Excel が「VBA プロジェクトを読み込めません」 → `.xlsm.backup` から作り直す
- どちらのブックを直したか分からなくなった → `git diff --stat` で `vba/` と `vba-共通/` のどちらが変わったか見る
