# Consulting-Q レイアウト（高級コンサル向け outline types）

`outline_to_pptx.js` の `"palette": "consulting"` 専用の型と共通装飾。決定論的に描画するので、
内容の良し悪しは outline 側で決める。数値・コントラスト・余白の正本は [design-system.md](design-system.md)、
文体と TPO の選び方は [deck-types.md](deck-types.md) を参照。

```bash
node scripts/outline_to_pptx.js examples/outline.consulting.sample.json /tmp/consulting.pptx
python3 scripts/lint_deck.py /tmp/consulting.pptx
```

## 1. 共通装飾と palette

まず palette を選ぶ。他 palette の挙動（四角マーカー、26pt Cambria 見出し）は変わらない。

| token | 推奨値 | 役割 |
|---|---|---|
| `paper` | `FFFFFF` | 背景。全面白 |
| `soft` | `EEF2F8` | so-what 帯・回答帯・グループ行の淡い塗り（その上は `ink`） |
| `brand` | `002060` | 表紙・区切り・矢羽・ブリッジの棒。白文字 15.3:1 |
| `accent` | `C00000` | 番号・三角など小さな強調 1 箇所。白背景比 6.5:1 |
| `ink` | `1F3864` | タイトル・本文。白背景比 11.6:1 |
| `grey` | `595959` | kicker・出典・補足。白背景比 7.0:1 |
| `line` | `D9D9D9` | 行間の hairline のみ。文字色に使わない |
| `onBrand` / `onBrandSoft` | `FFFFFF` / `D6E0F0` | 濃紺スライド上の文字 / 補助文字 |

`ink` ・`accent` ・`grey` は白背景で 4.5:1 を超える（算出方法は design-system.md §1）。
色だけで状態を伝えず、必ず記号と言葉を添える。

全スライドに同じ装飾が入る。作者が座標を指定する必要はない。

- 右上 `kicker` 11pt（`def.kicker` → `meta.kicker` の順で解決、表紙は `meta.date`）
- タイトル 23pt bold の結論文、最大 2 行（超えると 2 行目末尾を省略）
- 左下 `source` 10pt（`def.source` → `meta.source`、表紙は author・date）
- 右下ページ番号 `NN / NN`、スライド順に自動採番
- スピーカーノートは `note`（従来の `notes` も可）
- 表紙・section・closing は濃紺背景。白系の文字に切り替わるだけで、kicker・出典・番号の位置は同じ

タイトルは「何を変えるか」が分かる 1 文にする。装飾バー・下線・タイトル下の罫線は使わない。

## 2. 型一覧

### exec_summary — エグゼクティブサマリー

```json
{ "type": "exec_summary", "kicker": "結論", "title": "受注単価は18%改善できる",
  "points": [{ "headline": "割引が利益を削る", "detail": "42%の案件で10%以上の値引き" }],
  "soWhat": "だから、審査基準を変える。", "nextStep": "来週までにログを揃える。" }
```

主張 3〜4 件を番号付きで縦に並べ、下部の淡い帯に so-what と次の一手を置く。
**使いどころ**: 資料の最初、または意思決定者の時間が 30 秒しかない場面。
**避けるべき**: 主張が 5 件以上、根拠が 2 行を超える、so-what が抽象的なスローガン。

### kpi_row — 大型数値の横並び

```json
{ "type": "kpi_row", "title": "3つの数字が利益漏れの場所を示す",
  "kpis": [{ "value": "-18%", "label": "受注単価", "readout": "数量は伸びたが単価が落ちた", "source": "受注データ" }] }
```

数値 → ラベル → 読み取り 1 行 → 任意の個別出典を 1 つのボックスにまとめる（2〜4 列）。
**使いどころ**: 既存 `stats` では「数字の意味」が語り切れないとき。
**避けるべき**: 数値だけ並べる、読み取りが数値の言い換え、列間で単位が違うのに注記がない。

### axis_table — 軸のある表

```json
{ "type": "axis_table", "title": "値引きの原因は判断軸で分かれる",
  "table": { "headers": ["判断軸", "割引率", "影響"], "widths": [2.6, 1.6, 7.73],
    "rows": [{ "group": "新規案件" }, ["大型", "18%", "利益率を8pt押し下げ"]] } }
```

先頭列を太字の軸にし、データセルは中央揃え。ヘッダー下だけ濃罫、行間は hairline、
最終行の下は罫線なし。`{ "group": "..." }` で全幅のグループ行を挿入できる（最大 7 行）。
**使いどころ**: 比較軸が先にあり、セルは短い言葉のとき。
**避けるべき**: セルに文章を詰める、縦罫を増やす、軸なしの単なる一覧表。

### premise_conclusion — 前提 → 帰結

```json
{ "type": "premise_conclusion", "title": "割引は基準の欠如から生じている",
  "premise": { "header": "事実", "rows": [["価格監査", "42%に承認記録がない"]] },
  "conclusion": { "header": "だから、判断を変える", "bullets": ["承認ルールを1枚に集約する"] } }
```

左に事実の行リスト、列間中央に右向き三角 1 つ、右に「だから、〜」の帰結。
**使いどころ**: 観測から示唆を一意に導く論理の可視化。
**避けるべき**: 左と右の対応関係がない、右が事実の繰り返し、三角を 2 つ以上置く。

### issue_action — 課題と打ち手

```json
{ "type": "issue_action", "title": "課題ごとに打ち手とオーナーを対応させる",
  "rows": [{ "issue": { "label": "基準が曖昧", "text": "上限が案件ごとに違う" },
             "action": { "label": "統一ルール", "text": "3段階の承認ライン", "owner": "販売企画", "date": "12/15" } }] }
```

左の課題と右の打ち手を同じ行で対応させる。右列に owner と期限を添える（最大 4 行）。
**使いどころ**: アクションアイテムを役割まで落とし込むとき。
**避けるべき**: 行の対応が崩れる、owner が「関係者」のような曖昧な表現、期限なし。

### chevron_flow — 矢羽（プロセス）

```json
{ "type": "chevron_flow", "title": "4ステップで価格ガバナンスを再設計する",
  "steps": [{ "label": "01", "title": "現状把握", "text": "割引ログを分類する" }] }
```

3〜5 個の homePlate / chevron を横並びにし、濃紺塗り・白文字にする。
**使いどころ**: 手順・変遷・ロールアウトの順序そのものがメッセージのとき。
**避けるべき**: 順序が意味を持たない項目の並べ替え、1 つの矢羽に 3 行以上の説明。

### status_heatmap — 状態ヒートマップ

```json
{ "type": "status_heatmap", "title": "準備はデータで進み、運用で遅れている",
  "heatmap": { "columns": ["領域", "データ", "運用"],
    "rows": [{ "label": "承認ログ", "values": ["warn", "risk"] }],
    "comments": ["記録様式が3種類。統一が先決。"] } }
```

`ok` / `warn` / `risk` を ● 正常 / ▲ 注意 / ※ リスク として色と言語の両方で表示する。
`comments` を渡すと右にコメント列が付く（最大 5 行）。凡例は表の下に自動で入る。
**使いどころ**: 定点観測、複数領域の進捗差を一枚で見せるとき。
**避けるべき**: 記号なしの色分け、状態の定義が資料内で違う、コメントが状態の繰り返し。

### two_by_two — 2x2 マトリクス

```json
{ "type": "two_by_two", "title": "着手順は影響度と容易性で決める",
  "matrix": { "xAxis": { "name": "影響度", "low": "小", "high": "大" },
    "yAxis": { "name": "実行容易性", "low": "難", "high": "易" },
    "quadrants": [{ "label": "最初に着手", "items": ["割引上限の統一"] }] } }
```

`quadrants` は左上・右上・左下・右下の順でちょうど 4 つ。各象限に項目を最大 3 件。
**使いどころ**: 優先順位付け、ポートフォリオの切り分け。
**避けるべき**: 軸の定義が曖昧、4 象限に均等に項目を配って結論が消える、項目が 4 件以上。

### bridge — 増減ブリッジ（ウォーターフォール）

```json
{ "type": "bridge", "title": "単価18pt改善の半分は割引削減で生まれる",
  "chart": { "kind": "waterfall", "unit": "指数（現在=100）",
    "steps": [{ "label": "現在", "delta": 100, "total": true },
              { "label": "割引削減", "delta": 9 }, { "label": "改善後", "delta": 118, "total": true }] } }
```

`{ "type": "chart", "chart": { "kind": "waterfall", ... } }` でも同じ結果になる。
非表示の base 系列と可視の stacked bar で増減を描き、符号付きの値はカテゴリ軸の 2 行目に出す。
**使いどころ**: 開始値 → 複数の増減 → 終了値を 1 本の線で説明するとき。
**避けるべき**: 増減の合計が終値と一致しない、項目が 8 つ以上、割合と絶対値を混在させる。

### recommendation — 推奨とネクストステップ

```json
{ "type": "recommendation", "title": "価格ガバナンス再設計を承認する",
  "answer": "割引ルールの統一を12月から全営業に適用する。",
  "reasons": ["年額2.4億円の利益漏れ", "既存データで7割が実行可能"],
  "nextSteps": [{ "action": "承認基準を策定", "owner": "販売企画", "date": "12/15" }] }
```

結論の帯、根拠 3 点、owner と日付付きの next steps 表（最大 5 行）。
**使いどころ**: 資料の締め。承認してほしい決定を 1 つに絞る。
**避けるべき**: 根拠が結論と無関係、next steps に owner か日付がない、複数の決定を並べる。

## 3. 組み立ての原則

1 枚 1 結論。タイトルで答えを言い、本文でその理由だけを置く。空白は約 20% 残す。
`exec_summary` で答え → `kpi_row` / `axis_table` で根拠 → `issue_action` / `chevron_flow` で着手順 →
`status_heatmap` / `two_by_two` で優先順位 → `bridge` で効果 → `recommendation` で決定、という流れが
 examples/outline.consulting.sample.json に実装例としてある。連続するスライドで同じ型を使わないこと。
