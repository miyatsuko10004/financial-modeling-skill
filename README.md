# financial-modeling-skill

[![CI](https://github.com/miyatsuko10004/financial-modeling-skill/actions/workflows/ci.yml/badge.svg)](https://github.com/miyatsuko10004/financial-modeling-skill/actions)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)

**AIに投資銀行・ファンド水準の「まじ」な財務モデルを構築させるためのClaude Code / Codex / Antigravityスキル。**

財務モデリング規約正典（約70項目）、規約違反を0.1秒で検出する機械チェックバリデータ、**6つの財務モデル型カタログ**、数式連動計算ガイド、そして規約を100%パスする標準テンプレート（3表連動・DCF・感応度マトリクス）をパッケージングしたものです。

進め方は、**規約（modeling-rules.md）を読む → 型カタログからモデル構造を選ぶ → たたき台を生成・数式連動を組む → 機械チェック（check_model.py）を通す（FAIL 0） → 第三者エージェントに監査させる**、の順です。

> A Claude Code / Codex / Antigravity skill for generating investment banking and private equity grade financial models: an institutional modeling rulebook (~70 rules), an automated rule checker (`check_model.py`), 6 institutional model archetypes, and formula-linked templates (3-statement, DCF, sensitivity matrix).

---

## 本質は `references/modeling-rules.md`

このリポジトリでいちばん価値があるのは、スクリプトやテンプレ単体ではなく、**[references/modeling-rules.md](references/modeling-rules.md)** という規約の正典テキストです。

- **「青字＝手入力値、黒字＝計算式、緑字＝他シート参照の100%徹底」**
- **「式内定数ハードコード（`*1.05` 等のMagic Number）の絶対排除」**
- **「循環参照オン（反復計算）頼みの禁止 —— 期首残高金利による循環ゼロ設計」**
- **「BS現預金はCF計算書からの完全プラグ連動（直接入力の厳禁）」**
- **「全期間で貸借バランス差額 0.00 を常時検証するチェック行の常設」**
- **「行・列の非表示（Hide）禁止 —— グループ化（折りたたみ）の徹底」**

使い方は3つだけです：
1. **AIにモデルを作らせる前に毎回このファイルを読ませる。**
2. **出力後に `python3 scripts/check_model.py model.xlsx` で違反を機械検出する（FAIL 0 を必須化）。**
3. **最後に `references/model-review-prompt.md` の指示文で、作り方を知らない別のエージェントにモデルを渡し、投資委員会・取締役会目線で冷徹に監査させる。**

AIはセッションごとにコンテキストがリセットされるため、口頭で「綺麗に作って」「数式を連動させて」と指示しても定着しません。ルールをファイルにして毎回読ませ、機械バリデータでゲートを敷くのが品質を担保する唯一の方法です。

自社や案件で使うときは、`references/modeling-rules.md` に固有のルールや指摘を追記して育ててください。

---

## 6つの財務モデル型カタログ（Archetypes）

詳細は **[references/model-archetypes.md](references/model-archetypes.md)** に記載されています。目的に応じて最適な型を選択します：

| 型ID | モデル型名 | 構造・特徴 | 主な用途・シーン |
| :--- | :--- | :--- | :--- |
| **Type-01** | **単一シート3表連動モデル** | 前提・PL・BS・CF・DCF・感応度が縦1シートに完結。A列ジャンプ対応 | スピーディな事業価値評価、感応度分析、経営会議 |
| **Type-02** | **マルチシート詳細3表連動モデル** | 前提、PL、BS、CF、PP&E、Debt、WCを別シートにモジュール化 | 大規模事業会社、中計策定、複雑な事業別売上計画 |
| **Type-03** | **DCFバリュエーションモデル** | アンレバードFCF、WACC算定、TV（永久成長率／Exit倍率）、感応度マトリクス | M&Aバリュエーション、株式価値算定、投資委員会 |
| **Type-04** | **事業計画・感応度シミュレーション** | 複数シナリオ（Base/Best/Worst）スイッチ、KPIツリー連動 | スタートアップ資金調達、新規事業投資採算性検証 |
| **Type-05** | **LBOモデル** | ソーシング/使途、多層デットスケジュール、リターン（IRR/MoIC）分析 | プライベート・エクイティ（PE）投資、買収ファイナンス |
| **Type-06** | **M&A合算・財務統合モデル** | 買収ストラクチャー、連結消去・のれん償却、EPS希薄化（Accretion/Dilution） | 上場企業M&A、資本政策、シナジー分析 |

---

## セマンティックカラー規約（Wall Street / IBD Standard）

投資銀行・PEファンドの実務において、色分けは単なる装飾ではなく「監査可能性（Auditability）」の根幹です：

| 文字色 | カラーコード | 意味・用途 | 備考 |
| :--- | :--- | :--- | :--- |
| **青字** | `#0000FF` | **手入力値・前提パラメータ・定数** | 外部から入力・変更可能なセル |
| **黒字** | `#000000` | **同一シート内の計算式** | SUM、四則演算、同一シートセル参照 |
| **緑字** | `#008000` | **別シートからの参照** | マルチシート時の連携（他シート参照） |
| **赤字** | `#FF0000` | **外部ファイル参照 / 警告** | リンク切れリスクがあるため原則最小化 |
| **背景色** | `#FDE9D9` / `#FFFFCC` | **入力可能セル（Assumptions）** | ユーザーが編集すべきセルを明示 |

---

## 規約バリデータ（`scripts/check_model.py`）

規約違反や潜在バグを0.1秒で網羅検査するPythonバリデータです：

```bash
python3 scripts/check_model.py my_model.xlsx
```

```
======================================================================
  FINANCIAL MODEL INTEGRITY REPORT (FAIL: 0, WARN: 0)
======================================================================
  [PASS] 01. Excel Formula Errors       : No #REF!, #DIV/0!, #VALUE! found
  [PASS] 02. Hidden Elements            : No hidden rows or columns detected
  [PASS] 03. Circular References        : No direct circular references found
  [PASS] 04. Balance Sheet Check Row    : Check rows found and valid
  [PASS] 05. Balance Sheet True Balance : Evaluated values balanced across all forecast periods (Diff: 0.00)
  [PASS] 06. Cash Flow Linkage          : BS Cash links to CF ending cash
  [PASS] 07. Retained Earnings Linkage  : Net Income correctly closed to Retained Earnings
  [PASS] 08. PP&E Schedule Linkage      : Depreciation links to PL & CF
  [PASS] 09. Debt Schedule Linkage      : Interest links to PL
  [PASS] 10. Direct Constants Check     : Zero hardcoded constants inside formulas
  [PASS] 11. Negative Cost Convention   : Cost lines consistently configured
  [PASS] 12. Year Row Uniformity        : Forecast period timeline uniform
  [PASS] 13. Color Coding Uniformity    : Inputs and formulas strictly color-coded
  [PASS] 14. Sensitivity Matrix Check   : Sensitivity table verified
----------------------------------------------------------------------
  AUDIT RESULT: PERFECT PASS (FAIL: 0, WARN: 0)
======================================================================
```

---

## インストール・セットアップ

### 1. Claude Code で使う
```bash
git clone https://github.com/miyatsuko10004/financial-modeling-skill.git ~/.claude/skills/financial-modeling-skill
```
またはプロジェクト内のみで有効にする場合：
```bash
git clone https://github.com/miyatsuko10004/financial-modeling-skill.git .claude/skills/financial-modeling-skill
```

### 2. Codex で使う
```bash
git clone https://github.com/miyatsuko10004/financial-modeling-skill.git ~/.codex/skills/financial-modeling-skill
```

### 3. Google Antigravity (AGY) / Gemini CLI で使う
```bash
git clone https://github.com/miyatsuko10004/financial-modeling-skill.git ~/.gemini/config/skills/financial-modeling-skill
```

### 4. 依存ライブラリのインストール
機械バリデータでBS貸借の実数値検証を行うため、`openpyxl` および `formulas` をインストールします：
```bash
pip install -r requirements.txt
```

---

## 使い方（ワークフロー）

### 1. AIエージェントに指示を出す
チャットやプロンプトで以下のように指示します：
> 「`financial-modeling-skill` を使って、受領したPL/BS実績から今後5年間の3表連動モデルとDCFバリュエーションExcelを作成して。規約を満たして check_model.py で FAIL 0 に仕上げて。」

### 2. CLIからたたき台モデルを直接生成する
スクリプトで即座に規約完全準拠のExcelモデルを生成することも可能です：
```bash
python3 scripts/new_model.py --type single --title "新規事業 財務予測モデル" -o output.xlsx
```

### 3. モデルを検証する
```bash
python3 scripts/check_model.py output.xlsx
```

### 4. 第三者監査（フレッシュアイレビュー）
モデルが完成したら、`references/model-review-prompt.md` の内容を別のAIエージェントに渡し、セカンドオピニオン監査を実施します。

---

## リポジトリ構成

```
financial-modeling-skill/
├── README.md               # 本ドキュメント（全体概要・使い方・規約サマリ）
├── SKILL.md                # 各AIエージェント用スキル定義（手順・トリガー）
├── LICENSE                 # MIT License
├── pyproject.toml          # プロジェクト定義
├── requirements.txt        # 依存ライブラリ（openpyxl, formulas）
├── references/             # 【規約・設計正典】
│   ├── modeling-rules.md   # ★規約の正典（約70項目、投資銀行・ファンド水準）
│   ├── model-archetypes.md # 6つのモデル型カタログと選定デシジョンツリー
│   ├── schedule-logic-guide.md # スケジュール連動計算ガイド（PP&E, WC, Debt, CF）
│   ├── formula-smell-lexicon.md # 数式の悪癖・AI臭アンチパターン辞典
│   ├── audit-checklist.md  # 納品前4段階セルフ監査チェックリスト
│   └── model-review-prompt.md # フレッシュアイ第三者監査プロンプト
├── scripts/                # 【自動化ツール群】
│   ├── new_model.py        # 財務モデル自動生成ジェネレータ
│   └── check_model.py      # 規約整合性バリデータ（機械検査エンジン）
├── templates/              # 【正本テンプレート】
│   └── single_sheet_3statement_template.xlsx # FAIL 0 / PERFECT PASS 検証済み正本
└── .github/
    └── workflows/
        └── ci.yml          # 自動検査CI（GitHub Actions）
```

---

## ライセンス

本リポジトリは [MIT License](LICENSE) の下で公開されています。商用・非商用問わずご自由にお使いいただけます。
