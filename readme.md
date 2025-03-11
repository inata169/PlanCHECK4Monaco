# Enhanced Monaco Plan Check

放射線治療計画検証ツール（Elekta Monaco用）

## 概要

このツールは Elekta Monaco 治療計画システム用のプランチェックスクリプトで、治療計画の品質保証のための自動チェック機能を提供します。Monaco Scripting API を用いて開発されており、ビーム設定、処方線量、計算設定など多岐にわたる治療計画パラメータの検証を自動化します。

**注意**: このリポジトリは教育・研究目的のみを意図しており、実際の臨床使用前には適切な検証と承認が必要です。

## 公開範囲について

このリポジトリには、一般的な放射線治療計画検証ロジックのみが含まれており、以下の点に注意してください：

- すべての患者データやセンシティブな情報は削除されています
- 実際の臨床データは含まれていません
- 特定の施設固有の設定値や設定基準は一般化されています
- コードは教育・研究目的での参照用であり、実際の臨床環境での使用には適さない場合があります

## 主な機能

- **基本プラン情報の検証**
  - 患者情報（ID、名前、クリニック）
  - プランID形式（3D/VMAT/DCAT + 3桁の数字）

- **処方と線量のチェック**
  - 処方線量の最小値検証
  - 分割線量の最小値検証
  - 最大線量比率のチェック

- **ビームプロパティの検証**
  - アイソセンター一貫性
  - フィールドID形式（数字4桁）
  - 最小MU
  - 計算アルゴリズム
  - エネルギー

- **ジオメトリ設定のチェック**
  - ガントリ/コリメータ/カウチ角度
  - アーク方向および長さ
  - コリメータサイズ

- **治療補助具のチェック**
  - カウチの有効化状態
  - ボーラスの使用状況
  - ウェッジIDの確認

- **DVH統計情報の検証**
  - 適合度指数（CI）
  - 不均一性指数（HI）

- **計算設定の検証**
  - 線量計算アルゴリズム
  - 最終計算アルゴリズム
  - グリッド間隔
  - ビームあたりの最大粒子数

## 使用方法

1. Monacoで患者を開き、チェックしたい治療計画を選択します
2. このスクリプトを実行します
3. 結果は以下の形式で表示されます：
   - 画面上のデータグリッドビュー形式の結果表示
   - デスクトップの「MonacoPlanCheck」フォルダ内にテキストファイルで保存

## 結果の解釈

チェック結果は以下の重要度で表示されます：
- **エラー**：修正が必要な重大な問題
- **警告**：確認が必要な潜在的な問題
- **情報**：参考情報

## システム要件

- Elekta Monaco 治療計画システム
- Elekta.MonacoScripting.API
- .NET Framework 4.5以上

## 設定可能なパラメータ

コード内の定数を編集することで、チェック基準を調整できます：
```csharp
private const double MIN_MU = 10.0;             // 最小MU値
private const double MIN_DOSE = 1.0;            // 最小線量
private const double MAX_DOSE_RATIO = 1.17;     // 最大許容線量比率
private const double MIN_COLLIMATOR_SIZE = 3.0; // 最小コリメータサイズ
```

## 注意事項

- このツールは治療計画の検証を支援するものであり、医学物理士や放射線治療専門放射線技師、放射線治療医の専門的判断に代わるものではありません
- 治療計画の最終承認前に、すべての警告とエラーを適切に評価してください
- このコードは所属機関の正式な製品やサービスを代表するものではありません
- 実際の臨床使用は承認されていません

## 開発者向け情報

このスクリプトは Monaco Scripting API を使用しています。拡張または修正を行う場合は、以下の名前空間の理解が必要です：
- Elekta.MonacoScripting.API
- Elekta.MonacoScripting.API.General
- Elekta.MonacoScripting.API.Planning
- Elekta.MonacoScripting.API.DICOMImport
- Elekta.MonacoScripting.API.Beams
- Elekta.MonacoScripting.DataType
- Elekta.MonacoScripting.Log

## コントリビューション

本リポジトリは教育・研究目的の参照用として公開しています。イシューの報告や改善提案は歓迎しますが、プルリクエストについては慎重に検討させていただきます。

## 免責事項

このコードは特定の臨床環境で作成されたものの一部であり、他の環境での動作は保証されません。使用に伴うリスクはすべてユーザー自身の責任となります。実際の放射線治療計画の検証には、適切な資格を持つ医学物理士による確認が必要です。

## ライセンス

MIT License

Copyright (c) 2025 169@inata169.com

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.


## 連絡先

質問や問い合わせは、169@inata169.com までお願いします。
