# TestSheetMaker
試験票作成ツール。たぶん汎用的に使えそうなので公開。

## Usage

`pipenv install`後、以下のコマンドを実行

```powershell
> pipenv run python .\src\main.py -o [出力するXLSXファイルのパス] -c [コンフィグファイル(YML形式)のパス] [試験票(Markdownファイル)のパス]
```

### コンフィグファイル

YAML形式。sampleフォルダにもあるがサンプルにない設定もある。
```yaml
Headers:
  TestResult: ## 試験実施結果用の列を作る設定。要らない場合は配下ごと削除
    PrintCount: 2 ## 出力する試験結果列の個数
    Title: "{}回目" ## 列の上に出力する回数のキャプション(1回目、2回目 など)。{}が回数
    Labels: ## 試験実施結果列に追加する内容
      - 実施担当
      - 確認担当
      - 実施日
      - 結果
  TestItemsLabel: ## テスト項目の見出し数
    - ステップ
    - 中項目
    - 小項目
    - 詳細項目
  BackColor: "002060" ## ヘッダの背景色
  TextColor: "FFFFFF" ## ヘッダの文字色
Sheet:
  Name: "試験票" ## 試験票のシート名
  Caption: "テスト一覧" ## 試験票A1セルに追加する文字列
  FontName: "游ゴシック" ## 〃文字フォント
  FontSize: 14 ## 〃文字サイズ
  Height: 18.75 ## 1行目の高さ
Consts: ## 定数。名前: 値で定数を指定することができる。要らない場合は配下ごと削除
  Environment: "試験環境"
Rearrange: ## 試験票の結果並び替えを行う場合ここに並びを指定。要らない場合は配下ごと削除
  - "no"　## ダブルクオーテーションを外すとFalseとみなされエラーになるので注意
  - "itemname"
  - "content"
  - "results"
ColumnSet: ## カラムの表示スタイル調整。要らない場合は配下ごと削除
  Common: ## すべての列に共通の項目
    Header: ## ヘッダ行の設定 
      FontName: "游ゴシック" ## フォント名
      FontSize: 10 ## フォントサイズ
      FontColor: "FFFFFF" ## 文字色(ColumnSetを使うとこちらのテキスト色が優先されるのでTextColorを設定した場合はこちらも必須)
      AlignVertical: top ## 縦位置
      AlignHorizontal: center ## 横位置
    Body: ## ヘッダ以外のすべてのセルに適用する設定
      FontName: "游ゴシック"
      FontSize: 10
      AlignVertical: top
      AlignHorizontal: left
      AlignWrapText: True ## テキストの折り返しをおこなうかどうか
  "No": ## No 列に設定する項目
    Header:
      Replace: "No." # ヘッダ行の文字列入れ替え
      AlignHorizontal: left
    Body:
      Replace: "=ROW()-3" # ボディ部分の文字列入れ替え。数式も使用可能
  TestResultHeader:
    AlignVertical: center
    AlignHorizontal: center
    Height: 24 ## テスト結果行(2行目)の高さ
  HeaderRow:
    Height: 36 ## ヘッダ行(3行目)の高さ
```

### 試験票

試験票はMarkdown形式。試験項目の見出しを見出し記法(`#`)で表す。Markdownパーサーを使ってるわけではないので`====`や`----`を使って見出しを定義することはできない。

また、コンフィグファイルのTestItemsLabelより深い見出しを作るとエラーとなる。

```markdown
# STEP 1
## おおまかな内容
### こまかな内容
#### 正常動作
:: 条件 
* AAA
* BBB
* CCC
:: 手順
1. AAA
2. BBB
3. CCC
:: 結果
* AAAであること
:: 備考

#### 準正常動作
:: 条件
* AAA
* BBB
* CCC
:: 手順
1. AAA
2. BBB
3. CCC
:: 結果
* AAAであること
:: 備考

#### 異常動作
:: 条件
* AAA
* BBB
* CCC
:: 手順 &&
:: 結果
* BBBであること
:: 備考
```
条件や手順は`::`ではじめる。半角スペースはあってもなくてもよい。

項目の名称は**いままでの項目とあっていれば**問題ないし、急に項目を増やしてもいい(一つ目の試験項目が「条件、手順、結果」で、次の項目が「条件、手順、結果、備考」などでもいい)。

また、項目のうち、ないものは空白で埋められる。たとえば一つ目の試験項目が「条件、手順、結果」で、次の項目が「条件、備考」の場合、二つ目の試験項目の手順と結果は空欄となる。

なお、一つ上の試験項目と項目の内容が同じ場合は、最後に`&&`を付けることで、記述を省略可能。
