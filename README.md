# generateTables

Excel ファイルからフィールド定義を読み取り、FileMaker のテーブルオブジェクト XML を生成してクリップボードにコピーするツールです。

## 対応 OS

| OS | 備考 |
|---|---|
| macOS | Apple Silicon / Intel |
| Windows | |

---

## 使い方

### 1. Excel ファイルを用意する

各シートが FileMaker の 1 テーブルに対応します。

- シート名 `#SAMPLE` はスキップされます（サンプル用）
- 空のシートはスキップされます
- データ開始行は `config.xml` の `Field id` セル参照で決まります

### 2. 実行する

```bash
# 通常実行（結果をクリップボードにコピー）
./generateTables /path/to/Book.xlsx

# デバッグ実行（debug.log と output.xml も出力）
./generateTables -debug /path/to/Book.xlsx
```

### 3. FileMaker に貼り付ける

FileMaker の「データベースの管理」→「テーブル」タブを開き、クリップボードの内容を貼り付けます。

---

## ファイル構成

実行ファイルと同じディレクトリに以下を配置してください。

```
build/
├── generateTables        # 実行ファイル
├── config.xml            # フィールドマッピング設定（必須）
├── Sample.xlsx           # Excel ファイル
├── debug.log             # ログ（-debug 時のみ生成）
└── output.xml            # 生成 XML（-debug 時のみ生成）
```

---

## config.xml

Excel のどのセル列をどのフィールドプロパティに割り当てるかを定義します。  
セル参照（例: `A10`）の列部分（`A`）がマッピング先の列を示し、行番号はデータ行に合わせて自動でずれます。

```xml
<fmxmlsnippet type="FMObjectList">
  <BaseTable name="K3">
    <Field id="A10" name="C10" fieldType="H10" dataType="K10">
      <Calculation table="N10"><![CDATA[Q10]]></Calculation>
      <Validation ...>...</Validation>
      <AutoEnter constant="AI10" ...>
        <ConstantData>AL10</ConstantData>
        <Calculation table=""><![CDATA[AL10]]></Calculation>
      </AutoEnter>
      <Storage ...></Storage>
      <Comment>BA10</Comment>
    </Field>
  </BaseTable>
</fmxmlsnippet>
```

---

## Excel シートの列定義

### BaseTable

| config.xml 属性 | 内容 |
|---|---|
| `BaseTable name` | テーブル名が入力されたセル |

### Field（フィールド基本情報）

| config.xml 属性 | 内容 | デフォルト値 | 許可される値 |
|---|---|---|---|
| `Field id` | フィールド ID | 行インデックス番号 | 任意の文字列 |
| `Field name` | フィールド名 | `Field#{行番号}` | 任意の文字列 |
| `Field fieldType` | フィールドタイプ | `Normal` | 下表参照 |
| `Field dataType` | データタイプ | `Text` | 下表参照 |
| `Comment` | コメント | 空 | 任意の文字列 |

**fieldType の許可値**

| Excel 入力値 | 生成される値 |
|---|---|
| 通常タイプ（または空） | `Normal` |
| 計算タイプ | `Calculated` |
| 集計タイプ | `Summary` |

**dataType の許可値**

| Excel 入力値 | 生成される値 |
|---|---|
| テキスト型（または空） | `Text` |
| 数字型 | `Number` |
| 日付型 | `Date` |
| 時刻型 | `Time` |
| タイムスタンプ型 | `TimeStamp` |
| オブジェクト型 | `Binary` |

---

### Calculation（計算式）※ fieldType が 計算タイプ の場合のみ出力

| config.xml 属性 | 内容 | デフォルト値 |
|---|---|---|
| `Calculation table` | 参照テーブル名 | 空 |
| `Calculation`（CDATA） | 計算式 | 空 |

---

### AutoEnter（自動入力）

| config.xml 属性 | 内容 | デフォルト値 | 許可される値 |
|---|---|---|---|
| `AutoEnter constant` | 自動入力の種別 | 空（自動入力なし） | 下表参照 |
| `AutoEnter overwriteExistingValue` | 既存値を上書き | `True` | `True` / `False` |
| `AutoEnter allowEditing` | 編集を許可 | `False` | `True` / `False` |
| `AutoEnter alwaysEvaluate` | 常に評価 | `False` | `True` / `False` |
| `AutoEnter furigana` | ふりがな | `False` | `True` / `False` |
| `AutoEnter lookup` | ルックアップ | `False` | `True` / `False` |
| `ConstantData` | 固定値・計算式の内容 | 空 | 任意の文字列 |

**AutoEnter constant の許可値**

| Excel 入力値 | 動作 |
|---|---|
| 空 | 自動入力なし |
| 固定値 | 指定した固定値を入力（`ConstantData` の値を使用） |
| 計算値 | 計算式で自動入力（`ConstantData` の値を計算式として使用） |
| 作成TS | 作成タイムスタンプを自動入力 |
| 作成者 | 作成者アカウント名を自動入力 |
| 修正TS | 修正タイムスタンプを自動入力 |
| 修正者 | 修正者アカウント名を自動入力 |

---

### Validation（入力値の制限）

| config.xml 属性 | 内容 | デフォルト値 | 許可される値 |
|---|---|---|---|
| `Validation message` | 検証エラーメッセージ表示 | `False` | `True` / `False` |
| `Validation valuelist` | 値一覧による制限 | `False` | `True` / `False` |
| `Validation calculation` | 計算式による制限 | `False` | `True` / `False` |
| `Validation alwaysValidateCalculation` | 常に計算を検証 | `False` | `True` / `False` |
| `Unique value` | 値の一意性を検証 | `True` | `True` / `False` |
| `NotEmpty value` | 空を許可しない | `True` | `True` / `False` |
| `MaxDataLength value` | 最大文字数（空の場合は制限なし） | 空 | 数値 |
| `StrictDataType value` | 入力値のデータ型を制限 | 空（制限なし） | 下表参照 |

**StrictDataType の許可値**

| Excel 入力値 | 生成される値 | 意味 |
|---|---|---|
| 空 | （StrictDataType 要素なし） | 制限なし |
| 数字のみ | `Numeric` | 数値のみ |
| 日付のみ | `FourDigitYear` | 4桁年の日付のみ |
| 時刻のみ | `TimeOfDay` | 時刻のみ |

---

### Storage（保存オプション）

| config.xml 属性 | 内容 | デフォルト値 | 許可される値 |
|---|---|---|---|
| `Storage autoIndex` | 自動インデックス | `True` | `True` / `False` |
| `Storage index` | インデックス | `None` | `None` / `All` / `Minimal` |
| `Storage indexLanguage` | インデックス言語 | `Japanese` | `Japanese` など |
| `Storage global` | グローバルフィールド | `False` | `True` / `False` |
| `Storage maxRepetition` | 繰り返し数 | `1` | 数値 |

---

## ビルド

```bash
make build          # Linux 用（build/generateTables）
make darwin-arm64   # macOS Apple Silicon 用（build/generateTables_darwin_arm64）
make darwin-amd64   # macOS Intel 用（build/generateTables_darwin_amd64）
make windows        # Windows 用（build/generateTables.exe）
make clean          # ビルド成果物を削除
```

### 要件

- Go 1.24 以上
