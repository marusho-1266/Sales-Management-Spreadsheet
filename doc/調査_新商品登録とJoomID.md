# 調査：新商品登録処理におけるJoomID入力項目

**調査日**: 2026-03-06  
**目的**: 在庫管理シートのC列（JoomID）と新商品登録処理の整合性確認

---

## 1. 在庫管理シートのC列（JoomID）について

### 1.1 定義・設計
- **Constants.gs** (23行目): `JOOM_ID: 3` → C列がJoomID用に定義されている
- **InventoryManagement.gs** (42行目): ヘッダーに `'JoomID'` が含まれる
- **SheetInitializer.gs** (1169行目): 「A: 商品ID, B: 商品名, C: JoomID, D: SKU, ...」と列構成が定義されている

### 1.2 JoomID列の用途
- Joomに出品した商品のJoom側ID（Joomが生成する製品識別子）を保持
- Joom APIとの連携や、既存出品商品の更新時に利用される想定

---

## 2. 新商品登録処理（ProductInputForm.gs）の現状

### 2.1 バックエンド（saveNewProduct）
**JoomIDの受け取り・保存は実装済み**

```74:79:src/ProductInputForm.gs
    // 新しい行のデータを準備（Joom対応フィールド含む）
    const newRowData = [
      formData.productId,           // 商品ID
      formData.productName,         // 商品名
      formData.joomId || '',        // JoomID
```

- `formData.joomId` を受け取り、無ければ空文字 `''` でC列に保存するロジックがある
- JoomID列が在庫管理シートに存在しない場合は `addJoomIdColumnToExistingSheet()` で列を追加する処理もある（37-40行目）

### 2.2 フロントエンド（HTMLフォーム）
**JoomIDの入力項目を追加済み（2026-03-06、任意項目）**

| 項目 | 状態 |
|------|------|
| 基本情報 | 商品ID, 商品名, **JoomID（任意）**, SKU, ASIN |
| 仕入れ情報 | 仕入れ元, 仕入れ元URL, 仕入れ価格, 販売価格, 重量, 販売価格(USD) |
| 容積重量 | 高さ, 長さ, 幅 |
| Joom対応 | 商品説明, メイン画像URL, 通貨, 配送価格, 在庫数量 |
| 在庫 | 在庫ステータス, 商品カテゴリー |
| 備考 | 備考・メモ |
| **JoomID** | **入力項目あり（任意）** |

### 2.3 JavaScript（formData収集）
**joomIdがformDataに含まれる（2026-03-06実装）**

```javascript
        var formData = {
          productId: parseInt(document.getElementById('productId').value),
          productName: document.getElementById('productName').value,
          joomId: document.getElementById('joomId').value || '',
          sku: document.getElementById('sku').value,
          asin: document.getElementById('asin').value,
          supplier: document.getElementById('supplier').value,
          // ... 中略 ...
          notes: document.getElementById('notes').value
        };
```

- `joomId` が formData に含まれており、未入力時は空文字が送信される

---

## 3. 結論

| 項目 | 状況 |
|------|------|
| 在庫管理シートC列 | JoomID用に設計・実装済み ✓ |
| saveNewProduct（バックエンド） | joomId受け取り・保存のロジックあり ✓ |
| HTMLフォーム | JoomID入力項目あり（任意） ✓ |
| formData収集（JS） | joomIdを含む ✓ |

**現状（2026-03-06実装後）**: 新商品登録フォームからJoomIDを任意で入力でき、在庫管理シートのC列に保存される。

---

## 4. 補足：JoomIDの入力タイミング

- **新規商品**を在庫管理に登録する段階では、まだJoomに出品していないためJoomIDは存在しない
- Joomに出品し、Joom側でIDが発行された後、**後から在庫管理のC列に手動で入力する**運用が想定される可能性がある
- あるいは、**Joom出品CSV連携やAPI連携でJoomIDを取得・反映する**機能が別途ある想定かもしれない

このため、「新商品登録フォームにJoomID項目がない」ことは、
- 仕様として「新規時は空でよい」としている可能性
- 単なる漏れ（実装し忘れ）の可能性

のいずれも考えられる。

---

## 5. 改善案（実装済み 2026-03-06）

フォームにJoomID入力を任意項目として追加した。

1. **HTML**: 基本情報セクション（商品名の直後）に JoomID の input を追加（`form-group full-width`）
2. **JavaScript**: formData に `joomId: document.getElementById('joomId').value || ''` を追加
3. 任意項目として扱い、既存ロジック（`formData.joomId || ''`）でそのまま利用

---

**作成**: 2026-03-06  
**更新**: 2026-03-06（JoomID入力項目を任意項目として実装）
