# Joom for Merchants API v3 - Get Order（単一注文取得）詳細資料

## 概要

この資料は、Joom for Merchants API v3の「Get Order」エンドポイントについて詳細に説明します。このAPIは、指定された注文IDを使用して単一の注文情報を取得するために使用されます。

## エンドポイント情報

### 基本情報
- **エンドポイント名**: Get Order
- **URL**: `https://api-merchant.joom.com/api/v3/orders`
- **HTTPメソッド**: GET
- **認証**: OAuth 2.0 Bearer Token
- **Content-Type**: application/json

### サンドボックス環境
- **サンドボックスURL**: `https://sandbox-202506-api-merchant.joomdev.net/api/v3/orders`

## リクエスト仕様

### クエリパラメータ

| パラメータ名 | 型 | 必須 | 説明 | パターン | 例 |
|-------------|----|----|------|----------|-----|
| `id` | string | ✓ | 注文のJoom ID | `^[A-Z\\d]{8}$` | `ABCD1234` |

### リクエスト例

#### cURL例
```bash
curl --request GET \
  --url "https://api-merchant.joom.com/api/v3/orders?id=ABCD1234" \
  --header "Accept: application/json" \
  --header "Authorization: Bearer YOUR_ACCESS_TOKEN"
```

#### JavaScript例
```javascript
const response = await fetch('https://api-merchant.joom.com/api/v3/orders?id=ABCD1234', {
  method: 'GET',
  headers: {
    'Accept': 'application/json',
    'Authorization': 'Bearer YOUR_ACCESS_TOKEN'
  }
});

const data = await response.json();
```

## レスポンス仕様

### HTTPステータスコード
- **200**: 成功
- **400**: 不正なリクエスト
- **401**: 認証エラー
- **403**: アクセス拒否
- **404**: 注文が見つからない
- **429**: レート制限超過
- **500**: サーバーエラー

### レスポンス構造

```json
{
  "data": {
    // 注文データ（Orderオブジェクト）
  }
}
```

## Orderオブジェクト詳細仕様

### 基本情報フィールド

| フィールド名 | 型 | 必須 | 説明 | 例 |
|-------------|----|----|------|-----|
| `id` | string | ✓ | 注文ID | `"ABCD1234"` |
| `currency` | string | ✓ | 通貨コード（ISO形式） | `"USD"` |
| `status` | string | ✓ | 注文ステータス | `"approved"` |
| `orderTimestamp` | string<date-time> | ✓ | 注文日時 | `"2023-12-01T10:30:00Z"` |
| `updateTimestamp` | string<date-time> | ✓ | 最終更新日時 | `"2023-12-01T15:45:00Z"` |
| `quantity` | integer | ✓ | 数量（1以上） | `2` |

### 配送関連フィールド

| フィールド名 | 型 | 必須 | 説明 |
|-------------|----|----|------|
| `allowedShippingTypes` | string | - | 許可される配送タイプ |
| `onlineShippingRequirement` | string | - | オンライン配送要件（非推奨） |
| `daysToFulfill` | integer | - | 履行までの日数 |
| `fulfillTimeHours` | integer | - | 履行にかかった時間（時間） |
| `hoursToFulfill` | integer | - | 履行までの時間 |

### 価格情報（priceInfo）

| フィールド名 | 型 | 説明 | 例 |
|-------------|----|------|-----|
| `orderPrice` | string | 注文価格 | `"29.99"` |
| `unitPrice` | string | 単価 | `"14.99"` |
| `shippingPrice` | string | 配送料 | `"5.99"` |
| `commissionAmount` | string | 手数料 | `"3.00"` |
| `buyerGmv` | string | 購入者GMV | `"29.99"` |
| `joomShippingPrice` | string | Joom配送料 | `"5.99"` |
| `joomShippingPriceUsed` | boolean | Joom配送料使用フラグ | `true` |
| `vat` | string | 付加価値税 | `"2.40"` |
| `refundAmount` | string | 返金額 | `"0.00"` |

### 割引情報（discounts）

| フィールド名 | 型 | 説明 |
|-------------|----|------|
| `details` | array | 割引詳細の配列 |
| `sum` | string | 割引合計額 |

#### 割引詳細（details）

| フィールド名 | 型 | 説明 | 例 |
|-------------|----|------|-----|
| `amount` | string | 割引額 | `"4.52"` |
| `type` | string | 割引タイプ | `"blogger"` |

### 製品情報（product）

| フィールド名 | 型 | 必須 | 説明 |
|-------------|----|----|------|
| `id` | string | ✓ | 製品ID |
| `sku` | string | ✓ | 製品SKU |
| `variant` | object | ✓ | バリアント情報 |

#### バリアント情報（variant）

| フィールド名 | 型 | 説明 |
|-------------|----|------|
| `id` | string | バリアントID |
| `sku` | string | バリアントSKU |

### カスタマイズ情報（customisation）

| フィールド名 | 型 | 説明 |
|-------------|----|------|
| `productImage` | object | 製品画像情報 |
| `customerImage` | object | 顧客画像情報 |
| `customerImagePosition` | object | 顧客画像位置 |
| `printArea` | object | 印刷エリア |
| `type` | string | カスタマイズタイプ |

### 配送先住所（shippingAddress）

| フィールド名 | 型 | 説明 | 例 |
|-------------|----|------|-----|
| `name` | string | 受取人名 | `"田中太郎"` |
| `email` | string | メールアドレス | `"user@example.com"` |
| `phoneNumber` | string | 電話番号 | `"+81-90-1234-5678"` |
| `country` | string | 国コード | `"JP"` |
| `state` | string | 都道府県 | `"東京都"` |
| `city` | string | 市区町村 | `"渋谷区"` |
| `street` | string | 住所 | `"道玄坂1-2-3"` |
| `streetAddress1` | string | 住所1 | `"道玄坂1-2-3"` |
| `streetAddress2` | string | 住所2 | `"マンション101"` |
| `zipCode` | string | 郵便番号 | `"150-0043"` |
| `building` | string | 建物名 | `"サンプルマンション"` |
| `flat` | string | 部屋番号 | `"101"` |
| `house` | string | 家番号 | `"1-2-3"` |
| `block` | string | ブロック | `"A"` |
| `taxNumber` | string | 税務番号 | `"123456789"` |

### 配送オプション（shippingOption）

| フィールド名 | 型 | 説明 |
|-------------|----|------|
| `shippingPrice` | string | 配送料 |
| `tierId` | string | 配送ティアID |
| `tierName` | string | 配送ティア名 |
| `tierType` | string | 配送ティアタイプ |
| `warehouseId` | string | 倉庫ID |
| `warehouseName` | string | 倉庫名 |
| `warehouseType` | string | 倉庫タイプ |

### 出荷情報（shipment）

| フィールド名 | 型 | 説明 |
|-------------|----|------|
| `provider` | string | 配送業者 |
| `providerId` | string | 配送業者ID |
| `shippingOrderNumber` | string | 配送注文番号 |
| `trackingNumber` | string | 追跡番号 |
| `timestamp` | string<date-time> | 出荷日時 |
| `shippedTimestamp` | string<date-time> | 出荷完了日時 |
| `fulfilledTimestamp` | string<date-time> | 履行完了日時 |
| `shipmentTimeHours` | integer | 出荷時間（時間） |

### 返金情報（refund）

| フィールド名 | 型 | 説明 |
|-------------|----|------|
| `by` | string | 返金者 |
| `cost` | string | 返金コスト |
| `price` | string | 返金価格 |
| `fraction` | integer | 返金割合 |
| `reason` | string | 返金理由 |
| `timestamp` | string<date-time> | 返金日時 |

### 返品情報（return）

| フィールド名 | 型 | 説明 |
|-------------|----|------|
| `declineReason` | string | 拒否理由 |
| `timestamp` | string<date-time> | 返品日時 |

### その他のフィールド

| フィールド名 | 型 | 説明 |
|-------------|----|------|
| `giftProduct` | object | ギフト製品情報 |
| `reviewRating` | integer | レビュー評価（1-5） |
| `isFbj` | boolean | FBJ（Fulfillment by Joom）フラグ |
| `storeId` | string | ストアID |
| `transactionId` | string | トランザクションID |
| `canBeConsolidated` | boolean | 統合可能フラグ |
| `isFulfillmentAllowed` | boolean | 履行許可フラグ |
| `fulfillmentAllowedTimestamp` | string<date-time> | 履行許可日時 |
| `shippingAddressHash` | string | 配送先住所ハッシュ |
| `shippingAddressNative` | object | ユーザー言語での配送先住所 |

## 注文ステータス一覧

| ステータス | 説明 |
|-----------|------|
| `approved` | 承認済み - マーチャントのアクション待ち |
| `cancelled` | キャンセル済み - マーチャントの過失によりキャンセル |
| `fulfilledOnline` | オンライン履行済み - まだ配送業者に出荷マークされていない |
| `shipped` | 出荷済み - マーチャントまたはJoom Logisticsによる出荷完了 |
| `paidByJoomRefund` | Joom返金済み - Joomによってキャンセル、マーチャントは支払い受取 |
| `refunded` | 返金済み |
| `returnInitiated` | 返品開始 - 顧客が返品を決定 |
| `returnExpired` | 返品期限切れ - 返品が時間内に到着しなかった |
| `returnArrived` | 返品到着 - 返品がマーチャントに到着 |
| `returnCompleted` | 返品完了 - マーチャントが返品を受け入れ |
| `returnDeclined` | 返品拒否 - マーチャントが返品を拒否 |

## 配送タイプ（allowedShippingTypes）

| 値 | 説明 |
|----|------|
| `offlineOnly` | オフライン配送のみ |
| `offlineOrOnline` | オフラインまたはオンライン配送 |
| `onlineOnly` | オンライン配送のみ |

## エラーレスポンス

### 400 Bad Request
```json
{
  "errors": [
    {
      "code": "invalid_argument",
      "field": "id",
      "message": "Invalid argument `id`: expected valid order ID, got `INVALID123`"
    }
  ]
}
```

### 401 Unauthorized
```json
{
  "errors": [
    {
      "code": "unauthorized",
      "message": "Invalid or expired access token"
    }
  ]
}
```

### 404 Not Found
```json
{
  "errors": [
    {
      "code": "not_found",
      "message": "Order with ID 'ABCD1234' not found"
    }
  ]
}
```

## 使用例

### 基本的な注文取得
```javascript
async function getOrder(orderId) {
  try {
    const response = await fetch(`https://api-merchant.joom.com/api/v3/orders?id=${orderId}`, {
      method: 'GET',
      headers: {
        'Accept': 'application/json',
        'Authorization': `Bearer ${accessToken}`
      }
    });

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    return data.data; // Orderオブジェクト
  } catch (error) {
    console.error('Error fetching order:', error);
    throw error;
  }
}

// 使用例
const order = await getOrder('ABCD1234');
console.log('Order Status:', order.status);
console.log('Order Price:', order.priceInfo.orderPrice);
```

### 注文ステータスチェック
```javascript
function checkOrderStatus(order) {
  switch (order.status) {
    case 'approved':
      return '注文承認済み - 履行待ち';
    case 'fulfilledOnline':
      return 'オンライン履行済み - 出荷待ち';
    case 'shipped':
      return '出荷済み';
    case 'cancelled':
      return 'キャンセル済み';
    default:
      return `不明なステータス: ${order.status}`;
  }
}
```

## 注意事項

1. **認証**: すべてのリクエストには有効なOAuth 2.0アクセストークンが必要です
2. **レート制限**: 一般的に2000 rpm（1分あたりのリクエスト数）の制限があります
3. **HTTPS必須**: すべてのリクエストはHTTPS経由で行う必要があります
4. **日時形式**: すべての日時はRFC3339形式（UTC）で返されます
5. **価格形式**: 価格は文字列として返され、小数点区切りはドット（.）を使用します
6. **通貨**: 通貨はISO 4217コード（例：USD、EUR、JPY）で指定されます

## 関連リンク

- [Joom for Merchants API v3 公式ドキュメント](https://merchant.joom.com/docs/api/v3)
- [OAuth 2.0認証ガイド](https://merchant.joom.com/docs/api/v3#authentication)
- [エラーハンドリング](https://merchant.joom.com/docs/api/v3#api-errors)
- [レート制限](https://merchant.joom.com/docs/api/v3#request-rate-limits)

---

*この資料は、Joom for Merchants API v3の公式ドキュメントに基づいて作成されています。最新の情報については、必ず公式ドキュメントを参照してください。*
