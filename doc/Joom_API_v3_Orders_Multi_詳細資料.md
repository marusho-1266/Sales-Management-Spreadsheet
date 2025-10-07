# Joom for Merchants API v3 - Orders Multi（複数注文取得）詳細資料

## 概要

この資料は、Joom for Merchants API v3の「Orders Multi」エンドポイントについて詳細に説明します。このAPIは、複数の注文情報を一度に取得するために使用され、効率的な注文管理を可能にします。

## エンドポイント情報

### 基本情報
- **エンドポイント名**: Orders Multi
- **URL**: `https://api-merchant.joom.com/api/v3/orders/multi`
- **HTTPメソッド**: GET
- **認証**: OAuth 2.0 Bearer Token
- **Content-Type**: application/json
- **レート制限**: 50 rpm（1分あたり50リクエスト）

### サンドボックス環境
- **サンドボックスURL**: `https://sandbox-202506-api-merchant.joomdev.net/api/v3/orders/multi`

## リクエスト仕様

### クエリパラメータ

| パラメータ名 | 型 | 必須 | 説明 | 例 |
|-------------|----|----|------|-----|
| `limit` | integer | - | 1リクエストあたりの最大取得件数（デフォルト: 100、最大: 500） | `100` |
| `after` | string | - | ページネーション用のトークン | `"eyJpZCI6IkFCQ0QxMjM0In0="` |
| `updatedFrom` | string<date-time> | - | 更新日時の開始日時（RFC3339形式） | `"2023-12-01T00:00:00Z"` |
| `updatedTo` | string<date-time> | - | 更新日時の終了日時（RFC3339形式） | `"2023-12-31T23:59:59Z"` |
| `status` | string | - | 注文ステータスでフィルタリング | `"approved"` |
| `storeId` | string | - | ストアIDでフィルタリング | `"store123"` |

### 注文ステータス一覧

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

### リクエスト例

#### 基本的なリクエスト
```bash
curl --request GET \
  --url "https://api-merchant.joom.com/api/v3/orders/multi" \
  --header "Accept: application/json" \
  --header "Authorization: Bearer YOUR_ACCESS_TOKEN"
```

#### ステータスフィルタリング
```bash
curl --request GET \
  --url "https://api-merchant.joom.com/api/v3/orders/multi?status=approved&limit=50" \
  --header "Accept: application/json" \
  --header "Authorization: Bearer YOUR_ACCESS_TOKEN"
```

#### 日時範囲フィルタリング
```bash
curl --request GET \
  --url "https://api-merchant.joom.com/api/v3/orders/multi?updatedFrom=2023-12-01T00:00:00Z&updatedTo=2023-12-31T23:59:59Z" \
  --header "Accept: application/json" \
  --header "Authorization: Bearer YOUR_ACCESS_TOKEN"
```

#### JavaScript例
```javascript
const response = await fetch('https://api-merchant.joom.com/api/v3/orders/multi?status=approved&limit=100', {
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
- **429**: レート制限超過
- **500**: サーバーエラー

### レスポンス構造

```json
{
  "data": [
    {
      // Orderオブジェクト1
    },
    {
      // Orderオブジェクト2
    }
    // ... 最大500件まで
  ],
  "paging": {
    "next": "https://api-merchant.joom.com/api/v3/orders/multi?after=eyJpZCI6IkFCQ0QxMjM0In0=&limit=100"
  }
}
```

### ページネーション

複数注文取得APIは、大量のデータを効率的に処理するためにページネーションをサポートしています。

#### ページネーションの仕組み
- **デフォルト制限**: 100件/リクエスト
- **最大制限**: 500件/リクエスト
- **nextトークン**: 次のページのデータを取得するためのトークン
- **afterパラメータ**: ページネーション用のトークン

#### ページネーション使用例
```javascript
async function getAllOrders(accessToken) {
  let allOrders = [];
  let nextToken = null;
  
  do {
    const url = nextToken 
      ? `https://api-merchant.joom.com/api/v3/orders/multi?after=${nextToken}&limit=500`
      : 'https://api-merchant.joom.com/api/v3/orders/multi?limit=500';
    
    const response = await fetch(url, {
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
    allOrders = allOrders.concat(data.data);
    
    // 次のページがあるかチェック
    nextToken = data.paging?.next ? extractAfterToken(data.paging.next) : null;
    
  } while (nextToken);
  
  return allOrders;
}

function extractAfterToken(nextUrl) {
  const url = new URL(nextUrl);
  return url.searchParams.get('after');
}
```

## Orderオブジェクト詳細仕様

複数注文取得APIで返される各Orderオブジェクトは、単一注文取得APIと同じ構造を持ちます。

### 基本情報フィールド

| フィールド名 | 型 | 必須 | 説明 | 例 |
|-------------|----|----|------|-----|
| `id` | string | ✓ | 注文ID | `"ABCD1234"` |
| `currency` | string | ✓ | 通貨コード（ISO形式） | `"USD"` |
| `status` | string | ✓ | 注文ステータス | `"approved"` |
| `orderTimestamp` | string<date-time> | ✓ | 注文日時 | `"2023-12-01T10:30:00Z"` |
| `updateTimestamp` | string<date-time> | ✓ | 最終更新日時 | `"2023-12-01T15:45:00Z"` |
| `quantity` | integer | ✓ | 数量（1以上） | `2` |

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
| `zipCode` | string | 郵便番号 | `"150-0043"` |

## エラーレスポンス

### 400 Bad Request
```json
{
  "errors": [
    {
      "code": "invalid_argument",
      "field": "updatedFrom",
      "message": "Invalid argument `updatedFrom`: expected date/time, got `invalid_date`"
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

### 429 Too Many Requests
```json
{
  "errors": [
    {
      "code": "rate_limit_exceeded",
      "message": "Request rate quota exceeded"
    }
  ]
}
```

## 使用例

### 基本的な複数注文取得
```javascript
async function getOrders(accessToken, options = {}) {
  const params = new URLSearchParams();
  
  if (options.limit) params.append('limit', options.limit);
  if (options.status) params.append('status', options.status);
  if (options.updatedFrom) params.append('updatedFrom', options.updatedFrom);
  if (options.updatedTo) params.append('updatedTo', options.updatedTo);
  if (options.storeId) params.append('storeId', options.storeId);
  
  const url = `https://api-merchant.joom.com/api/v3/orders/multi?${params.toString()}`;
  
  try {
    const response = await fetch(url, {
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
    return data;
  } catch (error) {
    console.error('Error fetching orders:', error);
    throw error;
  }
}

// 使用例
const orders = await getOrders(accessToken, {
  status: 'approved',
  limit: 100
});
console.log(`取得した注文数: ${orders.data.length}`);
```

### 承認済み注文の一括取得
```javascript
async function getApprovedOrders(accessToken) {
  return await getOrders(accessToken, {
    status: 'approved',
    limit: 500
  });
}
```

### 特定期間の注文取得
```javascript
async function getOrdersByDateRange(accessToken, startDate, endDate) {
  return await getOrders(accessToken, {
    updatedFrom: startDate,
    updatedTo: endDate,
    limit: 500
  });
}

// 使用例
const startDate = '2023-12-01T00:00:00Z';
const endDate = '2023-12-31T23:59:59Z';
const orders = await getOrdersByDateRange(accessToken, startDate, endDate);
```

### 注文ステータス別の集計
```javascript
async function getOrderStatusSummary(accessToken) {
  const statuses = ['approved', 'fulfilledOnline', 'shipped', 'cancelled'];
  const summary = {};
  
  for (const status of statuses) {
    const orders = await getOrders(accessToken, { status, limit: 500 });
    summary[status] = orders.data.length;
  }
  
  return summary;
}
```

### 効率的なページネーション処理
```javascript
async function processAllOrders(accessToken, processor) {
  let processedCount = 0;
  let nextToken = null;
  
  do {
    const url = nextToken 
      ? `https://api-merchant.joom.com/api/v3/orders/multi?after=${nextToken}&limit=500`
      : 'https://api-merchant.joom.com/api/v3/orders/multi?limit=500';
    
    const response = await fetch(url, {
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
    
    // 各注文を処理
    for (const order of data.data) {
      await processor(order);
      processedCount++;
    }
    
    // 次のページがあるかチェック
    nextToken = data.paging?.next ? extractAfterToken(data.paging.next) : null;
    
    // レート制限を考慮して少し待機
    await new Promise(resolve => setTimeout(resolve, 100));
    
  } while (nextToken);
  
  return processedCount;
}

// 使用例：すべての注文を処理
const processedCount = await processAllOrders(accessToken, async (order) => {
  console.log(`処理中: 注文ID ${order.id}, ステータス: ${order.status}`);
  // ここで注文データの処理を行う
});
```

## パフォーマンス最適化

### レート制限の考慮
- **制限**: 50 rpm（1分あたり50リクエスト）
- **推奨**: リクエスト間に適切な間隔を設ける
- **実装例**:
```javascript
async function rateLimitedRequest(url, options) {
  const response = await fetch(url, options);
  
  // レート制限ヘッダーをチェック
  const remaining = response.headers.get('X-Rate-Limit-Remaining');
  if (remaining === '0') {
    const resetTime = response.headers.get('X-Rate-Limit-Reset');
    const waitTime = parseInt(resetTime) * 1000;
    console.log(`レート制限に達しました。${waitTime}ms待機します。`);
    await new Promise(resolve => setTimeout(resolve, waitTime));
  }
  
  return response;
}
```

### バッチ処理の最適化
```javascript
async function batchProcessOrders(accessToken, batchSize = 100) {
  const allOrders = [];
  let nextToken = null;
  
  do {
    const url = nextToken 
      ? `https://api-merchant.joom.com/api/v3/orders/multi?after=${nextToken}&limit=${batchSize}`
      : `https://api-merchant.joom.com/api/v3/orders/multi?limit=${batchSize}`;
    
    const response = await rateLimitedRequest(url, {
      method: 'GET',
      headers: {
        'Accept': 'application/json',
        'Authorization': `Bearer ${accessToken}`
      }
    });
    
    const data = await response.json();
    allOrders.push(...data.data);
    
    nextToken = data.paging?.next ? extractAfterToken(data.paging.next) : null;
    
  } while (nextToken);
  
  return allOrders;
}
```

## 注意事項

1. **レート制限**: 50 rpmの制限があるため、大量のデータを取得する際は適切な間隔を設ける
2. **ページネーション**: 大量のデータを取得する場合は、必ずページネーションを使用する
3. **日時形式**: すべての日時はRFC3339形式（UTC）で指定する
4. **認証**: すべてのリクエストには有効なOAuth 2.0アクセストークンが必要
5. **HTTPS必須**: すべてのリクエストはHTTPS経由で行う
6. **エラーハンドリング**: 適切なエラーハンドリングを実装し、リトライ機能を考慮する
7. **データの整合性**: 取得したデータの整合性をチェックし、必要に応じて再取得する

## 関連リンク

- [Joom for Merchants API v3 公式ドキュメント](https://merchant.joom.com/docs/api/v3)
- [Joom API v3 Get Order（単一注文取得）詳細資料](./Joom_API_v3_Get_Order_詳細資料.md)
- [Joom for Merchants API v3 日本語版](./Joom_for_Merchants_API_v3_日本語版.md)
- [OAuth 2.0認証ガイド](https://merchant.joom.com/docs/api/v3#authentication)
- [エラーハンドリング](https://merchant.joom.com/docs/api/v3#api-errors)
- [レート制限](https://merchant.joom.com/docs/api/v3#request-rate-limits)

---

*この資料は、Joom for Merchants API v3の公式ドキュメントに基づいて作成されています。最新の情報については、必ず公式ドキュメントを参照してください。*
