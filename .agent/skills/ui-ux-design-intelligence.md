---
name: ui-ux-design-intelligence
notify_on_use: true
description: |
  專業 UI/UX 設計知識與規範。當執行任何 UI 設計、前端開發或介面優化任務時自動套用。
  包含配色系統、字體搭配、互動規範、無障礙設計，以及交付前品質檢查清單。
triggers:
  - UI 設計
  - 前端開發
  - 介面優化
  - 樣式調整
  - 元件開發
---

# UI/UX 設計智慧 (Design Intelligence)

## 知識庫搜尋能力

當需要特定設計參考時，可使用以下搜尋工具：

```bash
python3 .shared/ui-ux-pro-max/scripts/search.py "<關鍵字>" --domain <領域>
```

### 可搜尋領域

| 領域 | 用途 | 範例關鍵字 |
|------|------|-----------|
| `product` | 產品類型建議 | SaaS, e-commerce, healthcare, dashboard |
| `style` | UI 風格 | glassmorphism, minimalism, dark mode |
| `typography` | 字體搭配 | elegant, playful, professional |
| `color` | 配色方案 | healthcare, fintech, beauty |
| `landing` | 頁面結構 | hero, testimonial, pricing |
| `chart` | 圖表選型 | trend, comparison, funnel |
| `ux` | 最佳實踐 | animation, accessibility, loading |

---

## 核心設計規範

### 圖示與視覺元素
- ✅ 使用 SVG 圖示 (Lucide, Heroicons)
- ❌ 禁止使用 Emoji 作為 UI 圖示
- ✅ Hover 狀態使用 color/opacity 過渡
- ❌ 避免 scale 變換導致佈局偏移

### 互動與游標
- ✅ 可點擊元素加上 `cursor-pointer`
- ✅ 過渡效果 150-300ms
- ❌ 無視覺回饋的互動元素

### 亮/暗模式對比
- ✅ 亮色模式文字使用 `slate-900` (#0F172A)
- ✅ 玻璃卡片使用 `bg-white/80` 以上透明度
- ❌ 亮色模式避免 `gray-400` 或更淺的文字
- ✅ 邊框在兩種模式皆需可見

### 佈局與間距
- ✅ 浮動導覽列加上 `top-4 left-4 right-4` 間距
- ✅ 預留固定元素高度避免內容遮擋
- ✅ 統一使用 `max-w-6xl` 或 `max-w-7xl`

---

## 交付前檢查清單

### 視覺品質
- [ ] 無 Emoji 作為圖示
- [ ] 圖示來自統一圖示庫
- [ ] Hover 狀態無佈局偏移

### 互動
- [ ] 可點擊元素有 `cursor-pointer`
- [ ] 過渡效果流暢 (150-300ms)
- [ ] 鍵盤 Focus 狀態可見

### 亮/暗模式
- [ ] 亮色模式文字對比度 ≥ 4.5:1
- [ ] 玻璃/透明元素在亮色模式可見
- [ ] 兩種模式下邊框皆可見

### 佈局
- [ ] 浮動元素有適當邊緣間距
- [ ] 無內容被固定元素遮擋
- [ ] 響應式斷點: 320px, 768px, 1024px, 1440px
- [ ] 無水平滾動條

### 無障礙
- [ ] 圖片有 alt 文字
- [ ] 表單輸入有 label
- [ ] 顏色非唯一資訊指示
- [ ] 尊重 `prefers-reduced-motion`

---

## 設計技巧提示

1. **關鍵字要具體** - "healthcare SaaS dashboard" > "app"
2. **多次搜尋** - 不同關鍵字揭示不同洞見
3. **組合領域** - Style + Typography + Color = 完整設計系統
4. **始終檢查 UX** - 搜尋 "animation", "z-index", "accessibility"
5. **檔案拆分** - 元件獨立檔案，每個檔案 200-300 行以內
