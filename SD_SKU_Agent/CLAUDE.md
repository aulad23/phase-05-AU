# AGENT — CLAUDE AUTO-READ INSTRUCTIONS

## PRIORITY 1: CHAT FOLDER CHECK
প্রতিটা conversation শুরুতে Claude **সবার আগে** এই folder চেক করবে:
`d:/phase-05 (AU)/Agent/Chat/`
- `.txt` বা `.md` file থাকলে → পড়ে task শুরু করো
- পড়া শেষে `claude_init/` এ save করো

---

## FOLDER STRUCTURE (Vendor-based)
```
Agent/
├── [VendorName]/
│   ├── instruction.json   ← categories + strategy (user confirm করে)
│   ├── Code/
│   │   └── scraper.py
│   ├── Demo/
│   │   └── [VendorName]_demo.xlsx
│   └── Data/
│       └── [VendorName].xlsx
├── Memory/
├── Chat/
└── Vendor_List/
```

---

## WORKFLOW — নতুন Vendor বললে

### STEP 1 — Vendor List থেকে খোঁজো
- `Vendor_List/SD_Web Scraping - Status Tracker.xlsx` খোলো
- Vendor name দিয়ে search করো → URL + sheet data নাও
- Vendor-এর Excel sheet থেকে categories + links পড়ো

### STEP 2 — Website scan করো
- Website-এ যাও → সব categories + links collect করো
- Site type detect করো: Shopify / WordPress / Custom
- Shopify হলে → `/products.json` API test করো
- Sample product data fetch করো → কোন fields আছে দেখো

### STEP 3 — Memory পড়ো
- `Memory/` folder থেকে similar vendor patterns পড়ো
- আগে এই ধরনের site কীভাবে scrape করেছি দেখো
- Best approach নাও

### STEP 4 — Brain (চিন্তা করো)
**`Brain.md` পড়ো** এবং প্রতিটা decision নাও:

1. Site type কী? (Shopify / WordPress / Next.js / Custom)
2. API আছে? → Test করো আগে, Selenium শেষ option
3. কোন columns API-তে আছে, কোনটা page scrape দরকার?
4. Pagination কীভাবে কাজ করে?
5. Duplicate handle করতে হবে কিনা?
6. Rate limiting কতটুকু দরকার?
7. Category name কি Vendor Excel-এর সাথে exact match?

**Brain.md-এর BRAIN OUTPUT FORMAT অনুযায়ী সব analyze করো।**

### STEP 5 — instruction.json তৈরি করো
`Agent/[VendorName]/instruction.json` এ save করো (Brain.md output format):
```json
{
  "vendor": "Brand Name",
  "url": "https://...",
  "site_type": "shopify / wordpress / custom",
  "scraping_method": "Shopify API + Requests+BS4 (product page specs)",
  "pagination": "api_page_param / query_param / button_click / infinite_scroll",
  "rate_limit_seconds": 0.5,
  "demo_per_category": 3,
  "dedup": true,
  "categories": [
    {"name": "Exact Name from Vendor Excel", "handle": "collection-handle", "link": "https://..."}
  ],
  "columns_available": {
    "from_api": ["SKU", "Price", "Weight", "Image URL", "Product Name", "Description"],
    "from_product_page": ["Width", "Depth", "Height", "Finish", "Materials"],
    "not_available": []
  },
  "spec_selector": "CSS selector for specs on product page",
  "notes": "কোনো বিশেষ বিষয় থাকলে এখানে"
}
```

### STEP 6 — User-কে দেখাও
- instruction.json এর summary দেখাও
- বলো: "চেক করুন — update থাকলে বলুন, **Done** বললে code শুরু হবে"

### STEP 7 — User "Done" বললে → Code লিখো
- `Agent/[VendorName]/Code/scraper.py` তৈরি করো
- instruction.json অনুযায়ী সব কিছু implement করো

---

## Demo / Confirm / Done

| Command | Action |
|---------|--------|
| **Demo** | DEMO_MODE=True → 3 products/category → `Demo/` folder |
| **Confirm** | DEMO_MODE=False → full run → `Data/` folder |
| **Done** | git push → শুধু `.py` files |

**গুরুত্বপূর্ণ:** Code update হলে সবসময় Demo আগে — user confirm ছাড়া Full run না।

---

## OUTPUT EXCEL FORMAT (Standard)
- Row 1: Brand name | Row 2: URL | Row 3: Empty | Row 4: Headers | Row 5+: Data
- প্রতি category = আলাদা sheet (sheet name = category name)
- Columns: Index, Category, Manufacturer, Source, Image URL, Product Name, SKU, Product Family Id, Description, Weight, Width, Depth, Diameter, Height, Seat Width, Seat Depth, Seat Height, Arm Height, Price, Finish, Special Order, Location, Materials, Tags, Notes

## GIT RULE
- শুধু "Done" বললে push — শুধু `.py` files, Excel কখনো না
