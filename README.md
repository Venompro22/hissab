# Hissab — Smart Accounting CLI with Claude AI

> **Hissab** is an intelligent accounting application for the terminal. It helps you track revenue and expenses, analyze your finances with Claude AI, and export your data — all from a beautiful interactive menu.

---

> **حساب** هو تطبيق محاسبة ذكي يعمل في سطر الأوامر. يساعدك على تتبع الإيرادات والمصاريف، وتحليل وضعك المالي بالذكاء الاصطناعي، وتصدير بياناتك بسهولة.

---

> **Hissab** (بالدارجة المغربية: حساب) — واجهة تفاعلية كتخليك تتتبع المداخيل والمصاريف، تحلل الوضع المالي ب Claude AI، وتصدر البيانات — كل هاد الشي من التيرمينال.

---

## ✨ Features

- 🗄️ **SQLite database** — persistent storage, data survives app restarts
- 🤖 **Claude AI analysis** — financial insights powered by Anthropic's Claude Opus 4.8
- 📄 **CSV & Excel export** — export your data with one click (3 sheets in Excel)
- 🌐 **Multi-language support** — switch between Darija 🇲🇦 / English 🇬🇧 / Arabic 🇸🇦 (full UI translation)
- 💬 **AI chat** — ask questions about your finances in natural language
- 🎨 **Beautiful CLI** — rich panels, tables, spinners, and emojis

---

## 🚀 Installation

### 1. Clone the repository

```bash
git clone https://github.com/Venompro22/hissab.git
cd hissab
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Setup your API key

```bash
cp .env.example .env
```

Then open `.env` and add your Anthropic API key:

```
ANTHROPIC_API_KEY=sk-ant-...
```

Get your key at [console.anthropic.com](https://console.anthropic.com)

---

## 📦 Usage

```bash
python3 hissab.py
```

You'll see an interactive menu:

```
╭─────────────────────────── ✨ HISSAB — Menu Principal ───────────────────────╮
│                                                                              │
│      0.  🌐 Bdel langue  (🇲🇦 Darija)                                        │
│      1.  💰 Zid revenu                                                       │
│      2.  💸 Zid mssaref                                                      │
│      3.  📊 Chouf rapport                                                    │
│      4.  🤖 Tahlil Claude AI (mock)                                          │
│      5.  🤖 Tahlil Claude AI (real)                                          │
│      6.  💬 Chat m3a AI                                                      │
│      7.  📁 Export CSV                                                       │
│      8.  📁 Export Excel                                                     │
│      9.  🚪 Khrej                                                            │
│                                                                              │
╰──────────────────────────────────────────────────────────────────────────────╯
```

### Menu options

| Option | Action |
|--------|--------|
| `0` | Switch language (Darija → English → Arabic → ...) |
| `1` | Add revenue entry |
| `2` | Add expense entry |
| `3` | View financial report (table with totals and profit) |
| `4` | AI analysis — mock mode (no API key needed) |
| `5` | AI analysis — real Claude AI (requires API key + credits) |
| `6` | Chat with AI in natural language |
| `7` | Export all data to CSV |
| `8` | Export all data to Excel (3 sheets: Revenue, Expenses, Summary) |
| `9` | Exit |

### Language switching

Press `0` at any time to cycle through all three languages. The **entire interface** switches instantly — menu, prompts, confirmations, report labels, and AI responses.

### Chat examples

```
Su'al dialk: شحال ربحت؟
→ 💰 ربحتي 18,300.00 DH — هامش الربح 79.6%. مزيان!

Your question: what's my profit margin?
→ 📊 Your profit margin is 79.6% — Excellent 🟢

سؤالك: ما هو وضعي المالي؟
→ 📊 الوضع المالي: ممتاز 🟢 | الإيرادات: 23,000 | الربح: 18,300
```

---

## 🗂️ Project Structure

```
hissab/
├── hissab.py          # Main application (class Hissab + menu)
├── .env               # Your API key (not committed to git)
├── .env.example       # Template — copy to .env and fill in your key
├── .gitignore         # Ignores .env, DB, exports, cache
├── requirements.txt   # Python dependencies
├── LICENSE            # MIT License
└── README.md          # This file
```

> `hissab.db` (SQLite) and export files (`*.csv`, `*.xlsx`) are generated at runtime and excluded from git.

---

## 🤝 Contributing

Contributions are welcome! Feel free to:

- Add new languages to the `TRANSLATIONS` dictionary in `hissab.py`
- Suggest new chat keywords or AI prompt improvements
- Open an issue for bugs or feature requests

---

## 📄 License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for details.
