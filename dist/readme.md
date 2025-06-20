# README.md

# 📊 Stock Risk Optimizer

A PyQt6-based application to calculate optimized investment portfolio weights using historical stock data from Yahoo Finance. The optimization minimizes risk while estimating expected return.

---

## 🚀 Features
- ✅ PyQt6 GUI with splash screen and fade-in animation
- 📈 Risk and return calculation using historical data
- 🧮 Optimization using `scipy.optimize`
- 🧪 Modular structure with test suite
- 🧭 Help menu with About and embedded README viewer
- 📦 Easily packaged via `pyinstaller`

---

## 🛠️ Installation
```bash
# Clone the repo
https://github.com/senthilrom/stock-risk-optimizer.git
cd stock-risk-optimizer

# Install dependencies
pip install -r requirements.txt
# or use pyproject
pip install .
```

---

## 🖥️ GUI Usage
```bash
python gui/main_window.py
```

### GUI Components:
- 📋 Enter NSE tickers (comma-separated)
- 📆 Select historical start date
- ▶️ Run optimization to compute weights
- 📤 Save results as Excel (.xlsx)
- 📑 Help menu → "About" and "View README"

---

## 💻 CLI Usage
```bash
python cli.py --tickers RELIANCE INFY TCS --start-date 2023-01-01
```

Outputs to console and saves Excel results to `/output` folder.

---

## 🧪 Run Tests
```bash
pytest tests/
```

---

## 📦 Packaging as EXE
```bash
pyinstaller StockRiskOptimizer.spec
```

Includes:
- 📁 Assets (icon, splash)
- 📄 README.md
- 💼 Equity CSV templates

Run the EXE from `dist/StockRiskOptimizer/StockRiskOptimizer.exe`.

---

## 🗂 Project Structure
```
.
├── gui/                # PyQt6 GUI interface
├── services/           # Core logic: data fetching + optimization
├── tests/              # pytest test cases
├── assets/             # app.ico, splash.png, input CSVs
├── cli.py              # CLI runner
├── README.md           # Displayed in GUI Help → README
├── pyproject.toml      # Project build file
└── StockRiskOptimizer.spec  # PyInstaller spec
```

---

## 📜 License
MIT License © 2025 senthilrom