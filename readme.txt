# 📊 Stock Risk Optimizer

A PyQt6-based application to calculate optimized investment portfolio weights using historical stock data from Yahoo Finance. The optimization minimizes risk while estimating expected return.

---

## 🚀 Features
- ✅ PyQt6 GUI with background threading
- 📈 Risk and return calculation using historical data
- 🧮 Optimization using `scipy.optimize`
- 🧪 Modular structure with test suite
- 📦 Easily packaged via `pyproject.toml`

---

## 🛠️ Installation
```bash
# Clone the repo
https://github.com/yourusername/stock-risk-optimizer.git
cd stock-risk-optimizer

# Install dependencies
pip install -r requirements.txt
# or build from pyproject
pip install .
```

---

## 🖥️ GUI Usage
```bash
python -m gui.main_window
```

Enter:
- NSE stock tickers (e.g., `RELIANCE, INFY, TCS`)
- Start date (via calendar)
- Click **Run Optimization** to compute weights

📁 Save the optimized weights to Excel after execution.

---

## 💻 CLI Usage (Optional)
```bash
python cli.py --tickers RELIANCE INFY TCS --start-date 2023-01-01
```
(Planned feature: CLI script for headless batch processing)

---

## 🧪 Run Tests
```bash
pytest tests/
```

---

## 📂 Project Structure
```
.
├── gui/                # PyQt6 GUI interface
├── services/           # Core logic: data fetching + optimization
├── tests/              # pytest test cases
├── output/             # Excel exports
├── pyproject.toml      # Project build file
└── README.md
```

---

## 📜 License
MIT License © [Your Name]
