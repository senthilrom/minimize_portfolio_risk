from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QLabel, QListWidget,
    QLineEdit, QFileDialog, QVBoxLayout, QWidget, QMessageBox, QDateEdit, QProgressBar,
    QSplashScreen )
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QDate, QPropertyAnimation, QEasingCurve
from PyQt6.QtGui import QPixmap
import sys
import pandas as pd
import time
from datetime import date
from pathlib import Path
from typing import List

# Ensure the top-level path is accessible
sys.path.append(str(Path(__file__).resolve().parent.parent))

from services.data_fetch import StockDataFetcher
from services.portfolio import PortfolioOptimizer, OptimizationResult


class OptimizerThread(QThread):
    finished = pyqtSignal(OptimizationResult)
    error = pyqtSignal(str)

    def __init__(self, tickers: List[str], start_date: date):
        super().__init__()
        self.tickers = tickers
        self.start_date = start_date

    def run(self):
        try:
            fetcher = StockDataFetcher(self.tickers, self.start_date)
            df = fetcher.fetch_all()
            optimizer = PortfolioOptimizer(df)
            result = optimizer.optimize_weights()
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Stock Risk Optimizer")
        self.setFixedSize(500, 450)

        self.layout = QVBoxLayout()

        self.label = QLabel("Enter NSE Tickers (comma separated):")
        self.layout.addWidget(self.label)

        self.input_line = QLineEdit()
        self.layout.addWidget(self.input_line)

        self.label2 = QLabel("Select Start Date:")
        self.layout.addWidget(self.label2)

        self.date_picker = QDateEdit()
        self.date_picker.setDate(QDate.currentDate().addYears(-1))
        self.date_picker.setCalendarPopup(True)
        self.layout.addWidget(self.date_picker)

        self.button = QPushButton("Run Optimization")
        self.button.clicked.connect(self.run_optimization)
        self.layout.addWidget(self.button)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # Indeterminate until finished
        self.progress_bar.setVisible(False)
        self.layout.addWidget(self.progress_bar)

        self.result_label = QLabel("")
        self.result_label.setWordWrap(True)
        self.layout.addWidget(self.result_label)

        container = QWidget()
        container.setLayout(self.layout)
        self.setCentralWidget(container)

    def run_optimization(self):
        tickers = [x.strip() + ".BO" for x in self.input_line.text().split(',') if x.strip()]
        start_date = self.date_picker.date().toPyDate()

        self.progress_bar.setVisible(True)
        self.thread = OptimizerThread(tickers, start_date)
        self.thread.finished.connect(self.display_result)
        self.thread.error.connect(self.show_error)
        self.thread.start()
        self.result_label.setText("Running optimization...")

    def display_result(self, result: OptimizationResult):
        self.progress_bar.setVisible(False)
        weights_str = result.weights[['weights_rounded']].to_string()
        self.result_label.setText(
            f"Optimization Complete:\n\nRisk: {result.risk:.4f}\nExpected Return: {result.expected_return:.4f}\n\nWeights:\n{weights_str}"
        )

        file_path, _ = QFileDialog.getSaveFileName(self, "Save Results", "Stock-Risk.xlsx", "Excel Files (*.xlsx)")
        if file_path:
            path = Path(file_path)
            path.parent.mkdir(parents=True, exist_ok=True)
            df = result.weights.copy()
            df.index = df.index.tz_localize(None) if getattr(df.index, 'tz', None) else df.index
            with pd.ExcelWriter(path) as writer:
                df.to_excel(writer, sheet_name='Optimized-Weights')

    def show_error(self, message: str):
        self.progress_bar.setVisible(False)
        QMessageBox.critical(self, "Error", f"An error occurred: {message}")


def resource_path(relative_path):
    import sys
    if hasattr(sys, '_MEIPASS'):
        return Path(sys._MEIPASS) / relative_path
    return Path(relative_path)

if __name__ == '__main__':
    app = QApplication(sys.argv)

    splash_pix = QPixmap(str(resource_path("assets/splash.png")))
    splash = QSplashScreen(splash_pix)
    splash.setWindowFlag(Qt.WindowType.FramelessWindowHint)
    splash.setWindowOpacity(0.0)

    # Fade-in
    splash.show()
    animation = QPropertyAnimation(splash, b"windowOpacity")
    animation.setDuration(1000)
    animation.setStartValue(0.0)
    animation.setEndValue(1.0)
    animation.setEasingCurve(QEasingCurve.Type.InOutQuad)
    animation.start()

    # Dynamic messages
    splash.showMessage("Initializing modules...", Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignCenter, Qt.GlobalColor.white)
    app.processEvents()
    time.sleep(0.5)

    splash.showMessage("Setting up user interface...", Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignCenter, Qt.GlobalColor.white)
    app.processEvents()
    time.sleep(0.5)

    splash.showMessage("Loading Stock Risk Optimizer...", Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignCenter, Qt.GlobalColor.white)
    app.processEvents()
    time.sleep(0.5)

    window = MainWindow()
    window.show()
    splash.finish(window)

    sys.exit(app.exec())
