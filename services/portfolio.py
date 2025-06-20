from dataclasses import dataclass
from typing import Tuple
import numpy as np
import pandas as pd
from scipy.optimize import minimize
from pathlib import Path

@dataclass
class OptimizationResult:
    weights: pd.DataFrame
    risk: float
    expected_return: float

class PortfolioOptimizer:
    def __init__(self, stock_data_df: pd.DataFrame):
        self.stock_data = stock_data_df.dropna()
        self.returns_df = self.stock_data.pct_change(1).dropna()
        self.vcv = self.returns_df.cov()

    def portfolio_risk(self, weights: np.ndarray) -> float:
        var_p = np.dot(weights.T, np.dot(self.vcv, weights))
        sd_p = np.sqrt(var_p)
        sd_p_annual = sd_p * np.sqrt(250)
        return sd_p_annual

    def expected_return(self, weights: np.ndarray) -> float:
        expected_return_daily = self.returns_df.mean()
        expected_return_portfolio = weights @ expected_return_daily
        expected_return_annual = (1 + expected_return_portfolio) ** 250 - 1
        return expected_return_annual

    def optimize_weights(self) -> OptimizationResult:
        num_assets = self.stock_data.shape[1]
        init_weights = np.array([1.0 / num_assets] * num_assets)
        bounds = tuple((0, 1) for _ in range(num_assets))
        constraints = ({'type': 'eq', 'fun': lambda x: np.sum(x) - 1})

        result = minimize(
            fun=self.portfolio_risk,
            x0=init_weights,
            bounds=bounds,
            constraints=constraints
        )

        if not result.success:
            raise RuntimeError(f"Optimization failed: {result.message}")

        optimized_weights = pd.DataFrame(result.x, index=self.stock_data.columns, columns=['weights'])
        optimized_weights['weights_rounded'] = optimized_weights['weights'].round(3)

        risk = self.portfolio_risk(result.x)
        expected_ret = self.expected_return(result.x)

        return OptimizationResult(weights=optimized_weights, risk=risk, expected_return=expected_ret)

    def save_to_excel(self, path: str | Path) -> None:
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(path) as writer:
            self.stock_data.to_excel(writer, sheet_name='Stock-Data')
            result = self.optimize_weights()
            result.weights.to_excel(writer, sheet_name='Optimized-Weights')