# tests/test_portfolio.py
import pytest
import pandas as pd
import numpy as np
from services.portfolio import PortfolioOptimizer, OptimizationResult

@pytest.fixture
def mock_data():
    dates = pd.date_range(start="2023-01-01", periods=250, freq="B")
    data = {
        'STOCK1-Close': np.random.normal(loc=100, scale=10, size=len(dates)),
        'STOCK2-Close': np.random.normal(loc=120, scale=15, size=len(dates)),
        'STOCK3-Close': np.random.normal(loc=80, scale=5, size=len(dates))
    }
    return pd.DataFrame(data, index=dates)

def test_portfolio_optimizer_output(mock_data):
    optimizer = PortfolioOptimizer(mock_data)
    result: OptimizationResult = optimizer.optimize_weights()

    assert isinstance(result, OptimizationResult)
    assert isinstance(result.weights, pd.DataFrame)
    assert result.weights.shape[0] == mock_data.shape[1]
    assert 0 <= result.risk < 10  # Risk should be a small positive number
    assert -1 < result.expected_return < 10  # Expected return reasonable range

def test_portfolio_weights_sum(mock_data):
    optimizer = PortfolioOptimizer(mock_data)
    result = optimizer.optimize_weights()
    total_weight = result.weights['weights'].sum()
    np.testing.assert_almost_equal(total_weight, 1.0, decimal=5)

def test_portfolio_risk_method(mock_data):
    optimizer = PortfolioOptimizer(mock_data)
    weights = np.array([1/3] * 3)
    risk = optimizer.portfolio_risk(weights)
    assert isinstance(risk, float)
    assert risk > 0
