from PyQt5 import QtCore, QtWidgets, QtGui
from support.stockCalculation import Ui_MainWindow
from support.pandasModel import PandasModel
import sys
import pandas as pd
import yfinance as yf
import numpy as np
from scipy.optimize import minimize

class TickerSelection(QtWidgets.QMainWindow, Ui_MainWindow):

    def __init__(self, *args, **kwargs):
        super( TickerSelection, self).__init__(*args, **kwargs)
        self.setupUi(self)
        #self.update_list()
        self.update_buttons_status()
        self.connections()

        colnames = ['bsecode', 'nsecode', 'compname']
        data = pd.read_csv( '../inputData/Equity.csv', names=colnames )

        BSE_Code = data.bsecode.tolist()
        NSE_Code = data.nsecode.tolist()
        Comp_Name = data.compname.tolist()

        self.widgets = NSE_Code[1:]
        self.listWidget.addItems( NSE_Code[1:] )
        self.lcdNumber.display(self.listWidget.count())

        # Adding Completer.
        self.completer = QtWidgets.QCompleter( self.widgets )
        #self.completer.setCaseSensitivity( QtCore.Qt.CaseInsensitive )

        self.search_le.setCompleter( self.completer )

        self.search_le.textChanged.connect( self.searchItem)

        self.pushButton.clicked.connect(self.optimizedWeights)

    @QtCore.pyqtSlot()
    def update_buttons_status(self):
        self.up_pb.setDisabled(not bool(self.listWidget_2.selectedItems()) or self.listWidget_2.currentRow() == 0)
        self.down_pb.setDisabled(not bool(self.listWidget_2.selectedItems()) or self.listWidget_2.currentRow() == (self.listWidget_2.count() -1))
        self.right_pb.setDisabled(not bool(self.listWidget.selectedItems()) or self.listWidget_2.currentRow() == 0)
        self.left_pb.setDisabled(not bool(self.listWidget_2.selectedItems()))

    def connections(self):
        self.listWidget.itemSelectionChanged.connect(self.update_buttons_status)
        self.listWidget_2.itemSelectionChanged.connect(self.update_buttons_status)
        self.right_pb.clicked.connect(self.on_mBtnMoveToAvailable_clicked)
        self.left_pb.clicked.connect(self.on_mBtnMoveToSelected_clicked)
        self.up_pb.clicked.connect(self.on_mBtnUp_clicked)
        self.down_pb.clicked.connect(self.on_mBtnDown_clicked)

    @QtCore.pyqtSlot()
    def on_mBtnMoveToAvailable_clicked(self):
        self.listWidget_2.addItem(self.listWidget.takeItem(self.listWidget.currentRow()))
        self.lcdNumber_2.display(self.listWidget_2.count())
        self.lcdNumber.display(self.listWidget.count())

    @QtCore.pyqtSlot()
    def on_mBtnMoveToSelected_clicked(self):
        self.listWidget.addItem(self.listWidget_2.takeItem(self.listWidget_2.currentRow()))
        self.lcdNumber_2.display( self.listWidget_2.count() )
        self.lcdNumber.display(self.listWidget.count())

    @QtCore.pyqtSlot()
    def on_mBtnUp_clicked(self):
        row = self.listWidget_2.currentRow()
        currentItem = self.listWidget_2.takeItem(row)
        self.listWidget_2.insertItem(row - 1, currentItem)
        self.listWidget_2.setCurrentRow(row - 1)

    @QtCore.pyqtSlot()
    def on_mBtnDown_clicked(self):
        row = self.listWidget_2.currentRow()
        currentItem = self.listWidget_2.takeItem(row)
        self.listWidget_2.insertItem(row + 1, currentItem)
        self.listWidget_2.setCurrentRow(row + 1)

    def get_left_elements(self):
        r = []
        for i in range(self.listWidget.count()):
            it = self.listWidget.item(i)
            r.append(it.text())
        return r

    def get_right_elements(self):
        r = []
        for i in range(self.listWidget_2.count()):
            it = self.listWidget_2.item(i)
            r.append(it.text())
        return r

    def searchItem( self, text):
        #search_string = self.search_le.toPlainText()
        match_items = self.listWidget.findItems( text, QtCore.Qt.MatchContains )
        for i in range( self.listWidget.count() ):
            it = self.listWidget.item( i )
            it.setHidden( it not in match_items )

    def get_data( self, ticker ):
        try:
            data = yf.Ticker( ticker )
            start_date_qt = self.start_date.date()
            start_date = start_date_qt.toPyDate()
            end_date_qt = self.end_date.date()
            end_date = end_date_qt.toPyDate()
            datadf = data.history( period='1d', start=start_date, end=end_date )
            stock_col = ticker + '-close'
            stock_data = datadf[['Close']]
            stock_data.rename( columns={ 'Close': str( stock_col ) }, inplace=True )
            return stock_data
        except:
            pass

    def getPortRisk( self, weights ):

        labels = [self.listWidget_2.item(i).text() for i in range(self.listWidget_2.count())]
        bsecode = ".BO"
        tickers = [x + bsecode for x in labels]

        df = pd.concat( [self.get_data(ticker) for ticker in tickers], axis=1 )
        df = df.dropna()

        with pd.ExcelWriter('Stock-Risk.xlsx') as writer:
            df.to_excel(writer, sheet_name='Stock-Data')

        returns_df = df.pct_change( 1 ).dropna()  # estimate returns for each asset

        vcv = returns_df.cov()  # being the variance covariance matrix

        var_p = np.dot( np.transpose( weights ), np.dot( vcv, weights ) )  # variance of the multi-asset portfolio
        sd_p = np.sqrt( var_p )  # standard deviation of the multi-asset portfolio
        sd_p_annual = sd_p * np.sqrt( 250 )  # annualised standard deviation of the multi-asset portfolio

        return sd_p_annual

    def optimizedWeights( self ):

        num_stocks = self.listWidget_2.count()  # being the number of stocks (this is a 'global' variable)
        init_weights = [1 / num_stocks] * num_stocks  # initialise weights (x0)

        # Constraint that weights in any asset j must be between 0 and 1 inclusive
        bounds = tuple( (0, 1) for i in range( num_stocks ) )

        # Constraint that the sum of the weights of all assets must equate to 1
        cons = ({ 'type': 'eq', 'fun': lambda x: np.sum( x ) - 1 })

        results = minimize( fun=self.getPortRisk, x0=init_weights, bounds=bounds, constraints=cons )

        # Check total risk of the equal weighted portfolio
        self.getPortRisk( init_weights )

        df = pd.read_excel( 'Stock-Risk.xlsx' )
        df.set_index( 'Date', inplace=True )

        # Explore optimised weights
        optimised_weights = pd.DataFrame( results['x'] )
        optimised_weights.index = df.columns
        optimised_weights.rename( columns={ optimised_weights.columns[0]: 'weights' }, inplace=True )

        # Clean format of the weights so it's more readable
        optimised_weights['weights_rounded'] = optimised_weights['weights'].apply( lambda x: round( x, 3 ) )
        self.model = PandasModel(optimised_weights)
        self.tableView.setModel(self.model)

        with pd.ExcelWriter('../output/Stock-Risk.xlsx') as writer:
            df.to_excel(writer, sheet_name='Stock-Data')
            optimised_weights.to_excel(writer, sheet_name='optimized-weights')


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv )
    qt_app = TickerSelection()
    qt_app.show()
    app.exec_()