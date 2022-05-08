from cProfile import label
import matplotlib.pyplot as plt
import pandas as pd
from openpyxl import load_workbook

# Date range A13-A1960

class ExchangeRate:
    def __init__(self):
        self.wb = load_workbook(filename='data/exchange_rate.xlsx')
        self.ws = self.wb.active
        self.date_col = self.ws['A']
        pass

class AUD(ExchangeRate):
    '''
    Australian Dollar x Canadian Dollar
    '''
    def __init__(self):
        super().__init__()
        self.name = self.ws['B10'].value
        self.AUD_col = self.ws['B']
        self.AUD_db = pd.DataFrame({'Date': [i.value for i in self.date_col[13:1346]], 'Dollars': [i.value for i in self.AUD_col[13:1346]]})

    def plot_AUD(self):
        self.fig, self.ax = plt.subplots(figsize=(18,9))
        self.ax.plot(self.AUD_db['Date'], self.AUD_db['Dollars'], label='1 AUD')
        self.ax.set_title(self.name)
        self.ax.set_xlabel('Date')
        self.ax.set_ylabel('Dollars')
        self.ax.set_title('AUD/CAD Exchange Rate')
        self.ax.grid(True)
        self.ax.legend()

class YUAN(ExchangeRate):
    '''
    Chinese Renminbi x Canadian Dollar
    '''
    def __init__(self):
        super().__init__()
        self.name = self.ws['C10'].value
        self.YUAN_col = self.ws['C']
        self.YUAN_db = pd.DataFrame({'Date': [i.value for i in self.date_col[13:1346]], 'Dollars': [i.value for i in self.YUAN_col[13:1346]]})
    def plot_YUAN(self):
        self.fig, self.ax = plt.subplots(figsize=(18,9))
        self.ax.plot(self.YUAN_db['Date'], self.YUAN_db['Dollars'], label='1 Yuan')
        self.ax.set_title(self.name)
        self.ax.set_xlabel('Date')
        self.ax.set_ylabel('Dollars')
        self.ax.set_title('YUAN/CAD Exchange Rate')
        self.ax.grid(True)
        self.ax.legend()

class INR(ExchangeRate):
    '''
    Indian Rupee x Canadian Dollar
    '''
    def __init__(self):
        super().__init__()
        self.name = self.ws['D10'].value
        self.rupee_col = self.ws['D']
        self.rupee_db = pd.DataFrame({'Date': [i.value for i in self.date_col[13:1346]], 'Dollars': [i.value for i in self.rupee_col[13:1346]]})
    def plot_INR(self):
        self.fig, self.ax = plt.subplots(figsize=(18,9))
        self.ax.plot(self.rupee_db['Date'], self.rupee_db['Dollars'], label='1 INR')
        self.ax.set_title(self.name)
        self.ax.set_xlabel('Date')
        self.ax.set_ylabel('Dollars')
        self.ax.set_title('INR/CAD Exchange Rate')
        self.ax.grid(True)
        self.ax.legend()

class RUB(ExchangeRate):
    '''
    Russian Ruble x Canadian Dollar
    '''
    def __init__(self):
        super().__init__()
        self.name = self.ws['E10'].value
        self.RUB_col = self.ws['E']
        self.RUB_db = pd.DataFrame({'Date': [i.value for i in self.date_col[13:1346]], 'Dollars': [i.value for i in self.RUB_col[13:1346]]})
    def plot_RUB(self):
        self.fig, self.ax = plt.subplots(figsize=(18,9))
        self.ax.plot(self.RUB_db['Date'], self.RUB_db['Dollars'], label='1 RUB')
        self.ax.set_title(self.name)
        self.ax.set_xlabel('Date')
        self.ax.set_ylabel('Dollars')
        self.ax.set_title('RUB/CAD Exchange Rate')
        self.ax.grid(True)
        self.ax.legend()