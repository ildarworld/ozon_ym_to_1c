import logging
import os
from os import listdir
from os.path import isfile, join
import pandas as pd


class Transaction:

    def __init__(self, row):
        self.sku_number = row[('Код товара продавца', 'Unnamed: 6_level_1')]
        self.sell_qty = row[('Реализовано', 'Кол-во')]
        self.sell_price = row[('Реализовано', 'Цена')]
        self.sell_sum_price = row[('Реализовано', 'Сумма, руб.')]
        self.sell_comission = row[('Реализовано', 'Ком-я, руб.')]
        self.returned_qty = row[('Возвращено клиентом', 'Кол-во')]
        self.returned_price = row[('Возвращено клиентом', 'Цена')]
        self.returned_sum_price = row[('Возвращено клиентом', 'Сумма, руб.')]
        self.return_comission = row[('Возвращено клиентом', 'Ком-я, руб.')]

    def __str__(self):
        return f'Sku {self.sku_number} - {self.sell_qty} - {self.sell_comission} - {self.sell_sum_price}'\
            f' Return {self.returned_qty} - {self.returned_sum_price} - {self.return_comission}'

class Sku:
    
    def __init__(self, transaction):
        self.sku_number = transaction.sku_number
        self.sell_qty = transaction.sell_qty - transaction.returned_qty
        self._sell_price = transaction.sell_price
        self._return_comission = transaction.return_comission
        self.sell_sum_price = 0
        self._update_sell_sum_price_by_transaction(transaction)

    def _update_sell_sum_price_by_transaction(self, transaction):
        self.sell_sum_price += transaction.sell_sum_price - transaction.sell_comission \
            - transaction.returned_sum_price + transaction.return_comission

    @property
    def sell_price(self):
        if self.sell_qty:
            return self.sell_sum_price / self.sell_qty
        else:
            return self.sell_sum_price

    def __str__(self):
        return f'{self.sku_number} - {self.sell_qty} - {self.sell_price}'

    def updateByTransaction(self, transaction):
        if transaction.sku_number != self.sku_number:
            logging.error('Sku number is not matching')
            raise Exception
        self.sell_qty += (transaction.sell_qty - transaction.returned_qty)
        self._return_comission = transaction.return_comission
        self._update_sell_sum_price_by_transaction(transaction)


class OzonExcelLoader:

    def __init__(self, filename):
        self.filename = filename
        self._mapping = {}
        self._load_mapping_table()
        self.sku_list = {}

    def _load_mapping_table(self):
        mapping = pd.read_excel('data/sku-mapping-table.xlsx',
                                engine="openpyxl",
                                sheet_name=0,
                                index_col=None,
                                converters={
                                    'ASKU': str,
                                    'артикул': str
                                }
        )
        for _, row in mapping.iterrows():
            self._mapping[row['артикул']] = row['ASKU']

    def load_excel(self):
        print(f'Filename: {self.filename}')
        try:
            payments = pd.read_excel(self.filename,
                                    header=[0,1],
                                    sheet_name=0,
                                    # index_col=0,
                                    skiprows=11,
                                    converters={
                                        'Код товара продавца': str,
                                    }
            ).reset_index()
        except Exception as ex:
            print(f'Error {ex}')
            return
        payments.drop('Unnamed: 0_level_0', axis=1, level=0, inplace=True)
        payments.dropna(axis=0, how="all", subset=[('№ п/п', 'Unnamed: 1_level_1')], inplace=True)

        for _, row in payments.iterrows():
            transaction = Transaction(row)
            if transaction.returned_qty or transaction.sell_qty:
                self.updateSkuByTransaction(transaction)


    def updateSkuByTransaction(self, transaction):
        if transaction.sku_number in self.sku_list.keys():
            self.sku_list[transaction.sku_number].updateByTransaction(
                transaction = transaction
            )
        else:
            sku = Sku(transaction)
            self.sku_list[transaction.sku_number] = sku

    def getListFor1C(self):

        sku_list = []
        qty_list = []
        price_list = []
        return_comission_list = []
        sum_price_list = []
        askus = []

        for sku_number, sku in self.sku_list.items():
            if sku_number in self._mapping.keys():
                sku_list.append(self._mapping[sku_number])
            else:
                sku_list.append(sku.sku_number)
            askus.append(sku.sku_number)
            qty_list.append(sku.sell_qty)
            price_list.append(sku.sell_price)
            sum_price_list.append(sku.sell_sum_price)
            return_comission_list.append(sku._return_comission)

        df = pd.DataFrame.from_dict({
            'Артикул': sku_list,
            'Артикул поставщика': askus,
            'Количество': qty_list,
            'Цена': price_list,
            'Комиссия за возвраты': return_comission_list,
            'Сумма': sum_price_list}
        )
        output_filename = '1C-' + os.path.basename(self.filename)
        df.to_excel(os.path.join(f'output/ozon/', output_filename), header=True, index=False)


class OzonFilesLoader:

    def __init__(self, filepath):
        files = [f for f in listdir(filepath) if isfile(join(filepath, f))]
        print(files)
        for file in  files:
            if not file.startswith('~'):
                ope = OzonExcelLoader(filename=os.path.join(filepath, file))
                ope.load_excel()
                ope.getListFor1C()

oz = OzonExcelLoader(filename='data/ozon/декабрь 2019.xlsx')
oz.load_excel()
oz.getListFor1C()
# OzonFilesLoader(filepath='data/ozon/')

logging.basicConfig(level=logging.INFO)
 


