from datetime import datetime
import os
import logging
import pandas as pd


class Transaction:

    def __init__(self, row):
        self.transactionDate = row['Дата транзакции']
        self.transactionID = row['ID транзакции']
        self.orderNumber = row['Номер заказа']
        self.orderDate = row['Дата оформления']
        self.skuNumber = row['Ваш SKU']
        self.sku_name = row['Название товара']
        self.qty = row['Количество']
        self.transactionPrice = row['Сумма транзакции, руб.'] 
        self.transactionType = row['Тип транзакции'] 
        self.transactionSource = row['Источник транзакции'] 
        self.paymentDate = row['Дата платёжного поручения'] 
        self.paymentNumber = row['Номер платёжного поручения'] 
        self.paymentSum = row['Сумма платёжного поручения'] 


class Sku:

    def __init__(self, sku_number, qty, sum_price, order_number):
        self.list_of_orders = []
        self.sku_number = sku_number
        self.list_of_orders.append(order_number)
        self.qty = qty
        self.sum_price = sum_price

    def __str__(self):
        return f'{self.sku_number} - {self.qty} - {self.sum_price}'


    @property
    def average_sum(self):
        if self.qty != 0:
            return self.sum_price / self.qty
        else:
            return 0
    
    def update_sku_by_order(self, transaction):
        if transaction.skuNumber != self.sku_number:
            logging.info('ALert, different SKUS')
            return
        if transaction.orderNumber not in self.list_of_orders:
            self.list_of_orders.append(transaction.orderNumber)
            if 'Начисление' == transaction.transactionType:
                self.qty += transaction.qty
            elif 'Возврат' == transaction.transactionType and transaction.qty:
                self.qty += transaction.qty * -1
            elif not transaction.qty:
                logging.info(f'Qty is not here order {transaction.orderNumber}')
            logging.info(f'Order Number {transaction.orderNumber} and qty {transaction.qty}')
            self.sum_price += transaction.transactionPrice
        else:
            if 'Начисление' == transaction.transactionType:
                self.sum_price += transaction.transactionPrice
            elif 'Возврат' == transaction.transactionType:
                self.sum_price += transaction.transactionPrice        



class SkuList:

    def __init__(self, month):
        self.skus = {}
        self.month = ''

    def __str__(self):
        skus_str = ''
        for sku in self.skus.values():
            skus_str += f'{sku} '
        return skus_str

    def addSkuByTransaction(self, transaction):
        if transaction.skuNumber in self.skus.keys():
            self.skus[transaction.skuNumber].update_sku_by_order(
                transaction = transaction
            )
        else:
            sku = Sku(
                sku_number=transaction.skuNumber,
                qty=transaction.qty,
                sum_price=transaction.transactionPrice,
                order_number=transaction.orderNumber
            )
            self.skus[transaction.skuNumber] = sku


class YandexPaymentExcelTo1C:

    def __init__(self, filename):
        self.month_skus = {}
        self._mapping = {}
        self.filename = filename
        self._load_mapping_table()

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


    def loadExcelFile(self):

        payments = pd.read_excel(self.filename,
                                engine="openpyxl",
                                sheet_name=0,
                                index_col=None,
                                converters={
                                    'Ваш SKU': str,
                                    'Дата платёжного поручения': str
                                }
        )
        # print(f'Payments: {payments}')
        fillna = {'Дата платёжного поручения': '', 'Количество': 0}
        payments = payments.fillna(value=fillna)
        for _, row in payments.iterrows():
            transaction = Transaction(row)
            if transaction.paymentDate:
                transactionDate = datetime.strptime(transaction.paymentDate, '%d.%m.%Y')
                month_str = f'{transactionDate.month}-{transactionDate.year}' 
                if month_str in self.month_skus.keys():
                    self.month_skus[month_str].addSkuByTransaction(transaction)
                else:
                    month_sku_list = SkuList(month_str)
                    month_sku_list.addSkuByTransaction(transaction)
                    self.month_skus[month_str] = month_sku_list

        for key, value in self.month_skus.items():
            self.save_to_excel(skus_list=value, filename=key)


    def save_to_excel(self, skus_list, filename):
        skus = []
        askus = []
        qtys = []
        avg_prices = []
        sums = []
        for number, sku in skus_list.skus.items():
            if sku.sku_number in self._mapping.keys():
                skus.append(self._mapping[sku.sku_number])
            else:
                skus.append(sku.sku_number)
            askus.append(sku.sku_number)
            qtys.append(sku.qty)
            avg_prices.append(sku.average_sum)
            sums.append(sku.sum_price)

            df = pd.DataFrame.from_dict(
                {'Артикул': skus,
                 'Артикул поставщика': askus,
                 'Количество': qtys,
                 'Цена': avg_prices,
                 'Сумма': sums}
            )
            df.to_excel(f'output/{filename}.xlsx', header=True, index=False)



logging.basicConfig(level=logging.INFO)
 
ype = YandexPaymentExcelTo1C(filename='data/payments_01-10-2019_31-12-2019.xlsx')
ype.loadExcelFile()
