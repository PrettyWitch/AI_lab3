from PySide2.QtCore import QFile, Qt
from PySide2.QtUiTools import QUiLoader
from PySide2.QtWidgets import QApplication, QCheckBox
from PySide2.QtGui import QStandardItemModel, QStandardItem
import xlrd as xlrd
import utils

# store user search history
history = []


class Search:
    def __init__(self):
        # Load UI definition from file
        qfile_stats = QFile('untitled.ui')
        qfile_stats.open(QFile.ReadOnly)
        qfile_stats.close()

        # Dynamically create a corresponding window object from the UI definition
        # Note: The control object inside has also become a property of the window object
        # such as self.ui.button , self.ui.textEdit
        self.ui = QUiLoader().load(qfile_stats)
        self.model = QStandardItemModel(4, 10)
        self.model.setHorizontalHeaderLabels(['name', 'id', 'Import', 'weight', 'price',
                                              'discount', 'shelfLife', 'sweetness', 'hardness', 'food'])

        self.ui.lineEdit.setPlaceholderText("Please enter")
        self.ui.button1.clicked.connect(self.search)
        # self.ui.label.setText("no such fruit")
        self.ui.button2.clicked.connect(self.add_sub)
        self.model2 = QStandardItemModel()
        print(self.search())

        date = self.readData()
        values = date[0]
        nrows = date[1]
        ncols = date[2]

        # Set table row height and column width
        # self.ui.tableView.resizeRowsToContents()
        for column in range(ncols - 3):
            self.ui.tableView.setColumnWidth(column, 30)
        # self.ui.tableView.setColumnWidth(0, 30)
        # self.ui.tableView.setColumnWidth(8, 120)
        # self.ui.tableView.setColumnWidth(9, 100)
        # self.ui.tableView.setColumnWidth(10, 70)
        # for column in range(1, ncols - 3):
        #     self.ui.tableView.setColumnWidth(column, 70)

        self.ui.tableView.setModel(self.model)
        self.ui.listView.setModel(self.model2)

    def readData(self):
        # Read table data
        book = xlrd.open_workbook('fruit.xlsx')
        sheet1 = book.sheets()[0]
        nrows = sheet1.nrows
        ncols = sheet1.ncols
        values = []
        for row in range(nrows):
            row_values = sheet1.row_values(row)
            values.append(row_values)
        return values, nrows, ncols

    def search(self):
        global history
        itemNum = -1
        name = self.ui.lineEdit.text()
        print(name)
        history.append(name)
        print(history)
        data = self.readData()
        nrows = data[1]
        values = data[0]
        ncols = data[2]
        # Find the row of the searched fruit -  row
        for row in range(nrows):
            cell = values[row][1]
            # cell = sheet1.cell(row, 1)
            if name.lower() == cell.lower():
                itemNum = row
                # return itemNum
                print(itemNum)
                print(values[itemNum])
                break
            else:
                continue
        if itemNum < 1:
            self.ui.label.setText("no such fruit")
            # print("error")
        else:
            # Use dictionary to store line number and distance information
            dict = {}
            for row in range(1, nrows):
                if itemNum != row:
                    a = utils.tra(values[itemNum], values[row])
                    b = utils.tra2(values[row])
                    row_num = values[row][0]
                    distance = utils.PearsonCorrelation(a, b)
                    dict[row_num] = distance
                    # print(f"{name} - {name2} distance = {distance}")
                else:
                    continue
            # Dictionary sort
            newDis = sorted(dict.items(), key=lambda d: d[1], reverse=True)
            print(newDis[0], newDis[1], newDis[2])
            fruit1 = int(newDis[0][0])
            fruit2 = int(newDis[1][0])
            fruit3 = int(newDis[2][0])
            print(fruit1, fruit2, fruit3)
            index = [itemNum, fruit1, fruit2, fruit3]
            # distance = [1, newDis[0][1]]
            for row in range(len(index)):
                for col in range(1, ncols):
                    item = QStandardItem('%s' % values[index[row]][col])
                    self.model.setItem(row, col - 1, item)
                    # self.ui.tableView.resizeRowsToContents()
            return index, name

    def add_sub(self):
        data = self.readData()
        nrows = data[1]
        values = data[0]
        ncols = data[2]
        index = self.search()[0]
        for row in range(len(index)):
            item = QStandardItem('%s' % values[index[row]][1])
            self.model2.setItem(row, item)


app = QApplication([])
search = Search()
search.ui.show()
app.exec_()
