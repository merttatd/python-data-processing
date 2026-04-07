import sys
import sqlite3
import random
from datetime import datetime

from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

from openpyxl import Workbook


class Card(QFrame):
    def __init__(self, title):
        super().__init__()
        self.setObjectName("card")

        layout = QVBoxLayout()
        self.title = QLabel(title)
        self.value = QLabel("0")

        self.title.setStyleSheet("color: #aaa;")
        self.value.setStyleSheet("font-size: 22px; font-weight: bold;")

        layout.addWidget(self.title)
        layout.addWidget(self.value)
        self.setLayout(layout)


class ChartWidget(QFrame):
    def __init__(self):
        super().__init__()
        self.setObjectName("card")

        layout = QVBoxLayout()
        self.figure = Figure()
        self.canvas = FigureCanvas(self.figure)

        layout.addWidget(self.canvas)
        self.setLayout(layout)


class DataApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Advanced BI Dashboard")
        self.setGeometry(100, 100, 1400, 800)

        self.conn = sqlite3.connect("data.db")
        self.cursor = self.conn.cursor()

        self.category_colors = {}

        self.create_table()
        self.init_ui()
        self.refresh_all()

    def create_table(self):
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS data (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                category TEXT,
                value REAL,
                time TEXT
            )
        """)
        self.conn.commit()

    def init_ui(self):
        main_layout = QHBoxLayout()

        sidebar = QVBoxLayout()

        title = QLabel("BI Dashboard")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        sidebar.addWidget(title)

        self.category_filter = QComboBox()
        self.category_filter.currentIndexChanged.connect(self.refresh_all)
        sidebar.addWidget(self.category_filter)

        self.category_input = QLineEdit()
        self.category_input.setPlaceholderText("Kategori")

        self.value_input = QLineEdit()
        self.value_input.setPlaceholderText("Değer")

        sidebar.addWidget(self.category_input)
        sidebar.addWidget(self.value_input)

        add_btn = QPushButton("Ekle")
        add_btn.clicked.connect(self.add_data)

        delete_btn = QPushButton("Sil")
        delete_btn.clicked.connect(self.delete_selected)

        delete_all_btn = QPushButton("Tümünü Sil")
        delete_all_btn.clicked.connect(self.delete_all)

        export_btn = QPushButton("Excel Export")
        export_btn.clicked.connect(self.export_excel)

        sidebar.addWidget(add_btn)
        sidebar.addWidget(delete_btn)
        sidebar.addWidget(delete_all_btn)
        sidebar.addWidget(export_btn)

        main_panel = QVBoxLayout()

        kpi_layout = QHBoxLayout()
        self.total_card = Card("Toplam")
        self.count_card = Card("Kayıt")
        self.max_card = Card("Max")

        kpi_layout.addWidget(self.total_card)
        kpi_layout.addWidget(self.count_card)
        kpi_layout.addWidget(self.max_card)

        main_panel.addLayout(kpi_layout)

        grid = QGridLayout()

        self.pie_chart = ChartWidget()
        self.bar_chart = ChartWidget()
        self.trend_chart = ChartWidget()

        grid.addWidget(self.pie_chart, 0, 0)
        grid.addWidget(self.bar_chart, 0, 1)
        grid.addWidget(self.trend_chart, 1, 0, 1, 2)

        main_panel.addLayout(grid)

        self.data_list = QListWidget()
        main_panel.addWidget(self.data_list)

        main_layout.addLayout(sidebar, 1)
        main_layout.addLayout(main_panel, 4)

        self.setLayout(main_layout)

        self.setStyleSheet("""
            QWidget {
                background-color: #1e1f26;
                color: white;
            }
            QPushButton {
                background-color: #3a3f5c;
                padding: 8px;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #5c638a;
            }
            QLineEdit, QComboBox {
                background-color: #2c2f33;
                padding: 6px;
                border-radius: 5px;
            }
            #card {
                background-color: #2c2f33;
                border-radius: 10px;
                padding: 10px;
            }
            #card:hover {
                background-color: #3a3f5c;
            }
        """)

    def get_color(self, category):
        if category not in self.category_colors:
            self.category_colors[category] = (
                random.random(),
                random.random(),
                random.random()
            )
        return self.category_colors[category]

    def add_data(self):
        try:
            category = self.category_input.text()
            value = float(self.value_input.text())
            time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

            self.cursor.execute(
                "INSERT INTO data (category, value, time) VALUES (?, ?, ?)",
                (category, value, time)
            )
            self.conn.commit()
            self.refresh_all()
        except:
            QMessageBox.warning(self, "Hata", "Geçersiz veri")

    def get_data(self):
        return self.cursor.execute("SELECT * FROM data").fetchall()

    def refresh_all(self):
        rows = self.get_data()

        self.data_list.clear()

        total = 0
        count = len(rows)
        max_val = 0
        max_cat = "-"

        categories = set()

        for r in rows:
            self.data_list.addItem(f"{r[1]} - {r[2]} ({r[3]})")

            total += r[2]
            if r[2] > max_val:
                max_val = r[2]
                max_cat = r[1]

            categories.add(r[1])

        self.total_card.value.setText(str(total))
        self.count_card.value.setText(str(count))
        self.max_card.value.setText(max_cat)

        current = self.category_filter.currentText()

        self.category_filter.blockSignals(True)
        self.category_filter.clear()
        self.category_filter.addItem("Tümü")

        for c in categories:
            self.category_filter.addItem(c)

        index = self.category_filter.findText(current)
        if index >= 0:
            self.category_filter.setCurrentIndex(index)

        self.category_filter.blockSignals(False)

        self.draw_all_charts()

    def delete_selected(self):
        row_index = self.data_list.currentRow()
        if row_index < 0:
            return

        record = self.get_data()[row_index]
        self.cursor.execute("DELETE FROM data WHERE id=?", (record[0],))
        self.conn.commit()
        self.refresh_all()

    def delete_all(self):
        confirm = QMessageBox.question(
            self,
            "Onay",
            "Tüm veriler silinsin mi?",
            QMessageBox.Yes | QMessageBox.No
        )

        if confirm == QMessageBox.Yes:
            self.cursor.execute("DELETE FROM data")
            self.conn.commit()
            self.refresh_all()

    def export_excel(self):
        rows = self.get_data()

        wb = Workbook()
        ws = wb.active
        ws.title = "Veriler"

        ws.append(["Kategori", "Değer", "Tarih"])

        for r in rows:
            ws.append([r[1], r[2], r[3]])

        file_name = f"export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        wb.save(file_name)

        QMessageBox.information(self, "Başarılı", f"{file_name} oluşturuldu!")

    def aggregate(self):
        rows = self.get_data()
        data = {}
        for r in rows:
            data[r[1]] = data.get(r[1], 0) + r[2]
        return data

    def draw_all_charts(self):
        data = self.aggregate()

        categories = list(data.keys())
        values = list(data.values())
        colors = [self.get_color(cat) for cat in categories]

        self.pie_chart.figure.clear()
        ax1 = self.pie_chart.figure.add_subplot(111)
        ax1.pie(values, labels=categories, colors=colors)
        self.pie_chart.canvas.draw()

        self.bar_chart.figure.clear()
        ax2 = self.bar_chart.figure.add_subplot(111)
        ax2.bar(categories, values, color=colors)
        self.bar_chart.canvas.draw()

        self.trend_chart.figure.clear()
        ax3 = self.trend_chart.figure.add_subplot(111)

        cat = self.category_filter.currentText()

        if cat != "Tümü":
            rows = self.cursor.execute(
                "SELECT value, time FROM data WHERE category=? ORDER BY time",
                (cat,)
            ).fetchall()

            dates = [datetime.strptime(r[1], "%d/%m/%Y %H:%M:%S") for r in rows]
            values = [r[0] for r in rows]

            ax3.plot(dates, values, marker='o', color=self.get_color(cat))

        self.trend_chart.canvas.draw()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DataApp()
    window.show()
    sys.exit(app.exec_())