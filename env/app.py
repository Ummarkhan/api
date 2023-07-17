from flask import Flask, render_template, request, send_file
import mysql.connector
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
from flask import send_file
import os

# Replace the values with your MySQL database credentials
db = mysql.connector.connect(
    host='localhost',
    user='root',
    password='root123',
    database='exel'
)
cursor = db.cursor()
cursor.execute("CREATE TABLE IF NOT EXISTS orders (order_id INT PRIMARY KEY, product_name VARCHAR(255), product_price FLOAT, shipped varchar(200))")
app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        if file and (file.filename):
            filename = file.filename
            file.save(filename)

            wb = openpyxl.load_workbook(filename)
            sheet = wb.active

            for row in sheet.iter_rows(min_row=2, values_only=True):
                order_id, product_name, product_price, shipped = row

                # Insert data into the database
                query = "INSERT INTO orders (order_id, product_name, product_price, shipped) VALUES (%s, %s, %s, %s)"
                values = (order_id, product_name, product_price, shipped)
                cursor.execute(query, values)
                db.commit()

            wb.close()
            return 'File uploaded and data inserted into the database.'
        else:
            return "upload the file input is empty."

    return render_template('index.html')


@app.route('/download')
def download_file():
    query = "SELECT * FROM orders"
    cursor.execute(query)
    results = cursor.fetchall()

    # Create a new workbook and select the active sheet
    wb = Workbook()
    sheet = wb.active

    # Write the column headers
    headers = ['order_id', 'product_name', 'product_price', 'shipped']
    for col_num, header in enumerate(headers, 1):
        col_letter = get_column_letter(col_num)
        sheet[f'{col_letter}1'] = header

    # Write the data rows
    for row_num, row in enumerate(results, 2):
        for col_num, value in enumerate(row, 1):
            col_letter = get_column_letter(col_num)
            sheet[f'{col_letter}{row_num}'] = value

    # Save the workbook
    file_path = os.path.join(app.root_path, 'orders.xlsx')

    wb.save(file_path)

    return send_file(file_path, as_attachment=True, download_name='orders.xlsx')
@app.route('/delete', methods=['GET'])
def delete_orders():
    cursor = db.cursor()
    cursor.execute("SELECT COUNT(*) FROM orders")
    count = cursor.fetchone()[0]
    if count == 0:
        return "No orders found. The database is already empty."

    cursor.execute("DELETE FROM orders")
    db.commit()
    return "All orders deleted from Database."

if __name__ == '_main_':
    app.run(debug=True)
