import mysql.connector
import pandas as pd

# Create a connection object to the database
connection = mysql.connector.connect(
    host='localhost',
    user='root',
    password='password',
    database='mydatabase',
)

# Create a cursor object
cursor = connection.cursor()

# Execute an SQL query to select the data you want to extract
cursor.execute('SELECT * FROM customers INNER JOIN orders ON customers.id = orders.customer_id')

# Use the cursor object to fetch the data from the database
customers_orders = cursor.fetchall()

import xlsxwriter

# Create a workbook object
workbook = xlsxwriter.Workbook('customers_orders.xlsx')

# Create a worksheet object
worksheet = workbook.add_worksheet('Customers_Orders')

# Write the header row
header = ['Customer ID', 'Customer Name', 'Order ID', 'Order Date', 'Order Amount']
worksheet.write_row(header)

# Write the data rows
for customer in customers:
    for order in orders:
        if customer['id'] == order['customer_id']:
            row = [customer['id'], customer['name'], order['id'], order['date'], order['amount']]
            worksheet.write_row(row)

# Close the workbook
workbook.close()


# Close the connection to the database
connection.close()
