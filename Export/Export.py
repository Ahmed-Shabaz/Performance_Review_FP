import sqlite3
import pandas as pd

# Connect to the SQLite database (use raw string to handle backslashes)
conn = sqlite3.connect(r'C:\\Users\\SHABHAZ AHMED\\OneDrive\\Desktop\\Performance_Review_Final_Plain\\instance\\reviews.db')

# Query the data from the review table
query = 'SELECT * FROM review'

# Use pandas to read the SQL query into a DataFrame
df = pd.read_sql_query(query, conn)

# Write the DataFrame to an Excel file
excel_file = 'Performance Reviews Export.xlsx'
df.to_excel(excel_file, index=False)

# Close the database connection
conn.close()

print(f'Data has been written to {excel_file}')
