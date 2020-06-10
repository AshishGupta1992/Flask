import sqlite3

conn = sqlite3.connect('revenue.db')
print("Opened database successfully")

conn.execute('CREATE TABLE pl (department TEXT, region TEXT, service_line TEXT, actual_mtd TEXT, budget_mtd TEXT, prior_mtd TEXT, actual_ytd TEXT, budget_ytd TEXT, prior_ytd TEXT)')
print("Table created successfully")
conn.close()