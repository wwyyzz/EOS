import sqlite3
conn = sqlite3.connect("device.db")

c = conn.cursor()

c.execute('''CREATE TABLE DEVICE
    (id integer PRIMARY KEY autoincrement,
    series_belong text,
    catalong text,
    module_type text, 
    module_sn text, 
    bom text);''')
    
conn.commit()

conn.close()
