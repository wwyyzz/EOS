import sqlite3
conn = sqlite3.connect("device.db")

c = conn.cursor()

c.execute('''CREATE TABLE DEVICE
    (id integer PRIMARY KEY autoincrement,
    device_type text, 
    module_type text, 
    module_sn text, 
    bom text);''')
    
conn.commit()

conn.close()
