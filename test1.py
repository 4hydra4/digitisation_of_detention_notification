import MySQLdb, openpyxl

wb = openpyxl.load_workbook('fortnight.xlsx')
sheet = wb['Sheet1']

db = MySQLdb.connect(host="localhost", user="dan", passwd="password", db="fortnight")
cursor = db.cursor()

#cursor.execute("insert into fortnight1 (rollno) values (1001)")
for row_cells in sheet.iter_rows(min_row=5, max_row=6):
	list1 = []
	for cell in row_cells:
		list1.append(cell.value)
	print(list1)
	sql_insert_query = """INSERT INTO fortnight1 (rollno, name, email, os, mp) VALUES (%s, %s, %s, %s, %s)"""
	#cursor.executemany("INSERT INTO fortnight1 (rollno, name, email, os, mp) VALUES ({0}, '{1}', '{2}', {3}, {4})".format(*list1))
	#cursor.executemany("""INSERT INTO fortnight1 (rollno, name, email, os, mp) VALUES (%s, '%s', '%s', %s, %s)""" %(list1))
	#cursor.executemany("INSERT INTO fortnight1 (rollno, name, email, os, mp) VALUES ('%s', '%s', '%s', '%s', '%s')" %list1)
	cursor.execute(sql_insert_query, list1)
db.commit()
db.close()