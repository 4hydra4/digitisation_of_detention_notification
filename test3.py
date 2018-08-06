import MySQLdb

db = MySQLdb.connect(host="localhost", user="dan", passwd="password", db="fortnight")
cursor = db.cursor()

list_to_store_email = []

#query = "SELECT rollno FROM fortnight1 WHERE (mp/%s) < 0.75"

cursor.execute("SELECT rollno FROM fortnight1 WHERE (mp/12) < 0.75")
rows = cursor.fetchmany(size=2)
for i in rows:
	cursor.execute("SELECT * FROM fortnight1 WHERE rollno = %s" %i)
	result = cursor.fetchone()
	name = result[1]
	email = result[2]
	attendance = result[-1]
	#list_to_store_email.append(email)
	print(result)
	print('%s, %s' %(name, attendance))

#for i in list_to_store_email:
#	print('%s' %i)

db.close()