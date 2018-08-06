import openpyxl, MySQLdb, smtplib

#loading excel worksheet
wb = openpyxl.load_workbook('fortnight.xlsx')
sheet = wb['Sheet1']

#db connection
db = MySQLdb.connect(host="localhost", user="dan", passwd="password", db="fortnight")
cursor = db.cursor()

#populating database
for row_cells in sheet.iter_rows(min_row=5, max_row=14):
	list1 = []
	for cell in row_cells:
		list1.append(cell.value)
	print(list1)
	sql_insert_query = """INSERT IGNORE INTO fortnight1 (rollno, name, email, os, mp, sm, rdbms) VALUES (%s, %s, %s, %s, %s, %s, %s)"""
	cursor.execute(sql_insert_query, list1)

#-----------------------------------------------------------------------
gmail_user = input('Enter your email id : ')
gmail_pwd = input('Enter password : ')
SUBJECT = 'YOU ARE DETAINED'

server = smtplib.SMTP('smtp.gmail.com', 587)
server.ehlo()
server.starttls()
server.login(gmail_user, gmail_pwd)
#------------------------------------------------------------------------

list_to_store_attendance = []
dict_to_store_email = {}

#for os------------------------------------------------------------------
cursor.execute("SELECT rollno FROM fortnight1 WHERE (os/%s) < 0.75" %sheet['D3'].value)
rows = cursor.fetchmany(size=10)
for i in rows:
	cursor.execute("SELECT * FROM fortnight1 WHERE rollno = %s" %i)
	result = cursor.fetchone()
	name = result[1]
	email = result[2]
	list_to_store_attendance.append(result[3])
	dict_to_store_email[name] = email

print(dict_to_store_email)
print(list_to_store_attendance)

i = 0

#sending email
for name, email in dict_to_store_email.items():
	body = "Subject: %s. \n%s, \nYou are detained in the subject : OS.\nYour attendance is %s out of 8.\n(THIS IS A TEST. JUST IGNORE. -DANISH)" %(SUBJECT, name, list_to_store_attendance[i])
	print(body)
	i = i+1
	print('Sending email to %s...' %email)
	try:
		server.sendmail(gmail_user, email, body)
	except:
		print('Failed to send email to %s' %email)


list_to_store_attendance.clear()
dict_to_store_email.clear()


#for mp------------------------------------------------------------------
cursor.execute("SELECT rollno FROM fortnight1 WHERE (mp/%s) < 0.75" %sheet['E3'].value)
rows = cursor.fetchmany(size=10)
for i in rows:
	cursor.execute("SELECT * FROM fortnight1 WHERE rollno = %s" %i)
	result = cursor.fetchone()
	name = result[1]
	email = result[2]
	list_to_store_attendance.append(result[4])
	dict_to_store_email[name] = email

print(dict_to_store_email)
print(list_to_store_attendance)

i = 0

#sending email
for name, email in dict_to_store_email.items():
	body = "Subject: %s. \n%s, \nYou are detained in the subject : MP.\nYour attendance is %s out of 12.\n(THIS IS A TEST. JUST IGNORE. -DANISH)" %(SUBJECT, name, list_to_store_attendance[i])
	print(body)
	i = i+1
	print('Sending email to %s...' %email)
	try:
		server.sendmail(gmail_user, email, body)
	except:
		print('Failed to send email to %s' %email)

list_to_store_attendance.clear()
dict_to_store_email.clear()		


#for sm------------------------------------------------------------------
cursor.execute("SELECT rollno FROM fortnight1 WHERE (sm/%s) < 0.75" %sheet['F3'].value)
rows = cursor.fetchmany(size=10)
for i in rows:
	cursor.execute("SELECT * FROM fortnight1 WHERE rollno = %s" %i)
	result = cursor.fetchone()
	name = result[1]
	email = result[2]
	list_to_store_attendance.append(result[5])
	dict_to_store_email[name] = email

print(dict_to_store_email)
print(list_to_store_attendance)

i = 0

#sending email
for name, email in dict_to_store_email.items():
	body = "Subject: %s. \n%s, \nYou are detained in the subject : SM.\nYour attendance is %s out of 10.\n(THIS IS A TEST. JUST IGNORE. -DANISH)" %(SUBJECT, name, list_to_store_attendance[i])
	print(body)
	i = i+1
	print('Sending email to %s...' %email)
	try:
		server.sendmail(gmail_user, email, body)
	except:
		print('Failed to send email to %s' %email)

list_to_store_attendance.clear()
dict_to_store_email.clear()


#for rdbms---------------------------------------------------------------
cursor.execute("SELECT rollno FROM fortnight1 WHERE (rdbms/%s) < 0.75" %sheet['G3'].value)
rows = cursor.fetchmany(size=10)
for i in rows:
	cursor.execute("SELECT * FROM fortnight1 WHERE rollno = %s" %i)
	result = cursor.fetchone()
	name = result[1]
	email = result[2]
	list_to_store_attendance.append(result[-1])
	dict_to_store_email[name] = email

print(dict_to_store_email)
print(list_to_store_attendance)

i = 0

#sending email
for name, email in dict_to_store_email.items():
	body = "Subject: %s. \n%s, \nYou are detained in the subject : RDBMS.\nYour attendance is %s out of 11.\n(THIS IS A TEST. JUST IGNORE. -DANISH)" %(SUBJECT, name, list_to_store_attendance[i])
	print(body)
	i = i+1
	print('Sending email to %s...' %email)
	try:
		server.sendmail(gmail_user, email, body)
	except:
		print('Failed to send email to %s' %email)

list_to_store_attendance.clear()
dict_to_store_email.clear()




db.commit()
db.close()
server.quit()




