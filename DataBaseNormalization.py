import sqlite3 as sq
import re
import openpyxl


sqlite_connection = sq.connect('ltx.db')
cursor = sqlite_connection.cursor() 

cursor.execute('CREATE TEMPORARY TABLE "dictionary" (fcdata	NUMERIC, nmdata	NUMERIC)')
cursor.execute('INSERT INTO dictionary(fcdata, nmdata) SELECT PurposeOfTheAssignment, NULL FROM mon')

reader = open(r'C:\Users\3d\Desktop\innertext1.txt', 'r', encoding='utf-8')
text = reader.read()

filter_file = r'[+]?\d{2}\.\d{2}'

normalized_data_from_file_list = re.findall(filter_file, text)

normalized_data_from_file_list = list(dict.fromkeys(normalized_data_from_file_list))
result = ''.join(normalized_data_from_file_list)

float_normalized_data_from_file_list = []
for number in normalized_data_from_file_list:
    float_number = float(number)
    if float_number <= 19:
        float_normalized_data_from_file_list.append(float_number)

for i, a in enumerate(float_normalized_data_from_file_list[1:], start=2):
    cursor.execute(f"UPDATE dictionary SET nmdata = ? WHERE ROWID = {i};", (a,))

cursor.execute("""CREATE TEMPORARY TABLE mon2(
	Cudnum TEXT,
	Koatyy TEXT,
	Area TEXT,
	District TEXT,
	Settlement TEXT,
	City TEXT,
	Street TEXT,
	TGName TEXT,
	PurposeOfTheAssignment REAL,
    PurposeOfTheAssignmentNormal REAL,
	NameOfTheSite TEXT,
	TransactionType TEXT,
	Price TEXT,
	OwnershipType TEXT,
	RegistrationDate TEXT,
	RegistrationNumber TEXT,
	ValueNGO TEXT,
	EvaluationDate TEXT
)""")

cursor.execute("""INSERT INTO mon2(
    Cudnum,
    Koatyy,
    Area,
    District,
    Settlement,
    City,Street,
    TGName,
    PurposeOfTheAssignment,
    PurposeOfTheAssignmentNormal,
    NameOfTheSite,
    TransactionType,
    Price,
    OwnershipType,
    RegistrationDate,
    RegistrationNumber,
    ValueNGO,
    EvaluationDate) SELECT 
    Cudnum,
    Koatyy,
    Area,
    District,
    Settlement,
    City,Street,
    TGName,
    PurposeOfTheAssignment,
    PurposeOfTheAssignmentNormal,
    NameOfTheSite,
    TransactionType,
    Price,
    OwnershipType,
    RegistrationDate,
    RegistrationNumber,
    ValueNGO,
    EvaluationDate
    FROM mon""")

cursor.execute('CREATE TABLE IF NOT EXISTS monitoring AS SELECT * FROM mon2')

cursor.execute('''UPDATE monitoring SET PurposeOfTheAssignmentNormal = (SELECT nmdata FROM dictionary WHERE monitoring.PurposeOfTheAssignment LIKE '%' || dictionary.nmdata || '%') WHERE PurposeOfTheAssignmentNormal IS NULL''')

reader.close()

cursor.execute('SELECT PurposeOfTheAssignmentNormal, Area, Price, ValueNGO, AVG(Price/Area) AS Pa, AVG(Price/ValueNGO) AS Pv FROM monitoring GROUP BY PurposeOfTheAssignmentNormal')

rows = cursor.fetchall()
for row in rows:
    purpose = row[0]
    area = row[1]
    price = row[2]
    value = row[3]
    print(purpose, area, price, value)

workbook = openpyxl.load_workbook('F:\Prog\Py\excel\excel draft\opus3.xlsx')
sheet = workbook['123']

row_num = 8
for row in rows:
    purpose = row[0]
    area = row[1]
    price = row[2]
    value = row[3]
    pa = row[4]
    pv = row[5]
    sheet.cell(row=row_num, column=2).value = purpose
    sheet.cell(row=row_num, column=5).value = area
    sheet.cell(row=row_num, column=6).value = price
    sheet.cell(row=row_num, column=7).value = value
    sheet.cell(row=row_num, column=8).value = pa
    sheet.cell(row=row_num, column=9).value = pv
    row_num += 1

workbook.save('F:\Prog\Py\excel\excel draft\opus3.xlsx')
sqlite_connection.commit()
cursor.close()