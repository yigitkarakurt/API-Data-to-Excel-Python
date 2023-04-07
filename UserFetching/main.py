import requests
import xlwings as xw
import time

workbook = xw.Book('bilgeDb.xlsx')
sheet = workbook.sheets['Sheet1']
titlesCoordinates = 'A1'
elementsCoordinates = 'B1'

while True:
    # Verileri API'den alın
    link = "https://randomuser.me/api/?results=1"
    response = requests.get(link)
    data = response.json()['results'][0]
    gender = data['gender']
    name = data['name']['title'] + " " + data['name']['first'] + " " + data['name']['last']
    country = data['location']['country']
    mail = data['email']
    age = data['dob']['age']
    titles = ["Gender", "Name", "Country", "Mail", "Age"]
    elements = [gender, name, country, mail, age]

    i = 0
    for i in range(len(elements)):
        newTitlesCoordinates = titlesCoordinates.replace('1', str(i + 1))
        newElementsCoordinates = elementsCoordinates.replace('1', str(i + 1))

        sheet.range(newTitlesCoordinates).value = titles[i]
        sheet.range(newElementsCoordinates).value = elements[i]

        i += 1

    # 3 saniye bekle
    time.sleep(0.01)

def update_excel_sheet():
    # Replace this with your real-time data source
    #data = get_real_time_data()

    # Write data to Excel sheet
    # sheet.range(jokeCoordinates).options(transpose=True).value = joke
    # sheet.range(answerCoordinates).options(transpose=True).value = answer

    rowNum1 = 1

    for j in range(len(elements)):
        sheet.range(rowNum1, 1).value = titles[i]
        sheet.range(rowNum1 + 1, 2).value = elements[i]
        rowNum1 += 1
        j += 1


while True:
    update_excel_sheet()
    time.sleep(0.01)  # Update data every second
