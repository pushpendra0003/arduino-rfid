import serial
import gspread
import datetime
from pymongo import MongoClient



from oauth2client.service_account import ServiceAccountCredentials

from pprint import pprint as pp


ser = serial.Serial('COM7', 9600)




scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json",scope)
client = gspread.authorize(creds)

sheet = client.open("RFID_Testing").sheet1

#Getting data from the sheet
data = sheet.get_all_records()

pp(data)


#Get the specific row, column and cell from the sheet

row = sheet.row_values(3)
col = sheet.col_values(3)
cell = sheet.cell(1,2).value

pp(cell)

#Inserting data in your sheet 





print("The row has been added")
# sheet.delete_rows(2)
# pp("The row has been deleted")

ser.close()

# Define the serial port and baud rate
serial_port = 'COM7'  # Change this to your serial port
baud_rate = 9600  # Change this to match your Arduino's baud rate

# MongoDB Atlas configuration
# Replace the connection URI with your MongoDB Atlas connection string
mongodb_uri = "mongodb+srv://nishantkumar32435:Nishant1194@cluster0.pfguow8.mongodb.net/"

# Initialize the MongoDB client with the Atlas connection URI
client = MongoClient(mongodb_uri)

# Specify the database and collection
database_name = 'test'
collection_name = 'systems'

# Access the specified database and collection
db = client[database_name]
collection = db[collection_name]

# Initialize the serial connection
ser = serial.Serial(serial_port, baud_rate)

my_dict = {
    "example_key": 0,
    "192":1,
    "245":1,
}


i=2

try:
    while True:
        current_time = datetime.datetime.now()
        formatted_date = current_time.strftime("%Y-%m-%d") 
        formatted_samay = current_time.strftime("%H:%M:%S")

        if(i==2):
            sample = ser.readline().strip().decode()

        name = ser.readline().strip().decode()
        status= ser.readline().strip().decode()
        uniqueid= ser.readline().strip().decode()


        pp(my_dict)


            
        # my_dict[uniqueid] = (int(my_dict[uniqueid])+1)%2 



        status_dict = {
                   0: "Checked Out",
                   1: "Checked In",
                  }
        
        data = {
            "day": formatted_date,
            "time": formatted_samay,
            "name": name,
            "status":status_dict[my_dict[uniqueid]],
            "uniqueid":uniqueid,
            }
        collection.insert_one(data)
        # data4 = ser.readline().strip().decode()
        if i%2 == 0 :
            my_dict[uniqueid]+=1
            my_dict[uniqueid]%=2
            insertRow = [name,status_dict[my_dict[uniqueid]],formatted_date,formatted_samay,uniqueid]
        else :
            my_dict[uniqueid]+=1
            my_dict[uniqueid]%=2
            insertRow = [name,status_dict[my_dict[uniqueid]],formatted_date,formatted_samay,uniqueid]


        sheet.insert_row(insertRow,i)   
        i+=1
except KeyboardInterrupt:
    # Close the serial connection and MongoDB client when Ctrl+C is pressed
    ser.close() 
    client.close()