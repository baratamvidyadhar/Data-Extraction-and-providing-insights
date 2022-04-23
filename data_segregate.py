# importing the required packages
from directory import prerequisitedirectories
import os
import ssl
import smtplib
import pandas as pd
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

prerequisitedirectories()

# fetching data from excel sheet
li = os.listdir("./data")
if li[1][-5:] == '.xlsx' or li[1][-4:] == '.xls':
    filename = li[1]

excel_data = pd.read_excel("./data/" + filename)

# Dropping the unnecssary columns in the data frame 'excel_data' it'll be done inplace
excel_data.drop([0, 1, 2, 3, 4, 5, 6, 7, 8], inplace=True)

# [ Date	Particulars	NaN	NaN	Vch Type	Vch No.	Debit	Credit] renaming the colunms
excel_data.rename(
    columns={'Unnamed: 1': 'By', 'Unnamed: 2': 'items_with_shop_name', 'Unnamed: 3': 'quan', 'Unnamed: 4': 'price',
             'Unnamed: 5': 'cum_price', 'Unnamed: 6': 'Debit', 'Unnamed: 7': 'Credit'}, inplace=True)

# Length of the excel_data (gives the number of rows)
print(len(excel_data))
excel_data.head()

# reading everest items from .txt file
with open("everest.txt", "r") as fp:
    everest = fp.read()

everest = everest.split("\n")
everest = list(set(everest))

# Length of everest items
print(len(everest))
print(int(excel_data.head()['price'][10]))


# writing the data to the file
def writetofile(shopname, li, filepointer):
    filepointer.write(shopname)
    for i in li:
        filepointer.write(i)


# Storing the everest records in the data structure
def fetchrecords(index, shopname, filepointer):
    li = []
    # print(shopname)
    temp = index + 1
    while index != 0:
        if pd.isna(excel_data['quan'][temp]) == False:
            if excel_data['items_with_shop_name'][temp] in everest:
                li.append(excel_data['items_with_shop_name'][temp] + "," + str(excel_data['quan'][temp]) + "," + str(
                    excel_data['price'][temp]) + "," + str(
                    float(excel_data['quan'][temp]) * float(excel_data['price'][temp])) + "," + str(
                    excel_data['cum_price'][temp]) + "\n")
            temp += 1
        else:
            index = 0
    if len(li) > 0:
        writetofile(shopname, li, filepointer)


# Serving the depencies methods
with open("final_info.txt", "w") as fp2:
    for j in excel_data.index:
        if pd.isna(excel_data['quan'][j]):
            if pd.isna(excel_data['items_with_shop_name'][j]) == False:
                fetchrecords(j, str(excel_data['MARUTHI AGENCIES'][j])[0:10] + "," + str(
                    excel_data['items_with_shop_name'][j]) + "," + str(excel_data['cum_price'][j]) + "," + str(
                    excel_data['price'][j]) + "," + str("Nan") + "\n", fp2)

# Converting the .txt file to .csv
final1 = pd.read_csv("final_info.txt", header=None)
final1.columns = ["date and item_names", "shop_name and quantity", "voucher_number and price",
                  "Total_Price and bill type", "Discounted_total_price"]
final1.to_csv('./result/' + filename.split(".")[0] + ".csv", index=None)

# initializing attachement path with name extenion to send email
mailing_filename_with_path = './result/' + filename.split(".")[0] + ".csv"
print(mailing_filename_with_path)

# deleting .xls and .xlsx files in data folder
files = os.listdir("./data")
for i in files:
    if i[-5:] == '.xlsx' or i[-4:] == '.xls':
        os.remove("./data/" + i)

print("Data transformation completed \ncheck the ./MA/result/{} file".format(filename))
input("Press Enter to close....")

subject = "Everest monthly data"
body = "Please,find the below attachment \n\n\nThanks and Regards,\n Maruthi Agencies,\n Srikakulam."
sender_email = input("Enter sender email")
receiver_email = input("Enter receiver email")
password = input("Enter Sender email password")

# Create a multipart message and set headers
message = MIMEMultipart()
message["From"] = sender_email
message["To"] = receiver_email
message["Subject"] = subject
# message["Bcc"] = receiver_email  #for mass emails

# Add body to email
message.attach(MIMEText(body, "plain"))

# Opening the attachment file in binary mode
with open(mailing_filename_with_path, "rb") as attachment:
    # Add file as application/octet-stream
    # Email client can usually download this automatically as attachment
    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())

# Encode file in ASCII characters to send by email
encoders.encode_base64(part)

# Add header as key/value pair to attachment part
mailing_filename_without_path = mailing_filename_with_path.split("/")[2]
part.add_header(
    "Content-Disposition",
    f"attachment; filename= {mailing_filename_without_path}",
)

# Add attachment to message and convert message to string
message.attach(part)
text = message.as_string()


def sendemail():
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, text)


# Log in to server using secure context and send email
context = ssl.create_default_context()

try:
    sendemail()
    print("Email sent to {} successfully".format(receiver_email))
    os.remove(mailing_filename_with_path)
except:
    print("Failure while sending email to {}".format(receiver_email))

input("Press any key")
