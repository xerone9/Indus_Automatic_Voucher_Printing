# Indus_Automatic_Voucher_Printing
 Fetch Data from API and print voucher instantly

# Automatic Voucher priting

During examinations there are a lot of students are approaching account windows for fee vouchers. Created a system in which computers with RFID card readers are placed into various sections

First we download the report from ERP and save it on google sheets (_which has integrated API that returns a JSON object_). That sheet is read by the application and every time the student scans his card his credentials will be searched in the sheet if the credentials found the app will read the relevant information from the sheet and create a URL. That URL (_possess a pdf file_) will then be downloaded and then rotated and then immediately sent to printer for printing.

## Just punch the card and have the voucher
