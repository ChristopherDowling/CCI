import os
import csv
import re
import sys
import shutil
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

pdfmetrics.registerFont(TTFont('Lucida', 'C:\Windows\Fonts\LTYPE.TTF'))

consignees = []
consignees_full = []

page                  = 1
consignee_name        = 0
consignee_address1    = 0
consignee_city        = 0
consignee_province    = 0
consignee_postal_code = 0
consignee_id          = 0
product_description   = 0
quantity              = 0
packaging_type        = 0
net_weight            = 0
weight_unit           = 0
shipment_id           = 0
number_of_packages    = 25 # TODO: Kludge
vendor_name           = 37 # TODO: Kludge
vendor_address1       = 0
vendor_zip_code       = 0
vendor_city           = 0
vendor_state          = 0
country_of_origin     = 0
hs_code               = 0
price                 = 0
total_price           = 0


def processed_yet(consignee):
    new = False
    for name in consignees:
        if name == consignee:
            new = True
    return new

#Finds the index for each label eg. "Consignee Name" at index 7
def define_labels(line):
    labels = [
        "consignee_name",
        "consignee_address1",
        "consignee_city",
        "consignee_province",
        "consignee_postal_code",
        "consignee_id",
        "product_description",
        "quantity",
        "packaging_type",
        "net_weight",
        "weight_unit",
        "shipment_id"
        "number_of_packages"
        "vendor_name",
        "vendor_address1",
        "vendor_zip_code",
        "vendor_city",
        "vendor_state",
        "country_of_origin",
        "hs_code",
        "price",
        "total_price"
        ]

    global shipment_id
    global consignee_name
    global consignee_address1
    global consignee_city
    global consignee_province
    global consignee_postal_code
    global consignee_id
    global product_description
    global quantity
    global packaging_type
    global net_weight
    global weight_unit
    global number_of_packages
    global vendor_name
    global vendor_address1
    global vendor_zip_code
    global vendor_city
    global vendor_state
    global country_of_origin
    global hs_code
    global price
    global total_price
    
    print(line)
    print(" ")
    for part in line:
        for label in labels:
            if part == label:
                if part == "consignee_name":
                    consignee_name = line.index(part)
                if part == "consignee_address1":
                    consignee_address1 = line.index(part)
                if part == "consignee_city":
                    consignee_city = line.index(part)
                if part == "consignee_province":
                    consignee_province = line.index(part)
                if part == "consignee_postal_code":
                    consignee_postal_code = line.index(part)
                if part == "consignee_id":
                    consignee_id = line.index(part)
                if part == "number_of_packages":
                    number_of_packages = line.index(part)
                if part == "product_description":
                    product_description = line.index(part)
                if part == "quantity":
                    quantity = line.index(part)
                if part == "packaging_type":
                    packaging_type = line.index(part) # Unused?
                if part == "net_weight":
                    net_weight = line.index(part)
                if part == "weight_unit":
                    weight_unit = line.index(part) #Unused. Always "LBR"
                if part == "shipment_id":
                    shipment_id = line.index(part)
                if part == "vendor_name":
                    vendor_name = line.index(part)
                if part == "vendor_address1":
                    vendor_address1 = line.index(part)
                if part == "vendor_zip_code":
                    vendor_zip_code = line.index(part)
                if part == "vendor_city":
                    vendor_city = line.index(part)
                if part == "vendor_state":
                    vendor_state = line.index(part)
                if part == "country_of_origin":
                    country_of_origin = line.index(part)
                if part == "hs_code":
                    hs_code = line.index(part)
                if part == "price":
                    price = line.index(part)
                if part == "total_price":
                    total_price = line.index(part)

#BEGIN

for arg in sys.argv[1:]:
    if not arg.endswith('.csv'): # Exit if a zip file wan't specified as an argument
        print('Error: Please use a .csv file')
        sys.exit()

lines = list(csv.reader(open(arg)))
CCIs_folder = os.getcwd() + "\\CCIs\\"
if os.path.exists(CCIs_folder):
    shutil.rmtree(CCIs_folder)
os.mkdir(CCIs_folder)

define_labels(lines[0])
del lines[0] #First column is label

#ask for date
date_pattern = re.compile(r'\d{2}-\d{2}-\d{2}')
done = False
while done == False:
    date = input("Please enter the arrival date in the following format: YY-MM-DD")
    if date_pattern.match(date):
        done = True
        #print(shipment)
    else:
        print("Date entered in improper format")

#Find all the consignees
for line in lines:
    new = ''
    if not processed_yet(line[consignee_name]):
        consignees.append(line[consignee_name])
        consignees_full.append(line)

#For each Consignee, create a CCI for them
for consignee in consignees:
    #print(consignee)
    index = consignees.index(consignee)
    gross_weight = 0.0
    total_invoice_price = 0.0
    consignee = consignee.replace('\\', '')
    consignee = consignee.replace('/', '')
    
    #CCI = open(CCIs_folder + consignee + ".txt", "w+")
    CCI_text = []
    CCI_text.append("CANADIAN CUSTOMS INVOICE                                                          PAGE " + str(page).zfill(2) + "")
    CCI_text.append("1. VENDOR (NAME AND ADDRESS)                2. DATE OF DIRECT SHIPMENT TO CANADA         ")
    CCI_text.append("STALLION EXPRESS                            20" + str(date).ljust(43)                 + "")
    CCI_text.append("7676 WOODBINE, UNIT #2                      3. OTHER REFERENCES                          ")
    CCI_text.append("MARKHAM, ON, CANADA                                                                      ")
    CCI_text.append("L3R 2N2                                     5. PURCHASER'S NAME AND ADDRESS              ")
    CCI_text.append("4. CONSIGNEE (NAME AND ADDRESS)                                                          ")
    CCI_text.append(consignees_full[index][consignee_name].ljust(44)            + "6. COUNTRY OF TRANSHIPMENT".ljust(45) + "")
    CCI_text.append(consignees_full[index][consignee_address1].ljust(89)                                   + "")
    CCI_text.append((consignees_full[index][consignee_city] + ", " + consignees_full[index][consignee_province]).ljust(44) + "7. COUNTRY OF ORIGIN".ljust(45) + "")
    CCI_text.append(consignees_full[index][consignee_postal_code].ljust(44)      + consignees_full[index][country_of_origin].ljust(45) + "")
    CCI_text.append("8. TRANSPORTATION (GIVE MODE AND            9. CONDITION OF SALE                         ")
    CCI_text.append("PLACE OF DIRECT SHIPMENT TO CANADA          SALE                                         ")
    CCI_text.append("HIGHWAY TRUCK 327618 / PORT 0427            10. CURRENCY OF SETTLEMENT: USD              ")
    CCI_text.append("                                                                                         ")
    CCI_text.append("11. NO.  12. SPECIFICATION                              13. QUANTITY  14. UNIT  15. TOTAL")
    CCI_text.append("    PKGS     OF COMMODITIES                                               PRICE          ")
    # for each shipment, create a CCI entry with that product
    for line in lines:
        if line[consignee_name] == consignee:
            gross_weight += float(line[net_weight])
            total_invoice_price += float(line[price])
            desc = (line[product_description][0:35] + " - HS " + line[hs_code]).ljust(51)
            CCI_text.append("    " + line[number_of_packages].ljust(9) + desc + line[quantity].rjust(4) + str(round(float(line[price]),2)).rjust(11) + str(round(float(line[total_price]),2)).rjust(10) + "") # TODO: 2-pad prices
    CCI_text.append("                                                                                         ")
    CCI_text.append(("16. TOT WT GROSS LB " + str(round(gross_weight,2)).ljust(6) + "  NET LB " + str(round(gross_weight,2))).ljust(44) + ("17. INVOICE TOTAL " + str(round(total_invoice_price,2))).ljust(45) + "")
    CCI_text.append("18. IF ANY OF FIELDS 1 TO 17 ARE INCLUDED ON AN ATTACHED INVOICE CHECK THIS BOX [ ]      ")
    CCI_text.append("19. EXPORTER'S NAME AND ADDRESS             20. ORIGINATOR                               ")
    CCI_text.append("                                            " + consignees_full[index][vendor_name].ljust(45) + "")
    CCI_text.append("                                            " + consignees_full[index][vendor_address1].ljust(45) + "")
    CCI_text.append("                                            " + (consignees_full[index][vendor_city] + ", " + consignees_full[index][vendor_state]).ljust(45) + "")
    CCI_text.append("                                            " + consignees_full[index][vendor_zip_code].ljust(45)  + "")
    CCI_text.append("21. AGENCY RULING          22. IF FIELDS 23 TO 25 ARE NOT APPLICABLE, CHECK THIS BOX [ ] ")
    CCI_text.append("23. IF INCLUDED IN FIELD 17 INDICATE AMOUNT 24. IF NOT INCLUDED IN FIELD 17 INDICATE AMNT")
    CCI_text.append("(I) TRANSPORTATION CHARGES, EXPENSES AND    (I) TRANSPORTATION CHARGES, EXPENSES AND     ")
    CCI_text.append("INSURANCE FROM THE PLACE OF DIRECT SHIPMENT INSURANCE FROM THE PLACE OF DIRECT SHIPMENT  ")
    CCI_text.append("TO CANADA:                                  TO CANADA:                                   ")
    CCI_text.append("(II) COSTS FOR CONSTRUCTION, ERECTION AND   (II) COSTS FOR CONSTRUCTION, ERECTION AND    ")
    CCI_text.append("ASSEMBLY INCURRED AFTER IMPORTATION INTO    ASSEMBLY INCURRED AFTER IMPORTATION INTO     ")
    CCI_text.append("CANADA:                                     CANADA:                                      ")
    CCI_text.append("(III) EXPORT PACKING:                       (III) EXPORT PACKING:                        ")
    CCI_text.append("25. CHECK IF APPLICABLE:                                                                 ")
    CCI_text.append("(I) ROYALTY PAYMENTS OR SUBSEQUENT PROCEEDS (II) THE PURCHASER HAS SUPPLIED THE GOODS OR ")
    CCI_text.append("ARE PAID OR PAYABLE BY THE PURCHASER [X]    SERVICES FOR USE IN THE PRODUCTION ON THESE  ")
    CCI_text.append("                                            GOODS [ ]                                    ")
    CCI_text.append("                                                                                         ")
    
    #CCI.write(''.join(CCI_text))
    #CCI.close()
    c = canvas.Canvas(CCIs_folder + "\\" + consignees_full[index][consignee_name].replace('/', '') + ".pdf", bottomup = 0)
    c.setFont("Lucida", 10)
    for i, line in enumerate(CCI_text):
        c.drawString(20, 30 + (i * 12), line.upper())
    c.save()
    

#print(out_text)
input("Done! Press Enter to exit")
