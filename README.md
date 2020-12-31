# Invoice-Automation

Automated the invoice the process for a local business.

Program uses the python package fitz to process pdf data, runs the data through a processing algorithm and stores the data into excel through python. 

This repository does not contain invoice data due to company security. 

The code was written in Python and starts on line 20 of the repository. 
Enjoy:)











import fitz
import xlsxwriter
import time
from datetime import date
today = str(date.today())
Invoice_List = ['Invoice 30165841.pdf', 'Invoice 30165828.pdf', 'Invoice 30165831.pdf']

#Opening Excel
book = xlsxwriter.Workbook('Today.xlsx')
sheet = book.add_worksheet()


p = 0
b = 0
for x in Invoice_List:
    p += 1
    file = fitz.open(x)
    sheet.write(p, 0, x)
#Loading In PDF
    page = file.loadPage(0)
    text = page.getText('text')
    text = text.split()

#Seperating Text Into Easy To Handle & Hard To Handle
    Start_Chaos = text.index('vary***')
    End_Chaos = text.index('Document')

    #Defining Easy to Handle Data
#Obtaining Invoice number
    Invoice_Count_text = int(text.count('Invoice'))
    text_Invoice_Trial = text
    for y in range(0, Invoice_Count_text):
        Reference = int(text_Invoice_Trial.index('Invoice'))
        if text_Invoice_Trial[Reference] + (text_Invoice_Trial[Reference+1]) == 'Invoice#':
            Invoice_Number = text_Invoice_Trial[Reference-1]
            sheet.write(0, 1, 'Invoice Number')
            sheet.write(p, 1, Invoice_Number)
            break
        else:
            text_Invoice_Trial = text_Invoice_Trial[(Reference+1):]


    Printed_Reference = text.index('Printed')
    Phone_Reference = text.index('Phone:')
    text_Invoice_Trial = text
    for y in range(0, Invoice_Count_text):
        Reference = int(text_Invoice_Trial.index('Invoice'))
        if text_Invoice_Trial[Reference] + (text_Invoice_Trial[Reference + 1]) == 'InvoiceDate':
            Invoice_Date = text_Invoice_Trial[Reference +2]
            sheet.write(0, 2, 'Invoice Date')
            sheet.write(p, 2, Invoice_Date)
            break
        else:
            text_Invoice_Trial = text_Invoice_Trial[(Reference + 1):]

    Invoice_Count_Quantity = int(text.count('Quantity'))
    text_Invoice_Trial = text
    for y in range(0, Invoice_Count_Quantity):
        Reference = int(text_Invoice_Trial.index('Quantity'))
        if text_Invoice_Trial[Reference] + (text_Invoice_Trial[Reference + 1]) == 'QuantityProduct':
            Customer_Number = text_Invoice_Trial[Reference - 1]
            sheet.write(0, 3, 'Customer#')
            sheet.write(p, 3, Customer_Number)
            break
        else:
            text_Invoice_Trial = text_Invoice_Trial[(Reference + 1):]

    Customer_Number_Reference = int(text.index(Customer_Number))
    SalesPerson = (text[(Customer_Number_Reference-2):(Customer_Number_Reference)])
    SalesPerson_List = SalesPerson
    SalesPerson=(' '.join(SalesPerson))
    sheet.write(0, 4, 'Sales Person')
    sheet.write(p, 4, SalesPerson)

    ShipTo_Reference = text.index('Telephone:')
    ShipTo_Reference2 = text.index('USA')
    Ship_To_Address = text[ShipTo_Reference2+2:ShipTo_Reference]
    Ship_To_Address = (' '.join(Ship_To_Address))
    sheet.write(0, 5, 'Ship To Address')
    sheet.write(p, 5, Ship_To_Address)

    #Payment Terms
    Invoice_Count_Quantity = int(text.count('Customer'))
    text_Invoice_Trial = text
    for y in range(0, Invoice_Count_Quantity):
        Reference = int(text_Invoice_Trial.index('Customer'))
        if text_Invoice_Trial[Reference] + (text_Invoice_Trial[Reference + 1]) + (text_Invoice_Trial[Reference+2]) == 'CustomerPO#2':
            Salesperson_Reference = int(text_Invoice_Trial.index(SalesPerson_List[0]))
            Payment_Terms = text_Invoice_Trial[Reference+2:Salesperson_Reference]
            Payment_Terms = (' '.join(Payment_Terms))
            sheet.write(0, 6, 'Payment_Terms')
            sheet.write(p, 6, Payment_Terms)
            break
        else:
            text_Invoice_Trial = text_Invoice_Trial[(Reference + 1):]

    #SalesOrder#
    Invoice_Count_Quantity = int(text.count('Sales'))
    text_Invoice_Trial = text
    for y in range(0, Invoice_Count_Quantity):
        Reference = int(text_Invoice_Trial.index('Sales'))
        if text_Invoice_Trial[Reference] + (text_Invoice_Trial[Reference + 1]) + (text_Invoice_Trial[Reference+2]) == 'SalesOrder#':
            Sales_Order_Number = text_Invoice_Trial[Reference - 1]
            sheet.write(0, 7, 'Sales Order Number')
            sheet.write(p, 7, Sales_Order_Number)
            break
        else:
            text_Invoice_Trial = text_Invoice_Trial[(Reference + 1):]

    #OrderDate
    Invoice_Count_Quantity = int(text.count('Order'))
    text_Invoice_Trial = text
    for y in range(0, Invoice_Count_Quantity):
        Reference = int(text_Invoice_Trial.index('Order'))
        if text_Invoice_Trial[Reference] + (text_Invoice_Trial[Reference + 1]) == 'OrderDate':
            Order_Date = text_Invoice_Trial[Reference +2]
            sheet.write(0, 8, 'Order Date')
            sheet.write(p, 8, Order_Date)
            break
        else:
            text_Invoice_Trial = text_Invoice_Trial[(Reference + 1):]

    #Shipment Date
    Invoice_Count_Quantity = int(text.count('Shipment'))
    text_Invoice_Trial = text
    for y in range(0, Invoice_Count_Quantity):
        Reference = int(text_Invoice_Trial.index('Shipment'))
        if text_Invoice_Trial[Reference] + (text_Invoice_Trial[Reference + 1]) == 'ShipmentDate':
            Shipment_Date = text_Invoice_Trial[Reference + 2]
            sheet.write(0, 9, 'Shipment_Date')
            sheet.write(p, 9, Shipment_Date)
            break
        else:
            text_Invoice_Trial = text_Invoice_Trial[(Reference + 1):]

    #CustomerPO#
    Invoice_Count_Quantity = int(text.count('Customer'))
    text_Invoice_Trial = text
    for y in range(0, Invoice_Count_Quantity):
        Reference = int(text_Invoice_Trial.index('Customer'))
        if text_Invoice_Trial[Reference] + (text_Invoice_Trial[Reference + 1]) == 'CustomerPO#':
            text_Invoice_Trial = text_Invoice_Trial[(Reference + 1):]
            Customer2 = int(text_Invoice_Trial.index('Customer'))
            Customer_PO_Number = text_Invoice_Trial[1:Customer2]
            Customer_PO_Number = (' '.join(Customer_PO_Number))
            sheet.write(0, 10, 'Customer PO Number')
            sheet.write(p, 10, Customer_PO_Number)
            break
        else:
            text_Invoice_Trial = text_Invoice_Trial[(Reference + 1):]

    #Carrier/Unit
    PRSCO_Number_Reference = text.index('215-785-3141')
    Carrier_List = text[PRSCO_Number_Reference:]
    Terms_Reference = Carrier_List.index('Terms')
    Carrier_Unit = Carrier_List[1:Terms_Reference]
    Carrier_Unit = (' '.join(Carrier_Unit))
    sheet.write(0, 11, 'Carrier Unit')
    sheet.write(p, 11, Carrier_Unit)

    #FreightTerms
    Invoice_Count_Quantity = int(text.count('Freight'))
    text_Invoice_Trial = text
    for y in range(0, Invoice_Count_Quantity):
        Reference = int(text_Invoice_Trial.index('Freight'))
        if text_Invoice_Trial[Reference] + (text_Invoice_Trial[Reference + 1]) == 'FreightTerms':
            Freight_Terms = text_Invoice_Trial[Reference+2]
            sheet.write(0, 12, 'Freight Terms')
            sheet.write(p, 12, Freight_Terms)
            break
        else:
            text_Invoice_Trial = text_Invoice_Trial[(Reference + 1):]



    #ShipmentOrder#
    Invoice_Count_Quantity = int(text.count('Shipment'))
    text_Invoice_Trial = text
    for y in range(0, Invoice_Count_Quantity):
        Reference = int(text_Invoice_Trial.index('Shipment'))
        if text_Invoice_Trial[Reference] + (text_Invoice_Trial[Reference + 1]) == 'Shipment#':
            Shipment_Order_Number = text_Invoice_Trial[Reference-1]
            sheet.write(0, 13, 'Shipment Order Number')
            sheet.write(p, 13, Shipment_Order_Number)
            break
        else:
            text_Invoice_Trial = text_Invoice_Trial[(Reference + 1):]


    Printed_Date = text[Printed_Reference + 1]
    Address = text[Printed_Reference+2:Phone_Reference]
    Address = (' '.join(Address))
    sheet.write(0, 14, 'Address')
    sheet.write(p, 14, Address)
    sheet.write(0, 15, 'Printed Date')
    sheet.write(p, 15, Printed_Date)

    Document_Count_In_text = int(text.count('Document'))
    for x in range(0, Document_Count_In_text):
        if text[End_Chaos] + text[End_Chaos + 1] == 'DocumentTotals':
            End_Chaos = text.index('Document')
            print('Document Totals Found!')
            break
        else:
            print('Document Totals Not Found :(')
            continue
    # 3rd Part of the TRILOGY
    Invoice_Count_Quantity = int(text.count('Document'))
    text_Invoice_Trial = text
    for y in range(0, Invoice_Count_Quantity):
        Reference = int(text_Invoice_Trial.index('Document'))
        if text_Invoice_Trial[Reference] + (text_Invoice_Trial[Reference + 1]) == 'DocumentTotals':
            Total_Weight = text_Invoice_Trial[Reference + 2]
            sheet.write(0, 16, 'Total Weight')
            sheet.write(p, 16, Total_Weight)
            break
        else:
            text_Invoice_Trial = text_Invoice_Trial[(Reference + 1):]

    Invoice_Count_Quantity = int(text.count('EST.'))
    text_Invoice_Trial = text
    for y in range(0, Invoice_Count_Quantity):
        Reference = int(text_Invoice_Trial.index('EST.'))
        if text_Invoice_Trial[Reference] + (text_Invoice_Trial[Reference + 1]) == 'EST.WT.':
            Total_Quantity = text_Invoice_Trial[Reference + 3]
            sheet.write(0, 17, 'Total_Quantity')
            sheet.write(p, 17, Total_Quantity)
            break
        else:
            text_Invoice_Trial = text_Invoice_Trial[(Reference + 1):]

        # Remit To
    Invoice_Count_Quantity = int(text.count('Remit'))
    text_Invoice_Trial = text
    for y in range(0, Invoice_Count_Quantity):
        Reference = int(text_Invoice_Trial.index('Remit'))
        if text_Invoice_Trial[Reference] + (text_Invoice_Trial[Reference + 1]) == 'Remitto:':
            text_Invoice_Trial = text_Invoice_Trial[Reference:]
            Sales_Tax = text_Invoice_Trial
            End_Reference = text_Invoice_Trial.index('USA')
            Remit_To1 = text_Invoice_Trial[2:End_Reference + 1]
            Remit_To1 = (' '.join(Remit_To1))

            break
        else:
            text_Invoice_Trial = text_Invoice_Trial[(Reference + 1):]

    Invoice_Count_Quantity = int(text.count('No'))
    text_Invoice_Trial = text
    for y in range(0, Invoice_Count_Quantity):
        Reference = int(text_Invoice_Trial.index('No'))
        if text_Invoice_Trial[Reference] + (text_Invoice_Trial[Reference + 1]) + (
        text_Invoice_Trial[Reference + 2]) + (text_Invoice_Trial[Reference + 3]) == 'NoDiscountonFreight':
            End_Reference = int(text_Invoice_Trial.index('540-948-6801'))
            Remit_To2 = text_Invoice_Trial[Reference + 5:End_Reference]
            Remit_To2 = (' '.join(Remit_To2))
            break
        else:
            text_Invoice_Trial = text_Invoice_Trial[(Reference + 1):]

    Remit_To = [Remit_To1, Remit_To2]
    Remit_To = (' '.join(Remit_To))
    Remit_To = Remit_To.replace(')', '')
    Remit_To = Remit_To.replace('\'', '')
    sheet.write(0, 18, 'Remit To:')
    sheet.write(p, 18, Remit_To)

        # Discount Date
    Invoice_Count_Quantity = int(text.count('If'))
    text_Invoice_Trial = text
    for y in range(0, Invoice_Count_Quantity):
        Reference = int(text_Invoice_Trial.index('If'))
        if text_Invoice_Trial[Reference] + (text_Invoice_Trial[Reference + 1]) + (text_Invoice_Trial[Reference + 2]) + (text_Invoice_Trial[Reference + 3]) + text_Invoice_Trial[Reference + 4] == 'Ifpaidonorbefore':
            Discount_Date = text_Invoice_Trial[Reference + 5]
            sheet.write(0, 19, 'Discount Date')
            sheet.write(p, 19, Discount_Date)
            break
        else:
            text_Invoice_Trial = text_Invoice_Trial[(Reference + 1):]
        # Discount
    Invoice_Count_Quantity = int(text.count('allowed'))
    text_Invoice_Trial = text
    for y in range(0, Invoice_Count_Quantity):
        Reference = int(text_Invoice_Trial.index('allowed'))
        if text_Invoice_Trial[Reference] + (text_Invoice_Trial[Reference + 1]) + (text_Invoice_Trial[Reference + 2]) + (text_Invoice_Trial[Reference + 3]) + text_Invoice_Trial[Reference + 4] + text_Invoice_Trial[Reference + 5] + text_Invoice_Trial[Reference + 6] == 'allowedonFreightNoDiscountonFreight':
            Discount = text_Invoice_Trial[Reference + 7]
            sheet.write(0, 20, 'Discount')
            sheet.write(p, 20, Discount)
            break
        else:
            text_Invoice_Trial = text_Invoice_Trial[(Reference + 1):]
        # SUBTOTAL AMOUNT
    Invoice_Count_Quantity = int(text.count('deduct'))
    text_Invoice_Trial = text
    for y in range(0, Invoice_Count_Quantity):
        Reference = int(text_Invoice_Trial.index('deduct'))
        if text_Invoice_Trial[Reference] + (text_Invoice_Trial[Reference + 1]) + (text_Invoice_Trial[Reference + 2]) == 'deductadiscount.':
            SubTotal_Amount = text_Invoice_Trial[Reference + 3]
            sheet.write(0, 21, 'SubTotal_Amount')
            sheet.write(p, 21, SubTotal_Amount)
            break
        else:
            text_Invoice_Trial = text_Invoice_Trial[(Reference + 1):]
        # SALES TAX
    SalesTax = Sales_Tax.index('USA')
    SalesTax = Sales_Tax[SalesTax + 1]
    sheet.write(0, 22, 'Sales Tax')
    sheet.write(p, 22, SalesTax)
        # FREIGHT(INLCUDED)
    Invoice_Count_Quantity = int(text.count('Phone:'))
    text_Invoice_Trial = text
    for y in range(0, Invoice_Count_Quantity):
        Reference = int(text_Invoice_Trial.index('Phone:'))
        if text_Invoice_Trial[Reference] + (text_Invoice_Trial[Reference + 1]) == 'Phone:Fax:':
            Freight = text_Invoice_Trial[Reference + 2]
            sheet.write(0, 23, 'Freight')
            sheet.write(p, 23, Freight)
            break
        else:
            text_Invoice_Trial = text_Invoice_Trial[(Reference + 1):]
        # INVOICE_TOTAL
    Invoice_Count_Quantity = int(text.count('INVOICE'))
    text_Invoice_Trial = text
    for y in range(0, Invoice_Count_Quantity):
        Reference = int(text_Invoice_Trial.index('INVOICE'))
        if text_Invoice_Trial[Reference] + (text_Invoice_Trial[Reference + 1]) == 'INVOICETOTAL:':
            INVOICE_TOTAL = text_Invoice_Trial[Reference + 2]
            sheet.write(0, 24, 'Invoice Total')
            sheet.write(p, 24, INVOICE_TOTAL)
            break
        else:
            text_Invoice_Trial = text_Invoice_Trial[(Reference + 1):]

    Chaotic_List = text[(Start_Chaos+1):End_Chaos]
    print(Chaotic_List)

    Count_Consistent = Chaotic_List.count('MBF')
    for i in range(0, Count_Consistent):
        p += 1
        Find_Pieces = Chaotic_List.index('PCS')
        Product_Description_End = (Find_Pieces + 1)
        Quantity_Reference = (Find_Pieces + 1)
        Price_Reference = (Find_Pieces + 2)
        Price = Chaotic_List[Price_Reference]
        Product_Description = (Chaotic_List[1:Find_Pieces])
        Price = Price.replace('/', '')
        Price = Price.replace(',', '')
        Quantity = float(Chaotic_List[Quantity_Reference])
        Amount = float(float(Quantity) * float(Price))
        Amount = round(Amount, 2)
        Product_Description = (' '.join(Product_Description))
        print(Price)
        print(Product_Description)
        print(Quantity)
        print(Amount)
        sheet.write(p, 25, Product_Description)
        sheet.write(p, 26, Quantity)
        sheet.write(p, 27, Price)
        sheet.write(p, 28, Amount)
        Raw_Amount = Chaotic_List[0]
        Raw_Amount = Raw_Amount.replace(',', '')
        Raw_Amount = float(Raw_Amount)
        if Amount == Raw_Amount:
            print('Great')
        else:
            print('Program is Indicating An Amount Error for iteration', i)
        MBF_index = Chaotic_List.index('MBF')
        Chaotic_List = Chaotic_List[(MBF_index + 1):]
    sheet.write(0, 25, 'Product Description')
    sheet.write(0, 26, 'Quantity of Product')
    sheet.write(0, 27, 'Price of Product')
    sheet.write(0, 28, 'Amount of Product')


book.close()
