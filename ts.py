import os
import pandas as pd
import pdfkit
import inflect



#load data frame in the empty df
excel_file = r"D:/Brave_download/code/code/maintanence/INVOICE_GENERATER/sample.xlsx"
df = pd.read_excel(excel_file)


# access html
with open(r"D:/Brave_download/code/code/maintanence/INVOICE_GENERATER/main.html", 'r') as file:
    template = file.read()

# output path
output_folder = r'D:/Brave_download/code/code/maintanence/INVOICE_GENERATER/pdf'

current_invoice_number = 1001
# to itterrate each row
for index, row in df.iterrows():
    # Replace placeholders in the template with data from the Excel file
    html_content = template.replace('[Client Name / Flat Number]', str(row['FLAT NO']))
    html_content = html_content.replace('[Bill Date]', str(row['Bill Date ']))
    html_content = html_content.replace('[Month]', str(row['month ']))
    html_content = html_content.replace('[no.of.month]', str(row['no.of.month']))
    html_content = html_content.replace('[Payment Done By]', str(row['payment done by']))
    html_content = html_content.replace('[Cash / Cheque / UPI or Net Banking]', str(row['Payment Method ']))
    html_content = html_content.replace('[Transaction Details]', str(row['transiction id ']))
    
    html_content = html_content.replace('[WATER CHARGE]', str(row['WATER CHARG']))
    html_content = html_content.replace('[SINKING FUND]', str(row['SINKING FUND']))
    html_content = html_content.replace('[REPAIR AND MAINT]', str(row['REPAIR AND MAINT']))
    html_content = html_content.replace('[MAJ.Rep maint]', str(row['MAJ.PERAIR AND MAINT']))
    html_content = html_content.replace('[SERVICE CHARGE]', str(row['SERVICE CHARGE']))
    html_content = html_content.replace('[Non Occupancy charges]', str(row['Non occupancy charges']))
    html_content = html_content.replace('[Festival Charges]', str(row['Festival Charges']))
    html_content = html_content.replace('[Bike Parking]', str(row['Bike Parking']))
    html_content = html_content.replace('[3/4 Wheeler Parking]', str(row['3/4 Wheeler Parking']))

    invoice_number = current_invoice_number
    current_invoice_number += 1
    html_content = html_content.replace('[Invoice Number]', str(invoice_number))

    #  to calc the sum of all 
    total_amount = (
        
        row['WATER CHARG'] +
        row['SINKING FUND'] +
        row['REPAIR AND MAINT'] +
        row['MAJ.PERAIR AND MAINT'] +
        row['SERVICE CHARGE']+
        row['Non occupancy charges'] +
        row['Festival Charges'] +
        row['Bike Parking'] +
        row['3/4 Wheeler Parking']
    )

    html_content = html_content.replace('[Total Amount]', '₹ '+ str(total_amount))
    # using inflect library
    inflector = inflect.engine()

    total_amount_in_words = inflector.number_to_words(total_amount)
    html_content = html_content.replace('[Total Amount in Words]', total_amount_in_words.title()+" Rupees Only")
    

    # to genrate pdf accoring to flat no.
    pdf_output_path = os.path.join(output_folder, f'{row["FLAT NO"]}_invoice.pdf')

    # Convert HTML to PDF using pdfkit
    pdfkit.from_string(html_content, pdf_output_path)

print("PDFs generated successfully.")
