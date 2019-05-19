'''
#Parse an HTML table and write to CSV
#Author: Kien Tran
#email: kien.tran@hella.com
#---------------------------------------
#Author         Date            Version     Description
#Kien Tran      May 10, 2019    1.0         Init tool
#Kien Tran      May 10, 2019    2.0         Update Gcape,bcm project, add TestVerdict and Change Status Fully or Auto
#---------------------------------------
#Python37
#BeautifulSoup
'''


from bs4 import BeautifulSoup
import csv


'''
This function is defined for modify the text from HTML, some data have define /n, /t in data. We have to modify the 
data again for write to csv file.
'''
def xulydata(text):
    text = text.replace('\n', '')
    return text

'''
This is function parse html file to csv
Input: html file, name of csv file
Output: csv file 
Description: The function will be parse html file, get the table in html. The data is second table (OverviewTable)
I define 4 colums for output template ( ID, Test Name, Test Result and Test Implemented). The table input have 5 column
we just use 4 columns for csv output.
If the result is pass/Passed => The test implemented is 'Implemented' else 'Not Implemented' 
'''
def ParseTable(html_input_file, csv_output_name, testverdict, project):
    ##Create output file
    csvoutput_file = open(csv_output_name + '.csv', 'w', newline='')
    csvwriter = csv.writer(csvoutput_file)

    csv_title = []
    csv_id = "Absolute Number"
    csv_status = testverdict
    csv_automation = "TS_Degree of Automation"
    csv_implemented = "TS_State of Automation"
    ##
    csv_title.append(csv_id)
    csv_title.append(csv_status)
    csv_title.append(csv_automation)
    csv_title.append(csv_implemented)
    ##
    csvwriter.writerow(csv_title)

    html = open(html_input_file).read()
    soup = BeautifulSoup(html, 'html.parser')
    table = soup.find_all('table', class_="OverviewTable")
    table_data = table[1]
    output_rows = []

    for table_row in table_data.findAll('tr'):
        columns = table_row.findAll('td')
        output_row = []
        id = ""
        result = ""
        count_f = 0
        for column in columns:
            column_text = xulydata(column.text)
            if (count_f == 1):
                id = column_text

            if (count_f == 3):
                result = column_text

            count_f += 1

        if (id != "") and (result != ""):
            output_row.append(id)
            if (result == "pass"):
                output_row.append("passed")
                output_row.append(project)
                output_row.append("Implemented")
            else:
                output_row.append("")
                output_row.append(project)
                output_row.append("In work")

            output_rows.append(output_row)

    csvwriter.writerows(output_rows)
    #Close the csv output file
    csvoutput_file.close()
