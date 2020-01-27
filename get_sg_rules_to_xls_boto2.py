#--------------- IMPORTING MODULES and FUNCTION -------------------#

import re
import datetime
import boto.ec2
import xlwt
import xlrd
from xlutils.copy import copy


#----------------------- CREATING CONNECTION -----------------------#

# Connecting to US-East-1 Region -  N.Virginia
ec2_client = boto.ec2.connect_to_region('us-east-1')


#----------------------- GETTING SG to FIND ------------------------#

# Getting the list of all the Security Groups in the region
list_of_sg = ec2_client.get_all_security_groups()

# Getting the Security Group to search
sg_to_find = raw_input("Please enter the security group name you want to search : ")
print sg_to_find


#----------- FUNCTION to WRITE FOUND SG INFO to XLS File ------------#

def writeToExcel(found_sg):

    # Creating Excel Sheet styles
    h1_style = xlwt.easyxf('font: name Times New Roman, bold on, height 380;'
                 'borders: left thick, right thick, top thick, bottom thick;'
                 'pattern: pattern solid, fore_colour light_orange;'
                'align: vertical center, horizontal center;')
    h2_style = xlwt.easyxf('font: name Times New Roman, bold on, height 340;'
                 'borders: left thin, right thin, top thin, bottom thin;'
                 'pattern: pattern solid, fore_colour light_turquoise;'
                'align: vertical center, horizontal center;')
    h3_style = xlwt.easyxf('font: name Times New Roman, bold on, height 300;'
                 'borders: left thin, right thin, top thin, bottom thin;'
                 'pattern: pattern solid, fore_colour light_yellow;'
                'align: vertical center, horizontal center;')
    ingress_data_style = xlwt.easyxf('font: name Times New Roman, height 280;'
                            'borders: left thin, right thin, top thin, bottom thin;'
                            'pattern: pattern solid, fore_colour light_green;'
                            'align: wrap on;')
    egress_data_style = xlwt.easyxf('font: name Times New Roman, height 280;'
                            'borders: left thin, right thin, top thin, bottom thin;'
                            'pattern: pattern solid, fore_colour gray25;'
                            'align: wrap on;')

    # Opening an existing Excel workbook
    wb = copy(xlrd.open_workbook("SecurityGroupTemplate.xls", formatting_info=True))

    # Adding a worksheet to the workbook
    ws = wb.add_sheet(found_sg.name)

    # Writing Main Heading
    ws.write_merge(0, 0, 0, 9, found_sg.name, h1_style)

    # Writing 2nd Heading
    ws.write_merge(1, 1, 0, 4, "INBOUND RULES", h2_style)
    ws.write_merge(1, 1, 5, 9, "OUTBOUND RULES", h2_style)

    # Writing 3rd Heading
    ws.write(2, 0, "Rule #", h3_style)
    ws.write(2, 1, "IP Protocol", h3_style)
    ws.write(2, 2, "Port", h3_style)
    ws.write(2, 3, "Source/Target", h3_style)
    ws.write(2, 4, "Description", h3_style)
    ws.write(2, 5, "Rule #", h3_style)
    ws.write(2, 6, "IP Protocol", h3_style)
    ws.write(2, 7, "Port", h3_style)
    ws.write(2, 8, "Source/Target", h3_style)
    ws.write(2, 9, "Description", h3_style)

    # Writing Data - Ingress Rules
    for index, rule in enumerate(found_sg.rules):
        row = 3+index
        ws.write(row, 0, index+1, ingress_data_style)
        ws.write(row, 1, rule.ip_protocol, ingress_data_style)
        port = "From Port: " + str(rule.from_port or "") + "\nTo Port: " + str(rule.to_port or "")
        ws.write(row, 2, port, ingress_data_style)
        ws.write(row, 3, str(rule.grants)[1:-1], ingress_data_style)
        desc_formula = 'VLOOKUP(D'+str(row+1)+',master!$A$2:$C$100'+',3,FALSE)'
        ws.write(row, 4, xlwt.Formula(desc_formula), ingress_data_style)

    # Writing Data - Egress Rules
    for index, rule in enumerate(found_sg.rules_egress):
        row = 3+index
        ws.write(row, 5, index+1, egress_data_style)
        ws.write(row, 6, rule.ip_protocol, egress_data_style)
        port = "From Port: " + str(rule.from_port or "") + "\nTo Port: " + str(rule.to_port or "")
        ws.write(row, 7, port, egress_data_style)
        ws.write(row, 8, str(rule.grants)[1:-1], egress_data_style)
        desc_formula = 'VLOOKUP(I'+str(row+1)+',master!$A$2:$C$100'+',3,FALSE)'
        ws.write(row, 9, xlwt.Formula(desc_formula), egress_data_style)
    
    # Saving workbook
    xls_output_name = "SecurityGroup-" + datetime.datetime.now().strftime("%Y%m%d-%H%M") + ".xls"
    wb.save(xls_output_name)


#------- SEARCHING SG and INVOKE FUNCTION to WRITE FOUND SG INFO to XLS File --------------#
 
flag = False
for sg in list_of_sg:
    # Searching for the input pattern in the list of security groups
    if re.search(sg_to_find, sg.name):
        flag = True
        writeToExcel(sg)
else:
    if not flag:
        print "Sorry..!! Could not find the security group - ", sg_to_find
        print "The list of security groups in US-East-1 region are :"
        for sg in list_of_sg:
            print sg
