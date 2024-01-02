
import xlsxwriter
import xml.etree.ElementTree as ET
from datetime import datetime

#tree = ET.parse('cda-tmp.xml')


namespaces = {'': 'urn:hl7-org:v3'}

tree = ET.parse('export_cda.xml')
#tree = ET.parse('country_data.xml')

root = tree.getroot()
print(root)
print(f'root tag: {root.tag}')
for child in root:
    print(f'{child.tag}  <==> {child.attrib}')


# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('cda.xlsx', {'strings_to_numbers': True})
worksheet = workbook.add_worksheet()
worksheet50 = workbook.add_worksheet('50')

components = root.findall('{urn:hl7-org:v3}entry/{urn:hl7-org:v3}organizer/{urn:hl7-org:v3}component')

# components = root.findall('entry/organizer/component', namespaces=namespaces)

worksheet.write(0, 0, 'effective time low')
worksheet.write(0, 1, 'Heart rate')
worksheet.write(0, 2, 'effective time high')

i = 1
for component in components:
    display_name = component.find('{urn:hl7-org:v3}observation/{urn:hl7-org:v3}code').attrib['displayName']

    if display_name != 'Heart rate':
        continue

    value = component.find('{urn:hl7-org:v3}observation/{urn:hl7-org:v3}value').attrib['value']

    if int(float(value)) > 50:
        continue

    effective_time_low = component.find('{urn:hl7-org:v3}observation/{urn:hl7-org:v3}effectiveTime/{urn:hl7-org:v3}low').attrib['value'].split('+')[0]
    effective_time_high = component.find('{urn:hl7-org:v3}observation/{urn:hl7-org:v3}effectiveTime/{urn:hl7-org:v3}high').attrib['value'].split('+')[0]

    date1 = datetime.strptime(effective_time_low, '%Y%m%d%H%M%S')
    date2 = datetime.strptime(effective_time_high, '%Y%m%d%H%M%S')
    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss'})
    worksheet.write(i, 0, date1, date_format)
    # worksheet.write(i, 1, display_name)
    worksheet.write(i, 1, value)
    worksheet.write(i, 2, date2, date_format)
    # print(f'{effective_time_low}, {effective_time_high}, {display_name}, {value}')
    i = i + 1
print('=========================')
print(len(components))

workbook.close()



