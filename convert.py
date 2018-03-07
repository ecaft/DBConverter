import xlrd
from collections import OrderedDict
import simplejson as json
 
# Open the workbook and select the first worksheet
wb = xlrd.open_workbook('excel-xlrd-sample.xls')
sh = wb.sheet_by_index(0)
 
# List to hold dictionaries
companies_list = []
 
# Iterate through each row in worksheet and fetch values into dict
for rownum in range(1, sh.nrows):
    names = OrderedDict()
    row_values = sh.row_values(rownum)
    company_name = u"{}".format(row_values[0]).replace(".","")
    companies = OrderedDict()
    companies["id"] = company_name.lower()
    companies['description'] =  u"{}".format(row_values[2])
    companies['jobtitles'] =  u"{}".format(row_values[16])
    companies['jobtypes'] =  u"{}".format(row_values[20])
    companies['location'] = ""
    companies['majors'] =  u"{}".format(row_values[18])
    companies['name'] =  u"{}".format(row_values[0])
    companies['optcpt'] =  u"{}".format(row_values[23])
    companies['sponsor'] =  u"{}".format(row_values[24])
    companies['website'] =  u"{}".format(row_values[3])
    names[company_name] = companies
    companies_list.append(names)

for i in names:

	names[i] = companies

poop = {}
poop["companies"] = companies_list
# Serialize the list of dicts to JSON
j = json.dumps(poop)
 
# Write to file
with open('data20.json', 'w') as f:
    f.write(j)