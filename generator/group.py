from models.group import Group
from System.Runtime.InteropServices import COMException
import random
import string
import getopt
import sys
import os.path
import clr
import time
__author__ = 'pzqa'

clr.AddReferenceByName(' Microsoft.Office.Interop.Excel, '
                       'Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel

try:
    opts, args = getopt.getopt(sys.argv[1:], "n:f:", ["number of groups", "file"])
except getopt.GetoptError as err:
    getopt.usage()
    sys.exit(2)

n = 5
f = "data/group.xlsx"

for o, a in opts:
    if o == "-n":
        n = int(a)
    elif o == "-f":
        f = a


def random_string(prefix, maxlen):
    symbols = string.ascii_letters + string.digits + string.punctuation + " "*10
    return prefix + "".join([random.choice(symbols) for i in range(random.randrange(maxlen))])


test_data = [Group(name="")] + [Group(name=random_string("name", 10)) for i in range(n)]

file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", f)

excel = Excel.ApplicationClass()
excel.Visible = True

workbook = excel.Workbooks.Add()
sheet = workbook.ActiveSheet

if os.path.exists(file):
    os.remove(file)

for i in range(len(test_data)):
    sheet.Range['A%s' % (i+1)].Value2 = test_data[i].name


try:
    workbook.SaveAs(file)
except COMException:
    os.makedirs(file)


time.sleep(10)
excel.Quit()
