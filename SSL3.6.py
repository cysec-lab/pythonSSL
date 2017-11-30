import socket
import ssl
import xlrd
import xlwt
from urllib.parse import urlparse

x=0
y=1
z=2
n=0
m=0
bookr=xlrd.open_workbook('go.jp1.xlsx')
bookw = xlwt.Workbook()
newsheet_1 = bookw.add_sheet('NewSheet_1')
sheet = bookr.sheet_by_index(0)
context = ssl.SSLContext()
context.set_ciphers("ALL:NULL")
for row in range(1,6):
     sock = socket.socket()
     ssl_sock = context.wrap_socket(sock)
     o = urlparse(sheet.cell(row, 0).value)
     o=o.netloc
     print()
     print(o)
     ssl_sock.connect((o, 443))
     der = ssl_sock.getpeercert(binary_form=True)
    # print (ssl.DER_cert_to_PEM_cert(der))
     filename=o+'.der'
     f = open(filename,'w')
     f.write(ssl.DER_cert_to_PEM_cert(der))
     f.close()
     list=ssl_sock.shared_ciphers()
     newsheet_1.write(n,0,o)
     n=n+1
     for i in list:
      lists=i
      newsheet1_row = newsheet_1.row(n)
      print(lists[x],end=' ')  
      newsheet1_row.write(1,lists[x])
      print(lists[y],end=' ')
      newsheet1_row.write(2,lists[y])
      print(lists[z])
      newsheet1_row.write(3,lists[z])
      bookw.save('gopro.xls')
      x+3
      y+3
      z+3
      n=n+1
      

