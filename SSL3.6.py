import socket
import ssl
import xlrd
import xlwt
from urllib.parse import urlparse


x=0
y=1
z=2
n=1
m=7
ssl.OP_ALL
bookr=xlrd.open_workbook('lg.jp.xls')
bookw = xlwt.Workbook()
sheet1 = bookw.add_sheet('test')
sheet = bookr.sheet_by_index(0)
for row in range(1,2):
     sock = socket.socket()
     ssl_sock = ssl.wrap_socket(sock)
     o = urlparse(sheet.cell(row, 5).value)
     o=o.netloc
     print(o)
     ssl_sock.connect((o, 443))
     der = ssl_sock.getpeercert(binary_form=True)
     print (ssl.DER_cert_to_PEM_cert(der))
     filename=o+'.der'
     f = open(filename,'w')
     f.write(ssl.DER_cert_to_PEM_cert(der))
     f.close()
     list=ssl_sock.shared_ciphers()
     
     for i in list:
      lists=i
  #sheet1_row = sheet1.row(n)
      print('暗号方式の名前',end='  ')
     #sheet1_row.write(m,lists[x])
      #m=m+1
      print(lists[x])  
      print('プロトコルのバージョン',end='  ')
    #sheet1_row.write(m,lists[x])
      #m=m+1
      print(lists[y])
      print('秘密鍵のビット長',end='  ')
      #sheet1_row.write(m,lists[x])
      print(lists[z])
      print()
      x+3
      y+3
      z+3
    #  n=n+1
     # m=m+1

