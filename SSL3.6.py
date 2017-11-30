import socket
import ssl
import xlrd
import xlwt
from urllib.parse import urlparse

x=0
y=1
z=2
n=0
m=7
bookr=xlrd.open_workbook('go.jp1.xlsx')
#bookw = xlwt.Workbook()
#newsheet_1 = bookw.add_sheet('NewSheet_1')
sheet = bookr.sheet_by_index(0)
context = ssl.SSLContext()
context.set_ciphers("AES:3DES:RSA:ECDHE:DH:ECDSA:CAMELLIA:DSS:SEED:SHA:RC4:MD5:PSK:GCM:CFB:OFB:CTR:ECB:IDEA")
for row in range(1,50):
     sock = socket.socket()
     ssl_sock = context.wrap_socket(sock)
     o = urlparse(sheet.cell(row, 0).value)
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
    # newsheet_1.write(n,0,o)
     n=n+1
     
     for i in list:
      lists=i
     # newsheet1_row = sheet1.row(n)
      m=m+1
      print(lists[x],end=' ')  
     # newsheet1_row.write(m,lists[x])
      m=m+1
      print(lists[y],end=' ')
     # newsheet1_row.write(m,lists[x])
      print(lists[z])
    #  book.save('gopro.xls')
      x+3
      y+3
      z+3
      n=n+1
      m=m+1

