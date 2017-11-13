import socket
import ssl
import xlrd

x=0
y=1
z=2
book=xlrd.open_workbook('lg.jp.xls')
sheet = book.sheet_by_index(0)
for row in range(1,15):
     sock = socket.socket()
     ssl_sock = ssl.wrap_socket(sock)
     print(sheet.cell(row, 5).value)
     ssl_sock.connect((sheet.cell(row, 5).value, 443))
     der = ssl_sock.getpeercert(binary_form=True)
     print (ssl.DER_cert_to_PEM_cert(der))
     f = open('証明書.der','w')
     f.write(ssl.DER_cert_to_PEM_cert(der))
     f.close()
     list=ssl_sock.shared_ciphers()
     for i in list:
      lists=i    
      print('暗号方式の名前',end='  ')
      print(lists[x])  
      print('プロトコルのバージョン',end='  ')
      print(lists[y])
      print('秘密鍵のビット長',end='  ')
      print(lists[z])
      print()
      x+3
      y+3
      z+3
      

  

 
   
    
         

     
     

