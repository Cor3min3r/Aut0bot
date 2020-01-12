#Author Cor3min3r
import pdfkit
import smtplib
import os
import sys
import getpass
import pyfiglet
from termcolor import colored
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

my_name = colored("Cor3min3r","red",attrs=['concealed','blink','bold'])

print (colored('[+] Welcome to AutoBot Certicate Generator & emailer by '+my_name,"blue"))

fig_name = pyfiglet.figlet_format("Aut0b()t")
#fig_name = pyfiglet.figlet_format("Cor3min3r")


print("                       ")
print(colored(fig_name,"red"))

print "[*] Initializing the connection...."

server = smtplib.SMTP('smtp.gmail.com',587)

print(colored("[+] Connection Established","green"))

Email_address = raw_input("[+] Enter your Email_address: ")

password = getpass.getpass(prompt='[+] Enter your Password: ', stream=None)  
	

print(colored("[*] starting TLS...","red"))

server.starttls()

print(colored("[+] TLS Started ","green"))


try:
	server.login(Email_address,password)
except Exception:
        print(colored("[-] There was some Error while trying to connect make sure you enterd the correct credentials","red"))
	print(colored("[!!] Make sure you enabled less secure apps using below link","red"))
	print(colored("https://myaccount.google.com/lesssecureapps","blue"))
	sys.exit()
else:
        server = smtplib.SMTP('smtp.gmail.com',587)
	server.starttls()
	print(colored("[+] Connected Successfully...","green"))
	


cert_temp = raw_input("[+] Enter the Template(.html) with path: ").split('.')	
cert_temp = cert_temp[0]+".html"	
csv_filename = raw_input("[+] Enter the CSV filename with path: ").split('.')
csv_filename = csv_filename[0]+".csv"

color_email = colored("Email address","red")
color_name = colored("Name","red")
field = colored("1","red")

column_1 = raw_input("[+] Enter the "+color_email+" column no starting from "+field+": ")	
column_2 = raw_input("[+] Enter the "+color_name+" column no starting from "+field+": ")

os_command_1 = '''awk -F "\ "*,\ "*" '{print $'''+str(column_1)+'''}' '''+ csv_filename+" > email.txt"

os_command_2 = '''awk -F "\ "*,\ "*" '{print $'''+str(column_2)+'''}' '''+ csv_filename+" > name.txt" 

os.system(os_command_1)
os.system(os_command_2)
os.system("paste email.txt name.txt > combined.txt")
os.system("rm email.txt name.txt")
header = "sed 1d combined.txt > Combined_2.txt"
os.system(header)


Sub_to_enter = raw_input("[+] Enter the Subject for this email: ")
content_to_deliver = raw_input("[+] Enter the content to deliver: ")
message = MIMEMultipart()
message['From'] = Email_address
#message['To'] =  
message['Subject'] = Sub_to_enter
mail_content = content_to_deliver
message.attach(MIMEText(mail_content, 'plain'))



fil = open("Combined_2.txt","r")


for i in fil.readlines():
        
        server.login(Email_address,password)

	a = i.split("\t")
        a = map(lambda a: a.strip(), a)
        #map(str.strip, a)
        fil2 = open(cert_temp,"rwt")
        read = fil2.read()
        replace = read.replace("hii",a[1])
        fil2.close()
        fil2 = open(cert_temp,"wt")
	fil2.write(replace)
        fil2.close()
	pdfkit.from_file(cert_temp, a[1]+".pdf")
#	pdfkit.from_string(a[1],a[1]+".pdf")

	attach_file = open(a[1]+".pdf",'rb')
	payload = MIMEBase('application', 'octate-stream')
	payload.set_payload((attach_file).read())	
	encoders.encode_base64(payload)
	
	see = colored(a[0],"green")
        see2 = colored(a[1],"green")
	see3 = colored(".pdf","green")	

	payload.add_header('Content-Decomposition', 'attachment', filename=a[1]+".pdf")
	print(colored("[*] File Attached ===> "+see2+see3,"blue"))
	message.attach(payload)	

	text = message.as_string()
	print(colored("[+] Sending Email to ==> "+see,"blue"))
	print(colored("[+] Receiver Name ==> "+see2,"blue"))
	server.sendmail(Email_address,a[0],text)
	print(colored("[!] Sending Email","red"))	
	print(colored("[^] Email sent ","green"))
	os.system("rm "+a[1]+".pdf")	

        fil2 = open(cert_temp,"rwt")
        read = fil2.read()
        #print "Changing from "+a[1]+" to hii"
        replace = read.replace(a[1],"hii")
        fil2.close()        
	fil2 = open(cert_temp,"wt")
        fil2.write(replace)
        fil2.close()
	print(colored("[+] Quiting and starting Server again","red"))
  	server.quit()
	server = smtplib.SMTP('smtp.gmail.com',587)
	server.starttls()
	message = MIMEMultipart()
	message['From'] = Email_address
#message['To'] =
	message['Subject'] = Sub_to_enter
	mail_content = content_to_deliver
	message.attach(MIMEText(mail_content, 'plain'))
		

os.system("rm combined.txt Combined_2.txt") 
print(colored("[+] All Email have been sent...Quiting bii","yellow"))
server.quit()
