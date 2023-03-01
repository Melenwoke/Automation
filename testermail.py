import win32com.client as win32
import datetime

outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)
email.To ='elenwokemike@gmail.com'
email.Subject ='this is a test'
email.Body ='hi michael, i work'
email.Send()
print('email sent')

