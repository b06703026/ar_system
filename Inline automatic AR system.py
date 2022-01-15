#!/usr/bin/env python
# coding: utf-8

# In[33]:


import pandas as pd


# In[34]:


invoice_data = pd.read_excel('inline_data.xlsx', sheet_name = 'invoice_data')
bank_statement = pd.read_excel('inline_data.xlsx', sheet_name = 'bank_statement')
buyer_info = pd.read_excel('inline_data.xlsx', sheet_name = 'buyer_info')
invoice_data.head()


# In[35]:


bank_statement.head()


# In[36]:


buyer_info.head()


# In[37]:


unpaid = list()
total_invoice = list()
unpaid_rate = list()
for i in buyer_info.index:
    invoice = 0
    deposit = 0
    for j in invoice_data.index:
        if buyer_info['買方統一編號'][i] == invoice_data['買方統一編號'][j]:
            invoice = invoice_data['總開立金額'][j]
    for k in bank_statement.index:
        if buyer_info['買方帳號'][i] == bank_statement['買方帳號'][k]:
            deposit = bank_statement['總存入金額'][k]
    unpaid.append(invoice - deposit)
    total_invoice.append(invoice)
    unpaid_rate.append('{:.2%}'.format((invoice - deposit) / invoice)) 


# In[38]:


buyer_info['未付總金額'] = unpaid
buyer_info['總開立金額'] = total_invoice
buyer_info['欠款比例'] = unpaid_rate
buyer_info.head()


# In[39]:


from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib


# In[40]:


for i in buyer_info.index:
    if buyer_info['未付總金額'][i] > 0:
        mail = MIMEMultipart()  #建立MIMEMultipart物件
        mail["subject"] = "Inline payment notification"  #郵件標題
        mail["from"] = "chopper870907@gmail.com"  #寄件者
        mail["to"] = str(buyer_info['買方信箱'][i]) #收件者
        msg = "Dear " + str(buyer_info["買方名稱"][i]) + ","         + "\n\nYour accumulated payment is $" + str(buyer_info["未付總金額"][i]) + "."         + "\nPlease kindly remit the payment as soon as possible."         + "\nWe appreciate a lot for your cooperation. "         + "\n\nBest regards, \nInline"
        mail.attach(MIMEText(msg))  #郵件內容
        try:
            smtp = smtplib.SMTP(host="smtp.gmail.com", port="587")  # 設定SMTP伺服器
            smtp.ehlo()  # 驗證SMTP伺服器
            smtp.starttls()  # 建立加密傳輸
            smtp.login("chopper870907@gmail.com", "vsarcyhflctzldur")  # 登入寄件者gmail
            smtp.send_message(mail)  # 寄送郵件
            print("Complete!")
        except Exception as e:
            print("Error message: ", e)


# In[ ]:




