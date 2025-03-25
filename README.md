import pandas as pa
import datetime
import smtplib

gmail_id='sayalikate4039@gmail.com'
gmail_password='trfa yotr jgos cwjq'
def send_email(to,sub,msg):
    s=smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(gmail_id,gmail_password)
    s.sendmail(gmail_id,to,f"Subject: {sub}")
    s.quit()
    print(f"Email sent to {to} sent with {sub}")
    

if __name__=="__main__":
    #send_email(gmail_id,"subject","test message")
    #exit()
    df=pa.read_excel(r"C:\Users\Sayali Kate\Birthday_wisher\data.xlsx")
    #print(df)
    today=datetime.datetime.now().strftime("%d-%m")
    year=datetime.datetime.now().strftime("%Y")
    #print(today)
    for index,item in df.iterrows():
        #print(index,item["Birthday"])
        bday=item["Birthday"].strftime("%d-%m")
        #print(bday)
        writeInd=[]
        if (today==bday and str(year) not in str(item["Year"])):
            send_email(item["Email"],"Happy Birthday",item["Dialogue"])
            #print("Happy Birthday")
            writeInd.append(index)
    #print(writeInd)
    for i in writeInd:
        yr=df.loc[i,'Year']
        df['Year'] = df['Year'].astype(str)  # Ensure all values are strings
        df.loc[i,'Year']= str(yr) + ', '+ str(year)
        #print(df.loc[i,'Year'])
    #print(df)
    df.to_excel(r"C:\Users\Sayali Kate\Birthday_wisher\data.xlsx", index=False)
