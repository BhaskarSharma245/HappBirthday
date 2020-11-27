import pandas as pd
import datetime


def sendEmail(to,sub,msg):
    print(f'Mail to {to} sent with subject {sub} and message {msg}')

if __name__ == '__main__':
    df = pd.read_excel('data1.xlsx')
    #print(df)
    today= datetime.datetime.now().strftime('%d-%m')
    #print(today)
    yearNow=datetime.datetime.now().strftime('%Y')
    #print(yearNow)
    writeInd=[]

    for index,items in df.iterrows():
        #print(index, items['Name'])
        bday= items['Birthday'].strftime('%d-%m')
        #print(bday)
        if(today==bday) and yearNow not in str(items['Year']):
            sendEmail(items['Mail'],'happy birthday',items['Dialogue'])
            writeInd.append(index)

    print(writeInd)

    for i in writeInd:
        yr=df.loc[i,'Year']
        print(yr)
        df.loc[i,'Year'] = str(yr)+','+str(yearNow)
        print(df.loc[i,'Year'])
        print(df)

    df.to_excel("D:\data1.xlsx")






