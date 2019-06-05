import mailbox
import csv
import datetime
import re
import os
filesList = os.listdir()

number=0
lines=[]
attachHeaderLine=[]
repeatAttachment={}

def extractattachements(attachFile,message,messageId,row,header):
    fileNameList=[]
    rowLines=[]
    attachfilerow=[]
    if message.get_content_maintype() == 'multipart':
        for part in message.walk():
            attachfilerow=[]
            if part.get_content_maintype() == 'multipart': continue
            if part.get('Content-Disposition') is None: continue
            attachfilerow.append(messageId)
            attachfilerow.append('')
            filename = part.get_filename()
            try:
                f=open(filename,'r')
                if filename in repeatAttachment:
                    repeatAttachment[filename]=repeatAttachment[filename]+1
                    filename=filename.rsplit('.', 1)[0]+'_'+str(repeatAttachment[filename])+'.'+filename.rsplit('.', 1)[1]
                    
                else: 
                    repeatAttachment[filename]=1
                    filename=filename.rsplit('.', 1)[0]+'_1.'+filename.rsplit('.', 1)[1]
                    
            except FileNotFoundError:
                print('File not fond')
                
                
            fileNameList.append(filename)
            
            attachfilerow.append(filename)
            attachfilerow.append(part.get_content_maintype()+'/'+part.get_content_subtype())
            attachfilerow.append(filename)
            rowLines.append(attachfilerow)
            print(filename)
            fb = open(filename,'wb')
            fb.write(part.get_payload(decode=True))
            fb.close()
        if(len(fileNameList)>0):
            concatName = ','.join(fileNameList)
            row.append(concatName)
        else:
            row.append('')
        attachWriter = csv.writer(attachFile)
        attachWriter.writerows(rowLines)
        
    

with open('emailMessage.csv' ,'w') as writeFile:
    for file in filesList:
        no=re.findall(r'\_\d{0,2}\.',file)
        file=re.sub('\_\d{1,2}\.','.',file)
        endNo = len(no)
        if(endNo>0):
            if file in repeatAttachment:
                no[endNo-1]=re.sub('\_','',no[endNo-1])
                no[endNo-1]=re.sub('\.','',no[endNo-1])
                if(int(no[endNo-1])>int(repeatAttachment[file])):
                    repeatAttachment[file]=int(no[endNo-1])
            else:
                no[endNo-1]=re.sub('\_','',no[endNo-1])
                no[endNo-1]=re.sub('\.','',no[endNo-1])
                repeatAttachment[file]=int(no[endNo-1])
        else:
            if repeatAttachment.get(file) == None:
                repeatAttachment[file]=1
    attachFile= open("attachment.csv","w")
    attachWriter = csv.writer(attachFile)
    attachHeader=['MessageId','parentId','Name','ContentType','Body']
    attachHeaderLine.append(attachHeader)
    attachWriter.writerows(attachHeaderLine)
    for message in mailbox.mbox('Sample.mbox'):
        row=[]
        header=[]
        header.append('Attachemnt')
        header.append('Body')
        if message.is_multipart():
            extractattachements(attachFile,message,message['Message-ID'],row,header)
            messagePart=[]
            for part in message.walk():
                print(part.get_content_maintype())
                if part.get_content_maintype() == 'text':
                    if part.get_content_subtype() =='plain':
                        messagePart.append(str(part))
            row.append(','.join(messagePart))
        else:
            content = message.get_payload(decode=True)
            print('non multipart')
            row.append("")
            row.append(part)
        number=number+1
        print("---------------------");
        
        if 'Subject' in message:
            header.append('Subject')
            row.append(message['Subject'])
        else:
            header.append('Subject')
            row.append('')

        if 'From' in message:
            header.append('From')
            lst = re.findall('[A-Za-z0-9\.\-+_]+@[A-Za-z0-9]+\.[A-Za-z0-9]+',message['From'])
            print(lst)
            row.append(lst[0])
        else:
            header.append('From')
            row.append('')
        if 'To' in message:
            header.append('To')
            lst = re.findall('[A-Za-z0-9\.\-+_]+@[A-Za-z0-9]+\.[A-Za-z0-9]+',message['To'])
            row.append(lst[0])
        else:
            header.append('To')
            row.append('')
        if 'Date' in message:
            header.append('Date')
            datedring = message['Date']
            datedring =re.sub('\(\S+\)','',datedring)
            datedring=datedring.strip()
            date_time_obj = datetime.datetime.strptime(datedring, '%a, %d %b %Y %H:%M:%S %z')
            print(date_time_obj)
            row.append(str(date_time_obj.strftime('%Y-%m-%d %H:%M:%S')))
        else:
            header.append('Date')
            row.append('')
        if 'Message-ID' in message:
            header.append('Message-ID')
            row.append(message['Message-ID'])
        else:
            header.append('Message-ID')
            row.append('')
        
            
        if number==1:
            lines.append(header)
        lines.append(row)
    writer = csv.writer(writeFile)
    writer.writerows(lines)


writeFile.close()
attachFile.close()

