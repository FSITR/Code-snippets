import win32com.client as win32
import pandas as pd

#outlook = win32.gencache.EnsureDispatch("Outlook.Application").GetNamespace("MAPI")
#C:\Users\jbuck\AppData\Local\Temp
#https://stackoverflow.com/questions/33267002/why-am-i-suddenly-getting-a-no-attribute-clsidtopackagemap-error-with-win32com

outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI") # use this instead, or delete gen_.py in temp folder

folder_name = 'inbox'

folder_dict = {'inbox':6,'Deleted':3}
code = folder_dict[folder_name]
folder = outlook.GetDefaultFolder(code).Items

def sender_details(e):
    if e.Class == 43: #check email is Mail item
        if type(e.Sender) != None:
            if e.SenderEmailType == 'EX': #check where the email has come from (exchange server or SMTP)
                sndr = e.Sender.GetExchangeUser()
                if sndr != None: sender_email = sndr.PrimarySmtpAddress #the 'Address' it is a long unreadable string, so we get the user's smtp address instead
                else: sender_email = 'EX_sender_not_found'
##            elif e.SenderEmailType == 'SMTP': sender_email = e.Sender.Address
            elif e.SenderEmailType == 'SMTP': sender_email = e.SenderEmailAddress # new
            else: sender_email=None
            return e.SentOn.date(),e.SentOn.time(),sender_email
    else: return None,None,None

def recipient_details(e):
    if e.Class == 43:
        recips=[]
        recips_cc=[]
        for r in e.Recipients:
            if len(r.Address)>0:#a recipeint object with a blank address indicates (most likely) a Draft - these may be present in the 'Deleted items' folder
                if r.AddressEntry.Type == 'EX':
                    recip = r.AddressEntry.GetExchangeUser()
                    if recip != None: recip_email = recip.PrimarySmtpAddress
                    else: recip_email = 'EX_recip_not_found'
                elif r.AddressEntry.Type == 'SMTP': recip_email = r.AddressEntry.Address
                else: recip_email=None
                if r.Name in e.CC: recips_cc.append(recip_email)
                else: recips.append(recip_email)
        return [recips,recips_cc]
    else: return None

#######################################################################################################################
df=pd.DataFrame()
res_dict={}
for i,e in enumerate(folder):
    sent_dt, sent_time, sender_email = sender_details(e)
    recips_emails = recipient_details(e)
    #print([i,str(sent_dt),sender_email,recips_emails])

    if recips_emails != None:
        for l in recips_emails:
            if len(l)==1 and l[0]==None: #new
                print('***PASS***')                   #new
            elif True in ['support@wildeanalysis.co.uk' in r for r in l]: #changed if to elif because of the new lines above
                print(i,str(sent_dt),str(sent_time),sender_email,e.Subject)
                #df=df.append(pd.Series([sent_dt,sent_time,sender_email,e.Subject]),ignore_index=True)
                res_dict['SDate'] = sent_dt
                res_dict['STime'] = sent_time
                res_dict['Sender'] = sender_email
                res_dict['Subject'] = e.Subject
                res_dict['Body'] = e.Body
                df=df.append(res_dict,ignore_index=True)
df['Source'] = folder_name

save = input('Save to csv?\n')
if save:
    filename = folder_name+' support emails all time'+'.csv'
    df.to_csv(filename)


#Loop through whole inbox and filter based on email to Support AND does not have WA- in the Subject. Then get Body and Subject to identify keywords
#support = [e for e in folder if 'WA-' in e.Subject]






####################################################################################
##for i in range(10):
##    try :print(i,outlook.GetDefaultFolder(i).Name)
##    except: None

##for email in inbox:
##    print(email.Subject)

##print(support[0].Body)

##for e in support:
##    if e.Class == 43: #check email is Mail item
##        if e.SenderEmailType == 'EX': #check where the email has come from (exchange server or SMTP)
##            sender_email = e.Sender.GetExchangeUser().PrimarySmtpAddress #the 'Address' it is a long unreadable string, so we get the user's smtp address instead
##        else: sender_email = e.Sender.Address
##    print(e.SentOn,sender_email,'\t('+e.SenderEmailType+')')
##    #[print('\t'+rec.Address) for rec in e.Recipients]
