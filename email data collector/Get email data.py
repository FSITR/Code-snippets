import win32com.client as win32
import pandas as pd
'''
Run this script whilst you have Outlook open and it will begin extracting your emails from your inbox to a pandas dataframe. This can be written to a csv for further analysis.
'''
#outlook = win32.gencache.EnsureDispatch("Outlook.Application").GetNamespace("MAPI")
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
            elif e.SenderEmailType == 'SMTP': sender_email = e.SenderEmailAddress
            else: sender_email=None
            return e.SentOn.date(),e.SentOn.time(),sender_email
    else: return None,None,None

def recipient_details(e):
    if e.Class == 43:
        recips=[]
        recips_cc=[]
        for r in e.Recipients:
            if len(r.Address)>0:# a recipeint object with a blank address indicates (most likely) a Draft - these may be present in the 'Deleted items' folder
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

    if recips_emails != None:
        for l in recips_emails:
            if len(l)==1 and l[0]==None:
                print('***PASS***')
            elif True in ['support@wildeanalysis.co.uk' in r for r in l]: #changed if to elif because of the new lines above
                print(i,str(sent_dt),str(sent_time),sender_email,e.Subject)
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
