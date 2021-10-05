import os
import xml.etree.ElementTree as et
import pandas as pd
from numpy import nan
##################################################################
file=input('Enter name of xml file:\n')
directory = os.getcwd()
path = os.path.join(directory,file)
##################################################################
dom = et.parse(path)
root = dom.getroot()
tags = [child.tag for child in root]
try: index = tags.index('Incident')
except: index=None
##################################################################
#
def parse_levels(node,L,ID,data_dict):
    child_list = list(node)
    child_count = len(child_list)
    #process text for printing
    indent = '\t'*L
    tag = str(node.tag)
    attribute = str(node.attrib) if len(node.attrib)>0 else nan                     #Replaced '' with nan
    text = nan if node.text is None else node.text                                  #Replaced '' with nan
    ID_print = ''.join(['['+str(x)+']' for x in ID])
    #print(indent + tag +': '+ text + attribute,ID_print)
    data_dict.update({ID_print:[len(indent),tag,text,attribute]})                   #NEW
    #recursion for accessing all levels
    if child_count > 0:
        L+=1
        ID.append(0)
        for indx in range(child_count):
            node = child_list[indx]
            ID[-1]=indx
            parse_levels(node,L,ID,data_dict)
            if indx == child_count-1:
                ID.pop()
    return data_dict

##################################################################
#
def get_tags_from_location(Location):
	loc_list = [loc+']' for loc in Location.split(']')][:-1]
	dsc_list=[]
	for level in range(len(loc_list)):
		loc = ''.join(loc_list[:level+1])
		dsc_list.append(df[(df['Level']==level) & (df['Location']==loc)]['Tag'].values[0])
	print(Location,', '.join(dsc_list))
	return ', '.join(dsc_list)
    
##################################################################

def parse_xml(index):
    data_dict={}
    if index is None:
        r = root
    else: r = root[index]
    for i,child in enumerate(r):
        ID = [i]
        data_dict = parse_levels(child,0,ID,data_dict)
    df = pd.DataFrame.from_dict(data_dict,orient='index',columns=['Level','Tag','Text','Attribute']).reset_index().rename(columns={'index':'Location'})
    return data_dict, df

##################################################################

dd, df = parse_xml(index)

print(df.loc[:20].to_string())
print('***SUMMARY***\n___________')
print('Counts:\n',)
print(df.count())
print('\n***Level 0 items***\n___________')
##df.groupby('Level').apply(lambda g: g[g['Level']==0]['Tag']).reset_index(drop=True)
##df['Level'].value_counts().sort_index().plot.bar()
a= df[['Level','Tag']].apply(lambda x: x.name if x[0]==0 else nan,axis=1).ffill().value_counts().to_frame().rename(columns={0:'count_elmnts'})
HL_counts = a.merge(df['Tag'],how='left',left_index=True,right_index=True)[['Tag','count_elmnts']].reset_index().rename(columns={'index':'orig_index'}).sort_values(by='orig_index')
print(HL_counts)
#df['ID_dsc'] = df['Location'].apply(get_tags_from_location)
#print(df[df['Attribute'].notna()].to_string())
print('\n***Complete***')
input()

##################################################################
#Get fields which are populated
    #Level count summary
print('___________\n***Count summaries***')
print('Level count summary')
gb = df.groupby(['Level']).count()[['Text','Attribute']]
print(gb[gb.sum(axis=1)>0].to_string())
    #Tag count summary
print('Tag count summary')
gb = df.groupby(['Tag']).count()[['Text','Attribute']]
print(gb[gb.sum(axis=1)>0].sort_values(by='Text',ascending=False).to_string())
    #Level,Tag count summary
print('Level,Tag count summary')
gb = df.groupby(['Level','Tag']).count()[['Text','Attribute']] #remove 'Level' in groupby for a summary across all levels
print(gb[gb.sum(axis=1)>0].sort_values(['Level','Text'],ascending=[True,False]).to_string())

##################################################################
# df['Import']='EDIT MANUALLY'
# df['Alias']='EDIT MANUALLY'

file = file.split('.')[0]+' - flat XML.xlsx'
path = os.path.join(directory,file)
df.to_excel(path)
