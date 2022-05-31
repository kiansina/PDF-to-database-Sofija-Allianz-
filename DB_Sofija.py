import fitz
import pandas as pd
doc = fitz.open(*******)
page1 = doc[0]
words = page1.get_text("words")

mywords=[]
for i in range(0,len(words)):
     mywords.append((words[i][4]))

df=pd.DataFrame(mywords)
df.columns=['mywords']
dell=df[df['mywords']=='-'].index
df.drop(dell, inplace = True)

dP=pd.DataFrame(columns=['Certificate Number','Insured Entity','Vessel Name','Date of Sailing','From','To','Good description','Insured Value','Currency'])



## C0
dP['Certificate Number']=[df.iloc[-2][0]]   #93,5,0
## C1
start=df[df['mywords']=='Contraente'].index[0]
end=df[df['mywords']=='Applicant'].index[0]
A=df.loc[start+1][0]
for i in range(start+2,end-1):
    A+=' '+df.loc[i][0]

dP['Insured Entity']=[A] #5,2,0 + #5,2,1
## C2
start=df[df['mywords']=='CARRIED'].index[0]
end=df[df['mywords']=='Effective'].index[0]
A=df.loc[start+2][0]
for i in range(start+3,end-3):
    A+=' '+df.loc[i][0]

dP['Vessel Name']=[A]
## C3
start=df[df['mywords']=='Effective'].index[0]
end=df[df['mywords']=='MERCI'].index[0]
A=df.loc[start+2][0]
for i in range(start+3,end):
    A+=' '+df.loc[i][0]

dP['Date of Sailing']=[A]
## C4
start=df[df['mywords']=='Da/From'].index[0]
end=df[df['mywords']=='A/To'].index[0]
A=df.loc[start+1][0]
for i in range(start+1,end-1):
    A+=' '+df.loc[i][0]

dP['From']=[A]
## C5
start=df[df['mywords']=='A/To'].index[0]
end=df[df['mywords']=='Via'].index[0]
A=''
for i in range(start+1,end):
    A+=df.loc[i][0]+' '

dP['To']=[A.strip()]
## C6
start=df[df['mywords']=='GOODS'].index[0]
end=df[df['mywords']=='MARCHE'].index[0]
A=df.loc[start+1][0]
for i in range(start+2,end):
    A+=' '+df.loc[i][0]

dP['Good description']=[A]
# C7
start=df[df['mywords']=='INSURED'].index[0]
A=df.loc[start+1][0]
dP['Insured Value']=[A]
# C8
start=df[df['mywords']=='in'].index[0]
A=df.loc[start+1][0]
dP['Currency']=[A]
dP.to_excel('DBSOFIJA.xlsx')
