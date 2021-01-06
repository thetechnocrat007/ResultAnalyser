import re
import openpyxl
import requests
import pdfplumber
import pandas as pd
from collections import namedtuple





new_candidate_re=re.compile(r'\d.* [A-Z]{2}\d{6}.*')
score_line_re=re.compile(r'(\s.|\-|.\d.)* \d.* \d.* \d.* \d.* \d.* \d+$')
date_re=re.compile(r'^[0-9|/]+\b')
line_items=[]
entr = namedtuple('entr', 'rank roll_no name gender category dob m_des e_des me_obj g1 g2 g3 g4 wrt_tot interview total')


with pdfplumber.open(r'C:\Users\Pratik\Downloads\fr.pdf') as pdf:

    #page = pdf.pages[68]
    #text = page.extract_text()

    for page in pdf.pages[68:]:
        text = page.extract_text()

        for line in text.split('\n'):
            if new_candidate_re.match(line):
                op = re.split(r'[ ](?=[M|F]+\b)', line)


                splited_line1=op[0]
                splited_line2=op[1].split()

                rank, roll_no, garb1, garb2, *name = splited_line1.split()
                name = ' '.join(name)


                gender=splited_line2[0]
                category=splited_line2[1]


                #if date_re.match(splited_line2[3]):
                 #   dob=splited_line2[3]
                #else:dob=splited_line2[4]

                #dob=re.match('^[0-9|/]+\b',op[1])
                dob = re.search(r'[0-9]{2}[\/][0-9]{2}[\/][0-9]{4}\b', op[1]).group()

            if score_line_re.match(line):
                #print(line)
                splited_score=line.split()
                total=splited_score[-1]
                interview=splited_score[-2]
                wrt_tot=splited_score[-3]
                g4=splited_score[-4]
                g3=splited_score[-5]
                g2=splited_score[-6]
                g1=splited_score[-7]
                me_obj=splited_score[-8]
                e_des=splited_score[-9]
                m_des=splited_score[-10]

                line_items.append(entr(rank, roll_no, name, gender, category, dob, m_des, e_des, me_obj, g1, g2, g3, g4, wrt_tot, interview, total))


print('***********************************************************************************************')
try:
    df=pd.DataFrame(line_items)

    df.head()
    print(df)

    df[["rank", "m_des", "e_des", "me_obj", "g1", "g2", "g3", "g4", "wrt_tot", "interview", "total"]]=df[["rank", "m_des", "e_des", "me_obj", "g1", "g2", "g3", "g4", "wrt_tot", "interview", "total"]].apply(pd.to_numeric)

    #df["dob"]=pd.to_datetime(df["dob"])

    df["dob"]=pd.to_datetime(df["dob"],format='%d/%m/%Y')



    # addind post
    with pdfplumber.open(r'C:\Users\Pratik\Downloads\fr.pdf') as pdf:
        line_post = []
        entr_post = namedtuple('entr_post', 'r_no post')

        for page in pdf.pages[:67]:
            text = page.extract_text()

            for line in text.split('\n'):
                if new_candidate_re.match(line):
                    p = re.split(r'[0-9]+[ ][0-9]+[ ][0-9]+', line)
                    p_splited1 = p[0].split()

                    line_post.append(entr_post(p_splited1[1], p[1]))


        df_post = pd.DataFrame(line_post)

        print('-------------------NEW DF--------------------------')
        print(df)
        df.insert(loc=1, column='post', value='-')
        #df['post'] = '-'

        for i in df_post.index:


            print(i)
            r_temp = df_post['r_no'][i]
            post_temp = df_post['post'][i]
            print(r_temp,post_temp)

            df.loc[df.roll_no == r_temp, 'post'] = post_temp






    df.to_excel(r'C:\Users\Pratik\Downloads\ResExcMod.xlsx',index=False)
except Exception as e:print(e)

#print(rank, roll_no, name, gender, category,dob)

#print(m_des,e_des,me_obj,g1,g2,g3,g4,wrt_tot,interview,total)

