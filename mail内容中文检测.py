#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd


# In[3]:


##import data
datasource = pd.read_excel("D://Users//pjchang//Desktop/标准Q解决率报表明细-IBU.xlsx")
Dlist = datasource["标准Q"]


# In[5]:


Dlist


# In[112]:


## get 邮件主题 list
for i in Dlist[6:7]:
    print(i)


# In[113]:


import sys

all_consequence_list = []
## chinese detection function, once return with one character
def check_contain_chinese(check_str):
    one_sentence_list = []
    for ch in check_str.encode('utf-8').decode('utf-8'):
            if u'\u4e00' <= ch <= u'\u9fff':
                one_sentence_list.append(ch)
    return one_sentence_list
                

for i in Dlist:
    one_consequence_list=[]
    a = check_contain_chinese(i)
    ## union all character into one sentence
    for e in a:
        a = "".join(a)
    ##all sentence into consequence list
    all_consequence_list.append(a)
    
print(all_consequence_list)


# In[118]:


## add into dataframe
datasource['中文检测'] = all_consequence_list


# In[119]:


datasource


# In[122]:


## wrap into one file
datasource.to_csv('Result.csv',encoding="utf_8_sig")

