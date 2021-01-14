
# coding: utf-8

# In[77]:


import numpy as np
import pandas as pd


# ### Read data & add the header

# In[78]:


df = pd.read_excel('NZ_Admin_JOBS.xlsx',sheet_name='sheet1', header=None, names=["Job", "URL", "Company", "Location", "Posted Date", "Classification"])
df


# ### Refine the "Posted Date" column

# Split into "Posted Date" and "Posted Place"

# In[79]:


new = df["Posted Date"].str.split(",", n = -1, expand = True) 
df["Posted Date"] = new[0]
df["Posted Place"] = new[2]
df


# ### Refine the "Location" column

# 1. some rows contain the salary after "," and we need to remove them

# In[80]:


new = df["Location"].str.split(",", n = 1, expand = True) 
df["New Location"] = new[0]


# 2. Need to remove the "location:" at the beginning of the strings, and then separate the region and the city

# In[81]:


new = df["New Location"].str.split(": ", n = 3, expand = True) 


# In[82]:


df['Region'] = new[1]
df['City'] = new[2]
df = df.drop('Location',axis=1)
df = df.drop('New Location',axis=1)
df


# 3. Need to remove the duplication

# In[83]:


def cut_half(x):
    if x != None:
        if x[-4:] == "area":
            return (x[0:((len(x)-4)//2)])
        elif x == "Administration & Office SupportAdministration & Office SupportsubClassification":
            return "Administration & Office Support"
        else:
            return (x[0:(len(x)//2)])
    else:
        return None


# In[84]:


df["Region"] = df["Region"].apply(cut_half)
df["City"] = df["City"].apply(cut_half)
df


# ### Refine "Classification"

# 1. Some rows contain the "classification" (i.e., Administration & Office SupportAdministration & Office SupportsubClassification) and the "subclassification", others contain the salary

# let df2 be the dataframe contains "classification" and "subclassification":

# In[85]:


df2 = df[df.Classification.str.contains(':',case=False)]
df2


# let df3 be the dataframe contains "salary":

# In[86]:


df3 = df[~df.Classification.str.contains(':',case=True)]
df3=df3.rename(columns = {'Classification':'Salary'})
df3


# In[87]:


new2 = df2["Classification"].str.split(": ", n = 3, expand = True) 


# In[88]:


df2['Classification'] = new2[1]
df2['Sub-classification'] = new2[2]
# df2 = df2.drop('Classification',axis=1)
df2


# 2. Remove duplication

# In[89]:


df2["Classification"] = df2["Classification"].apply(cut_half)
df2["Sub-classification"] = df2["Sub-classification"].apply(cut_half)
df2


# In[90]:


df3


# ### Merge df2 and df3

# In[91]:


df4 = pd.merge(df2, df3, how='outer')
df5 = df4[["Job", "URL", "Company", "Posted Date", "Posted Place", "Classification", "Sub-classification", "Region", "City", "Salary"]]
df5


# ### Remove null data

# In[92]:


df5 = df5.fillna(value='NO DATA')
df5


# ### Export to excel

# In[93]:


df5.to_excel('Cleaned_NZ_Admin_JOBS.xlsx',sheet_name='Sheet1')


# ### To do:

# In[94]:


Refine the "Salary" column

