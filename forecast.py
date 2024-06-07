import pandas as pd
import matplotlib.pyplot as plt 
import numpy as np
from sklearn.model_selection import train_test_split


#Read document and save it as dataframe
df = pd.read_excel('SuperstoreDataset.xlsx')

train, test = train_test_split(df, test_size=0.2, random_state=0)

ids = train["Sub-Category"].unique()
cats = train["Segment"].unique()
dict = {}
#because we ran out of time, we took the mean of similar products and used it to estimate the sales
for id in ids:
    for cat in cats:
        dict[(id+"."+cat)] = np.mean(list(train[(train['Sub-Category']==id) & (train['Segment']==cat)]["Sales"]))

diff = 0
#we will test with root mean squared difference between the values
for item in test:
        pass