## 1. Introduction ##

import pandas as pd

data = pd.read_csv('~/Downloads/AmesHousing.txt', delimiter="\t")
train = data[0:1460]
test = data[1460:]

train_null_counts = train.isnull().sum()
print(train_null_counts)
cols_no_null = train_null_counts[ train_null_counts == 0].index
df_no_mv = train[cols_no_null]
print (cols_no_null)

## 2. Categorical Features ##

text_cols = df_no_mv.select_dtypes(include=['object']).columns
df_no_mv.info()


for col in text_cols:
    print(col+":", len(train[col].unique()))
    train[col] = train[col].astype('category')
    
print(train['Utilities'].cat.codes)
print(train['Utilities'].cat.codes.value_counts())

## 3. Dummy Coding ##

dummy_cols = pd.DataFrame()
for text in text_cols:
    dummy_values = pd.get_dummies(train[text])
    train = train.drop(text, axis = 1)
    train = pd.concat([train, dummy_values], axis = 1)
print (train.iloc[1])
#print (dummy_values)
train.shape




## 4. Transforming Improper Numerical Features ##

train['years_until_remod'] = train['Year Remod/Add'] - train['Year Built']
print (train['years_until_remod'])

## 5. Missing Values ##

import pandas as pd

data = pd.read_csv('~/Downloads/AmesHousing.txt', delimiter="\t")
train = data[0:1460]
test = data[1460:]

train_null_counts = train.isnull().sum()
cols_good = (train_null_counts[(train_null_counts<584) & (train_null_counts>0) ].index)
df_missing_values = train [cols_good]
df_missing_values.isnull().sum()
df_missing_values.info()


## 6. Imputing Missing Values ##

float_cols = df_missing_values.select_dtypes(include=['float'])
#for col in float_cols:
 #   col_mean = df_missing_values[col].mean()
  #  df_missing_values[col].fillna(col_mean)

   # df_missing_values[col].isnull()
float_cols = float_cols.fillna(float_cols.mean())
float_cols.info()
#print(float_cols.columns)   
#df_missing_values[float_cols]
