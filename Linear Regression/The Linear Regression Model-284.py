## 2. Introduction To The Data ##

import pandas as pd
data= pd.read_csv("~/Downloads/AmesHousing.txt", sep = "\t")
train = data [0 : 1460]
test = data [1460 : ]
target = "SalePrice"
data.info()


## 3. Simple Linear Regression ##

import matplotlib.pyplot as plt
# For prettier plots.
import seaborn
fig = plt.figure(figsize = (7,15))
ax1 = fig.add_subplot(3, 1, 1)
ax2 = fig.add_subplot(3, 1, 2)
ax3 = fig.add_subplot(3, 1 ,3)
ax1.scatter(train["Garage Area"], train["SalePrice"])
ax2.scatter(train["Gr Liv Area"], train["SalePrice"])
ax3.scatter(train["Overall Cond"], train["SalePrice"])
fig.show()


## 5. Using Scikit-Learn To Train And Predict ##

from sklearn.linear_model import LinearRegression
lm = LinearRegression ()

lm.fit(train[['Gr Liv Area']], train['SalePrice'])
a1 = lm.coef_
a0 = lm.intercept_
lm

## 6. Making Predictions ##

import numpy as np
from sklearn.metrics import mean_squared_error as mse

lr = LinearRegression()
lr.fit(train[['Gr Liv Area']], train['SalePrice'])
predictions_train = lr.predict(train[['Gr Liv Area']])
predictions_test = lr.predict(test[['Gr Liv Area']])

train_rmse = np.sqrt(mse(train['SalePrice'], predictions_train))
test_rmse = np.sqrt(mse(test['SalePrice'], predictions_test))



## 7. Multiple Linear Regression ##

from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error as mse
cols = ['Overall Cond', 'Gr Liv Area']
lm = LinearRegression()
lm.fit(train[cols], train['SalePrice'])
lm_1 = lm.predict(train[cols])
lm_2 = lm.predict(test[cols])
train_rmse_2 = np.sqrt(mse(train['SalePrice'], lm_1))
test_rmse_2 = np.sqrt(mse(test['SalePrice'], lm_2))

