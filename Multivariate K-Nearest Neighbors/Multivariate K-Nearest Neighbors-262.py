## 1. Recap ##

import pandas as pd
import numpy as np
np.random.seed(1)

dc_listings = pd.read_csv('~/Downloads/dc_airbnb.csv')
dc_listings = dc_listings.loc[np.random.permutation(len(dc_listings))]
stripped_commas = dc_listings['price'].str.replace(',', '')
stripped_dollars = stripped_commas.str.replace('$', '')
dc_listings['price'] = stripped_dollars.astype('float')
dc_listings.info()

## 2. Removing features ##

drop_columns = ['room_type', 'city', 'state', 'latitude', 'longitude', 'zipcode', 'host_response_rate', 'host_acceptance_rate', 'host_listings_count']
dc_listings = dc_listings.drop(drop_columns, axis=1)
print(dc_listings.isnull().sum())

## 3. Handling missing values ##

drop_columns = ['cleaning_fee', 'security_deposit']
dc_listings = dc_listings.drop(drop_columns, axis=1)
drop_rows = ['bedrooms', 'bathrooms', 'beds']
dc_listings = dc_listings.dropna(axis = 0)

print(dc_listings.isnull().sum())



## 4. Normalize columns ##

normalized_listings = (dc_listings  - dc_listings.mean()) / dc_listings.std()
normalized_listings['price'] = dc_listings ['price']
print (normalized_listings.iloc[0:4])

## 5. Euclidean distance for multivariate case ##

from scipy.spatial import distance
cols= ["accommodates", "bathrooms"]
first_listing = normalized_listings[cols].iloc[0]
fifth_listing = normalized_listings[cols].iloc[4]
first_fifth_distance = distance.euclidean(first_listing, fifth_listing)


## 7. Fitting a model and making predictions ##

from sklearn.neighbors import KNeighborsRegressor

train_df = normalized_listings.iloc[0:2792]
test_df = normalized_listings.iloc[2792:]

knn = KNeighborsRegressor(algorithm = 'brute')

train_features = train_df[['accommodates', 'bathrooms']]
# List-like object, containing just the target column, `price`.
train_target = train_df['price']
# Pass everything into the fit method.
knn.fit(train_features, train_target)
predictions = knn.predict(test_df[['accommodates', 'bathrooms']])
print (predictions)

## 8. Calculating MSE using Scikit-Learn ##

from sklearn.metrics import mean_squared_error
import numpy as np
train_columns = ['accommodates', 'bathrooms']
knn = KNeighborsRegressor(n_neighbors=5, algorithm='brute', metric='euclidean')
knn.fit(train_df[train_columns], train_df['price'])
predictions = knn.predict(test_df[train_columns])
two_features_mse = mean_squared_error (test_df["price"], predictions)
two_features_rmse = np.sqrt(two_features_mse)


## 9. Using more features ##

features = ['accommodates', 'bedrooms', 'bathrooms', 'number_of_reviews']
from sklearn.neighbors import KNeighborsRegressor
knn = KNeighborsRegressor(n_neighbors=5, algorithm='brute')


import numpy as np

knn.fit(train_df[features], train_df['price'])
four_predictions = knn.predict(test_df[features])
four_mse = mean_squared_error (test_df["price"], four_predictions)
four_rmse = np.sqrt(four_mse)




## 10. Using all features ##


features = ['accommodates', 'bedrooms', 'bathrooms', 'beds',
       'minimum_nights', 'maximum_nights', 'number_of_reviews']
from sklearn.neighbors import KNeighborsRegressor
knn = KNeighborsRegressor(n_neighbors=5, algorithm='brute')


import numpy as np

knn.fit(train_df[features], train_df['price'])
all_features_predictions = knn.predict(test_df[features])
all_features_mse  = mean_squared_error (test_df["price"], all_features_predictions)
all_features_rmse  = np.sqrt(all_features_mse )


