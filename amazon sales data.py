import pandas as pd
import numpy as np

#data extraction
#load dataset from excel file
df = pd.read_excel("Amazon.xlsx")
print(df.columns)


#rename columns to match the expected names
df = df.rename(columns={
    'Total Revenue': 'Sales',
    'Item Type': 'Product Category'
})

print(df.columns)


#data transformation and cleaning
#drop duplicates and handle missing values
df.drop_duplicates(inplace=True)
df.dropna(inplace=True)


#convert date columns to datetime format
df['Order Date'] = pd.to_datetime(df['Order Date'])

#extract relevant columns
df = df[['Order Date', 'Sales', 'Product Category', 'Region']]
print(df)

#extract year, month, and year-month from the date
df['Year'] = df['Order Date'].dt.year
df['Month'] = df['Order Date'].dt.month
df['Year_Month'] = df['Order Date'].dt.to_period('M')

#descriptive statistics
print(df.describe())

#save the cleaned data
#df.to_csv('cleaned_Amazon.csv', index=False)

#data analysis
#calculate monthwise sales trend
month_wise_sales = df.groupby('Month')['Sales'].sum()

#calculate yearwise sales trend
year_wise_sales = df.groupby('Year')['Sales'].sum()

#calculate yearly monthwise sales trend
yearly_month_wise_sales = df.groupby(['Year', 'Month'])['Sales'].sum().unstack()

#identify key metrics
total_sales = df['Sales'].sum()
average_sales = df['Sales'].mean()
num_transactions = df.shape[0]

#determine relationships between attributes
sales_by_category = df.groupby('Product Category')['Sales'].sum()
sales_by_region = df.groupby('Region')['Sales'].sum()


#visualization of data
import matplotlib.pyplot as plt
from matplotlib import style
style.use('Solarize_Light2')
#print(plt.style.available)
import seaborn as sns

#distribution of Sales
plt.figure(figsize=(10, 6))
sns.histplot(df['Sales'], bins=30, kde=True)
plt.title('Distribution of Sales')
plt.xlabel('Sales')
plt.ylabel('Frequency')
plt.show()

#box plot of Sales by Product Category
plt.figure(figsize=(12, 6))
sns.boxplot(x='Product Category', y='Sales', data=df)
plt.title('Box Plot of Sales by Product Category')
plt.xticks(rotation=45)
plt.show()

#correlation matrix
numeric_cols = df.select_dtypes(include=[np.number])
corr_matrix = numeric_cols.corr()
plt.figure(figsize=(10, 8))
sns.heatmap(corr_matrix, annot=True, cmap='coolwarm', fmt='.2f')
plt.title('Correlation Matrix')
plt.show()


#monthwise sales trend
plt.figure(figsize=(10, 6))
month_wise_sales.plot(kind='bar',color="orange")
plt.title('Month-wise Sales Trend')
plt.xlabel('Month')
plt.ylabel('Total Sales')
plt.show()

#year wise sales trend
plt.figure(figsize=(10, 6))
year_wise_sales.plot(kind='bar',color="green")
plt.title('Year-wise Sales Trend')
plt.xlabel('Year')
plt.ylabel('Total Sales')
plt.show()

#yearly monthwise sales trend
plt.figure(figsize=(12, 8))
sns.heatmap(yearly_month_wise_sales, annot=True, fmt='.0f', cmap='viridis')
plt.title('Yearly Month-wise Sales Trend')
plt.xlabel('Month')
plt.ylabel('Year')
plt.show()

#sales by product category
plt.figure(figsize=(10, 6))
sales_by_category.plot(kind='bar',color="pink")
plt.title('Sales by Product Category')
plt.xlabel('Product Category')
plt.ylabel('Total Sales')
plt.show()

#sales by region
plt.figure(figsize=(10, 6))
sales_by_region.plot(kind='bar',color="purple")
plt.title('Sales by Region')
plt.xlabel('Region')
plt.ylabel('Total Sales')
plt.show()

