# Deduplication Microsoft CE
This script can be used to deduplicate Companies listed in Microsoft CE (Dynamics). When working with multiple account managers labels and/or companies. The larger the company, the more difficult data quality becomes, but also the more essential. This script tries to find names of companies that are very similar, so they could be merged in CE. Leaving you with one entry per company. 

It works really well for minor differences in company names, like `the <company>` vs `<company>` vs `<company> LLC>`. This scripts outputs an Excel file with relations between CE-accounts which are similar and could be the same. A confidence score is added. Hand the output the CE administrators, they should be able to merge companies.  

Start by installing the `dedupe`


```python
! pip install pandas-dedupe
```


```python
import pandas as pd
import re
from unidecode import unidecode
import pandas_dedupe
```

## Load data

Download the data from CE:  
1. Open CE (Dynamics)  
2. Go to Customers > Accounts  
3. Select 'All Accounts'  
4. Click on 'Export to Excel'  

Replace the `filename` to the path you saved the data.


```python
filename = 'Accounts.xlsx'
dataset = pd.read_excel(filename)
```


```python
data = dataset[['Account Name', 'Account Number', 'Relationship Type']]
```

## Preprocessing
Do a little preprocessing on the names, like removing newlines, some punctuations.


```python
def preProcess(column):
    """Do a little bit of data cleaning with the help of Unidecode and Regex.
    Things like casing, extra spaces, quotes and new lines can be ignored.
    """

    column = unidecode(column)
    column = re.sub('\n', ' ', column)
    column = re.sub('-', '', column)
    column = re.sub('/', ' ', column)
    column = re.sub("'", '', column)
    column = re.sub(",", '', column)
    column = re.sub(":", ' ', column)
    column = re.sub('  +', ' ', column)
    column = column.strip().strip('"').strip("'").lower().strip()
    if not column:
        column = None
    return column
```


```python
data['name'] = data['Account Name'].apply(preProcess)
```

## Run Deduplication


```python
df_final = pandas_dedupe.dedupe_dataframe(data, ['name'])
```


```python
# save
ts = datetime.datetime.today().strftime("%Y%m%d%H%M%S")
df_final.to_csv(f'{ts}_deduplication_output.csv')
```

## Join additional useful information


```python
# select columns
df_final_ = df_final[["name","cluster id","Account Number"]]

# merge with itself so you'll get relationships (A is similar to B)
df_joined = pd.merge(df_final, df_final_, on=["cluster id"], how="left")

# remove the same accounts !(A == A)
results = df_joined[df_joined['Account Number_x'] != df_joined['Account Number_y'] ]

# merge account number of the joined accounts (B)
results = pd.merge(results,data[['Account Number', 'Account Name']] ,left_on=["Account Number_y"], right_on = ['Account Number'],how="left")

# save
results[['Account Number_x', 'Account Number_y', 'Account Name_x', 'Account Name_y', 'confidence']].to_excel(f'{ts}_Depuplication_results_CE.xlsx', index = False)
```
