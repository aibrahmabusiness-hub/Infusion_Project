import pandas as pd

df = pd.DataFrame(columns=['Order ID', 'MRN', 'Medication'])
print("Original cols:", df.columns.tolist())

# Series apply on empty column
series = df['Medication'].apply(lambda x: True)
print("Series:", series)
print("Series index:", series.index)

filtered = df[series]
print("Filtered cols:", filtered.columns.tolist())
print("Filtered df:", filtered)
