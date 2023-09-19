product_performance = df.groupby('product name').agg({
    'price with dph': 'sum',
    'quantity sold': 'sum',
    'faktura id': 'count'
}).reset_index()

# mean price per product
product_performance['Average Price'] = product_performance['price with dph'] / product_performance['quantity sold']
