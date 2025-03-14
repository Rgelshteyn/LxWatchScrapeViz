import os
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

def load_data(file_path):
    """Load data from an Excel file."""
    return pd.read_excel(file_path)

def clean_data(df):
    """Clean the data by removing currency symbols and converting prices to float."""
    df['Price'] = df['Price'].replace('[\$\€\£,]', '', regex=True).astype(float)
    return df

def plot_brand_distribution(df):
    """Plot the distribution of watch listings by brand."""
    brand_counts = df['Brand'].value_counts()
    plt.figure(figsize=(10, 8))
    sns.barplot(x=brand_counts.values, y=brand_counts.index)
    plt.xlabel('Number of Listings')
    plt.ylabel('Brand')
    plt.title('Distribution of Watch Listings by Brand')
    plt.show()

def plot_average_price_by_brand(df):
    """Plot the average price of watches by brand."""
    avg_price = df.groupby('Brand')['Price'].mean().sort_values()
    plt.figure(figsize=(10, 8))
    sns.barplot(x=avg_price, y=avg_price.index)
    plt.xlabel('Average Price')
    plt.ylabel('Brand')
    plt.title('Average Price by Brand')
    plt.show()

def plot_price_distribution_by_brand(df):
    """Plot the price distribution for each brand using box plots."""
    plt.figure(figsize=(12, 8))
    sns.boxplot(x='Brand', y='Price', data=df, order=df['Brand'].value_counts().index)
    plt.xticks(rotation=45)
    plt.xlabel('Brand')
    plt.ylabel('Price')
    plt.title('Price Distribution by Brand')
    plt.show()

def plot_price_trends_over_time(df):
    """Plot the average price trends over time."""
    df['Date'] = pd.to_datetime(df['Date'])
    avg_price_over_time = df.groupby(df['Date'].dt.date)['Price'].mean()
    plt.figure(figsize=(10, 6))
    plt.plot(avg_price_over_time.index, avg_price_over_time.values, marker='o')
    plt.xlabel('Date')
    plt.ylabel('Average Price')
    plt.title('Average Price Trends Over Time')
    plt.grid(True)
    plt.show()

if __name__ == "__main__":
    current_dir = os.path.dirname(os.path.abspath(__file__))
    excel_file_path = os.path.join(current_dir, 'WatchForumList.xlsx')
    #excel_file_path
    df = load_data(excel_file_path)
    df = clean_data(df)

    # Plotting the graphs
    plot_brand_distribution(df)
    plot_average_price_by_brand(df)
    plot_price_distribution_by_brand(df)
    plot_price_trends_over_time(df)
