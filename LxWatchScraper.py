import requests
from bs4 import BeautifulSoup
from datetime import datetime , timedelta
import pandas as pd

url = 'https://www.watchrecon.com'

def remove_duplicates(existing_file_path, new_data):
    try:
        existing_df = pd.read_excel(existing_file_path)
        existing_urls = existing_df['Link'].tolist() if 'Link' in existing_df.columns else []
    except FileNotFoundError:
        existing_df = pd.DataFrame()
        existing_urls = []

    # Filter out duplicate data
    filtered_new_data = [item for item in new_data if item.get('Link') not in existing_urls]
    return existing_df, filtered_new_data

desired_brands = ['Rolex', 'Omega', 'Patek Philippe', 'Audemars Piguet', 'Tag Heuer', 'Seiko', 'Sinn', 'Oris', 'Vacheron Constantin', 'Zenith', 'Girard Perregaux', 'Nomos', "Jaeger-LeCoultre",'Longines','Doxa' ,'A. Lange & Söhne',
'Cartier', 'F.P.Journe', 'Hamilton', 'Richard Mille', 'Tudor', 'Zodiac', 'Grand Seiko', 'Ulysse Nardin']

# Scrape the website for watch data
def scrape_watch_data(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')
    listings = soup.find_all('div', class_='galleryItemContainer med')
    watches = []
    current_date = datetime.now().strftime('%m/%d/%Y')

    for listing in listings:
        brand = listing.find('span', class_='brandInfo').text.strip() if listing.find('span', class_='brandInfo') else ''
        model = listing.find('span', class_='modelInfo').text.strip() if listing.find('span', class_='modelInfo') else ''
        price_text = listing.find('span', class_='priceInfo').text.strip() if listing.find('span', class_='priceInfo') else ''
        link = listing.find('a', class_='listingLink')['href'] if listing.find('a', class_='listingLink') else ''


        if brand in desired_brands and model and price_text:
            watches.append({'Brand': brand, 'Model': model, 'Price': price_text, 'Link': link, 'Date': current_date})

    return watches

if __name__ == "__main__":
    excel_file_path = 'B:\WatchReconScraper\WatchForumList.xlsx'


    # Scrape the watch data from the website
    new_watches_data = scrape_watch_data(url)

    try:
        existing_df = pd.read_excel(excel_file_path)
    except FileNotFoundError:
        existing_df = pd.DataFrame()

    # Remove any duplicates
    existing_df, filtered_data = remove_duplicates(excel_file_path, new_watches_data)

    if filtered_data:
        filtered_df = pd.DataFrame(filtered_data)

        if existing_df.empty:
            with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
                filtered_df.to_excel(writer, sheet_name='Sheet1', index=False)
        else:
            with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                start_row = len(existing_df) if not existing_df.empty else 0
                filtered_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=start_row)

        print(f'Data written to {excel_file_path}')
    else:
        print('No new data to write.')