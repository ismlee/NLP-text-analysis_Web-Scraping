import requests
import os
import pandas as pd
from bs4 import BeautifulSoup

# Load the Excel file
input_file = 'Input.xlsx'
df = pd.read_excel(input_file)

# Create a directory to save the extracted articles
output_directory = 'extracted_articles'
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# Loop through the URLs and extract article text
for index, row in df.iterrows():
    url_id = str(row['URL_ID'])
    url = row['URL']

    # Send a request to the URL and get the HTML content
    response = requests.get(url)
    html_content = response.text

    # Parse the HTML content with BeautifulSoup
    soup = BeautifulSoup(html_content, 'html.parser')

    # Find article title using one of the specified classes
    # entry-title or tdb-title-text

    possible_title_classes = ['entry-title', 'another-title-class', 'yet-another-title-class']
    article_title = None
    for class_name in possible_title_classes:
        article_title_element = soup.find('h1', class_=class_name)
        if article_title_element:
            article_title = article_title_element.get_text().strip()
            break

    if article_title is None:
        print(f"Article title not found for URL_ID {url_id}.")
        article_title = "Article Title Not Found"

    # Extract article text using one of the specified classes
    # td-post-content tagdiv-type or tdb-block-inner td-fix-index

    possible_text_classes = ['elementor-element-54f0702', 'another-text-class', 'yet-another-text-class']
    article_text = ''
    for class_name in possible_text_classes:
        article_text_elements = soup.find_all('div', class_=class_name)
        if article_text_elements:
            for element in article_text_elements:
                article_text += element.get_text().strip() + '\n'
            break

    # Create a text file with the article content
    output_filename = os.path.join(output_directory, f'{url_id}.txt')
    with open(output_filename, 'w', encoding='utf-8') as file:
        file.write(article_title + '\n\n')
        file.write(article_text)

    print(f'Extracted and saved article for URL_ID {url_id}.')

print('Extraction complete.')
