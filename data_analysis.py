import requests
import os
import pandas as pd
from bs4 import BeautifulSoup
from textstat import flesch_reading_ease, syllable_count

# Load the Excel file
input_file = 'Input.xlsx'
df = pd.read_excel(input_file)

# Create a directory to save the extracted articles
output_directory = 'extracted_articles'
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# Load positive and negative word lists with ISO-8859-1 encoding
positive_words = set(open('MasterDictionary/positive-words.txt', 'r', encoding='ISO-8859-1').read().split())
negative_words = set(open('MasterDictionary/negative-words.txt', 'r', encoding='ISO-8859-1').read().split())


# Load stopwords
stopwords = set()
stopword_files = ['StopWords_Auditor.txt', 'StopWords_Currencies.txt', 'StopWords_DatesandNumbers.txt',
                  'StopWords_GenericLong.txt', 'StopWords_Generic.txt', 'StopWords_Geographic.txt', 'StopWords_Names.txt']
for file in stopword_files:
    stopwords.update(open(os.path.join('StopWords', file), 'r', encoding='ISO-8859-1').read().split())


# Create a DataFrame to store the results
output_df = pd.DataFrame(columns=[
    'URL_ID', 'URL', 'POSITIVE SCORE', 'NEGATIVE SCORE', 'POLARITY SCORE', 'SUBJECTIVITY SCORE',
    'AVG SENTENCE LENGTH', 'PERCENTAGE OF COMPLEX WORDS', 'FOG INDEX',
    'AVG NUMBER OF WORDS PER SENTENCE', 'COMPLEX WORD COUNT', 'WORD COUNT',
    'SYLLABLE PER WORD', 'PERSONAL PRONOUNS', 'AVG WORD LENGTH'
])

# Loop through the URLs and perform textual analysis
for index, row in df.iterrows():
    url_id = str(row['URL_ID'])
    url = row['URL']
    output_filename = os.path.join(output_directory, f'{url_id}.txt')

    with open(output_filename, 'r', encoding='ISO-8859-1', errors='ignore') as file:
        content = file.read()

    # Perform textual analysis
    words = content.split()
    word_count = len(words)
    complex_word_count = sum(1 for word in words if syllable_count(word) > 2)
    sentence_count = content.count('.') + content.count('!') + content.count('?')
    avg_words_per_sentence = word_count / sentence_count if sentence_count > 0 else 0
    fog_index = 0.4 * (avg_words_per_sentence + complex_word_count)

    positive_score = sum(1 for word in words if word in positive_words)
    negative_score = sum(1 for word in words if word in negative_words)
    polarity_score = (positive_score - negative_score) / (positive_score + negative_score + 1e-10)
    subjectivity_score = (positive_score + negative_score) / (word_count + 1e-10)

    personal_pronouns = sum(1 for word in words if word.lower() in {'i', 'me', 'my', 'mine', 'we', 'us', 'our', 'ours'})

    avg_word_length = sum(len(word) for word in words) / (word_count + 1e-10)
    syllable_per_word = sum(syllable_count(word) for word in words) / (word_count + 1e-10)

    # Add results to the DataFrame
    output_df.loc[index] = [
        url_id, url, positive_score, negative_score, polarity_score, subjectivity_score,
        flesch_reading_ease(content), complex_word_count / word_count, fog_index,
        avg_words_per_sentence, complex_word_count, word_count, syllable_per_word,
        personal_pronouns, avg_word_length
    ]

    print(f'Performed textual analysis for URL_ID {url_id}.')

# Save the DataFrame to "Output.xlsx"
output_file = 'Output.xlsx'
output_df.to_excel(output_file, index=False)

print('Textual analysis complete. Results saved to Output.xlsx.')
