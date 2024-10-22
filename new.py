import os
import requests
from bs4 import BeautifulSoup
import pandas as pd
import re


def load_stop_words(stop_words_folder):
    stop_words = set()
    for filename in os.listdir(stop_words_folder):
        file_path = os.path.join(stop_words_folder, filename)
        with open(file_path, 'r') as f:
            stop_words.update(f.read().splitlines())
    return stop_words

def load_master_dictionary(master_dict_folder):
    positive_words = set()
    negative_words = set()

    positive_file = os.path.join(master_dict_folder, 'positive-words.txt')
    negative_file = os.path.join(master_dict_folder, 'negative-words.txt')

    
    with open(positive_file, 'r') as pos_file:
        positive_words.update(word.strip() for word in pos_file if word.strip())

   
    with open(negative_file, 'r') as neg_file:
        negative_words.update(word.strip() for word in neg_file if word.strip())

    return positive_words, negative_words


def count_syllables(word):
    word = word.lower()
    vowels = "aeiou"
    syllable_count = 0
    if word[0] in vowels:
        syllable_count += 1
    for index in range(1, len(word)):
        if word[index] in vowels and word[index-1] not in vowels:
            syllable_count += 1
    if word.endswith("e"):
        syllable_count -= 1
    if syllable_count == 0:
        syllable_count += 1
    return syllable_count


def count_personal_pronouns(text):
    pronouns = re.findall(r'\b(I|we|my|ours|us)\b', text, re.I)
    return len(pronouns)


def extract_article_text(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    
    title = soup.find('h1').text.strip() if soup.find('h1') else 'No Title'

    paragraphs = soup.find_all('p')
    article_text = ' '.join([p.text.strip() for p in paragraphs])

    return title, article_text


def analyze_sentiment(text, stop_words, positive_words, negative_words):
    words = text.split()
    words = [word for word in words if word.lower() not in stop_words]

    word_count = len(words)
    sentences = re.split(r'[.!?]', text)
    sentence_count = len(sentences) if sentences else 1
    
    avg_sentence_length = word_count / sentence_count if sentence_count > 0 else 0
    

    complex_words = [word for word in words if count_syllables(word) > 2]
    complex_word_count = len(complex_words)
    
    percentage_complex_words = (complex_word_count / word_count) * 100 if word_count > 0 else 0
    
    
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)

    avg_word_length = sum(len(word) for word in words) / word_count if word_count > 0 else 0

    total_syllables = sum(count_syllables(word) for word in words)
    syllables_per_word = total_syllables / word_count if word_count > 0 else 0

    positive_score = sum(1 for word in words if word.lower() in positive_words)
    negative_score = sum(1 for word in words if word.lower() in negative_words)
    
    polarity_score = (positive_score - negative_score) / (positive_score + negative_score + 1)
    subjectivity_score = (positive_score + negative_score) / (word_count + 1)


    personal_pronouns = count_personal_pronouns(text)

    return {
        'positive_score': positive_score,
        'negative_score': negative_score,
        'polarity_score': polarity_score,
        'subjectivity_score': subjectivity_score,
        'avg_sentence_length': avg_sentence_length,
        'percentage_complex_words': percentage_complex_words,
        'fog_index': fog_index,
        'avg_words_per_sentence': avg_sentence_length, 
        'complex_word_count': complex_word_count,
        'word_count': word_count,
        'syllables_per_word': syllables_per_word,
        'personal_pronouns': personal_pronouns,
        'avg_word_length': avg_word_length
    }

stop_words_folder = r''
master_dict_folder = r''

currency_stopwords = load_stop_words(stop_words_folder)
positive_words, negative_words = load_master_dictionary(master_dict_folder)


input_file = r''  
urls_df = pd.read_excel(input_file)
urls = urls_df['URL'].tolist() 


results = []
for url in urls:
    title, article_text = extract_article_text(url)
    analysis = analyze_sentiment(article_text, currency_stopwords, positive_words, negative_words)
    

    result = {
        'Title': title,
        'URL': url
    }
    result.update(analysis)
    results.append(result)

output_df = pd.DataFrame(results)

output_df = output_df[[
    'positive_score', 'negative_score', 'polarity_score', 'subjectivity_score', 'avg_sentence_length', 
    'percentage_complex_words', 'fog_index', 'avg_words_per_sentence', 'complex_word_count', 
    'word_count', 'syllables_per_word', 'personal_pronouns', 'avg_word_length'
]]

output_file = r''  
output_df.to_excel(output_file, index=False)

print("Sentiment analysis completed. Results saved to out3.xlsx.")
