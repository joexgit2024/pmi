import pandas as pd

# Load the data
df = pd.read_excel('input/2025 - PMI Sydney Chapter Project Management Day of Service (PMDoS) 2025 Professional Registration (Responses).xlsx')

print('LinkedIn URL analysis:')
urls = df['LinkedIn Profile URL'].dropna()
print(f'Total LinkedIn URLs provided: {len(urls)}')
print(f'Total PMP participants: {len(df)}')
print(f'Percentage with LinkedIn: {len(urls)/len(df)*100:.1f}%')

print('\nURL formats:')
for url in urls.head(10):
    print(f'- {url}')

print('\nURL patterns analysis:')
url_patterns = {}
for url in urls:
    if 'linkedin.com/in/' in str(url):
        url_patterns['linkedin.com/in/'] = url_patterns.get('linkedin.com/in/', 0) + 1
    elif 'au.linkedin.com' in str(url):
        url_patterns['au.linkedin.com'] = url_patterns.get('au.linkedin.com', 0) + 1
    elif 'www.linkedin.com' in str(url):
        url_patterns['www.linkedin.com'] = url_patterns.get('www.linkedin.com', 0) + 1
    else:
        url_patterns['other'] = url_patterns.get('other', 0) + 1

for pattern, count in url_patterns.items():
    print(f'{pattern}: {count}')