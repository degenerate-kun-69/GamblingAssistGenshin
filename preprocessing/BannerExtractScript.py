import pandas as pd
from bs4 import BeautifulSoup

# Assuming 'html_content' contains the HTML content of the page
soup = BeautifulSoup(html_content, 'html.parser')

rows = []

# Find all tables
tables = soup.find_all('table', class_='a-table')

for table in tables:
    version = table.find_previous('h3', class_='a-header--3').text.split()[-1]
    phases = table.find_all('tr')

    for i, phase in enumerate(phases[1:], start=1):  # Skip header row
        # Extract banner data
        banners = phase.find_all('td')[0].find_all('a', class_='a-link')
        banner_1 = banners[0].text if len(banners) > 0 else 'nil'
        banner_2 = banners[1].text if len(banners) > 1 else 'nil'

        # Extract rate-up characters
        rate_up_divs = phase.find_all('td')[1].find_all('div', class_='align')
        rate_up_1 = rate_up_divs[0].find_all('a') if len(rate_up_divs) > 0 else []
        rate_up_2 = rate_up_divs[1].find_all('a') if len(rate_up_divs) > 1 else []

        # Character names for each banner
        chars_1 = [a['alt'].replace('Genshin -', '') for a in rate_up_1 if 'alt' in a.attrs]
        chars_2 = [a['alt'].replace('Genshin -', '') for a in rate_up_2 if 'alt' in a.attrs]

        # Extract date range
        date_range = phase.find_all('td')[2].text.strip()

        # Append row data
        rows.append([
            version,
            i,
            banner_1,
            banner_2,
            date_range,
            *chars_1,
            *(chars_2 if chars_2 else ['nil'] * len(chars_1))
        ])

# Now 'rows' contains the extracted data
# Define column names based on the maximum character entries
max_chars = max(len(row) for row in rows) - 5
columns = ['Version', 'Phase', 'Banner 1', 'Banner 2', 'Date Range'] + [f'Character {i}' for i in range(1, max_chars + 1)]

# Save to Excel
df = pd.DataFrame(rows, columns=columns)
df.to_excel('preprocessing/CharacterBannerList.xlsx', index=False, startrow=1)