from bs4 import BeautifulSoup
import requests, openpyxl

url = 'https://www.imdb.com/chart/top/'
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
    }

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'IMDB top movies'

sheet.append(['Name', 'Rank', 'Year', 'Rating'])

try:
  source = requests.get(url, headers=headers)
  source.raise_for_status()

  soup = BeautifulSoup(source.text, 'html.parser')

  movies = soup.find_all('div', class_='sc-b0691f29-0 jbYPfh cli-children')

  for movie in movies:

    print('---------------------------------------------')
    print()

    name = movie.find('a', class_='ipc-title-link-wrapper').h3.get_text(strip=True).split('.')[1].strip(' ')

    rank = movie.find('a', class_='ipc-title-link-wrapper').h3.get_text(strip=True).split('.')[0].strip(' ')

    year = movie.find('div', class_='sc-b0691f29-7 hrgukm cli-title-metadata').find('span', class_='sc-b0691f29-8 ilsLEX cli-title-metadata-item').text

    rating = movie.find('span', class_='ipc-rating-star ipc-rating-star--base ipc-rating-star--imdb ratingGroup--imdb-rating').get_text(strip=True).split('(')[0]

    print('Name : ',name)
    print('Rank : ',rank)
    print('Year : ',year)
    print('Rating : ',rating)

    sheet.append([name, rank, year, rating])

except Exception as e:
  print(e)

excel.save('IMDB Movie.xlsx')