from cmath import e
import requests
from bs4 import BeautifulSoup
import openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = "Top Rated movies"
sheet.append(['MOVIE RANK','NAME','YEAR','RATING'])

try:
    source = requests.get('https://www.imdb.com/chart/top/')
    source.raise_for_status()
    soup = BeautifulSoup(source.text,'html.parser')
    movies = soup.find('tbody',class_='lister-list').find_all('tr')
    # mov=movies.find_all('tr')
    # print(len(movies))
    for movie in movies:
        name = movie.find('td',class_='titleColumn').a.text
        rank = movie.find('td',class_='titleColumn').get_text(strip=True).split('.')[0]
        year = movie.find('td',class_='titleColumn').span.text.strip('()')
        rating = movie.find('td',class_="ratingColumn imdbRating").strong.text
        print(name,rank,year,rating)
        sheet.append([name,rank,year,rating])
        #break

except Exception as e:
    print(e)
excel.save('IMDB RATINGS.XLSX')