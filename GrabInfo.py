from bs4 import BeautifulSoup
import requests,openpyxl

excel=openpyxl.Workbook()
print(excel.sheetnames)
sheet=excel.active
sheet.title='IMDB Top 250'
print(excel.sheetnames)
sheet.append(['Rank','Title','Year','Ranking'])

try:
    source=requests.get('https://www.imdb.com/chart/top')
    source.raise_for_status()

    soup=BeautifulSoup(source.text,'html.parser')

    movie=soup.find('tbody',class_="lister-list").find_all('tr')
    #print(len(movie))
    for i in movie:
        name=i.find('td',class_="titleColumn").a.text
        rank=i.find('td',class_="titleColumn").get_text(strip=True).split('.')[0]
        year=i.find('td',class_="titleColumn").span.text.strip('()')
        rating=i.find('td',class_="ratingColumn imdbRating").strong.text
        sheet.append([int(rank),name,int(year),float(rating)])
        print(rank,name,year,rating)
except Exception as e:
    print(e)
excel.save('Imdb rating1.xlsx')
