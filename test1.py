from bs4 import BeautifulSoup
import requests
import xlsxwriter
import itertools
from datetime import datetime, time

# --- Working with Excel block ---
workbook = xlsxwriter.Workbook('Top Movies 2019.xlsx')
worksheet = workbook.add_worksheet('2019 Feature Films')
topsheet = workbook.add_worksheet('New Top 2019')

for sheet in worksheet, topsheet:
    sheet.set_column('A:A', 50)
    sheet.set_row(0, 30)
    sheet.set_column('B:E', 15)

cellformat = workbook.add_format()
cellformat.set_align('center')
cellformat.set_align('vcenter')
cellformat.set_bold()
cellformat.set_bg_color('#b5f5cb')

headers1 = ['Name', 'IMBD Rating', 'Metascore', 'Votes']
for i in range(4):
    for header in headers1:
        worksheet.write(0, i, headers1[i], cellformat)

headers2 = ['Name', 'Final Score']
for i in range(2):
    for header in headers2:
        topsheet.write(0, i, headers2[i], cellformat)


# --- Creating a Beautiful Soup object ---
response = requests.get('https://www.imdb.com/search/title/?title_type=feature&year=2019-01-01,2019-12-31&ref_=adv_prv')
soup = BeautifulSoup(response.content, 'lxml')

movie_block = soup.find('div', attrs={'class': 'lister-item-content'})
movie_title = soup.find('h3', attrs={'class': 'lister-item-header'})


# --- Completing Excel Table ---
counter = itertools.count(1, 1)
i = 1
max_votes = 0
my_score = 0
total = []

while True:
    try:
        if i < 50:
            nxt = next(counter)
            movie_block = movie_block.find_next('div', attrs={'class': 'lister-item-content'})
            movie_title = movie_block.find('h3', attrs={'class': 'lister-item-header'})

            try:
                name = str(movie_title.find('a').text)
                imbdrating = float(movie_block.find('div', attrs={
                    'class': 'inline-block ratings-imdb-rating'})['data-value'])
                if imbdrating < 7.5:
                    i += 1
                    counter = itertools.count(nxt, 1)
                    continue
                metascore = int(movie_block.find('div', attrs={
                    'class': 'inline-block ratings-metascore'}).find('span').text)
                if metascore < 75:
                    i += 1
                    counter = itertools.count(nxt, 1)
                    continue
                votes = int(movie_block.find('span', attrs={
                    'name': 'nv'})['data-value'])
                if votes > max_votes:
                    max_votes = votes

                worksheet.write(nxt, 0, name)
                worksheet.write(nxt, 1, imbdrating)
                worksheet.write(nxt, 2, metascore)
                worksheet.write(nxt, 3, votes)

                total.append([name, imbdrating, metascore, votes])

                print(name)
            except AttributeError:
                counter = itertools.count(nxt, 1)
            except TypeError:
                counter = itertools.count(nxt, 1)
            i += 1
        else:
            link = soup.find('a', attrs={'class': 'lister-page-next next-page'})['href']
            response = requests.get('https://www.imdb.com{}'.format(link))
            soup = BeautifulSoup(response.content, 'lxml')
            movie_block = soup.find('div', attrs={'class': 'lister-item-content'})
            movie_title = soup.find('h3', attrs={'class': 'lister-item-header'})
            i = 1
    except:
        for i in range(len(total)):
            my_score = int(((total[i][1] * 10) + total[i][2] + (total[i][3] * 100 / max_votes)) / 3)
            total[i].append(my_score)

        total = sorted(total, key=lambda x: x[4], reverse=True)

        for i in range(0, len(total)):
            topsheet.write(i+1, 0, total[i][0])
            topsheet.write(i+1, 1, total[i][4])

        workbook.close()
