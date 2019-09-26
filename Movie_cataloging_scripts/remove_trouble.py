import os, winshell
import urllib2
from bs4 import BeautifulSoup
from mechanize import Browser
import re
from win32com.client import Dispatch

source_dir = r"F:\Movies\Sorted\Troublesome"
dest_dir = r"F:\Movies\Sorted"

def getunicode(soup):
  body=''
  if isinstance(soup, unicode):
    soup = soup.replace('\'',"'")
    soup = soup.replace('&quot;','"')
    soup = soup.replace('&nbsp;',' ')
    body = body + soup
  else:
    if not soup.contents:
      return ''
    con_list = soup.contents
    for con in con_list:
      body = body + getunicode(con)
  return body



def createshortcut(dest_folder, movie_name, source_folder, full_movie_name, rating):
    
    if rating != '' :
        path = os.path.join(dest_folder, movie_name + " (" + rating + ").lnk")
    else :
        path = os.path.join(dest_folder, movie_name + ".lnk")
    
    if os.path.exists(path):
        return
    target = source_folder + "\\" + full_movie_name 
    wDir = source_folder 
    icon = target  
    
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = target
    shortcut.WorkingDirectory = wDir
    shortcut.IconLocation = icon
    shortcut.save()

def makeshortcut(movie, full_name,condition):
  movie_search = '+'.join(movie.split())

  base_url = 'http://www.imdb.com/find?q='
  url = base_url+movie_search+condition
  title_search = re.compile('/title/tt\d+')

  br = Browser()
  br.open(url)
  link = br.find_link(url_regex = re.compile(r'/title/tt.*'))
  res = br.follow_link(link)
  print link

  soup = BeautifulSoup(res.read(),"html.parser")
  movie_title = getunicode(soup.find('title'))
  print movie_title

  rate = soup.find('span',itemprop='ratingValue')
  rating = getunicode(rate)
  print rating

  actors=[]
  actors_soup = soup.findAll('span',{'itemprop':'actors'})
  for i in range(len(actors_soup)):
    actors.append(getunicode(actors_soup[i].find('span',{'itemprop':'name'})))
  print actors
  
  genre=[]
  genrelist = soup.findAll('span',itemprop='genre')
  for i in range(len(genrelist)):
    genre.append(getunicode(genrelist[i]))
  print genre

  subtext = soup.find('div',{'class':'subtext'}).findAll('a')[-1]
  release_date = str(getunicode(subtext)).rstrip().split()
  year = str(release_date[-2].strip())
  z = 0
  while not year.isdigit():
      year = str(release_date[-2 - z].strip())
      z = z + 1
  decade = year[:-1]
  decade = decade + str('0s')

  print year
  for j in range(len(actors)):
      d = dest_dir + r'\\By Actors\\' + actors[j]
      if not os.path.exists(d):
          os.makedirs(d)
      createshortcut(d, movie, source_dir, full_name, rating)

  for j in range(len(genre)):
      d = dest_dir + '\\By Genres\\' + genre[j] + '\\' + decade + '\\' + year
      if not os.path.exists(d):
          os.makedirs(d)
      createshortcut(d, movie, source_dir, full_name, rating)


  rating_folder = ''
  if int(rating[2]) >= 5 : 
      rating_folder = rating[0] + '.5 - ' + str(int(rating[0]) + 1) + '.0'
  else :
      rating_folder = rating[0] + '.0 - ' + rating[0] + '.5'
  d = dest_dir + '\\By IMDB rating\\' + rating_folder + '\\' + decade + '\\' + year
  if not os.path.exists(d):
      os.makedirs(d)
  createshortcut(d, movie, source_dir, full_name, rating)

def shortcuts(list_movies,files):
  No_of_Movies = len(list_movies)

  for i in range(No_of_Movies):
    if files:
      temp = list_movies[i].rfind('.')
      movie = ''
      full_name = list_movies[i]
      if temp != -1:
          movie = list_movies[i][:temp]
      else :
          movie = list_movies[i]
    else:
      movie = full_name = list_movies[i]

    print movie + "\t" + str(float(i)/float(No_of_Movies)*100) + "\% done. " + str(i+1) + " of " + str(No_of_Movies) + " movies done."


    try:
      makeshortcut(movie,full_name,'&s=tt&ttype=ft&ref_=fn_ft')
    except:
      makeshortcut(movie,full_name,'&s=all')

shortcuts([name for name in os.listdir(source_dir) if os.path.isdir(os.path.join(source_dir, name))] , False)
shortcuts([name for name in os.listdir(source_dir) if not os.path.isdir(os.path.join(source_dir, name))] , True)