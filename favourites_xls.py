from twython import Twython, TwythonRateLimitError
from xlwt.Workbook import Workbook
from xlwt import easyxf, Formula
import time
import sys
# import argparse
import re
import csv
import sqlite3

RED='\033[1;31m'
BLUE='\033[1;34m'
GREEN='\033[1;32m'
CYAN='\033[1;36m'
MAGENTA='\033[1;35m'
GREY='\033[1;30m'
YELLOW='\033[1;33m'
NC='\033[0m'
MAX_FIB=233

# Uncomment these lines and fill in your details before running
access_token_key = '23733847-vL377cLdwxIOWuge9GIpYN7Kg5qUqF1OoLDhyr7aZ'
access_token_secret = 'pKPYzK5xvV1qVW4lQQ6G0UeFnapflDjvf7nzcq34iypYb'
consumer_key = 'hE1FtTwq8Ig2m9fTSzHWUHkTi'
consumer_secret = 'rPVUeVu23dEX2meHi4pwO6frjhFxPX3lZFnXiNrFqM7XZhqe19'

rangeRegex = re.compile('[1-9]*(:?[0-9]+)')

maxId=-1

global fib0
fib0 = 0
global fib1
fib1 = 1
def nextFib():
    global fib1, fib0
    fib = fib1;
    fib1 = fib0 + fib1;
    fib0 = fib
def resetFib():
    global fib1, fib0
    fib0=0;
    fib1=1;

def colorize(f):
    if (f < 3):
        sys.stdout.write(GREY)
    elif (f < 21):
        sys.stdout.write(BLUE)
    elif(f < 144):
        sys.stdout.write(CYAN)
    elif(f < 233):
        sys.stdout.write(GREEN)
    elif(f < 610):
        sys.stdout.write(YELLOW)
    elif(f < 987):
        sys.stdout.write(MAGENTA)
    elif(f <= 1597):
        sys.stdout.write(RED)


def getTwythonHandle():
    return Twython(consumer_key, consumer_secret, access_token_key, access_token_secret)

def getSQLiteConnection():
    return sqlite3.connect('D:\\vpt\\swift.db');


def setWidths(ws):
    # Column widths
    ws.col(0).width = 256 * 5
    ws.col(1).width = 256 * 5
    ws.col(2).width = 256 * 5
    ws.col(3).width = 256 * 5
    ws.col(4).width = 256 * 10
    ws.col(5).width = 256 * 10
    ws.col(6).width = 256 * 10
    ws.col(7).width = 256 * 30
    ws.col(8).width = 256 * 5
    ws.col(9).width = 256 * 5
    ws.col(10).width = 256 * 10
    ws.col(11).width = 256 * 10

def escapeStrings(key, value):
    if (key == 'Date' or key == 'ProbSensitive' or key == 'Name'
        or key == 'Handle' or key == 'Text' or key == 'DP' or key == 'Banner'
        or key == 'Links' or key == 'OtherLinks'):
        return "'" + value + "'"
    else:
        return value

def prepareWrite():
    return 'INSERT INTO twitter_favourites ("Id", "Date", "RtCnt", "FavCnt", "ProbSensitive", "Name", "Handle", "Text", "DP", "Banner", "Links", "OtherLinks") VALUES (?,?,?,?,?,?,?,?,?,?,?,?)'

def favourites_xls(opt):
    conn = getSQLiteConnection()
    cursor = conn.cursor();
    for dbRow in cursor.execute('select max(id) as topVal from twitter_favourites'):
        maxId=dbRow[0];
    if (maxId is None):
        maxId = -1;
    print (maxId);
    cursor.close();
    conn.close();

    t = getTwythonHandle()
    # Let's create an empty xls Workbook and define formatting
    wb = Workbook()
    ws = wb.add_sheet('0')

    # Stylez
    style_link = easyxf('font: underline single, name Arial, height 160, colour_index blue')
    style_heading = easyxf('font: bold 1, name Arial, height 160; pattern: pattern solid, pattern_fore_colour yellow, pattern_back_colour yellow')
    style_wrap = easyxf('align: wrap 1; font: height 160')
    # style_nowrap = easyxf('font: height 160')
    style_id = easyxf('align: wrap 1; font: height 160;', "#")
    style_date = easyxf('font: height 160;', "ddd mmm DD HH:MM:SS+")

    
    # Headings in proper MBA spreadsheet style - Bold with yellow background
    ws.write(0, 0, 'Id', style_heading)
    ws.write(0, 1, 'Date', style_heading)
    ws.write(0, 2, 'RtCnt', style_heading)
    ws.write(0, 3, 'FavCnt', style_heading)
    ws.write(0, 4, 'ProbSensitive', style_heading)
    ws.write(0, 5, 'Name', style_heading)
    ws.write(0, 6, 'Handle', style_heading)
    ws.write(0, 7, 'Text', style_heading)
    ws.write(0, 8, 'DP', style_heading)
    ws.write(0, 9, 'Banner', style_heading)
    ws.write(0, 10, 'Links', style_heading)
    ws.write(0, 11, 'OtherLinks', style_heading)
    

    with open('twitter_favourites.txt', 'w+', newline='', encoding='utf-8') as csvfile:
        fieldHeaders = ['Id', 'Date', 'RtCnt', 'FavCnt', 'ProbSensitive', 'Name', 'Handle', 'Text', 'DP', 'Banner', 'Links', 'OtherLinks'];
        cw = csv.DictWriter(csvfile, fieldnames=fieldHeaders)
        cw.writeheader();

        # Let's start at page 1 of your favourites because you know, it's a very
        # good place to start
        count = 1
        pagenum = 1
        # Now, let's start an infinite loop and I don't mean the one with Apple's HQ
        while True:
            # Get your favourites from Twitter
            try: 
                faves = t.get_favorites(page=pagenum)
                sql=prepareWrite()
                # If there's no favourites left from this page OR we've reached the
                # page number specified by our user, do the Di Caprio and jump out
                if len(faves) == 0 or (opt != 'all' and pagenum > int(opt)):
                    print ("\nDone.\n");
                    break
                rows = [];
                # Programmers have been doing inception for ages before Nolan did. Let's go deeper and get
                # into another loop now
                for fav in faves:
                    row = {}
                    user = fav['user']
                    ws.write(count, 0, fav['id'], style_id)
                    row['Id'] = ('%s' % fav['id']);
                    
                    ws.write(count, 1, fav['created_at'], style_date)
                    row['Date'] = ('%s' % fav['created_at']);
                    
        
                    ws.write(count, 2, fav['retweet_count'], style_wrap)
                    row['RtCnt'] = ('%s' % fav['retweet_count']);
        
                    ws.write(count, 3, fav['favorite_count'], style_wrap)
                    row['FavCnt'] = ('%s' % fav['favorite_count']);
        
                    if ('possibly_sensitive' in fav):
                        ws.write(count, 4, fav['possibly_sensitive'], style_wrap)
                        row['ProbSensitive'] = ('%s' % fav['possibly_sensitive']);
                    else:
                        row['ProbSensitive'] = ''
        
                    ws.write(count, 5, user['name'],style_wrap);
                    row['Name'] = ('%s' % user['name']);
        
                    ws.write(count, 6, user['screen_name'], style_wrap)
                    row['Handle'] = ('%s' % user['screen_name']);
        
                    ws.write(count, 7, fav['text'],style_wrap)
                    row['Text'] = ('%s' % fav['text']);
        
        
                    if ('profile_image_url_https' in user):
                        formattedLink = 'HYPERLINK("%s";"%s")' % (user['profile_image_url_https'],"DP")
                        ws.write(count, 8, Formula(formattedLink), style_link)
                        row['DP'] = ('%s' % user['profile_image_url_https']);
        
                    
                    if ('profile_banner_url' in user):
                        formattedLink = 'HYPERLINK("%s";"%s")' % (user['profile_banner_url'],"BNR")
                        ws.write(count, 9, Formula(formattedLink), style_link)
                        row['Banner'] = ('%s' % user['profile_banner_url']);
                    else:
                        row['Banner'] = '';
        
                    
                    # LINKS
                    links = fav['entities']['urls']
                    i = 0
                    row['Links'] = '';
                    row['OtherLinks']='';
                    for link in links:
                        formatted_link = 'HYPERLINK("%s";"%s")' % (link['url'],"link")
                        ws.write(count, 10 + i, Formula(formatted_link), style_link)
                        row['Links'] += ('%s ' % link['url']);
            
                        i += 1
                        formattedExpandedLink = 'HYPERLINK("%s";"%s")' % (link['expanded_url'],"expanded_link")
                        if (len(formattedExpandedLink) < 255):
                            ws.write(count, 10 + i, Formula(formattedExpandedLink), style_link)
                            row['OtherLinks'] += ('%s ' % link['expanded_url']);
                
                        i += 1
                    count += 1
                    colorize(fib0)

                    sys.stderr.write('\rPG:%d,CNT:%d'%(pagenum, count))
                    cw.writerow(row);
                    dbRow=(row["Id"], row["Date"], row["RtCnt"], row["FavCnt"], row["ProbSensitive"], row["Name"], row["Handle"], row["Text"], row["DP"], row["Banner"], row["Links"], row["OtherLinks"])
                    if (maxId < int(row["Id"])):
                        rows.append(dbRow);


                try: 
                    conn = getSQLiteConnection()
                    cursor = conn.cursor();
                    cursor.executemany(sql, rows);
                    conn.commit();
                    conn.close();
                except sqlite3.OperationalError as oe: 
                    sys.stderr.write(RED)
                    sys.stderr.write("\nERROR! %s\n" % oe)
                    sys.stderr.write('\n')
                    conn.close();
                except sqlite3.IntegrityError as ie:
                    sys.stderr.write(RED)
                    sys.stderr.write("\nERROR! %s\n" % ie)
                    conn.close();
                    
                pagenum += 1
                colorize(fib0)
                sys.stderr.write('\rPG:%d,CNT:%d' % (pagenum, count))
                if (fib0 == 0): nextFib();
                nextFib();
                secs = fib0;
                    
                time.sleep(secs)
                if (fib0 >= MAX_FIB):
                    resetFib();
            
            except TwythonRateLimitError as trle:
                sys.stderr.write('%s' % trle);
                rateLimits=t.get_application_rate_limit_status()
                
                favLimits=rateLimits['resources']['favorites']["/favorites/list"];
                if (favLimits['remaining'] == 0):
                    waitTime = favLimits['reset'] - int(time.time());
                    waitTime += 15;
                    print("\n Sleeping for: %s seconds" % waitTime);
                    time.sleep(waitTime)
                else:
                    print("Something weird happened: %s" % trle);
        
    # conn.close();
    # Now for the step that has caused untold misery and suffering to people who forget to do it at work
    wb.save('twitter_favourites.xls');

if __name__ == '__main__':
# =============================================================================
#     parser = argparse.ArgumentParser()
#     parser.add_argument('pageRange', help="pass the range of pages to download. from:to")
#     args = parser.parse_args()
#     print(args.pageRange)
#     if (rangeRegex.match(args.pageRange)):
#         pageRange=args.pageRange.split(':', 1)
# 
#     print ("pageRange=", pageRange)
# 
# =============================================================================
    start = time.time()

    if len(sys.argv) == 1:
        print ('Usage is python favourites_xls.py all OR python favourites_xls <number of pages of recent faves>')
    else:
        if (sys.argv[1].isdigit()):
            favourites_xls(sys.argv[1])
        else:
            if (sys.argv[1] == 'all'):
                favourites_xls(sys.argv[1])
            else:
                print ('Please use either all or a page number. This script is not Niels Bohr ok?')

    end = time.time()
    
    print("\nIt took this long:", end-start)
    print("\n")