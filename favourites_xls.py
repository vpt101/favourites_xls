from twython import Twython
from xlwt.Workbook import Workbook
from xlwt import easyxf, Formula
import time
import sys
import argparse
import re


# Uncomment these lines and fill in your details before running
access_token_key = '23733847-vL377cLdwxIOWuge9GIpYN7Kg5qUqF1OoLDhyr7aZ'
access_token_secret = 'pKPYzK5xvV1qVW4lQQ6G0UeFnapflDjvf7nzcq34iypYb'

consumer_key = 'hE1FtTwq8Ig2m9fTSzHWUHkTi'
consumer_secret = 'rPVUeVu23dEX2meHi4pwO6frjhFxPX3lZFnXiNrFqM7XZhqe19'

rangeRegex = re.compile('[1-9]*(:?[0-9]+)')

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
    


def getTwythonHandle():
    return Twython(consumer_key, consumer_secret, access_token_key, access_token_secret)


def favourites_xls(opt):

    t = getTwythonHandle()
    # Let's create an empty xls Workbook and define formatting
    wb = Workbook()
    ws = wb.add_sheet('0')

    # Column widths
    ws.col(0).width = 256 * 30
    ws.col(1).width = 256 * 30
    ws.col(2).width = 256 * 60
    ws.col(3).width = 256 * 30
    ws.col(4).width = 256 * 30

    # Stylez
    style_link = easyxf('font: underline single, name Arial, height 220, colour_index blue')
    style_heading = easyxf('font: bold 1, name Arial, height 220; pattern: pattern solid, pattern_fore_colour yellow, pattern_back_colour yellow')
    style_wrap = easyxf('align: wrap 1; font: height 220')

    # Headings in proper MBA spreadsheet style - Bold with yellow background
    ws.write(0, 0, 'Author', style_heading)
    ws.write(0, 1, 'Twitter Handle', style_heading)
    ws.write(0, 2, 'Text', style_heading)
    ws.write(0, 3, 'Embedded Links', style_heading)

    # Let's start at page 1 of your favourites because you know, it's a very
    # good place to start
    count = 1
    pagenum = 1

    # Now, let's start an infinite loop and I don't mean the one with Apple's HQ
    while True:
        # Get your favourites from Twitter
        faves = t.get_favorites(page=pagenum)

        # If there's no favourites left from this page OR we've reached the
        # page number specified by our user, do the Di Caprio and jump out
        if len(faves) == 0 or (opt != 'all' and pagenum > int(opt)):
            break

        # Programmers have been doing inception for ages before Nolan did. Let's go deeper and get
        # into another loop now
        for fav in faves:
            
            ws.write(count, 0, fav['user']['name'],style_wrap)
            ws.write(count, 1, fav['user']['screen_name'],style_wrap)
            ws.write(count, 2, fav['text'],style_wrap)
            links = fav['entities']['urls']
            i = 0
            for link in links:
                formatted_link = 'HYPERLINK("%s";"%s")' % (link['url'],"link")
                ws.write(count, 3+i, Formula(formatted_link), style_link)
                i += 1
                formatted_ExpandedLink = 'HYPERLINK("%s";"%s")' % (link['expanded_url'],"expanded_link")
                ws.write(count, 3+i, Formula(formatted_ExpandedLink), style_link)
                i += 1

            count += 1
            #print 'At count: [%d%%]\r'%count,
            sys.stderr.write('\rPG:%d,CNT:%d'%(pagenum, count))

        pagenum += 1
        sys.stderr.write('\rPG:%d,CNT:%d'%(pagenum, count))
        if (fib0 == 0): nextFib();
        nextFib();
        secs = fib0;
        time.sleep(secs)
        # print ('Sleeping for {} secs\n'.format(secs));
        if (fib0 >= 1597):
            resetFib();
        

    # Now for the step that has caused untold misery and suffering to people who forget to do it at work
    wb.save('twitter_favourites.xls')


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
