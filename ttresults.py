#!/usr/bin/env python
# -*- coding: utf-8 -*-

#   Copyright (C) 2011 Richard A. Johnson <Email>
#
#   This program is free software: you can redistribute it and/or modify
#   it under the terms of the GNU General Public License as published by
#   the Free Software Foundation, either version 3 of the License, or
#   (at your option) any later version.
#
#   This program is distributed in the hope that it will be useful,
#   but WITHOUT ANY WARRANTY; without even the implied warranty of
#   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#   GNU General Public License for more details.
#
#   You should have received a copy of the GNU General Public License
#   along with this program.  If not, see <http://www.gnu.org/licenses/>.
"""Downloads results of the Indoor TT Series from Google Docs, parses the
information, and then creates HTML ready pages to be used for displaying
results via gender and age groups. Output is an ascii type table output
similar to those seen in triathlon events."""

import datetime
import gdata.service
import gdata.spreadsheet
import gdata.spreadsheet.service
import getopt
import os
import re
import sqlite3
import sys

# Spreadsheet Name
SPREADSHEET = '2011_TTSeries_Reg_Results'
# Header text to use on ouput & width of output
HEADER = '2011 ABD INDOOR TIME TRIAL SERIES'
# Data needed for results
RESULT_KEYS = ('age', 'gender', 'ridername', 'city', 'state', 'club',
               'tt1results', 'tt2results', 'tt3results', 'tt4results',
               'cumulative2', 'cumulative3', 'ttseriestotal')
# Age groups for results
AGE_GROUPS = ('10-14', '15-19', '20-24', '25-29', '30-34', '35-39', '40-44',
              '45-49', '50-54', '55-59', '60-64', '65-69', '70-74', '75-79',
              '80-84', '85-89', '90-94', '95-99')

# Project Path
PROJECT_PATH = os.path.dirname(os.path.abspath(__file__))

# Update Date & Time
LAST_UPDATED = datetime.datetime.now()

# CONTACT_NAME & CONTACT_EMAIL globals - set via command line
CONTACT_NAME = None
CONTACT_EMAIL = None


def create_gdocs_client(username, password):
    """Create the Google Docs client, connect, and login.

    Keyword arguments:
    username -- Google username (ie. <username>@gmail.com)
    password -- Google password

    """
    gc = gdata.spreadsheet.service.SpreadsheetsService()
    gc.email = username
    gc.password = password
    gc.source = 'ABD TT Results Application'
    gc.ProgrammaticLogin()
    return gc


def get_worksheet_data(gc):
    """Get the results data from the Google Docs spreadsheet.

    Keyword arguments:
    gc -- Google Docs client

    """
    feed = gc.GetSpreadsheetsFeed()
    spreads = [i.title.text for i in feed.entry]
    snum = [i for i, j in enumerate(spreads) if j == SPREADSHEET][0]
    parts = feed.entry[snum].id.text.split('/')
    key = parts[len(parts) - 1]
    feed = gc.GetWorksheetsFeed(key)
    parts = feed.entry[0].id.text.split('/')
    wid = parts[len(parts) - 1]
    return gc.GetListFeed(key, wid)


def events_completed(riders):
    """Count the events that have been completed. Helps with output formatting.

    Keyword arguments:
    riders -- Dictionary of riders

    """
    count = None
    for gender in riders:
        for group in riders[gender]:
            for rider in riders[gender][group]:
                if rider['tt4results']:
                    return 4
                if rider['tt3results']:
                    count = 3
                elif rider['tt2results'] and not count > 2:
                    count = 2
                elif rider['tt1results'] and not count > 1:
                    count = 1
    return count


def create_riders(results):
    """Create a nested dictionary of riders via gender and age groups.

    Keyword arguments:
    results -- Results retrieved from the Google spreadsheet

    """
    riders = {'MEN': {}, 'WOMEN': {}}
    for entry in results.entry:
        rider = {}
        for key in entry.custom:
            if key in RESULT_KEYS:
                rider[key] = entry.custom[key].text
        if rider['ridername'] and 'RIDER NAME' not in rider['ridername']:
            for group in AGE_GROUPS:
                for age in range(int(group.split('-')[0]),
                        int(group.split('-')[1])+1):
                    if (rider['gender'] == 'M' and rider['age'] and
                            int(rider['age']) == age):
                        try:
                            riders['MEN'][group].append(rider)
                        except KeyError:
                            riders['MEN'][group] = []
                            riders['MEN'][group].append(rider)
                    elif (rider['gender'] == 'F' and rider['age'] and
                            int(rider['age']) == age):
                        try:
                            riders['WOMEN'][group].append(rider)
                        except KeyError:
                            riders['WOMEN'][group] = []
                            riders['WOMEN'][group].append(rider)
    return riders


def create_sql_tables(events, riders, sqldb):
    """Create the SQLite3 tables.

    Keyword arguments:
    events -- Number of events completed
    riders -- Nested dictionary of riders
    sqldb -- SQLite3 database connection

    """
    c = sqldb.cursor()
    for sex in riders:
        tname = None
        for group in riders[sex]:
            tname = '%s_%s' % (sex, group.replace('-', '_'))
            c.execute('''create table %s (ridername text, city text,
            state text, club text, tt1results text, tt2results text,
            tt3results text, tt4results text, cumulative2 text, cumulative3 text,
            ttseriestotal text)''' % tname)
            for rider in riders[sex][group]:
                c.execute('''insert into %s values (?, ?, ?, ?, ?, ?, ?, ?, ?,
                ?, ?)''' % tname, (rider['ridername'], rider['city'],
                    rider['state'], rider['club'], rider['tt1results'],
                    rider['tt2results'], rider['tt3results'],
                    rider['tt4results'], rider['cumulative2'],
                    rider['cumulative3'], rider['ttseriestotal']))
            if events == 1:
                c.execute('''delete from %s where exists (select * from %s t2
                where %s.ridername = t2.ridername and %s.tt1results >
                t2.tt1results)''' % (tname, tname, tname, tname))
                c.execute('''delete from %s where %s.tt1results="DNS" or
                %s.tt1results="DNF" or %s.tt1results isnull''' % (tname, tname,
                    tname, tname))
            elif events == 2:
                c.execute('''delete from %s where exists (select * from %s t2
                where %s.ridername = t2.ridername and %s.cumulative2 >
                t2.cumulative2)''' % (tname, tname, tname, tname))
                c.execute('''delete from %s where exists (select * from %s t2
                where %s.cumulative2 isnull)''' % (tname, tname, tname))
            elif events == 3:
                c.execute('''delete from %s where exists (select * from %s t2
                where %s.ridername = t2.ridername and %s.cumulative3 >
                t2.cumulative3)''' % (tname, tname, tname, tname))
                c.execute('''delete from %s where exists (select * from %s t2
                where %s.cumulative3 isnull)''' % (tname, tname, tname))
            elif events == 4:
                c.execute('''delete from %s where exists (select * from %s t2
                where %s.ridername = t2.ridername and %s.ttseriestotal >
                t2.ttseriestotal)''' % (tname, tname, tname, tname))
                c.execute('''delete from %s where exists (select * from %s t2
                where %s.ttseriestotal isnull)''' % (tname, tname, tname))

    sqldb.commit()
    c.close()


def create_html_tables(events, sqldb):
    """Create HTML tables from the SQLite3 tables to display on website.

    Keyword Arguments:
    events -- The number of events completed
    sqldb -- SQLite3 database connection

    """
    c = sqldb.cursor()
    tnames = sorted([i[1] for i in c.execute('select * from sqlite_master')])
    groups = []
    for tname in tnames:
        gender, age1, age2 = tname.split('_')
        ages = '%s-%s' % (age1, age2)
        groups.append('%s %s' % (gender, ages))
    sorted(groups)
    # Column Header Names
    chn = ['Place', 'Name', 'City', 'St.', 'Club', 'TT #1 Time', 'TT #2 Time',
         'TT #3 Time', 'TT #4 Time', 'Total Time']
    # Column Header Borders
    chb = {
        '3': '===',
        '5': '=====',
        '12': '============',
        '20': '====================',
        '35': '==================================='}
    men = []
    women = []
    for i, t in enumerate(tnames):
        table = '<div id="%s" class="center">\n<pre>\n' % t
        if events == 1:
            width = 100
            head = HEADER.center(width)
            table += head
            table += '\n'
            c.execute('select * from %s order by tt1results' % t)
            group = groups[i].center(width)
            table += group
            table += '\n\n'
            line = '%-5s %-20s %-20s %-3s %-35s %-12s'
            table += (line % (chn[0], chn[1], chn[2], chn[3], chn[4], chn[5]))
            table += '\n'
            table += (line % (chb['5'], chb['20'], chb['20'], chb['3'],
                chb['35'], chb['12']))
            table += '\n'
            for place, rider in enumerate(c):
                table += (line % (place + 1, rider[0], rider[1], rider[2],
                    rider[3], rider[4]))
                table += '\n'
        elif events == 2:
            width = 126
            head = HEADER.center(width)
            table += head
            table += '\n'
            c.execute('select * from %s order by tt1results' % t)
            group = groups[i].center(width)
            table += group
            table += '\n\n'
            line = '%-5s %-20s %-20s %-3s %-35s %-12s %-12s %-12s'
            table += (line % (chn[0], chn[1], chn[2], chn[3], chn[4], chn[5],
                chn[6], chn[9]))
            table += '\n'
            table += (line % (chb['5'], chb['20'], chb['20'], chb['3'],
                chb['35'], chb['12'], chb['12'], chb['12']))
            table += '\n'
            for place, rider in enumerate(c):
                table += (line % (place + 1, rider[0], rider[1], rider[2],
                    rider[3], rider[4], rider[5], rider[8]))
                table += '\n'
        elif events == 3:
            width = 139
            head = HEADER.center(width)
            table += head
            table += '\n'
            c.execute('select * from %s order by tt1results' % t)
            group = groups[i].center(width)
            table += group
            table += '\n\n'
            line = '%-5s %-20s %-20s %-3s %-35s %-12s %-12s %-12s %-12s'
            table += (line % (chn[0], chn[1], chn[2], chn[3], chn[4], chn[5],
                chn[6], chn[7], chn[9]))
            table += '\n'
            table += (line % (chb['5'], chb['20'], chb['20'], chb['3'],
                chb['35'], chb['12'], chb['12'], chb['12'], chb['12']))
            table += '\n'
            for place, rider in enumerate(c):
                table += (line % (place + 1, rider[0], rider[1], rider[2],
                    rider[3], rider[4], rider[5], rider[6], rider['9']))
                table += '\n'
        elif events == 4:
            width = 152
            head = HEADER.center(width)
            table += head
            table += '\n'
            c.execute('select * from %s order by tt1results' % t)
            group = groups[i].center(width)
            table += group
            table += '\n\n'
            line = '%-5s %-20s %-20s %-3s %-35s %-12s %-12s %-12s %-12s %-12s'
            table += (line % (chn[0], chn[1], chn[2], chn[3], chn[4], chn[5],
                chn[6], chn[7], chn[8], chn['9']))
            table += '\n'
            table += (line % (chb['5'], chb['20'], chb['20'], chb['3'],
                chb['35'], chb['12'], chb['12'], chb['12'], chb['12'],
                chb['12']))
            table += '\n'
            for place, rider in enumerate(c):
                table += (line % (place + 1, rider[0], rider[1], rider[2],
                    rider[3], rider[4], rider[5], rider[6], rider['7'],
                    rider['10']))
                table += '\n'
        table += '</pre><div id="updated"><br />Last updated on %s at %s<br /><br /></div>' % (LAST_UPDATED.strftime('%m/%d/%Y'), LAST_UPDATED.strftime('%I:%M %p'))
        table += '<div id="contact">Contact <a href="mailto:%s?subject=Age Group TT Results Issue">%s</a> for any issues.</div></div>' % (CONTACT_EMAIL, CONTACT_NAME)
        if t.startswith('M'):
            men.append(table)
        elif t.startswith('W'):
            women.append(table)
    c.close()
    return men, women


def create_html(men, women):
    """Create the HTML skeleton."""
    html_mkup = '''<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd"
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8">
<meta name="robots" content="all">
<link rel="icon" href="http://www.ABDcycling.com/favicon.ico" type="image/x-icon"/>
<link rel="shortcut icon" href="http://www.ABDcycling.com/favicon.ico" type="image/x-icon"/>
<link rel="apple-touch-icon" href="http://www.ABDcycling.com/apple-touch-icon.png" type="image/x-icon"/>
<title>2011 Indoor TT Results</title>
<link rel="stylesheet" type="text/css" href="http://yui.yahooapis.com/combo?3.3.0/build/cssreset/reset-min.css&3.3.0/build/cssfonts/fonts-min.css&3.3.0/build/cssbase/base-min.css"/>
<script type="text/javascript" src="https://ajax.googleapis.com/ajax/libs/jquery/1.4.4/jquery.min.js"></script>
<script type="text/javascript" src="js/ui.core.min.js"></script>
<script type="text/javascript" src="js/ui.tabs.min.js"></script>
<script type="text/javascript">
    $(function() {
        $('#container ul').tabs();
    });
</script>
<style type="text/css">
    #header {padding-bottom:75px;}
    #header #logo-floater {float:right;padding-right:10px;}
    ul li {list-style:none;}
    .center {text-align:center;}
    #updated {text-align:center;color:#a0a0a0;font-style:italic;}
    #contact {text-align:center;color:#a0a0a0;}
    #contact a:link, #contact a:vistied {color:#296cd8;}
    #footer {width:100%;position:absolute;bottom:1px;text-align:center;background-color:#ebebeb;border-top:1px solid #97a5b0;font-size:77%;}
    #footer a:link, #footer a:visited {color:#296cd8;}
</style>
<link type="text/css" rel="stylesheet" href="css/tab.css">
<link type="text/css" rel="stylesheet" href="css/yui.css">
</head>
<body>
<div id="header">
<div id="logo-floater"><a href="http://www.abdcycling.com"><img src="img/abdlogoglobal_120x92.gif" alt="ABD Cycling" /></a></div>
</div>
<div id="container">
<ul>
<li><a href="#men"><span>Men</span></a></li>
<li><a href="#women"><span>Women</span></a></li>
</ul>
<div id="men">
<ul>'''
    li = []
    for mkup in men:
        hrefid = re.findall(r'\<div\ id\=\"(.*)\"\ class\=\"center\"\>', mkup,
                re.MULTILINE)[0]
        age1, age2 = hrefid.split('_')[1:]
        li.append('\n<li><a href="#%s"><span>%s-%s</span></a></li>' % (hrefid,
            age1, age2))
    li.append('\n</ul>\n')
    for line in li:
        html_mkup += line
    for table in men:
        html_mkup += table + '\n'
    html_mkup += '\n</div>\n<div id="women">\n<ul>'
    li = []
    for mkup in women:
        hrefid = re.findall(r'\<div\ id\=\"(.*)\"\ class\=\"center\"\>', mkup,
                re.MULTILINE)[0]
        age1, age2 = hrefid.split('_')[1:]
        li.append('\n<li><a href="#%s"><span>%s-%s</span></a></li>' % (hrefid,
            age1, age2))
    li.append('\n</ul>\n')
    for line in li:
        html_mkup += line
    for table in women:
        html_mkup += table + '\n'
    html_mkup += '\n</div>\n</div>'
    html_mkup += '''
<div id="footer">
&copy; 2011 Athletes By Design | <a href="http://www.abdcycling.com/about.html">About</a> | <a href="mailto:abdcycling@gmail.com">Contact Us</a>
</div>
'''
    html_mkup += '</body>\n</html>'
    f = open(os.path.join(PROJECT_PATH, 'html', 'index.html'), 'w')
    for line in html_mkup:
        f.write(line)
    f.close()


def main():
    """Main function to run application."""
    global CONTACT_NAME
    global CONTACT_EMAIL
    # Get username and password for Google Docs from the command line
    try:
        opts, args = getopt.getopt(sys.argv[1:], '', ['user=', 'pw=', 'name=', 'email='])
        del args
    except getopt.error:
        print 'ttresults.py --user [username] --pw [password] --name [Your name] --email [Your Email Address]'
        sys.exit(2)
    user = None
    pswd = None
    for o, a in opts:
        if o == '--user':
            user = a
        elif o == '--pw':
            pswd = a
        elif o == '--name':
            CONTACT_NAME = a
        elif o == '--email':
            CONTACT_EMAIL = a
    if not user or not pswd or not CONTACT_NAME or not CONTACT_EMAIL:
        print 'ttresults.py --user [username] --pw [password] --name [Your Name] --email [Your Email Address]'
        sys.exit(2)
    # Connect to a SQLite3 database located in memory
    sqlcon = sqlite3.connect(':memory:')

    # Get the Google Docs data
    gc = create_gdocs_client(user, pswd)
    data = get_worksheet_data(gc)

    # Create the riders tables
    riders = create_riders(data)

    # Get number of events completed
    events = events_completed(riders)

    # Create SQLite3 tables & data
    create_sql_tables(events, riders, sqlcon)

    # Create HTML Tables
    men, women = create_html_tables(events, sqlcon)

    # Create HTML Markup
    create_html(men, women)


if __name__ == '__main__':
    try:
        import psyco
        psyco.full()
    except ImportError:
        print 'If you had python-psyco installed, this would be faster'
    main()
