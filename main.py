#
# GDG - Online Cloud Study Jams Vietnam - Quests Counter
# Author: VietKubers team
# Date: Aug 29, 2019
#

import datetime
import io
import os
import pprint
import random
import shutil
import six
import sys

import bs4
from console import fg, bg, fx, defx  # shortcut: sc
import openpyxl
import requests

DEBUG = False
COUNT_TOP_PEOPLE_ONLY = 5 if DEBUG else 0

INPUT_FILE = None
INPUT_DATA = {
    'doers': {},
    'result': {
        'ok': None,
        'error': None,
    },
    'excel': {
        'workbook': None,
        'worksheet': None,
    },
}

GDOCS_FILE_ID = '1VE2sH6zePhdwaSDir9ucUoXPYTXIjIR3eRFKQ-IVZcw'
GDOCS_SHEET_ID = '241580121'
GDOCS_DOWNLOAD_LINK = 'https://docs.google.com/feeds/download/spreadsheets/Export?key=%(file_id)s&exportFormat=xlsx&gid=%(sheet_id)s'

COL_RESULT_QUEST_COUNT = 'J'
COL_RESULT_ALL = 'K'
COL_RESULT_HANOI = 'L'
COL_RESULT_DANANG = 'M'
COL_RESULT_HCM = 'N'

# Colors
COLOR_ALL = [ 'black', 'red', 'green', 'yellow', 'blue', 'magenta',
              'cyan', 'white', 'lightblack', 'lightred',
              'lightgreen', 'lightyellow', 'lightblue', 'lightmagenta',
              'lightcyan', 'lightwhite' ]

MAX_LINE_LEN = int(shutil.get_terminal_size(fallback=(80, 30))[0] * 0.8)

SHOW_QUESTS_DETAIL = True
QUEST_COUNT_FROM = [datetime.date(2019, 7, 28), datetime.date(2019, 8, 30)]
SKIP_QUESTS = ['GCP Essentials']


def pp(*args):
    print(*args)
    
def pp_err(*args):
    pp()
    pp(bg.lightred + fg.white + fx.bold, 'ERROR', fx.end, *args)
    
def pp_warn(*args):
    pp()
    pp(bg.lightyellow + fg.blue + fx.bold, 'WARNING', fx.end, *args)
    
def random_bg():
    while True:
        color = random.choice(COLOR_ALL)
        # BG must not be black?
        if color not in ('black', 'lightblack'):
            return color

def random_fg(bg=None):
    while True:
        color = random.choice(COLOR_ALL)
        # FG must not be same (or similar) as BG
        if not bg or not (color.endswith(bg) or bg.endswith(color)):
            return color

def main():
    # Parse args
    global INPUT_FILE
    arg_1st = sys.argv[1] if len(sys.argv) > 1 else None
    if arg_1st:
        if arg_1st in ('--help', '-h', '?'):
            usage()
            return
        else:
            INPUT_FILE = arg_1st
    # Try download input file
    if not INPUT_FILE:
        download_input()
    # Parse input
    parse_input(INPUT_FILE)
    # Process input
    count_quests()
    
def usage():
    pp(fg.blue + bg.lightpurple + fx.bold + fx.frame,
       'GDG Quest Counter', fx.end)
    pp(fg.green, '  Usage: run.sh [Input (EXCEL xlsx file)]', fx.end)

def download_input():
    pp()
    pp(bg.lightyellow + fg.blue + fx.bold, 'DOWNLOADING GOOGLE DOCS FILE ...', fx.end)
    pp()
    
    filepath = 'result.xlsx'
    url = GDOCS_DOWNLOAD_LINK % {'file_id': GDOCS_FILE_ID, 'sheet_id': GDOCS_SHEET_ID}
    req = requests.get(url)
    
    with open(filepath, 'w+b') as file:
        file.write(req.content)
    
    pp('      Input file saved into', fg.cyan + filepath + fx.end)
    pp()
    
    global INPUT_FILE
    INPUT_FILE = filepath

def parse_input(input):
    wb = openpyxl.load_workbook(filename=input)
    sh = wb['Results']
    
    doers = INPUT_DATA['doers']
    rows_not_processed = []
    
    row_id = 0
    for row in sh.iter_rows():
        row_id += 1
        if row[0].value == 'Timestamp':
            # Skip header row
            pass
        elif not row[0].is_date:
            rows_not_processed.append(row)
        else:
            person = {
                'row_id': row_id,
                'timestamp': row[0].value,
                'email': row[1].value.strip().lower(),
                'name': row[2].value.strip(),
                'nick_name': row[3].value.strip(),
                'qwiklabs_link': row[4].value.strip(),
                'location': row[5].value.strip(),
                'quests': [],
                'legal_quests': [],
            }
            email = person['email']
            if email in doers:
                # Duplicated doer
                pp_warn('Dupplicated participant', fg.cyan + email + fx.end)
            doers[email] = person
    
    if len(rows_not_processed):
        show_ignored_rows(rows_not_processed)
        
    # Save the excel input
    INPUT_DATA['excel']['workbook'] = wb
    INPUT_DATA['excel']['worksheet'] = sh
        
def show_ignored_rows(rows):
    # TODO
    pp('IGNORED ROWS')
    
def count_quests():
    doers = INPUT_DATA['doers']
    person_index = 0
    for person in doers.values():
        try:
            count_quests_of(person)
        except Exception as ex:
            pp('Unable to parse QUESTS report for user %s' % person['email'])
        person_index += 1
        # For testing only
        if COUNT_TOP_PEOPLE_ONLY and person_index == COUNT_TOP_PEOPLE_ONLY:
            break
    
    # Track ERROR and OK reports
    error_list = []
    ok_list = []
    for person in doers.values():
        if person.get('error', None):
            error_list.append(person)
        else:
            ok_list.append(person)
    
    # Sort people by LEGAL quests
    ok_list.sort(key=lambda x: len(x['legal_quests']), reverse=True)
    
    # Filter and sort result by location
    hanoi_list = []
    danang_list = []
    hcm_list = []
    unknown_list = []
    for person in ok_list:
        loc = person['location'].lower()
        if loc in ['hà nội', 'ha noi']:
            hanoi_list.append(person)
        elif loc in ['đà nẵng', 'da nang']:
            danang_list.append(person)
        elif loc in ['thành phố hồ chí minh', 'hồ chí minh', 'ho chi minh']:
            hcm_list.append(person)
        else:
            unknown_list.append(person)
    
    # Filter and sort result by time submitting first quest
    def _first_quest_date_str(person):
        date = first_quest_date(person)
        return str(date) if date else 'z'  # Biggest string
        
    ok_list_by_time = list(ok_list)
    ok_list_by_time.sort(key=lambda p: _first_quest_date_str(p))
    hanoi_list_by_time = list(hanoi_list)
    hanoi_list_by_time.sort(key=lambda p: _first_quest_date_str(p))
    danang_list_by_time = list(danang_list)
    danang_list_by_time.sort(key=lambda p: _first_quest_date_str(p))
    hcm_list_by_time = list(hcm_list)
    hcm_list_by_time.sort(key=lambda p: _first_quest_date_str(p))
    
    INPUT_DATA['result'] = {
        'error': error_list,
        'ok': {
            'all': ok_list,
            'all_by_time': ok_list_by_time,
            'hanoi': hanoi_list,
            'hanoi_by_time': hanoi_list_by_time,
            'danang': danang_list,
            'danang_by_time': danang_list_by_time,
            'hcm': hcm_list,
            'hcm_by_time': hcm_list_by_time,
            'unknown': unknown_list,
        },
    }

    # Show final result on screen
    show_result_header(INPUT_DATA['result'])
    # Errors
    show_result_error(error_list)
    # Result all location/time
    show_result_by_loc('ALL LOCATION', ok_list)
    show_result_by_time('ALL LOCATION (BY TIME SUBMITTING FIRST QUEST)', ok_list_by_time)
    # Result by location/time
    show_result_by_loc('HÀ NỘI', hanoi_list)
    show_result_by_time('HÀ NỘI (BY TIME SUBMITTING FIRST QUEST)', hanoi_list_by_time)
    show_result_by_loc('ĐÀ NẴNG', danang_list)
    show_result_by_time('ĐÀ NẴNG (BY TIME SUBMITTING FIRST QUEST)', danang_list_by_time)
    show_result_by_loc('HỒ CHÍ MINH', hcm_list)
    show_result_by_time('HỒ CHÍ MINH (BY TIME SUBMITTING FIRST QUEST)', hcm_list_by_time)
    show_result_by_loc('UNKNOWN LOCATION', unknown_list)
    
    # Save result to text file
    save_result_txt()
    
    # Save result to excel file also
    save_result_excel()
    
    pp()
    pp('RESULT saved in', fg.cyan + 'result.txt' + fx.end, 'and',
       fg.cyan + 'result.xlsx' + fx.end)

def first_quest_date(person):
        quests = person['legal_quests']
        quest = quests[0] if len(quests) else None
        return quest['earned_date'] if quest else None
        
def show_result_header(result, outfile=None):
    if not outfile:
        pp()
        pp()
        pp(bg.lightyellow + fg.blue + fx.bold, 'FINAL RESULT', fx.end)
        pp()
    else:
        outfile.write('\nGDG - CLOUD STUDY JAMS RESULT\n')
        outfile.write('    Total participants: %d\n' % len(result['ok']['all']))
        outfile.write('        Hà Nội: %d\n' % len(result['ok']['hanoi']))
        outfile.write('        Đà Nẵng: %d\n' % len(result['ok']['danang']))
        outfile.write('        Hồ Chí Minh: %d\n' % len(result['ok']['hcm']))
        outfile.write('        Unknown location: %d\n' % len(result['ok']['unknown']))
        outfile.write('    Time period:\n')
        outfile.write('        From Date: %s\n' % str(QUEST_COUNT_FROM[0]))
        outfile.write('        To Date: %s\n' % str(QUEST_COUNT_FROM[1]))
        outfile.write('\n')

def show_result_error(plist, outfile=None):
    if not plist:
        return
        
    if not outfile:
        pp()
        pp(bg.lightred + fg.white + fx.bold, 'ERRORS ENCOUNTERED', fx.end)
        
        for person in plist:
            pp()
            pp_err(person['name'], '(' + fg.cyan + person['email'] + fx.end + ')',
                   fg.red, person['error'], fx.end)
    else:
        ordinal = 0
        outfile.write('\n' + 'ERRORS ENCOUNTERED' + '\n')
        for person in plist:
            ordinal += 1
            outfile.write('  %d. %s (%s) - %s\n' % (
                          ordinal, person['name'], person['email'],
                          person['error']))

def show_result_by_loc(title, plist, outfile=None):
    if not outfile:
        pp()
        pp()
        pp(bg.magenta + fg.white + fx.bold, title, fx.end)
        pp()
        
        ordinal = 0
        for person in plist:
            ordinal += 1
            pp()
            pp(bg.lightgreen + fg.yellow + fx.bold, 'QUEST BY LOC', fx.end,
               person['name'], '(' + fg.cyan + person['email'] + fx.end + ')')
            pp()
            pp('      ', bg.lightwhite+fg.lightblue, '%2d.' % ordinal, fx.end,
               bg.lightyellow+fg.green, '%3d LEGAL QUESTS' % len(person['legal_quests']), fx.end,
               bg.lightyellow+fg.green, '%3d TOTAL QUESTS' % len(person['quests']), fx.end)
    else:
        ordinal = 0
        outfile.write('\n' + title + '\n')
        for person in plist:
            ordinal += 1
            outfile.write('  %d. %s (%s) - %d legal quests (%d total)\n' % (
                          ordinal, person['name'], person['email'],
                          len(person['legal_quests']), len(person['quests'])))

def show_result_by_time(title, plist, outfile=None):
    if not outfile:
        pp()
        pp()
        pp(bg.magenta + fg.white + fx.bold, title, fx.end)
        pp()
        
        ordinal = 0
        for person in plist:
            ordinal += 1
            pp()
            pp(bg.lightgreen + fg.yellow + fx.bold, 'QUEST BY TIME', fx.end,
               person['name'], '(' + fg.cyan + person['email'] + fx.end + ')')
            pp()
            pp('      ', bg.lightwhite+fg.lightblue, '%2d.' % ordinal, fx.end,
               bg.lightyellow+fg.green, 'DATE SUBMITTED %s' % str(first_quest_date(person)), fx.end)
    else:
        ordinal = 0
        outfile.write('\n' + title + '\n')
        for person in plist:
            ordinal += 1
            outfile.write('  %d. %s (%s) - Time submitted %s\n' % (
                          ordinal, person['name'], person['email'],
                          str(first_quest_date(person))))

def save_result_txt():
    result = INPUT_DATA['result']
    error_list = result['error']
    ok_list = result['ok']['all']
    hanoi_list = result['ok']['hanoi']
    hanoi_list_by_time = result['ok']['hanoi_by_time']
    danang_list = result['ok']['danang']
    danang_list_by_time = result['ok']['danang_by_time']
    hcm_list = result['ok']['hcm']
    hcm_list_by_time = result['ok']['hcm_by_time']
    unknown_list = result['ok']['unknown']
    
    # Outfile for saving result
    with io.open('result.txt', 'w', encoding='utf-8') as outfile:
        # Header
        show_result_header(INPUT_DATA['result'], outfile=outfile)
        # Errors
        show_result_error(error_list, outfile=outfile)
        # Result all location
        show_result_by_loc('ALL LOCATION', ok_list, outfile=outfile)
        # Result by location/time
        show_result_by_loc('HÀ NỘI', hanoi_list, outfile=outfile)
        show_result_by_time('HÀ NỘI (BY TIME SUBMITTING FIRST QUEST)', hanoi_list_by_time, outfile=outfile)
        show_result_by_loc('ĐÀ NẴNG', danang_list, outfile=outfile)
        show_result_by_time('ĐÀ NẴNG (BY TIME SUBMITTING FIRST QUEST)', danang_list_by_time, outfile=outfile)
        show_result_by_loc('HỒ CHÍ MINH', hcm_list, outfile=outfile)
        show_result_by_time('HỒ CHÍ MINH (BY TIME SUBMITTING FIRST QUEST)', hcm_list_by_time, outfile=outfile)
        show_result_by_loc('UNKNOWN LOCATION', unknown_list, outfile=outfile)
    
def save_result_excel():
    wb = INPUT_DATA['excel']['workbook']
    sh = INPUT_DATA['excel']['worksheet']
    
    sh['%s1' % COL_RESULT_QUEST_COUNT] = 'LegalQuests'
    sh['%s1' % COL_RESULT_ALL] = 'ResultALL'
    sh['%s1' % COL_RESULT_HANOI] = 'Hà Nội'
    sh['%s1' % COL_RESULT_DANANG] = 'Đà Nẵng'
    sh['%s1' % COL_RESULT_HCM] = 'Hồ Chí minh'
    
    result = INPUT_DATA['result']['ok']
    result_all = result['all']
    for i in range(0, len(result_all)):
        person = result_all[i]
        sh['%s%d' % (COL_RESULT_QUEST_COUNT, person['row_id'])] = len(person['legal_quests'])
        sh['%s%d' % (COL_RESULT_ALL, person['row_id'])] = i+1
    result_hanoi = result['hanoi']
    for i in range(0, len(result_hanoi)):
        person = result_hanoi[i]
        sh['%s%d' % (COL_RESULT_HANOI, person['row_id'])] = i+1
    result_danang = result['danang']
    for i in range(0, len(result_danang)):
        person = result_danang[i]
        sh['%s%d' % (COL_RESULT_DANANG, person['row_id'])] = i+1
    result_hcm = result['hcm']
    for i in range(0, len(result_hcm)):
        person = result_hcm[i]
        sh['%s%d' % (COL_RESULT_HCM, person['row_id'])] = i+1
    # Save workbook
    wb.save(INPUT_FILE)

def count_quests_of(person):
    qwiklabs_link = person['qwiklabs_link']
    resp = requests.get(qwiklabs_link)
    if resp.status_code != 200:  # Not OK
        pp('UNABLE to load QUESTS report for user %s' % person['email'])
        person['error'] = 'UNABLE to load QUESTS report page'
    else:
        html = bs4.BeautifulSoup(resp.content, features="html.parser")
        div_all_quests = html.body.find_all('div', attrs={'class': 'public-profile__badge'})
        if not div_all_quests:
            pp_err('UNABLE to parse QUESTS report for user %s' % person['email'])
            person['error'] = 'UNABLE to parse QUESTS report (seems no quests at all)'
        else:
            quests_list = person['quests']
            # pp('Quest count = %d' % len(div_all_quests))
            for div in div_all_quests:
                child_tags = []
                for child in div.children:
                    if isinstance(child, bs4.element.Tag):
                        child_tags.append(child)
                if len(child_tags) != 3:
                    pp('UNEXPECTED quests report content')
                    person['error'] = 'UNEXPECTED quests report content'
                else:
                    title = child_tags[1].text.strip()
                    date_str = child_tags[2].text.strip().split('\n')[1]
                    date = datetime.datetime.strptime(date_str, '%b %d, %Y').date()
                    quest_info = {
                        'title': title,
                        'earned_date': date,
                    }
                    # pp(quest_info)
                    quests_list.append(quest_info)
        # pp(quests_list)
        show_quests_report_of(person)
        # Count legal quests
        legal_quests = []
        for quest in person['quests']:
            if (quest['title'] not in SKIP_QUESTS and
                QUEST_COUNT_FROM[0] <= quest['earned_date'] <= QUEST_COUNT_FROM[1]):
                legal_quests.append(quest)
        person['legal_quests'] = legal_quests
        

def show_quests_report_of(person):
    # Title line
    pp()
    pp(bg.lightgreen + fg.yellow + fx.bold, 'QUEST FOUND', fx.end,
       person['name'], '(' + fg.cyan + person['email'] + fx.end + ') -',
       str(len(person['quests'])), 'quests')
    # Show all quests
    if SHOW_QUESTS_DETAIL:
        est_line_len = 0
        quests_at_line = []
        for quest in person['quests']:
            title = quest['title']
            if len(title) + est_line_len <= MAX_LINE_LEN:
                est_line_len += len(title)
                quests_at_line.append(title)
            else:
                show_quests_at_line(person, quests_at_line)
                quests_at_line = [title]
                est_line_len = len(title)
        if len(quests_at_line):
            show_quests_at_line(person, quests_at_line)
        
def show_quests_at_line(person, quests_title):
    args = []
    for title in quests_title:
        bgc = random_bg()
        fgc = random_fg(bgc)
        args.append(getattr(bg, bgc) + getattr(fg, fgc))
        args.append(title)
        args.append(fx.end)
    pp()
    pp('      ', *args)

if __name__ == '__main__':
    main()
