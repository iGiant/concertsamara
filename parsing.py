def getafisha(isLog=True):
    def addtusa(str):
        temp = list(parsed_body.xpath(str))
        return '' if not temp else temp[0]

    def changequotes(text):
        if text[0]=='"':
            text='«'+text[1:]
        text = text.replace(' "', ' «')
        return text.replace('"', '»')

    def load_setting():
        with open('subscription.dat', 'r', encoding='utf-8') as file:
            subsriptions = file.read().split('\n')
        result = []
        for subsription in subsriptions:
            if subsription != '':
                mydict = {}
                parametrs = subsription.split(',')
                mydict['keys'] = list(parametrs[:-2])
                mydict['mail'] = parametrs[-2]
                mydict['count'] = parametrs[-1]
                result.append(mydict)
        return result


    def search(eventlist:list,subscription:list):
        def send_mail(mail:str, key:str, event:dict):
            import smtplib
            from email.mime.text import MIMEText
            text = """Сработал триггер на слово "<b>{}</b>"<br>
            Мероприятие <a href="{}">{}</a>  пройдет {} ({}) в {} в следующем месте: <i>{}</i>.<br><br>
            Билеты можно купить <a href="{}">здесь</a><br><br>
               {}""".format(key, event['url'], event['name'], event['date'], event['time'][:2], event['time'][3:],
                            event['place'], event['buy'], event['detail'], )
            msg = MIMEText(text, 'HTML', 'utf-8')
            msg['Subject'] = 'Культурные мероприятия Самары: сработал триггер {}'.format(key)
            msg['From'] = 'Concert@metrosamara.ru'
            msg['To'] = mail
            smtpObj = smtplib.SMTP('mail.sm', 587)
            smtpObj.ehlo()
            smtpObj.starttls()
            smtpObj.ehlo()
            smtpObj.login('services@metrosamara.ru', '123456')
            try:
                smtpObj.sendmail('Concert@metrosamara.ru', mail, msg.as_string())
            except:
                pass
            smtpObj.quit()

        for event in eventlist:
            for i, keys in enumerate(subscription):
                fbreak = False
                for key in keys['keys']:
                    if event['name'].lower().find(key.lower()) != -1 or event['detail'].lower().find(key.lower()) !=-1:
                        fbreak = True
                        send_mail(subscription[i]['mail'], key, event)
                        logger.info('Письмо на "{}", кодовое слово "{}"'.format(subscription[i]['mail'],key))
                        break;
                if fbreak:
                    if keys['count'] != '0':
                        subscription[i]['count'] = '-1' if keys['count'] == '1' else str(int(keys['count'])-1)
                    break
        with open('subscription.dat', 'w', encoding='utf-8') as file:
            for keys in subscription:
                if keys['count'] != '-1':
                    text = ','.join(keys['keys'])
                    file.write('{},{},{}\n'.format(text, keys['mail'], keys['count']))


    import requests
    from lxml import html
    import urllib.parse
    import logging

    logging.basicConfig(format='%(asctime)s - %(message)s', level=logging.INFO, filename=r'd:\python\logs\parsing.log')
    logger = logging.getLogger ( __name__ )
    logger.info('*'*25)
    http = 'http://koncertsamara.ru/afisha/'
    afisha = []
    page = 1
    while True:
        response = requests.get(http+'?a-page='+str(page-1))
        parsed_body = html.fromstring(response.text)
        if isLog:
            logger.info('Обрабатывается страница %d' % page)
        else:
            print('Обрабатывается страница %d' % page)
        for i in range (1, len(parsed_body.xpath('///*[@id="main"]/div/div/ul/li/div/div[2]/h3/text()'))+1):
            tusa = {}
            tusa['name'] = changequotes(addtusa('///*[@id="main"]/div/div/ul/li[%d]/div/div[2]/h3/text()' % i))
            tusa['date'] = addtusa('///*[@id="main"]/div/div/ul/li[%d]/div/div[1]/span[1]/text()' % i)
            tusa['time'] = addtusa('///*[@id="main"]/div/div/ul/li[%d]/div/div[1]/span[3]/text()' % i)
            tusa['place'] = changequotes(addtusa('///*[@id="main"]/div/div/ul/li[%d]/h4/a/text()' % i))
            tusa['url'] = addtusa('///*[@id="main"]/div/div/ul/li[%d]/div/div[4]/div/a[1]/@href' % i)
            tusa['buy'] = addtusa('///*[@id="main"]/div/div/ul/li[%d]/div/div[4]/div/a[2]/@href' % i)
            if addtusa('///*[@id="main"]/div/div/ul/li[%d]/div/div[4]/div/a[1]/text()' % i) in ['']:
                tusa['url'] = addtusa('///*[@id="main"]/div/div/ul/li[%d]/div/div[4]/div/a[2]/@href' % i)
                tusa['buy'] = addtusa('///*[@id="main"]/div/div/ul/li[%d]/div/div[4]/div/a[3]/@href' % i)
            tusa['url'] = urllib.parse.urljoin(http, tusa['url'])
            tusa['buy'] = urllib.parse.urljoin(http, tusa['buy'])
            response_detail = requests.get(tusa['url'])
            parsed_body_detail = html.fromstring(response_detail.text)
            tusa['detail'] = ''
            temp_detail = list(parsed_body_detail.xpath('//*[ @ id = "current-description"]/p/text()'))
            if temp_detail:
                for j in range(len(temp_detail)):
                    if len(tusa['detail']) < len(temp_detail[j]):
                        tusa['detail'] = temp_detail[j]
            afisha.append(tusa)
        response = requests.get(http + '?a-page=' + str(page - 1))
        parsed_body = html.fromstring(response.text)
        if parsed_body.xpath('//*[@id="main"]/div[2]/div[3]/ul/li/a/text()')[-1] == str(page): break
        page += 1
    search(afisha, load_setting())
    return afisha

def savetofile(afisha,file='koncert.xlsx'):
    import openpyxl
    from openpyxl.styles import fonts, alignment, Side, Border
    from openpyxl.styles.colors import COLOR_INDEX
    from openpyxl.comments import comments
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['Дата','Время','Событие (клик – подробно)','Место проведения (клик – бронирование)'])
    side = Side(style='thin', color=COLOR_INDEX[0])
    dside = Side(style='double', color=COLOR_INDEX[0])
    border = Border(left=side, right=side, top=side, bottom=side)
    hborder = Border(left=side, right=side, top=side, bottom=dside)
    for i, entry in enumerate(afisha):
        ws.append([afisha[i]['date'], afisha[i]['time'],
                   '=HYPERLINK("%s","%s")' % (afisha[i]['url'], afisha[i]['name']),
                   '=HYPERLINK("%s","%s")' % (afisha[i]['buy'], afisha[i]['place'])])
        if len(afisha[i]['detail']) > 10:
            ws['C' + str(i + 2)].comment =  comments.Comment(afisha[i]['detail'], '')
        for r in ('A', 'B', 'C', 'D'):
            ws[r+str(i+2)].border = border
            if r in ('A', 'B'):
                ws[r + str(i + 2)].alignment = alignment.Alignment(horizontal='center')
    for sym in ['A1', 'B1', 'C1', 'D1']:
        ws[sym].font = fonts.Font(size=12, bold=True)
        ws[sym].alignment = alignment.Alignment(horizontal='center')
        ws[sym].border = hborder
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 60
    ws.column_dimensions['D'].width = 60
    wb.save(file)

if __name__ == '__main__':
    import sys
    if len(sys.argv) > 1:
        if not sys.argv[1].endswith('.xlsx'):
            sys.argv[1]+='.xlsx'
        savetofile(getafisha(True), sys.argv[1])
    else:
        afisha = getafisha(False)
        for i in range(len(afisha)):
             print('%s - %s : %s (%s)' % (afisha[i]['date'], afisha[i]['time'], afisha[i]['name'], afisha[i]['place']))
        input()