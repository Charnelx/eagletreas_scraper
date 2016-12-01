import datetime
import requests
import re
import xlrd, xlsxwriter
from lxml import html
import argparse



class Spider:

    def __init__(self):
        self.base_url = 'https://eagletreas.mohavecounty.us/treasurer/treasurerweb/account.jsp'
        self.headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.116 Safari/537.36 OPR/40.0.2308.81",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Encoding": "gzip, deflate, lzma",
            "Connection": "keep-alive",
            "Connection-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Host": "eagletreas.mohavecounty.us",
            "Upgrade-Insecure-Requests": "1",
            "Accept-language": "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4",
        }

        self.pattern = re.compile(r'account=(\w+)&action=tx$', re.IGNORECASE)

    def __connect(self):
        r = requests.Session()

        resp = r.post('http://eagletreas.mohavecounty.us/treasurer/web/loginPOST.jsp', data={'guest': 'true',
                                                                                             'submit': 'I Have Read The Above Statement'})
        if resp.status_code == 200:
            return r

    def run(self, in_file, out_file='sample.xlsx'):

        self.con = self.__connect()

        rb = xlrd.open_workbook(in_file)
        sheet = rb.sheet_by_index(0)

        data_list = [sheet.row_values(rownum) for rownum in range(1, sheet.nrows)]

        data = self.__get_content(data_list)
        self.__write_to_file(data)


    def __get_content(self, data_list):
        data = dict()

        for item in data_list:
            try:
                url = item[2]
                id = re.search(self.pattern, url).group(1)

                payload = {
                    'account': id,
                    'action': 'tx',
                }
                resp = self.con.get(self.base_url, params=payload)

                root = html.fromstring(resp.content.decode())

                date = root.xpath('//table[@class="account"]/tbody/tr[last()]/td[1]/text()')[0]

                print('Working with ID: {0}'.format(id))
                data[id] = (item[0], date, item[2])
            except Exception as err:
                print('Error occurred with ID: {0}\n{1}'.format(item[0], err))

        return data


    def __write_to_file(self, data, out_file=None):

        if not out_file:
            out_file = datetime.datetime.now().strftime("%d.%m.%Y") + '_sample.xlsx'

        workbook = xlsxwriter.Workbook(out_file)
        worksheet = workbook.add_worksheet('sample')

        worksheet.set_column('A:A', 12)
        worksheet.set_column('B:B', 12)
        worksheet.set_column('C:C', 30)


        worksheet.write('A1', 'Account#')
        worksheet.write('B1', 'Earliest Summary Tax Year')
        worksheet.write('C1', 'Web Link')

        index = 2
        for key in data.keys():
            element = data[key]
            worksheet.write('A{0}'.format(index), element[0])
            worksheet.write('B{0}'.format(index), element[1])
            worksheet.write('C{0}'.format(index), element[2])
            index += 1

        workbook.close()



if __name__ == '__main__':
    parser = argparse.ArgumentParser()

    parser.add_argument('-i', '--input', help='Input filename to process')
    parser.add_argument('-o', help='Output filename to export')

    args = parser.parse_args()

    if not args.input:
        print('You mast provide input file name!')
        exit()

    obj = Spider()
    obj.run(args.input)