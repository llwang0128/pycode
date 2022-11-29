import requests
import xlwings

class Tmc:
    
    def __init__(self, filePath):
        try:
            self.app = xlwings.App(visible = False, add_book = False)
            self.book = self.app.books.open(filePath)
            self.__done = {}
        except:
            if self.book :
                self.book.close()
            if self.app :
                self.app.quit()
            print("Init error")
            raise IOError("Open excel file error")
 
    def url(self, sec_cd):
        pass
        if sec_cd[0] == '6':
            sec_cd = "s_sh" + sec_cd
        else:
            sec_cd = "s_sz" + sec_cd
        return "http://qt.gtimg.cn/q=" + sec_cd
    
    def fetch_tmc(self, code):
        tmc = -1.0
        if code in self.__done :
            tmc = self.__done[code]
        else :
            try :
                resp = requests.get(self.url(code), timeout=5)
                if resp.status_code == 200 :
                    if not resp.text.startswith("v_pv_none"):
                        tmc = float(resp.text.split("~")[-2])
                        self.__done[code] = tmc
            except :
                raise IOError("web job failed")
        print('fetch {} tmc Got {}'.format(code, tmc))
        return tmc

    def run(self):
        try:
            for i in range(self.book.sheets.count):
                sheet = self.book.sheets[i]
                for j in range(3, sheet.used_range.last_cell.row + 1):
                    r = '{}{}'.format('B', j)
                    code = '{0:0>6.0f}'.format(sheet.range(r).value)
                    w = '{}{}'.format('H', j)
                    sheet.range(w).value = self.fetch_tmc(code)
        finally:
            if self.book :
                self.book.save()
                self.book.close()
            if self.app :
                self.app.quit()


if __name__ == "__main__":
    file_path = r'C:\some\path\to\file.xlsx'
    try :
        Tmc(file_path).run()
        print("Job done")
    except :
        print("Sth error")
