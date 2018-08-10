from bs4 import BeautifulSoup
import sys
import xlwt
from progressbar import *
class GetKml:
    def __init__(self):
        self.progress = ProgressBar()
        self.gpus = sys.argv[1]
        self.name_list = []
        self.lat_lng = []
        self.lat_lng_2 = []
        self. lat_lng_3 = []
    def Get_kml_raw(self):
        soup = BeautifulSoup(open(str(self.gpus),'rb'),"xml")
        for i in soup.Folder.find_all('Placemark'):
            self.name_list.append(i.contents[1].get_text()) ##获取区块名称
            self.lat_lng.append(i.coordinates.get_text().strip()) ##获取区块经纬度
        for i in self.lat_lng:
            self.lat_lng_2.append(i.split(' '))
        for i in range(len(self.lat_lng_2)):
            for j in range(len(self.lat_lng_2[i])):
                tmp = self.name_list[i]+','+self.lat_lng_2[i][j]
                self.lat_lng_3.append(tmp)
    def WriteExcel(self):
        wb = xlwt.Workbook('utf-8')
        sh = wb.add_sheet('kml')
        sh.write(0,0,'name')
        sh.write(0,1,'lng')
        sh.write(0,2,'lat')
        for i in self.progress(range(len(self.lat_lng_3))):
            sh.write(i+1,0,self.lat_lng_3[i].split(',')[0])
            sh.write(i+1,1,self.lat_lng_3[i].split(',')[1])
            sh.write(i+1,2,self.lat_lng_3[i].split(',')[2])
            wb.save('kml.xls')

if __name__ == '__main__':
    kml = GetKml()
    kml.Get_kml_raw()
    kml.WriteExcel()
    print('Down')