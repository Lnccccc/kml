from bs4 import BeautifulSoup
import xlwt
from progressbar import *
import os
class GetKml:
    def __init__(self):

        self.name_list = []
        self.lat_lng = []
        self.lat_lng_2 = []
        self. lat_lng_3 = []

    def Get_kml_raw(self,name):
        soup = BeautifulSoup(open('kml/'+str(name),'rb'),"xml")
        for i in soup.find_all('Placemark'):
            self.name_list.append(i.contents[1].get_text()) ##获取区块名称
            self.lat_lng.append(i.coordinates.get_text().strip()) ##获取区块经纬度
        print(self.name_list)
        for i in self.lat_lng:
            self.lat_lng_2.append(i.split(' '))
        for i in range(len(self.lat_lng_2)):
            for j in range(len(self.lat_lng_2[i])):
                tmp = self.name_list[i]+','+self.lat_lng_2[i][j]+','+str(j)
                self.lat_lng_3.append(tmp)
    def WriteExcel(self,name):
        wb = xlwt.Workbook('utf-8')
        sh = wb.add_sheet('kml')
        sh.write(0,0,'name')
        sh.write(0,1,'order')
        sh.write(0,2,'lng')
        sh.write(0,3,'lat')
        for i in range(len(self.lat_lng_3)):
            sh.write(i+1,0,self.lat_lng_3[i].split(',')[0])
            sh.write(i+1,1,int(self.lat_lng_3[i].split(',')[4])+1)
            sh.write(i+1,2,self.lat_lng_3[i].split(',')[1])
            sh.write(i+1,3,self.lat_lng_3[i].split(',')[2])
            wb.save(str(name)[0:-4]+'.xls')
        self.name_list = []
        self.lat_lng = []
        self.lat_lng_2 = []
        self.lat_lng_3 = []
getkml = GetKml()
dir_list = os.listdir('kml/')

process = ProgressBar()
for i in process(range(len(dir_list))):
     getkml.Get_kml_raw(dir_list[i])
     getkml.WriteExcel(dir_list[i])
