from bs4 import BeautifulSoup
import sys
import xlwt

gpus = sys.argv[1]
name_list = []
lat_lng = []
lat_lng_2 = []
lat_lng_3 = []
soup = BeautifulSoup(open(str(gpus),'rb'),"xml")
for i in soup.Folder.find_all('Placemark'):
    name_list.append(i.contents[1].get_text()) ##获取区块名称
    lat_lng.append(i.coordinates.get_text().strip()) ##获取区块经纬度
for i in lat_lng:
    lat_lng_2.append(i.split(' '))

wb = xlwt.Workbook('utf-8')
sh = wb.add_sheet('kml')
sh.write(0,0,'name')
sh.write(0,1,'lng')
sh.write(0,2,'lat')
for i in range(len(lat_lng_2)):
    for j in range(len(lat_lng_2[i])):
        tmp = name_list[i]+','+lat_lng_2[i][j]
        lat_lng_3.append(tmp)
print(lat_lng_3)
for i in range(len(lat_lng_3)):
    sh.write(i+1,0,lat_lng_3[i].split(',')[0])
    sh.write(i+1,1,lat_lng_3[i].split(',')[1])
    sh.write(i+1,2,lat_lng_3[i].split(',')[2])
wb.save('kml.xls')