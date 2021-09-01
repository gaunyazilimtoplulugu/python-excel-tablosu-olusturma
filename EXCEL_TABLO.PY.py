import os
import xlsxwriter
workbook = xlsxwriter.Workbook('benim_dosyam.xlsx') # benim dosyam isminde bir excel dosyası oluşturuyoruz.
worksheet = workbook.add_worksheet()
print(""" 
***************************************************************
======> Excel'de dosya oluşturma programına hoşgeldiniz <======
***************************************************************
""") # karşılama yolladık
sutun = int(input('kaç sütun kullanacaksınız: ')) # sutun sayısını istedik
x,satırsayısı = 0,0
sutun_sırası_list,sutun_ismi_list,sutun_bas_harf_list = [],[],[] # 3 tane kullanacağımız boş liste oluşturduk
liste = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X']
for satırsayısı in liste: 
    zz = 1 
    sutun_bas_harf_list.append(satırsayısı)  
    sutun_sırası_list.append(f"{satırsayısı}{zz}")
    x += 1
    c = input(f"{x}.sütun:")
    sutun_ismi_list.append(f"{c} sütununu")
    worksheet.write(f"{str(sutun_sırası_list[x-1])}", f"{c.strip().upper()}")
    if x == sutun:
        break

kez = int(input("Kaç kişinin bilgisini gireceksiniz: "))
j,i = 2,0
for a in range(0,kez):
    for t in sutun_bas_harf_list:
        a = f"{str(sutun_bas_harf_list[i])}{j}"
        name = input(f"{j-1}. kişinin {sutun_ismi_list[i]} giriniz:")
        worksheet.write(f"{a}", f"{name.strip().title()}")
        i+=1
        if i == len(sutun_bas_harf_list):
            i=0       
    j+=1
workbook.close()

durum = input("Programı açmak için 'başlat' yazınız: ")
if durum == 'başlat':
    os.system("benim_dosyam.xlsx")
else:
    print("program bitti...")