from sympy import im
from to_docx import biletlar_ru
from gpt import savol_tuplam
# from temp import  savollar_from_xls 

#savol_file =  "savollar.xlsx"

bilet_soni = 10
fanimiz = "Общая и прикладная биотехнология"
savollarim = savol_tuplam(fan_nomi=fanimiz, bilet_soni=bilet_soni)
# savollarim = savollar_from_xls(savol_file)
semstr = str(5)
kaf =  "ИСиТ" 
tu = "Хамраева Д."
zav_ka = "Комилов С."

biletlar_ru(bilet_soni=bilet_soni, savollar=savollarim, fan=fanimiz, semestr=semstr, kafedra=kaf,  tuzuvchi  = tu, zav_kaf= zav_ka)

