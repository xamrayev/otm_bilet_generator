import g4f
from g4f.Provider import (
    Poe,
    You,
    GptGo,
    Bing,
    Aichat
)
def ai_savol(fan_nomi=str):
    '''
    fan_nomi=>str
    fan nomi bo'yicha 5 ta savolni generatsiya qiladi
    '''
    promt = f"ответь на русском. составь 5 вопросы для контрольной работы по предмету {fan_nomi}.  вопросы дай как - questions = ['1)','2)']"
    response = g4f.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": promt}],
        provider=g4f.Provider.Aichat,
        stream=True,
    )
    msg=""
    for message in response:
        # print(message, flush=True, end='')
        msg+=message
    msg = msg.split("\n")
    gen_savollar = []
    for i in msg:
        if i[:1]=="1" or i[:1]=="2" or i[:1]=="3" or i[:1]=="4" or i[:1]=="5":
            i=i[3:]
            gen_savollar.append(i)
    return gen_savollar


def savol_tuplam(bilet_soni=int, fan_nomi=str):
    """
    bilet_soni marta ai_savolni chaqiradi 
    """
    savollar = []
    schetchik = 0
    for i in range(bilet_soni):
        print(schetchik,"%")
        savol = ai_savol(fan_nomi)
        # print(savol)
        savollar.extend(savol)
        schetchik+=100/bilet_soni
    return savollar

# with open("file.txt", "r+") as file:
#     a = savol_tuplam(bilet_soni=10,  fan_nomi="органическая химия")
#     for i in a:
#         file.write(i)
#     file.close()
        

    
