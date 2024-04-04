import openpyxl
import random

def teslar_to_dict(filename=str):
    """
        file kiritilsa testlarni dict ga to'playdi
    """
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.active
    testlar = {}
    for row in sheet.iter_rows(min_row=2):
        question_number = row[0].value
        question_text = row[1].value
        answers = [cell.value for cell in row[2:]]
        testlar[question_number] = {
            "savol": question_text,
            "javob": answers,
        }
    return testlar

def savol_gen(variantlar_soni=int, testdagi_savollar_soni=int, bazada_testlar_soni=int):
    """
        variantlar_soni => kerakli variantlar soni
        bazada_testlar_soni => excel dagi test savollari soni
        testdagi_savollar_soni => Bir variantda nechta savol bo'lishi kerak
    """
    testlar = list(range(1,bazada_testlar_soni+1))
    savol_in_tests = []
    tayyor_variantlar = []
    for i in range(1,variantlar_soni+1):
        savol_in_tests = []
        while len(savol_in_tests)<testdagi_savollar_soni:
            savol = random.choice(testlar)
            if savol in savol_in_tests:
                continue
            else:
                savol_in_tests.append(savol)
        tayyor_variantlar.append(savol_in_tests)
    return tayyor_variantlar

def javob_gen(savollar):
    soni=len(savollar)
    """
        Test soniga ko'ra javoblarning random ko'rinishini generatsiya qiladi
    """
    variantlar = {}
    for i in range(1,soni+1):
        variantlar[i] = {
            "variant": random.sample(range(4), 4)
        }
    return variantlar


def javoblar_random_tayyor(testlar, variantlar):
    """
        javob_gen dagi variantlar va testlar_to_dict dan natijlarni olib javoblari almashgan test hosil qiladi
    """
    tayyor_test = {}
    for i in range(1, len(testlar)+1,1):
        a = variantlar[i]["variant"][0]
        b = variantlar[i]["variant"][1]
        c = variantlar[i]["variant"][2]
        d = variantlar[i]["variant"][3]
        test = testlar[i]
        test_answers_a = test["javob"][a]
        test_answers_b = test["javob"][b]
        test_answers_c = test["javob"][c]
        test_answers_d = test["javob"][d]
        tayyor_test[i] = {
            "savol": test["savol"],
            "javob": [test_answers_a, test_answers_b, test_answers_c, test_answers_d],
        }
    return tayyor_test


def savollar_tayyor(variantlar_soni=int, savollar_random=list, testdagi_savollar_soni=int, savollarim=dict):
    pechatga = []
    yigim = {}
    savollar = []
    for i in range(variantlar_soni):
        varianti = savollar_random[i]
        # print(savollar_random[i])
        for j in range(1,testdagi_savollar_soni):
            pechatga = {}
            for k in range(len(varianti)):
                a=savollarim[varianti[k]]["savol"]
                b=savollarim[varianti[k]]["javob"]
                pechatga[k+1]=({"savol":a, "javob":b})
        yigim[i]=pechatga
        savollar.append(yigim[i])
    return savollar

def tayyorlash(test_fayl=str, kerakli_variantlar_soni=int, variantda_test_soni=int):
    testlar = teslar_to_dict(test_fayl)
    variant = kerakli_variantlar_soni
    testdagi_savollar_soni = variantda_test_soni
    bazada_testlar_soni = len(testlar.keys())

    savollar_random = savol_gen(variantlar_soni=variant, testdagi_savollar_soni=testdagi_savollar_soni, bazada_testlar_soni=bazada_testlar_soni)

    a = savollar_tayyor(variantlar_soni=variant, savollar_random=savollar_random, testdagi_savollar_soni=testdagi_savollar_soni, savollarim=testlar)

    tayyor = {}

    for i in range(len(a)):
        variantlar = javob_gen(testlar)
        j = javoblar_random_tayyor(a[i],variantlar=variantlar)
        tayyor[i+1]=j

    return tayyor


# fayl="testlar.xlsx"
# kerakli_variantlar_soni=5
# variantda_test_soni=4


# print(tayyorlash(test_fayl=fayl,  kerakli_variantlar_soni=kerakli_variantlar_soni,variantda_test_soni=variantda_test_soni))
