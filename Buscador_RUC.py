from CAPTCHA_Final.CAPTCHA_object_detection import *
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from bs4 import BeautifulSoup
import time
from os import *
from PIL import Image
from io import BytesIO
import traceback
from webdriver_manager.chrome import ChromeDriverManager
pd.set_option('display.max_colwidth', 200)

############################################# Deprecated #############################################

RUC = []

URL = 'https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/frameCriterioBusqueda.jsp'
driver = webdriver.Chrome(ChromeDriverManager().install())
driver.get(URL)
path_img = r'C:\Users\Usuario\Desktop\Py\Machine Learning\Regression\imagenes_prueba'

TRAB = pd.DataFrame(columns=['Periodo', 'Num Trabajadores', ' Num Pensionistas', 'Num Prest_Servicios'])
EST = pd.DataFrame(columns=['Tipo de Establecimiento', 'Dirección', 'Actividad Económica'])
REP = pd.DataFrame(columns=['Documento', 'Nro.Documento', 'Nombre', 'Cargo', 'Desde'])
GEN = pd.DataFrame(columns=['Información', 'Data'])
UBI = pd.DataFrame(columns=['Región', 'Provincia', 'Distrito'])
ERRORES = []

def get_path_captcha(driver):

    ref_img = driver.find_element_by_tag_name('img')
    driver.execute_script("window.scrollTo(0,0);")

    location = ref_img.location
    size = ref_img.size
    png = driver.get_screenshot_as_png()
    try:
        with Image.open(BytesIO(png)) as img:
            left = location['x']
            top = location['y']
            right = location['x'] + size['width']
            bottom = location['y'] + size['height']
            img = img.crop((left, top, right, bottom))
            name = 'imagen_prueba.png'
            path_captcha = path.join(path_img, name)
            img.save(path_captcha)

        return path_captcha

    except Exception as e:
        print(traceback.format_exc())
        raise Exception(f'Error saving captcha image {e}')


def change_window(driver, index):
    windows = driver.window_handles
    if (index < len(windows)):
        driver.switch_to.window(windows[index])


def gen_info(driver):
    global GEN
    global UBI
    ubigeos = {
        'AMAZONAS': '01',
        'ANCASH': '02',
        'APURIMAC': '03',
        'AREQUIPA': '04',
        'AYACUCHO': '05',
        'CAJAMARCA': '06',
        'CALLAO': '07',
        'CUSCO': '08',
        'HUANCAVELICA': '09',
        'HUANUCO': '10',
        'ICA': '11',
        'JUNIN': '12',
        'LA LIBERTAD': '13',
        'LAMBAYEQUE': '14',
        'LIMA': '15',
        'LORETO': '16',
        'MADRE DE DIOS': '17',
        'MOQUEGUA': '18',
        'PASCO': '19',
        'PIURA': '20',
        'PUNO': '21',
        'SAN MARTIN': '22',
        'TACNA': '23',
        'TUMBES': '24',
        'UCAYALI': '25'
    }
    try:
        source = driver.page_source
        soup = BeautifulSoup(source,'lxml')
        table = soup.find(text='Número de RUC: ').find_parent('table')
        rows = table.findAll('tr')
        if GEN.empty:
            gen = []
            for row in rows:
                td = row.findAll('td')
                row = [i.text for i in td]
                gen.append(row)
            gen = pd.DataFrame(gen)
            gen = gen.iloc[1:10,0:2]
            gen.replace(r'\n              \n             ', '', regex=True, inplace=True)
            gen.replace(r'\n\n\n\n\n\n\n\n\n\n\n\n\n', '', regex=True, inplace=True)
            gen.replace(r'                                                                                                                                ', '@_', regex=True, inplace=True)
            gen.replace(r'                                                                                                                               ', '@_', regex=True, inplace=True)
            gen.replace(r'                                                                                                                              ', '@_', regex=True, inplace=True)
            gen.replace(r'                                                                                                                             ', '@_', regex=True, inplace=True)
            gen.replace(r'                                                                                                                            ', '@_', regex=True, inplace=True)
            gen.replace(r'                                                                                                                           ', '@_', regex=True, inplace=True)
            gen.replace(r'                                                                                                                          ', '@_', regex=True, inplace=True)
            gen.replace(r'                                                                                                                         ', '@_', regex=True, inplace=True)
            gen.replace(r'                                                                                                                        ', '@_', regex=True, inplace=True)
            gen.replace(r'                                                                                                                       ', '@_', regex=True, inplace=True)
            gen.replace(r'                                                                                                                      ', '@_', regex=True, inplace=True)
            gen.replace(r'                                                                                                                     ', '@_', regex=True, inplace=True)
            gen.replace(r'                                                                                                                    ', '@_', regex=True, inplace=True)
            gen.replace(r'                                                                                                                   ', '@_', regex=True, inplace=True)
            gen.replace(r'                                                                                                                  ', '@_', regex=True, inplace=True)
            gen.replace(r'                                                                                                                 ', '@_', regex=True, inplace=True)
            gen.replace(r'                                                                                                                ', '@_', regex=True, inplace=True)
            gen.replace(r'                                                                                                               ', '@_', regex=True, inplace=True)
            gen.replace(r'                                                                                                              ', '@_', regex=True, inplace=True)
            gen.replace(r'                                                                                                             ', '@_', regex=True, inplace=True)
            gen.replace(r'                                                                                                            ', '@_', regex=True, inplace=True)
            gen.replace(r'                                                                                                           ', '@_', regex=True, inplace=True)
            gen.replace(r'                                                                                                          ', '@_', regex=True, inplace=True)


            gen.columns = ['Información', 'Data']
            # gen.to_excel('error1.xlsx',index=False)
            split_1 = gen.apply(lambda x: x['Data'].split('@_-') if x['Información'] == 'Dirección del Domicilio Fiscal:' else ['na','na','na'], axis=1).tolist()
            split_1 = [i for i in split_1 if i != ['na', 'na', 'na']]
            split_1 = pd.DataFrame(split_1, columns=['Región', 'Provincia', 'Distrito'])
            split_2 = split_1.apply(lambda x: x['Región'].split(' '), axis=1).tolist()
            split_1['Región'] = split_1.apply(lambda x: split_2[-1][-1] if x['Región'] != 'na' else 'na', axis=1)
            split_1['Región'] = split_1.apply(lambda x: 'LA LIBERTAD' if x['Región'] == 'LIBERTAD' else x['Región'], axis=1)
            split_1['Región'] = split_1.apply(lambda x: 'MADRE DE DIOS' if x['Región'] == 'DIOS' else x['Región'], axis=1)
            split_1['Región'] = split_1.apply(lambda x: 'SAN MARTIN' if x['Región'] == 'MARTIN' else x['Región'], axis=1)
            split_1['Ubigeo'] = split_1.apply(lambda x: ubigeos[x['Región']], axis=1)
            split_1['RUC'] = ruc
            gen['RUC'] = ruc
            GEN = pd.concat([GEN, gen], axis=0)
            UBI = pd.concat([UBI, split_1], axis=0)
        else:
            gen2 = []
            for row in rows:
                td = row.findAll('td')
                row = [i.text for i in td]
                gen2.append(row)
            gen2 = pd.DataFrame(gen2)
            gen2 = gen2.iloc[1:10,0:2]
            gen2.replace(r'\n              \n             ', '', regex=True, inplace=True)
            gen2.replace(r'\n\n\n\n\n\n\n\n\n\n\n\n\n', '', regex=True, inplace=True)
            gen2.replace(r'                                                                                                                                ', '@_', regex=True, inplace=True)
            gen2.replace(r'                                                                                                                               ', '@_', regex=True, inplace=True)
            gen2.replace(r'                                                                                                                              ', '@_', regex=True, inplace=True)
            gen2.replace(r'                                                                                                                             ', '@_', regex=True, inplace=True)
            gen2.replace(r'                                                                                                                            ', '@_', regex=True, inplace=True)
            gen2.replace(r'                                                                                                                           ', '@_', regex=True, inplace=True)
            gen2.replace(r'                                                                                                                          ', '@_', regex=True, inplace=True)
            gen2.replace(r'                                                                                                                         ', '@_', regex=True, inplace=True)
            gen2.replace(r'                                                                                                                        ', '@_', regex=True, inplace=True)
            gen2.replace(r'                                                                                                                       ', '@_', regex=True, inplace=True)
            gen2.replace(r'                                                                                                                      ', '@_', regex=True, inplace=True)
            gen2.replace(r'                                                                                                                     ', '@_', regex=True, inplace=True)
            gen2.replace(r'                                                                                                                    ', '@_', regex=True, inplace=True)
            gen2.replace(r'                                                                                                                   ', '@_', regex=True, inplace=True)
            gen2.replace(r'                                                                                                                  ', '@_', regex=True, inplace=True)
            gen2.replace(r'                                                                                                                 ', '@_', regex=True, inplace=True)
            gen2.replace(r'                                                                                                                ', '@_', regex=True, inplace=True)
            gen2.replace(r'                                                                                                               ', '@_', regex=True, inplace=True)
            gen2.replace(r'                                                                                                              ', '@_', regex=True, inplace=True)
            gen2.replace(r'                                                                                                             ', '@_', regex=True, inplace=True)
            gen2.replace(r'                                                                                                            ', '@_', regex=True, inplace=True)
            gen2.replace(r'                                                                                                           ', '@_', regex=True, inplace=True)
            gen2.replace(r'                                                                                                          ', '@_', regex=True, inplace=True)

            gen2.columns = ['Información', 'Data']
            split_1 = gen2.apply(lambda x: x['Data'].split('@_-') if x['Información'] == 'Dirección del Domicilio Fiscal:' else ['na','na','na'], axis=1).tolist()
            split_1 = [i for i in split_1 if i != ['na', 'na', 'na']]
            split_1 = pd.DataFrame(split_1, columns=['Región', 'Provincia', 'Distrito'])
            split_2 = split_1.apply(lambda x: x['Región'].split(' '), axis=1).tolist()
            split_1['Región'] = split_1.apply(lambda x: split_2[-1][-1] if x['Región'] != 'na' else 'na', axis=1)
            split_1['Región'] = split_1.apply(lambda x: 'LA LIBERTAD' if x['Región'] == 'LIBERTAD' else x['Región'], axis=1)
            split_1['Región'] = split_1.apply(lambda x: 'MADRE DE DIOS' if x['Región'] == 'DIOS' else x['Región'], axis=1)
            split_1['Región'] = split_1.apply(lambda x: 'SAN MARTIN' if x['Región'] == 'MARTIN' else x['Región'], axis=1)
            split_1['Ubigeo'] = split_1.apply(lambda x: ubigeos[x['Región']], axis=1)
            split_1['RUC'] = ruc
            gen2['RUC'] = ruc
            GEN = pd.concat([GEN, gen2], axis=0)
            UBI = pd.concat([UBI, split_1], axis=0)
    except Exception as e:
        print(e)
        print('No hay Información para el ruc: {}'.format(ruc))


def trabajadores(driver):
    global TRAB
    try:
        bot_trab = WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, '//*[@id="div_estrep"]/table/tbody/tr[1]/td[4]/form/input[1]')))
        bot_trab.click()
        source = driver.page_source
        soup = BeautifulSoup(source,'lxml')
        table = soup.find('table',{'class': 'vista'})
        rows = table.findAll('tr')
        if TRAB.empty:
            trab = []
            for row in rows:
                td = row.findAll('td')
                row = [i.text for i in td]
                trab.append(row)
            trab = pd.DataFrame(trab)
            trab = trab.iloc[:,0:4]
            trab.drop(trab.index[0:3], inplace=True)
            trab.columns=['Periodo', 'Num Trabajadores', ' Num Pensionistas', 'Num Prest_Servicios']
            trab['RUC'] = ruc
            trab.replace('NE','0', inplace=True)
            TRAB = pd.concat([TRAB,trab], axis=0)
        else:
            trab2 = []
            for row in rows:
                td = row.findAll('td')
                row = [i.text for i in td]
                trab2.append(row)
            trab2 = pd.DataFrame(trab2)
            trab2 = trab2.iloc[:,0:4]
            trab2.drop(trab2.index[0:3], inplace=True)
            trab2.columns = ['Periodo', 'Num Trabajadores', ' Num Pensionistas', 'Num Prest_Servicios']
            trab2['RUC'] = ruc
            trab2.replace('NE', '0', inplace=True)
            TRAB = pd.concat([TRAB,trab2],axis=0)
        retornar_trab = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH,
                                                                                       '/html/body/table[1]/tbody/tr[1]/td/div/input')))
        retornar_trab.click()
    except Exception as e:
        print('El proveedor {} no reporta trabajadores'.format(ruc))


def establecimientos(driver):
    global EST
    try:
        bot_est = WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH,
                                                                                '//*[@id="div_estrep"]/table/tbody/tr[3]/td[2]/form/input[1]')))
        bot_est.click()
        source = driver.page_source
        soup = BeautifulSoup(source, 'lxml')
        table = soup.find(text='Código').find_parent('table')
        rows = table.findAll('tr')
        if EST.empty:
            est = []
            for row in rows:
                td = row.findAll('td')
                row = [i.text for i in td]
                est.append(row)
            est = pd.DataFrame(est)
            est = est.iloc[:,1:]
            est.drop(est.index[0], inplace = True)
            est.columns = ['Tipo de Establecimiento', 'Dirección', 'Actividad Económica']
            est['RUC'] = ruc
            est.replace(r'\n          ','', regex=True, inplace=True)
            est.replace(r'\t','', regex=True, inplace=True)
            EST = pd.concat([EST,est], axis=0)
        else:
            est2 = []
            for row in rows:
                td = row.findAll('td')
                row = [i.text for i in td]
                est2.append(row)
            est2 = pd.DataFrame(est2)
            est2 = est2.iloc[:, 1:]
            est2.drop(est2.index[0], inplace=True)
            est2.columns =['Tipo de Establecimiento', 'Dirección', 'Actividad Económica']
            est2['RUC'] = ruc
            est2.replace(r'\n          ', '', regex=True, inplace=True)
            est2.replace(r'\t', '', regex=True, inplace=True)
            EST = pd.concat([EST, est2], axis=0)
        retornar_est = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH,
                                                                                      '//*[@id="div_estrep"]/table/tbody/tr[1]/td/div/form/input')))
        retornar_est.click()
    except Exception as e:
        print('El proveedor {} no tiene establecimientos'.format(ruc))


def representantes(driver):
    global REP
    try:
        bot_rep = WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH,'//*[@id="div_estrep"]/table/tbody/tr[3]/td[1]/form/input[1]')))
        bot_rep.click()
        source = driver.page_source
        soup = BeautifulSoup(source, 'lxml')
        table = soup.find(text='Cargo').find_parent('table')
        rows = table.findAll('tr')
        if REP.empty:
            rep = []
            for row in rows:
                td = row.findAll('td')
                row = [i.text for i in td]
                rep.append(row)
            rep = pd.DataFrame(rep)
            rep.drop(rep.index[0], inplace=True)
            rep.replace(r'\n               ', '', regex=True, inplace=True)
            rep.columns = ['Documento', 'Nro.Documento', 'Nombre', 'Cargo', 'Desde']
            rep['RUC'] = ruc
            REP = pd.concat([REP, rep])
        else:
            rep2 = []
            for row in rows:
                td = row.findAll('td')
                row = [i.text for i in td]
                rep2.append(row)
            rep2 = pd.DataFrame(rep2)
            rep2.drop(rep2.index[0], inplace=True)
            rep2.replace(r'\n               ', '', regex=True, inplace=True)
            rep2.columns = ['Documento', 'Nro.Documento', 'Nombre', 'Cargo', 'Desde']
            rep2['RUC'] = ruc
            REP = pd.concat([REP, rep2])
        retornar_rep = WebDriverWait(driver,5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="div_estrep"]/table/tbody/tr[1]/td/div/input')))
        retornar_rep.click()
    except Exception as e:
        print('El proveedor {} no reporta representantes legales'.format(ruc))


def extraer(driver):
    ingresar_ruc = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH,
                                                                                   '//*[@id="s1"]/input')))
    ingresar_ruc.send_keys(ruc)
    obt_img = get_path_captcha(driver)
    CAPTCHA = Captcha_detection(obt_img)
    time.sleep(1)

    if len(CAPTCHA) == 4: #A veces cuando hay dos letras repetidas juntas, el código sólo envía una.
        ingresar_captcha = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH,
                                                                                           '/html/body/form/table/tbody/tr/td/table[2]/tbody/tr[1]/td[6]/input')))
        ingresar_captcha.send_keys(CAPTCHA)
        time.sleep(1)
        buscar = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH,
                                                                                 '/html/body/form/table/tbody/tr/td/table[2]/tbody/tr[1]/td[7]/input')))
        buscar.click()
    else:
        refresh_captcha = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/form/table/tbody/tr/td/table[2]/tbody/tr[2]/td[4]/a')))
        refresh_captcha.click()
        obt_img = get_path_captcha(driver)
        CAPTCHA = Captcha_detection(obt_img)
        ingresar_captcha = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH,
                                                                                          '/html/body/form/table/tbody/tr/td/table[2]/tbody/tr[1]/td[6]/input')))
        ingresar_captcha.send_keys(CAPTCHA)
        time.sleep(1)
        buscar = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH,
                                                                                '/html/body/form/table/tbody/tr/td/table[2]/tbody/tr[1]/td[7]/input')))
        buscar.click()

    change_window(driver, 1)
    try:
        error = driver.find_element_by_xpath('/html/body/div/p')
        if len(error)>0:
            change_window(driver, 0)
            refresh_captcha = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, '/html/body/form/table/tbody/tr/td/table[2]/tbody/tr[2]/td[4]/a')))
            refresh_captcha.click()
            obt_img = get_path_captcha(driver)
            CAPTCHA = Captcha_detection(obt_img)
            ingresar_captcha = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH,
                                                                                               '/html/body/form/table/tbody/tr/td/table[2]/tbody/tr[1]/td[6]/input')))
            ingresar_captcha.send_keys(CAPTCHA)
            buscar = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH,
                                                                                     '/html/body/form/table/tbody/tr/td/table[2]/tbody/tr[1]/td[7]/input')))
            buscar.click()
            change_window(driver, 1)

            gen_info(driver)
            trabajadores(driver)
            establecimientos(driver)
            representantes(driver)

            change_window(driver, 0)
            time.sleep(1)
            driver.find_element_by_tag_name('body').send_keys(Keys.COMMAND + 'r')
            refresh_captcha = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, '/html/body/form/table/tbody/tr/td/table[2]/tbody/tr[2]/td[4]/a')))
            refresh_captcha.click()
            ingresar_ruc.clear()
    except Exception as e:
        print("Correct CAPTCHA")

        gen_info(driver)
        trabajadores(driver)
        establecimientos(driver)
        representantes(driver)

        change_window(driver, 0)
        time.sleep(1)
        driver.find_element_by_tag_name('body').send_keys(Keys.COMMAND + 'r')
        refresh_captcha = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/form/table/tbody/tr/td/table[2]/tbody/tr[2]/td[4]/a')))
        refresh_captcha.click()
        ingresar_ruc.clear()



for ruc in RUC:
    try:
        extraer(driver)
    except Exception as e:
        print('No se encontro el ruc: {}'.format(ruc))
        ERRORES.append(ruc)
        continue
ERRORES = pd.DataFrame(ERRORES, columns=['RUC'])
driver.quit()

# GEN.to_excel('Info.xlsx', index=False)
#UBI.to_excel('Ubigeos2.xlsx', index=False)
# TRAB.to_excel('Trabajadores.xlsx', index=False)
# EST.to_excel('Establecimientos.xlsx', index=False)
# REP.to_excel('Representantes.xlsx', index=False)
# ERRORES.to_excel('Errores_sunat.xlsx', index=False)
