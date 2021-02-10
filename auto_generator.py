import logging as log
import os
import PyPDF2
import re
from decouple import config
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from time import sleep
from tkinter import Tk
from tkinter.filedialog import askopenfilename


def my_counter(matches: dict, fields: dict, values: dict) -> dict:
    for count, field in enumerate(fields):
        field = float(matches[fields[field]].replace(',', ''))
        keys = list(fields.keys())
        values[keys[count]] += field
    return values


def scan_axa_info(text: str = None) -> dict:
    """Axa Company provides information by email, quantities are updated in .txt file from ./file folder
    manually before running script."""
    try:
        pattern = re.compile(r"(Vida|No Vida|Acreditado|I.S.R.|Retenido)\s*:\s*\-*\$*(\d*\,?\d*.\d*)")
        matches = dict(re.findall(pattern, text))
    except Exception as err:
        log.error(f'Error while scanning Axa.txt file: {err}')
    return matches


def axa_process() -> dict:
    axa_content = open('./files/Axa.txt').read()
    # print(f'Content:{axa_content}')
    matches = scan_axa_info(axa_content)
    # print(f'Matches from Axa.txt: {matches}')
    axa_fields = {
        'damage': 'No Vida',
        'life': 'Vida',
        'iva_tras': 'Acreditado',
        'isr': 'I.S.R.',
        'iva_ret': 'Retenido'
    }
    empty_fields = {
        'damage': 0.00,
        'life': 0.00,
        'iva_tras': 0.00,
        'isr': 0.00,
        'iva_ret': 0.00
    }
    values = my_counter(matches, axa_fields, empty_fields)
    log.info(' | Facturando para Axa Seguros | ')
    log.info(f'Comisión de Daños: {values["damage"]}')
    log.info(f'Comisión de Vida: {values["life"]}')
    log.info(f'IVA Trasladado/Acreditado: {0.16 * values["damage"]}')
    log.info(f'IVA Retenido: {values["iva_ret"]}')
    log.info(f'ISR: {values["isr"]}')
    if values['damage'] != 0:
        log.info(f'Tasa de IVA Ret.: {round(values["iva_ret"] / values["damage"], 7)}')
        log.info(f'Tasa de ISR Daños: {round(values["isr"] / values["damage"], 7)}')
    else:
        log.info('Tasa de IVA Ret.: 0.00')
        log.info('Tasa de IVA Ret.: 0.00')
    log.info(f'Tasa de ISR Vida: {round(values["isr"] / values["life"], 7)}')
    log.info(f'Subtotal: {values["damage"] + values["life"]}')
    log.info(f'Total Impuestos Trasladados: {values["iva_tras"]}')
    log.info(f'Total Impuestos Retenidos: {values["isr"] + values["iva_ret"]}')
    log.info(f'Total: {values["life"] + values["damage"] + values["iva_tras"] - values["iva_ret"] - values["isr"]}')
    return values


def scan_qualitas_info(pdf: int, page: int = 0) -> [dict, int]:
    """Qualitas Company provides personal .pdf files with the neccesary information.
    In ./files folder you could find three .pdf samples to evaluate functionality, selection
    of files is meant to be user interaction."""
    log.info(f'PDF #{pdf+1} loading...')
    Tk().withdraw() # tkinter GUI to select files in order to be scanned
    pdf_file_obj = open(askopenfilename(), 'rb')  # creating a pdf file object
    pdf_reader = PyPDF2.PdfFileReader(pdf_file_obj)  # creating a pdf reader object
    num_pages = pdf_reader.numPages
    page_obj = pdf_reader.getPage(page)  # creating a page object
    pattern = re.compile(r"(IMPORTE|I.V.A.|TOTAL|I.S.R.|LEY|NETAS)\s+:(\d*\,?\d*.\d*)")
    text = page_obj.extractText()  # extracting text from page
    # print(f'\nVisualization of PDF´s text: {text}')
    pdf_file_obj.close()  # closing the pdf file object
    matches = dict(re.findall(pattern, text))
    # print(f'\nExtracted text from PDF: {matches}')
    if matches == dict():
        matches = {
            'import': 0.00,
            'iva': 0.00,
            'total': 0.00,
            'isr': 0.00,
            'iva_ret': 0.00,
            'commissions': 0.00
        }
    print("\nFile scanned successfully!")
    return matches, num_pages


def qualitas_process() -> dict:
    num_pdfs = int(input("\nNumber of pdf files to scan? "))
    qualitas_fields = {
        'import': 'IMPORTE',
        'iva': 'I.V.A.',
        'total': 'TOTAL',
        'isr': 'I.S.R.',
        'iva_ret': 'LEY',
        'commissions': 'NETAS'
    }
    empty_fields = {
        'import': 0.00,
        'iva': 0.00,
        'total': 0.00,
        'isr': 0.00,
        'iva_ret': 0.00,
        'commissions': 0.00
    }
    for pdf in range(num_pdfs):
        matches, num_pages = scan_qualitas_info(pdf)
        first_values = my_counter(matches, qualitas_fields, empty_fields)
        values = first_values
        if num_pages >= 2:
            log.info(f'Loading next page in PDF: ')
            log.info(f'Please select the file again to load next page')
            matches, num_pages = scan_qualitas_info(pdf, 1)
            second_values = my_counter(matches, qualitas_fields, first_values)
            values = second_values
            if num_pages >= 3:
                log.info(f'Loading next page in PDF: ')
                log.info(f'Please select the file again to load next page')
                matches, num_pages = scan_qualitas_info(pdf, 2)
                third_values = my_counter(matches, qualitas_fields, second_values)
                values = third_values
    log.info(' | Facturando para Quálitas |')
    log.info(f'Valor Unitario: {round(values["import"], 2)}')
    log.info(f'ISR: {round(values["isr"], 2)}')
    log.info(f'Tasa de ISR: {round(values["isr"] / values["import"], 6)}')
    log.info(f'Ret. I.V.A. Segun Ley: {round(values["iva_ret"], 2)}')
    log.info(f'Tasa de IVA Ret.: {round(values["iva_ret"] / values["import"], 6)}')
    log.info(f'IVA Trasladado: {0.16 * values["import"]}')
    log.info(f'Subtotal: {values["total"]}')
    log.info(f'Total Impuestos Trasladados: {values["iva"]}')
    log.info(f'Total Impuestos Retenidos: {values["isr"] + values["iva_ret"]}')
    log.info(f'Comisión neta: {values["commissions"]}')
    return values


def scan_potosi_info(text: str = None) -> [dict, str]:
    """Seguros el Potosi company provides information by agents portal, quantities have to be
    updated in .txt file from ./file folder manually before running script."""
    try:
        print("Available Tax Groups:\nD) Damage\nL) Life")
        tax_group = input("Tax Group to make bill?: ")
        if tax_group in ['D', 'd', 'Damage', 'damage']:
            pattern = re.compile(r"(D_Subtotal|D_I.V.A.|D_I.V.A. Ret.|D_I.S.R.|D_Total)\s*:\s*\$*(\d*?\d*.\d*)")
        else:
            pattern = re.compile(r"(V_Subtotal|V_I.V.A.|V_I.V.A. Ret.|V_I.S.R.|V_Total)\s*:\s*\$*(\d*?\d*.\d*)")
        matches = dict(re.findall(pattern, text))
    except Exception as err:
        log.error(f'Error while scanning Axa.txt file: {err}')
    return matches, tax_group


def potosi_process() -> [dict, str]:
    potosi_content = open("./files/SP.txt").read()
    matches, tax_group = scan_potosi_info(potosi_content)
    if tax_group in ['D', 'd', 'Damage', 'damage']:
        potosi_fields = {
            'subtotal': 'D_Subtotal',
            'iva': 'D_I.V.A.',
            'iva_ret': 'D_I.V.A. Ret.',
            'isr': 'D_I.S.R.',
            'total': 'D_Total'
        }
        group = 'Daños'
    else:
        potosi_fields = {
            'subtotal': 'V_Subtotal',
            'iva': 'V_I.V.A.',
            'iva_ret': 'V_I.V.A. Ret.',
            'isr': 'V_I.S.R.',
            'total': 'V_Total'
        }
        group = 'Vida'
    empty_fields = {
        'subtotal': 0.00,
        'iva': 0.00,
        'iva_ret': 0.00,
        'isr': 0.00,
        'total': 0.00
    }
    values = my_counter(matches, potosi_fields, empty_fields)
    log.info('| Facturando para Seguros el Potosí |')
    log.info(f'Comisión de {group}: {values["subtotal"]}')
    log.info(f'IVA Trasladado/Acreditado: {values["iva"]}')
    log.info(f'IVA Retenido: {values["iva_ret"]}')
    log.info(f'ISR: {values["isr"]}')
    if values['subtotal'] != 0:
        log.info(f'Tasa de IVA Ret.: {round(values["iva_ret"] / values["subtotal"], 7)}')
        log.info(f'Tasa de ISR {group}: {round(values["isr"] / values["subtotal"], 7)}')
    else:
        log.info('Tasa de IVA Ret.: 0.00')
        log.info('Tasa de IVA Ret.: 0.00')
    log.info(f'Subtotal: {values["subtotal"]}')
    log.info(f'Total Impuestos Trasladados: {values["iva"]}')
    log.info(f'Total Impuestos Retenidos: {values["isr"] + values["iva_ret"]}')
    log.info(f'Total: {values["subtotal"] + values["iva"] - values["iva_ret"] - values["isr"]}')
    return values, group


def automation_in_sat_webpage(company, date=None, damage=0.0, life=0.0, isr=0.0, iva_ret=0.0):
    # Driver access
    gc = webdriver.Chrome('./files/chromedriver')
    gc.maximize_window()
    gc.get('https://portalcfdi.facturaelectronica.sat.gob.mx/')

    # Login
    btn_e_firma = gc.find_element_by_id('buttonFiel').click()
    btn_cer = gc.find_element_by_id('btnCertificate').click()
    btn_key = gc.find_element_by_id('btnPrivateKey').click()
    sleep(15)
    password = gc.find_element_by_id('privateKeyPassword').send_keys(config('sat_password'))
    btn_send = gc.find_element_by_id('submit').click()

    # Accessing to CFDI generator
    cfdi = gc.find_element_by_xpath(
        "//a[@href='https://pacsat.facturaelectronica.sat.gob.mx/Comprobante/CapturarComprobante']").click()

    # Selecting client
    client = Select(gc.find_element_by_name('Receptor.RfcCargado'))
    if company in ['SP', 'sp', 'seguros el potosi', 'Seguros el Potosi', 'Potosi', 'potosi']:
        client.select_by_value('SPO830427DQ1 SEGUROS EL POTOSI, S.A.')
    elif company in ['Q', 'q', 'Qualitas', 'qualitas']:
        client.select_by_value('QCS931209G49 QUALITAS COMPAÑIA DE SEGUROS SA DE CV')
    elif company in ['A', 'a', 'Axa', 'axa']:
        client.select_by_value('ASE931116231 AXA SEGUROS SA DE CV')
    else:
        log.error('Error, compañía no registrada.')
    sleep(3)
    use = Select(gc.find_element_by_name('Receptor.UsoCFDIMoral'))
    use.select_by_value('G03 Gastos en general')
    sleep(3)
    next_tab = gc.find_element_by_xpath("//button[contains(@onclick,'clickTab')]").click()

    # Pay condition selection
    if company in ['SP', 'sp', 'seguros el potosi', 'Seguros el Potosi', 'Potosi', 'potosi']:
        pay_condition = gc.find_element_by_id('CondicionesDePago').send_keys('Al contado')
    else:
        pay_condition = gc.find_element_by_id('CondicionesDePago').send_keys('En una sola exhibición')

    if damage != 0.0:
        # Damages Section
        new_concept = gc.find_element_by_id('btnMuestraConcepto').click()
        service = gc.find_element_by_id('ClaveProdServ').send_keys('80141600 Actividades de ventas y promoción de negocios')
        sleep(1)
        if company in ['Q', 'q', 'Qualitas', 'qualitas']:
            description = gc.find_element_by_id('Descripcion').send_keys(' ' + date + ' Agente 05886')
        elif company in ['A', 'a', 'Axa', 'axa']:
            description = gc.find_element_by_id('Descripcion').send_keys(' ' + date + ' Agente 124109')
        else:
            description = gc.find_element_by_id('Descripcion').send_keys(' ' + date)
        gc.find_element_by_id('ValorUnitario').clear()
        gc.find_element_by_id('ValorUnitario').send_keys(str(damage))

        # I.S.R.
        taxes_damage = gc.find_element_by_xpath("//a[contains(@onclick,'deshabilitaBotonAceptar')]").click()
        sleep(1)
        edit_isr = gc.find_element_by_xpath("//*[@id='tablaConRetenciones']/tbody/tr[1]/td[6]/span[1]").click()

        # gc.find_element_by_id('Retenciones_Base').clear()
        # base = gc.find_element_by_id('Retenciones_Base').send_keys(str(damage+life))

        tax_type = Select(gc.find_element_by_xpath('//*[@id="Retenciones_Impuesto"]'))
        tax_type.select_by_value('001 ISR')

        gc.find_element_by_id('Retenciones_TasaOCuota').clear()
        percentage_isr = gc.find_element_by_id('Retenciones_TasaOCuota').send_keys(str(isr / damage))
        refresh_isr = gc.find_element_by_id('btnAgregaConImpuestoRetenido').click()

        # Retained I.V.A.
        edit_iva = gc.find_element_by_xpath("//*[@id='tablaConRetenciones']/tbody/tr[2]/td[6]/span[1]").click()
        tax_type = Select(gc.find_element_by_xpath('//*[@id="Retenciones_Impuesto"]'))
        tax_type.select_by_value('002 IVA')

        gc.find_element_by_id('Retenciones_TasaOCuota').clear()
        percentage_iva = gc.find_element_by_id('Retenciones_TasaOCuota').send_keys(str(iva_ret / damage))
        refresh_iva = gc.find_element_by_id('btnAgregaConImpuestoRetenido').click()

        # Finish edition of Damages Section
        move_to_concept = gc.find_element_by_xpath("//a[@id='tabConceptosPrincipal']")
        gc.execute_script("arguments[0].click();", move_to_concept)
        sleep(1)
        add_concept = gc.find_element_by_xpath("//div[@id='tabConceptos']//button[@id='btnAceptarModal']").click()
        sleep(3)
    else:
        log.error('There is no information in Damage Section.\nEvaluating Life Section')

    if company in ['SP', 'sp', 'seguros el potosi', 'Seguros el Potosi', 'Potosi', 'potosi', 'A', 'a', 'Axa', 'axa']:
        if life != 0.0:
            # Life Section
            gc.execute_script("scrollBy(0,-1000);")
            sleep(1)
            new_concept2 = gc.find_element_by_id('btnMuestraConcepto').click()
            service2 = gc.find_element_by_id('ClaveProdServ').send_keys('80141601 Servicios de promoción de ventas')
            sleep(1)
            if company in ['A', 'a', 'Axa', 'axa']:
                description = gc.find_element_by_id('Descripcion').send_keys(' ' + date + ' Agente 124109')
            else:
                description = gc.find_element_by_id('Descripcion').send_keys(' ' + date)
            gc.find_element_by_id('ValorUnitario').clear()
            gc.find_element_by_id('ValorUnitario').send_keys(str(life))

            if damage == 0.0:
                # I.S.R.
                taxes_life = gc.find_element_by_xpath("//*[@id='tabsConcepto']/li[2]/a").click()
                sleep(1)
                edit_isr2 = gc.find_element_by_xpath("//*[@id='tablaConRetenciones']/tbody/tr/td[6]/span[1]").click()
                gc.find_element_by_id('Retenciones_TasaOCuota').clear()
                percentage_isr2 = gc.find_element_by_id('Retenciones_TasaOCuota').send_keys(str(isr / life))
                refresh_isr2 = gc.find_element_by_id('btnAgregaConImpuestoRetenido').click()

            else:
                radio_btn_taxes = gc.find_element_by_xpath('//*[@id="AdicionalImpuestos"]')
                radio_btn_taxes.click()

            # Finish edition of Life Section
            move_to_concept2 = gc.find_element_by_xpath("//a[@id='tabConceptosPrincipal']")
            gc.execute_script("arguments[0].click();", move_to_concept2)
            sleep(1)
            add_concept2 = gc.find_element_by_xpath("//div[@id='tabConceptos']//button[@id='btnAceptarModal']").click()
            sleep(3)
        else:
            log.error('There is no information in Life Section.')

    # Time to verify quantities
    log.info("---Please verify quantities---")
    sleep(20)

    # Finish bill
    gc.find_element_by_xpath("//button[contains(@onclick,'sellar')]").click()

    # Sign
    gc.find_element_by_id('privateKeyPassword').send_keys(config('sat_password'))
    btns = gc.find_elements_by_xpath("//span[@class='btn btn-default']")
    [span.click() for span in btns]
    sleep(10)
    btn_confirm = gc.find_element_by_xpath("//button[@id='btnValidaOSCP']").click()
    sleep(10)
    log.info('Facturation done!!!!')
    btn_sign = gc.find_element_by_xpath("//button[@id='btnFirmar']").click()
    sleep(10)
    download_pdf = gc.find_element_by_xpath("//span[@class='glyphicon glyphicon-file']").click()
    gc.execute_script("scrollBy(0,-1000);")
    sleep(5)
    download_xml = gc.find_element_by_xpath("//span[@class='glyphicon glyphicon-download-alt']").click()
    sleep(10)


if __name__ == '__main__':
    try:
        log.basicConfig(level=log.INFO, format='%(asctime)s :: %(levelname)s :: %(message)s')
        print("Welcome to Bill Generator by Noriaki Mawatari")
        print('Available companies:\nQ) Quálitas\nSP) Seguros el Potosí\nA) Axa Seguros')
        company = input('For which company do you want to make the bill?: ')
        date = input('Bill Period:')

        if company in ['A', 'a', 'Axa', 'axa']:
            axa_values = axa_process()
            # print(axa_values)
            log.info(' | Running SAT automation for Axa.txt file | ')
            automation_in_sat_webpage(company, date, damage=axa_values['damage'], life=axa_values['life'],
                                      isr=axa_values['isr'], iva_ret=axa_values['iva_ret'])
        elif company in ['Q', 'q', 'Qualitas', 'qualitas']:
            qualitas_values = qualitas_process()
            # print(qualitas_values)
            log.info('| Running SAT automation for Quálitas .pdf files |')
            automation_in_sat_webpage(company, date, damage=qualitas_values['import'], isr=qualitas_values['isr'],
                                      iva_ret=qualitas_values['iva_ret'])
        elif company in ['SP', 'sp', 'seguros el potosi', 'Seguros el Potosi', 'Potosi', 'potosi']:
            potosi_values, group = potosi_process()
            # print(potosi_values, group)
            log.info('| Running SAT automation for SP.txt file |')
            if group == 'Daños':
                automation_in_sat_webpage(company, date, damage=potosi_values['subtotal'], isr=potosi_values['isr'],
                                          iva_ret=potosi_values['iva_ret'])
            else:
                automation_in_sat_webpage(company, date, life=potosi_values['subtotal'], isr=potosi_values['isr'],
                                          iva_ret=potosi_values['iva_ret'])
    except Exception as error:
        log.error('ERROR: ', error)
