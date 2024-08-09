
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
import getpass
import os
import csv
import sqlite3
import shutil
import time
import sys
import calendar
import pandas as pd
import mysql.connector
from mysql.connector import Error
from datetime import datetime, date
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.chrome.service import Service as ChromeService
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from time import sleep
from datetime import datetime, timedelta
from selenium.webdriver.support.ui import Select
from pynput.keyboard import Key, Controller
import re
from shareplum import Site
from shareplum.site import Version
import pyautogui
import chardet

navegador = webdriver.Chrome()
    #------------------------------ 

url = "https://www.google.com.br/"

url_sharpoint = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F1%C2%AA%20Quinzena&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"

## CABO E ROTA
url_cabo_as_1_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F1%C2%AA%20Quinzena%2FCabo%20AS%2FOS%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
url_cabo_as_2_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F2%C2%AA%20Quinzena%2FCabo%20AS%2FOS%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
url_cabo_rota_1_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F1%C2%AA%20Quinzena%2FCabo%20Rota%2FOS%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
url_cabo_rota_2_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F2%C2%AA%20Quinzena%2FCabo%20Rota%2FOS%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"



##CARUARU AS E ROTA 
url_caruaru_as_1_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F1%C2%AA%20Quinzena%2FCaruaru%20AS%2FOS%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
url_caruaru_as_2_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F2%C2%AA%20Quinzena%2FCaruaru%20AS%2FOS%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
url_caruaru_rota_1_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F1%C2%AA%20Quinzena%2FCaruaru%20Rota%2FOS%2FComplemento%2F355651%20%2D%20%20Rota%2Erar&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9&parent=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F1%C2%AA%20Quinzena%2FCaruaru%20Rota%2FOS%2FComplemento"
url_caruaru_rota_2_quinzena ="https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F2%C2%AA%20Quinzena%2FCaruaru%20Rota%2FOS%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"



#CDI ANAPOLIS 
url_anapolis_1_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F1%C2%AA%20Quinzena%2FCDI%20An%C3%A1polis%2F354561%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
url_anapolis_2_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F2%C2%AA%20Quinzena%2FCDI%20An%C3%A1polis%2F358967%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"



#CDR BELEM 
url_cdrBelem_1_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F1%C2%AA%20Quinzena%2FCDR%20Bel%C3%A9m%2F356369%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
url_cdrBelem_2_quinzena = "vazio"



##MACEIO AS E ROTA
url_maceio_as_1_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F1%C2%AA%20Quinzena%2FMacei%C3%B3%20AS%2FOS%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
url_maceio_as_2_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F2%C2%AA%20Quinzena%2FMacei%C3%B3%20AS%2FOS%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
url_maceio_rota_1_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F1%C2%AA%20Quinzena%2FMacei%C3%B3%20Rota%2FOS%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
url_maceio_rota_2_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F2%C2%AA%20Quinzena%2FMacei%C3%B3%20Rota%2FOS%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"



#MANAUS ROTA 
url_manaus_rota_1_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F1%C2%AA%20Quinzena%2FManaus%20Rota%2F354879%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9" 
url_manaus_rota_2_quinzena ="https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F2%C2%AA%20Quinzena%2FManaus%20Rota%2F358795%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
print(f"manaus rota 1 quinzera ERRO")

#OLINDA AS E ROTA
url_olinda_as_1_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F1%C2%AA%20Quinzena%2FOlinda%20AS%2FOS%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
url_olinda_as_2_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F2%C2%AA%20Quinzena%2FOlinda%20AS%2FOS%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
url_olinda_rota_1_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F1%C2%AA%20Quinzena%2FOlinda%20Rota%2FOS%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
url_olinda_rota_2_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F2%C2%AA%20Quinzena%2FOlinda%20Rota%2FOS%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"



##OPERALOG BARREIRAS 
url_opBarreiras_1_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F1%C2%AA%20Quinzena%2FOperalog%20Barreiras%2FOS%2FComplemento%2F354683%2F1%C2%AA%20QUINZENA%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
url_opBarreiras_2_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F2%C2%AA%20Quinzena%2FOperalog%20Barreiras%2FOS%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"



#OPERALOG IMPERATRIZ
url_operalog_imperatriz_1_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F1%C2%AA%20Quinzena%2FOperalog%20Imperatriz%2F354494%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
url_operalog_imperatriz_2_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F2%C2%AA%20Quinzena%2FOperalog%20Imperatriz%2F358195%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"



#SAOLUIZ AS, FRANQUIA E ROTA
url_saoLuiz_as_1_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F1%C2%AA%20Quinzena%2FS%C3%A3o%20Lu%C3%ADs%20AS%2F354592%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
url_saoLuiz_as_2_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F2%C2%AA%20Quinzena%2FS%C3%A3o%20Lu%C3%ADs%20AS%2F358397%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
url_saoLuiz_franquia_1_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F1%C2%AA%20Quinzena%2FS%C3%A3o%20Lu%C3%ADs%20Franquia%2F354595%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
url_saoLuiz_franquia_2_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F2%C2%AA%20Quinzena%2FS%C3%A3o%20Lu%C3%ADs%20Franquia%2F358400%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
url_saoLuiz_rota_1_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F1%C2%AA%20Quinzena%2FS%C3%A3o%20Lu%C3%ADs%20Rota%2F354596%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
url_saoLuiz_rota_2_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F2%C2%AA%20Quinzena%2FS%C3%A3o%20Lu%C3%ADs%20Rota%2F358401%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"


## UDC BELEM
url_udcBelem_1_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?newTargetListUrl=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados&viewpath=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2FForms%2FAllItems%2Easpx&id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F1%C2%AA%20Quinzena%2FUDC%20Bel%C3%A9m%2FOS%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
url_udcBelem_2_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F2%C2%AA%20Quinzena%2FUDC%20Bel%C3%A9m%2F358241%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"


##UDC SAO LUIZ 
url_udc_saoLuiz_1_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F1%C2%AA%20Quinzena%2FUDC%20S%C3%A3o%20Lu%C3%ADs%2F354597%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"
url_udc_saoLuiz_2_quinzena = "https://grupohorizonte.sharepoint.com/sites/SP_Controller/Documentos%20Partilhados/Forms/AllItems.aspx?id=%2Fsites%2FSP%5FController%2FDocumentos%20Partilhados%2F02%2E%20Custos%2F01%2E%20Faturamento%2F1%2E%20Fechamento%20SRTrans%2F2024%2F05%2E%20Maio%2F2%C2%AA%20Quinzena%2FUDC%20S%C3%A3o%20Lu%C3%ADs%2F358404%2FComplemento&viewid=6e14a596%2D3fd0%2D46ed%2D9d32%2Db3a0150e5dd9"


def config_options(usuario):
    download_directory = "C:\\Arquivos_Controladoria\\"
    
    options = Options() 
    options.add_argument('--disable-gpu')
    options.add_argument("--start-maximized")
    options.add_argument("--force-device-scale-factor=0.8")
    options.add_argument("--disable-javascript")
    
    prefs = {
        "download.default_directory": download_directory,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
        
    }
    options.add_experimental_option("prefs", prefs)
    return options


"--------------------------------------------------------------------------------------------------------------"
xpath_input_email_sharepoint = "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[2]/div[2]/div/input[1]"
xpath_input_avancar = "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[4]/div/div/div/div/input"
xpath_input_senha_sharepoint = "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[3]/div/div[2]/input "
xpath_input_entrar = "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[5]/div/div/div/div/input"
xpath_Input_sim = "/html/body/div/form/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[3]/div[2]/div/div/div[2]/input"
path_cabo_as = "(//div[contains(@class,'ms-DetailsRow-cell cell-342')])[2]"
path_cabo_rota = "(//div[contains(@class,'ms-DetailsRow-cell cell-342')])[2]"
button_abrir_relatorio_cabo = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
hover_cabo_as_1_quinzena = "(//div[contains(@class,'ms-DetailsRow-cell cell-209')])[2]"
button_baixar_relatorio = "//button[@name='Baixar']/div[1]/span[1]"
span_cabo_as_1_quinzena = "(//span[@class='signalFieldValue_b9e74371'])[3]"
span_cabo_as_2_quinzena = "(//span[@class='signalFieldValue_b9e74371'])[4]"
span_cabo_rota_1_quinzena = "(//span[@class='signalFieldValue_b9e74371'])[2]"
span_cabo_rota_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[4]/div/div/div[2]/div[1]"
span_anapolis_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div[2]/div[1]"
button_2_quinzena_cabo_as = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[4]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
button_1_quinzena_cabo_rota = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[2]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
button_2_quinzena_cabo_rota = "/html/body/div/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[4]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
button_1_quinzena_anapolis = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_anapolis_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[4]/div/div/div[2]/div[1]"
button_2_quinzena_anapolis = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[4]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_cdr_belem_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[4]/div/div/div[2]/div[2]"
button_cdr_belem_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[4]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_maceio_as_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div[1]/div[5]/div/div/div[2]/div[1]"
button_maceio_as_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div[1]/div[5]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_maceio_as_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div[1]/div[7]/div/div/div[2]/div[1]"
button_maceio_as_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div[1]/div[7]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_maceio_rota_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div[1]/div[5]/div/div/div[2]/div[1]"
button_maceio_rota_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div[1]/div[5]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_maceio_rota_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div[1]/div[7]/div/div/div[2]/div[1]"
button_maceio_rota_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div[1]/div[7]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_manaus_rota_1_quinzena = "/html/body/div/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[2]/div/div/div[2]/div[1]"
button_manaus_rota_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[2]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_manaus_rota_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div[2]/div[1]"
button_manaus_rota_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_olinda_as_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[7]/div/div/div[2]/div[1]"
button_olinda_as_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[7]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_olinda_as_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div[2]/div[1]"
button_olinda_as_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_olinda_rota_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[7]/div/div/div[2]/div[1]"
button_olinda_rota_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[7]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_olinda_rota_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div[2]/div[1]"
button_olinda_rota_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_Operalog_barreiras_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[7]/div/div/div[2]/div[1]"
button_Operalog_barreiras_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[7]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_Opelog_barreiras_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[5]/div/div/div[2]/div[1]"
button_Operalog_barreiras_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[5]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_Operalog_imperatriz_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div[1]/div[13]/div/div/div[2]/div[1]"
button_Operlog_imperatriz_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div[1]/div[13]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
div_imperatriz = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]"
span_Operalog_imperatriz_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div[1]/div[17]/div/div/div[2]/div[1]"
button_Operlog_imperatriz_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div[1]/div[17]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_Opelog_saoLuiz_as_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div[2]/div[1]"
button_Operalog_saoLuiz_as_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_Opelog_saoLuiz_as_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[4]/div/div/div[2]/div[1]"
button_Operalog_saoLuiz_as_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[4]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_Opelog_saoLuiz_franquia_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div[2]/div[1]"
button_Operalog_saoLuiz_franquia_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_Opelog_saoLuiz_franquia_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div[2]/div[1]"
button_Operalog_saoLuiz_franquia_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_Opelog_saoLuiz_rota_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div[2]/div[1]"
button_Operalog_saoLuiz_rota_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_Opelog_saoLuiz_rota_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[4]/div/div/div[2]/div[1]"
button_Operalog_saoLuiz_rota_2_quinzena ="/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[4]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_udc_belem_1_quinzena = "/html/body/div/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[1]/div/div/div[2]/div[1]"
button_udc_belem_1_quinzena = "/html/body/div/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[1]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_udc_belem_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[8]/div/div/div[2]/div[1]"
button_udc_belem_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[8]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_udc_saoluiz_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div[2]/div[1]"
button_udc_saoluiz_1_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"
span_udc_saoluiz_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[4]/div/div/div[2]/div[1]"
button_udc_saoluiz_2_quinzena = "/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div/div[4]/div/div/div[2]/div[2]/div/div[1]/div[2]/div/button[2]/span"








def abrir_navegador():
    global url,navegador
    navegador.get(url)
    
def change_tab_with_interactions(driver, url):
    # Salvar a janela atual
    current_window = driver.current_window_handle
    
    # Abrir uma nova aba e mudar para ela
    driver.execute_script("window.open('" + url + "', '_blank');")
    
    # Mudar o foco para a nova aba
    for window_handle in driver.window_handles:
        if window_handle != current_window:
            driver.switch_to.window(window_handle)
            break

def click_js(elemento):
    navegador.execute_script("arguments[0].click();", elemento)         


def wait(time):
    global navegador
    return WebDriverWait(navegador, time)


def click(elemento):
    try:
        wait(50).until(EC.visibility_of_element_located((By.XPATH, elemento))).click()
    except Exception as e:
        print(f"Erro ao clicar no elemento: {e}")

        
def escreva_digitando(elemento, valor):
    campo = wait(50).until(
        EC.visibility_of_element_located((By.XPATH, elemento))
    )
    
    # Limpar o campo
    campo.clear()

    # Escrever o valor letra por letra
    for letra in valor:
        campo.send_keys(letra)
        time.sleep(0.1)


def entrarIframe(xpath):
    wait(50).until(
        EC.frame_to_be_available_and_switch_to_it((By.XPATH, xpath))
    )

    
def sairIframe():
    navegador.switch_to.default_content()

def hoveraction(driver, elemento):
    # Use o ActionChains para passar o mouse sobre o elemento
    ActionChains(driver).move_to_element(elemento).perform()    
    

def change_tab_with_interactions(driver, url):
    # Salvar a janela atual
    current_window = driver.current_window_handle
    
    # Abrir uma nova aba e mudar para ela
    driver.execute_script("window.open('" + url + "', '_blank');")
    
    # Mudar o foco para a nova aba
    for window_handle in driver.window_handles:
        if window_handle != current_window:
            driver.switch_to.window(window_handle)
            break
    # Fechar a aba anterior //public back = $sql_inject


def hover(elemento):
    wait(50).until(
        EC.presence_of_element_located((By.XPATH, elemento))
    )
    elemento_hover = navegador.find_element(By.XPATH, elemento)
    ActionChains(navegador).move_to_element(elemento_hover).perform() 

def fechar_primeira_janela(driver):
    # Certificar-se de que existem pelo menos duas janelas abertas
    if len(driver.window_handles) >= 1:
        # Trocar para a primeira janela
        driver.switch_to.window(driver.window_handles[0])
        # Fechar a primeira janela
        driver.close()
        # Se houver mais de uma janela, trocar para a próxima janela disponível
        if len(driver.window_handles) > 0:
            driver.switch_to.window(driver.window_handles[0])
    else:
        print("Não há janelas abertas para fechar.") 


def scroll_down(driver):
    try:
        # Rola a página para baixo usando as teclas de seta
        body = driver.find_element(By.TAG_NAME, 'body')
        body.send_keys(Keys.ARROW_DOWN)
        time.sleep(1)  # Aguarda um segundo para a rolagem ser executada
        # print("Rolou a página para baixo.")
    except Exception as e:
        print(f"Erro ao rolar a página para baixo: {e}")
        
        

def verificar_arquivos_no_diretorio(nomes_arquivos, diretorio="C:\\Arquivos_Controladoria\\"):
    # Lista para armazenar os arquivos que estão presentes no diretório
    arquivos_presentes = []

    # Iterar sobre a lista de nomes de arquivos
    for nome_arquivo in nomes_arquivos:
        # Construir o caminho completo do arquivo
        caminho_arquivo = os.path.join(diretorio, nome_arquivo)
        
        # Verificar se o arquivo existe
        if os.path.isfile(caminho_arquivo):
            arquivos_presentes.append(nome_arquivo)
        # else:
        #     print(f"O arquivo {nome_arquivo} não foi encontrado no diretório {diretorio}")

    return arquivos_presentes


def click_at(x, y):
    pyautogui.click(x, y)
    

def join_arquivos(diretorio, output_file):
    # Lista todos os arquivos no diretório que terminam com '.csv'
    arquivos = [f for f in os.listdir(diretorio) if f.endswith('.csv')]
    
    # Lista para armazenar os dataframes e as colunas
    dataframes = []
    colunas_conjuntas = set()
    
    def tirar_espacos(df):
        """Remove os espaços em branco de todas as células do dataframe."""
        return df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    
    # Lê cada arquivo CSV e adiciona à lista de dataframes
    for arquivo in arquivos:
        caminho_arquivo = os.path.join(diretorio, arquivo)
        try:
            df = pd.read_csv(caminho_arquivo, encoding='utf-8')
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(caminho_arquivo, encoding='latin-1')
            except Exception as e:
                print(f"Erro ao ler o arquivo {caminho_arquivo}: {e}")
                continue
        
        # Remove espaços em branco
        df = tirar_espacos(df)
        
        # Atualiza o conjunto de colunas conjuntas
        colunas_conjuntas.update(df.columns)
        dataframes.append(df)
    
    # Alinha todos os dataframes às colunas conjuntas
    colunas_conjuntas = sorted(colunas_conjuntas)
    dataframes_alinhados = []
    for df in dataframes:
        df_alinhado = df.reindex(columns=colunas_conjuntas)
        dataframes_alinhados.append(df_alinhado)
    
    # Concatena todos os dataframes alinhados em um único dataframe
    if dataframes_alinhados:
        df_consolidado = pd.concat(dataframes_alinhados, ignore_index=True)
        
        # Remover duplicatas
        df_consolidado = df_consolidado.drop_duplicates()
        
        # Salva o dataframe consolidado em um único arquivo CSV
        output_path = os.path.join(diretorio, output_file)
        df_consolidado.to_csv(output_path, index=False)
        print(f"Arquivo consolidado salvo como {output_path}")
    else:
        print("Nenhum arquivo CSV válido foi encontrado no diretório.")
        
# Define o diretório e o nome do arquivo de saída
diretorio = 'C:\\Arquivos_Controladoria'
output_file = '03.08.15_Consolidado.csv'
    
# join_arquivos(diretorio, output_file)

# print(output_file)



    
    
    


usuario = getpass.getuser()
options = config_options(usuario)
navegador = webdriver.Chrome(options=options) ##DRIVER DA FUNÇÃO = NAVEGADOR 
abrir_navegador()
time.sleep(5) 
change_tab_with_interactions(navegador,url_sharpoint)
time.sleep(5)
click(xpath_input_email_sharepoint)
time.sleep(1)
escreva_digitando(xpath_input_email_sharepoint,"joao.tavares@grupohorizonte.com.br")
time.sleep(3)
click(xpath_input_avancar)
time.sleep(2)
escreva_digitando(xpath_input_senha_sharepoint,"Shinji@0305")
time.sleep(2)
click(xpath_input_entrar)
time.sleep(4)
click(xpath_Input_sim)
time.sleep(4)
fechar_primeira_janela(navegador)
time.sleep(2)

# #CABO AS 1 QUINZENA
change_tab_with_interactions(navegador, url_cabo_as_1_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10)
click(span_cabo_as_1_quinzena)
time.sleep(3)
hover(span_cabo_as_1_quinzena)
time.sleep(3)
click(button_abrir_relatorio_cabo)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10)
print("CABO AS 1 QUINZENA")

##CABO AS 2 QUINZENA

change_tab_with_interactions(navegador, url_cabo_as_2_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10)
click(span_cabo_as_2_quinzena)
time.sleep(3)
click(button_2_quinzena_cabo_as)
time.sleep(3)
click(button_baixar_relatorio)
time.sleep(10)
print(f"CABO AS 2 QUINZENA") 
 
 
# ##CABO ROTA 1 QUINZENA 

change_tab_with_interactions(navegador, url_cabo_rota_1_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10)
click(span_cabo_rota_1_quinzena)
time.sleep(3)
click(button_1_quinzena_cabo_rota)
time.sleep(3)
click(button_baixar_relatorio)
time.sleep(10)
print(f"CABO ROTA 1 QUINZENA")

# CABO ROTA 2 QUINZENA 

change_tab_with_interactions(navegador, url_cabo_rota_2_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10)
click(span_cabo_rota_2_quinzena) 
time.sleep(3)
click(button_2_quinzena_cabo_rota)
time.sleep(3)
click(button_baixar_relatorio)
time.sleep(10)
print(f"CABO ROTA 2 QUINZENA")


#ANAPOLIS 1 QUINZENA 

change_tab_with_interactions(navegador, url_anapolis_1_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10)
click(span_anapolis_1_quinzena)
time.sleep(3)
click(button_1_quinzena_anapolis)
time.sleep(3)
click(button_baixar_relatorio)
time.sleep(10)
print(f"ANAPOLIS 1 QUINZENA")

#ANAPOLIS 2 QUINZENA

change_tab_with_interactions(navegador, url_anapolis_2_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(2)
click(span_anapolis_2_quinzena)
time.sleep(3)
click(button_2_quinzena_anapolis)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10)
print(f"ANAPOLIS 2 QUINZENA")

#CDR BELEM 1 QUINZENA

change_tab_with_interactions(navegador, url_cdrBelem_1_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10)
click(span_cdr_belem_1_quinzena)
time.sleep(3)
click(button_cdr_belem_1_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10)
print(f"CDR BELÉM QUINZENA")


## MACEIO AS 1 QUINZENA 

change_tab_with_interactions(navegador, url_maceio_as_1_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10) 
click(span_maceio_as_1_quinzena)
time.sleep(3)
click(button_maceio_as_1_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10)
print(f"MACEIO AS 1 QUINZENA")

# ## MACEIO AS 2 QUINZENA


change_tab_with_interactions(navegador, url_maceio_as_2_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10) 
click(span_maceio_as_2_quinzena)
time.sleep(3)
click(button_maceio_as_2_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10)
print(f"MACEIO AS 2 QUINZENA")

# MACEIO ROTA 1 QUINZENA 


change_tab_with_interactions(navegador, url_maceio_rota_1_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10) 
click(span_maceio_rota_1_quinzena)
time.sleep(3)
click(button_maceio_rota_1_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10)
print(f"MACEIO ROTA 1 QUINZENA")

# MACEIO ROTA 2 QUINZENA

change_tab_with_interactions(navegador, url_maceio_rota_2_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10) 
click(span_maceio_rota_2_quinzena)
time.sleep(3)
click(button_maceio_rota_2_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10)
print(f"MACEIO ROTA 2 QUINZENA")

# MANAUS ROTA 1 QUINZENA


change_tab_with_interactions(navegador, url_manaus_rota_1_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10) 
click(span_manaus_rota_2_quinzena)
time.sleep(3)
click(button_manaus_rota_1_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(3)
print(f"MANAUS ROTA 1 QUINZENA")


# # MANAUS ROTA 2 QUINZENA

change_tab_with_interactions(navegador, url_manaus_rota_2_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(2) 
click(span_manaus_rota_2_quinzena)
time.sleep(3)
click(button_manaus_rota_2_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10)
print(f"MANAUS ROTA 2 QUINZENA")


# # OLINDA AS 1 QUINZENA

 
change_tab_with_interactions(navegador, url_olinda_as_1_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(1) 
click(span_olinda_as_1_quinzena)
time.sleep(3)
click(button_olinda_as_1_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10)
print(f"OLINDA AS 1 QUINZENA")

# #OLINDA AS 2 QUINZENA 


change_tab_with_interactions(navegador, url_olinda_as_2_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10) 
click(span_olinda_as_2_quinzena)
time.sleep(3)
click(button_olinda_as_2_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10)
print(f"OLINDA AS 2 QUINZENA")

#OLINDA ROTA 1 QUINZENA 

change_tab_with_interactions(navegador, url_olinda_rota_1_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10) 
click(span_olinda_rota_1_quinzena)
time.sleep(3)
click(button_olinda_rota_1_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10)
print(f"OLINDA ROTA 1 QUINZENA")


#OLINDA ROTA 2 QUINZENA 


change_tab_with_interactions(navegador, url_olinda_rota_2_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10) 
click(span_olinda_rota_2_quinzena)
time.sleep(3)
click(button_olinda_rota_2_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10)
print(f"OLINDA ROTA 2 QUINZENA")


##OPERALAOG BARREIRAS 1 QUINZENA 


change_tab_with_interactions(navegador, url_opBarreiras_1_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10) 
click(span_Operalog_barreiras_1_quinzena)
time.sleep(3)
click(button_Operalog_barreiras_1_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10)
print(f"OPERALOG BARREIRAS 1 QUINZENA")


# #OPELOG BARREIRAS 2 QUINZENA 


change_tab_with_interactions(navegador, url_opBarreiras_2_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10) 
click(span_Opelog_barreiras_2_quinzena)
time.sleep(3)
click(button_Operalog_barreiras_2_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10) 
print(f"OPERALOG BARREIRAS 2 QUINZENA")

# # OPERALOG IMPERATRIZ 1 QUINZENA 


change_tab_with_interactions(navegador, url_operalog_imperatriz_1_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
click(div_imperatriz)
time.sleep(1)
scroll_down(navegador)
time.sleep(1)
scroll_down(navegador)
time.sleep(3)
scroll_down(navegador)
time.sleep(1)  
scroll_down(navegador)
time.sleep(1)  
scroll_down(navegador)
time.sleep(1)  
scroll_down(navegador)
time.sleep(1)  
scroll_down(navegador)
time.sleep(1)  
scroll_down(navegador)
time.sleep(1)  
scroll_down(navegador)
time.sleep(1)  
scroll_down(navegador)
time.sleep(1)  
scroll_down(navegador)
time.sleep(5)  
click(span_Operalog_imperatriz_1_quinzena)
time.sleep(3)
click(button_Operlog_imperatriz_1_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10)
print(f"OPERALOG IMPERATRIZ 1 QUINZENA")

# # #OPERALOG IMPERATRIZ 2 QUINZENA

change_tab_with_interactions(navegador, url_operalog_imperatriz_2_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
click(div_imperatriz)
time.sleep(1)
scroll_down(navegador)
time.sleep(1)
scroll_down(navegador)
time.sleep(1)
scroll_down(navegador)
time.sleep(1)  
scroll_down(navegador)
time.sleep(1)  
scroll_down(navegador)
time.sleep(1)  
scroll_down(navegador)
time.sleep(1)  
scroll_down(navegador)
time.sleep(1)  
scroll_down(navegador)
time.sleep(1)  
scroll_down(navegador)
time.sleep(1)  
scroll_down(navegador)
time.sleep(1)  
scroll_down(navegador)
time.sleep(5)  
scroll_down(navegador)
time.sleep(5) 
click(span_Operalog_imperatriz_2_quinzena)
time.sleep(1)
click(button_Operlog_imperatriz_2_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10)
print(f"OPERALOG IMPERATRIZ 2 QUINZENA")



# OPERALOG SAO LUIZ AS 1 QUINZENA 


change_tab_with_interactions(navegador, url_saoLuiz_as_1_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10) 
click(span_Opelog_saoLuiz_as_1_quinzena)
time.sleep(3)
click(button_Operalog_saoLuiz_as_1_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10) 

print(f"OPERALOG SAO LUIZ AS 1 QUINZENA")

### OPERALOG SAO LUIZ AS 2 QUINZENA 

change_tab_with_interactions(navegador, url_saoLuiz_as_2_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10) 
click(span_Opelog_saoLuiz_as_2_quinzena)
time.sleep(3)
click(button_Operalog_saoLuiz_as_2_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10) 

print(f"OPERALOG SAO LUIZ AS 2 QUINZENA")

# ##OPERALOG SAO LUIZ FRANQUIA 1 QUINZENA 

change_tab_with_interactions(navegador, url_saoLuiz_franquia_1_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10) 
click(span_Opelog_saoLuiz_franquia_1_quinzena)
time.sleep(3)
click(button_Operalog_saoLuiz_franquia_2_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10) 
print(f"OPERALOG SAO LUIZ FRANQUIA 1 QUINZENA")


##OPERALOG SAO LUIZ FRANQUIA 2 QUINZENA 

change_tab_with_interactions(navegador, url_saoLuiz_franquia_2_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10) 
click(span_Opelog_saoLuiz_franquia_2_quinzena)
time.sleep(3)
click(button_Operalog_saoLuiz_franquia_2_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10) 
print(f"OPERALOG SAO LUIZ FRANQUIA AS 2 QUINZENA")


# # OPERALOG SAO LUIZ ROTA 1 QUINZENA 

change_tab_with_interactions(navegador, url_saoLuiz_rota_1_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10) 
click(span_Opelog_saoLuiz_rota_1_quinzena)
time.sleep(3)
click(button_Operalog_saoLuiz_rota_1_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10) 
print(f"OPERALOG SAO LUIZ ROTA 1 QUINZENA")

# #OPERALOG SAO LUIZ ROTA 2 QUINZENA 

change_tab_with_interactions(navegador, url_saoLuiz_rota_2_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10) 
click(span_Opelog_saoLuiz_rota_2_quinzena)
time.sleep(3)
click(button_Operalog_saoLuiz_rota_2_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10) 
print(f"OPERALOG SAO LUIZ ROTA 2 QUINZENA")



# #UDC BELEM 1 QUINZENA 

change_tab_with_interactions(navegador, url_udcBelem_1_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10) 
click(span_udc_belem_1_quinzena)
time.sleep(3)
click(button_udc_belem_1_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10) 
print(f"UDC BELÉM 1 QUINZENA")


# #UDC BELEM 2 QUINZENA 

change_tab_with_interactions(navegador, url_udcBelem_2_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10) 
click(span_udc_belem_2_quinzena)
time.sleep(3)
click(button_udc_belem_2_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10) 
print(f"UDC BELÉM 2 QUINZENA")

# #UDC SAO LUIZ 1 QUINZENA 

change_tab_with_interactions(navegador, url_udc_saoLuiz_1_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10) 
click(span_udc_saoluiz_1_quinzena)
time.sleep(3)
click(button_udc_saoluiz_1_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10) 
print(f"UDC SAO LUIZ 1 QUINZENA")


# ##UDC SAO LUIZ 2 QUINZENA 

change_tab_with_interactions(navegador, url_udc_saoLuiz_2_quinzena)
time.sleep(10)
fechar_primeira_janela(navegador)
time.sleep(2)
scroll_down(navegador)
time.sleep(10) 
click(span_udc_saoluiz_2_quinzena)
time.sleep(3)
click(button_udc_saoluiz_2_quinzena)
time.sleep(3) 
click(button_baixar_relatorio)
time.sleep(10) 
print(f"UDC SAO LUIZ 2 QUINZENA")


nomes_arquivos = ["03.08.15_as.csv", "03.08.15_rota","03.08.15_rota"] 
arquivos_encontrados = verificar_arquivos_no_diretorio(nomes_arquivos)

print("Arquivos encontrados:", arquivos_encontrados)

diretorio = "C:\\Arquivos_Controladoria\\"
diretorio_dest = "C:\\Arquivos_Controladoria\\consolidado\\"


# join_arquivos(diretorio, output_file)

# print(output_file)

# Lista para armazenar os dataframes
dfs = []

conteudo_total = []

# Função para detectar a codificação de um arquivo
def detectar_codificacao(arquivo):
    with open(arquivo, 'rb') as file:
        conteudo_raw = file.read()
    deteccao = chardet.detect(conteudo_raw)
    return deteccao['encoding']

# Iterar sobre cada arquivo na pasta
for indice, filename in enumerate(os.listdir(diretorio)):
    if filename.endswith('.csv'):  # Certifique-se de que é um arquivo de texto
        filepath = os.path.join(diretorio, filename)
        codificacao = detectar_codificacao(filepath)

        # Abrir e ler o arquivo com a codificação detectada
        with open(filepath, 'r', encoding=codificacao) as file:
            linhas = file.readlines()
            
            # Se não for o primeiro arquivo, remover a primeira linha
            if indice > 0 and linhas:
                linhas = linhas[1:]
            
            # Adicionar as linhas à lista total de conteúdo
            conteudo_total.extend(linhas)

# Salvar o conteúdo combinado em um novo arquivo
caminho_arquivo_saida = os.path.join(diretorio_dest, 'teste.csv')
with open(caminho_arquivo_saida, 'w', encoding='utf-8') as arquivo_saida:
    arquivo_saida.writelines(conteudo_total)


sys.exit



# time.sleep(400)
                                                               


