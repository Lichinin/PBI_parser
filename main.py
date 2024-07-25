import os

from dotenv import load_dotenv

from pbi_classes.dmm_2200p_s2 import DMM2200P
from pbi_classes.dmm_2410d_s2 import DMM2410D

load_dotenv()

location = 'Кораблино'

pbi_2200_list = {
    212: '192.168.1.212',
    213: '192.168.1.213',
    214: '192.168.1.214',
    215: '192.168.1.215',
    216: '192.168.1.216',
    217: '192.168.1.217',
    218: '192.168.1.218',
}

pbi_2410_list = {
    219: '192.168.1.219',
    220: '192.168.1.220',
}

for name, address in pbi_2200_list.items():
    pbi = DMM2200P(
        address,
        os.getenv('LOGIN'),
        os.getenv('KORABLINO_PASSWORD'),
        location
    )
    pbi.get_all_parameters()
    pbi.export_params_to_excel(name)

for name, address in pbi_2410_list.items():
    pbi = DMM2410D(
        address,
        os.getenv('LOGIN'),
        os.getenv('KORABLINO_PASSWORD'),
        location
    )
    pbi.get_all_parameters()
    pbi.export_params_to_excel(name)
