from DMM_2200P_S2 import DMM2200P
pbi_list = {
    62: '192.168.1.62',
    63: '192.168.1.63',
    64: '192.168.1.64',
    65: '192.168.1.65',
}
username = 'admin'
password = 'CrnCrjgby'
location = 'Скопин'

for name, address in pbi_list.items():
    pbi = DMM2200P(address, username, password, location)
    pbi.get_all_parameters()
    pbi.export_params_to_excel(name)
