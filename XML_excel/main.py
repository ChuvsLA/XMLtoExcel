import xmltodict
import os
import pandas as pd

def get_info(filename, values):
    with open(f'NFs/{filename}', 'rb') as xml_file:
        file_dict = xmltodict.parse(xml_file)
        if 'NFe' in file_dict:
            nf_info = file_dict['NFe']['infNFe']
        else:
            nf_info = file_dict['nfeProc']['NFe']['infNFe']
        receipt = nf_info['@Id']
        company_issuing = nf_info['emit']['xNome']
        client_name = nf_info['dest']['xNome']
        address = nf_info['dest']['enderDest']
        if 'vol' in nf_info['transp']:
            weight = nf_info['transp']['vol']['pesoB']
        else:
            weight = 'Not informed'
        values.append([receipt, company_issuing, client_name, address, weight])

file_list = os.listdir("NFs")

columns = ['receipt', 'company_issuing', 'client_name', 'address', 'weight']
values = []

for file in file_list:
    get_info(file, values)

chart = pd.DataFrame(columns=columns, data=values)
chart.to_excel('NotasFiscais.xlsx', index=False)