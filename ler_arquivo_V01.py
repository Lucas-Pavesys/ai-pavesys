# ------ IMPORTAÇÕES ------
import os

# ------ FUNÇÕES ------
def dicionario_arquivos(path):

    print('Lendo arquivos .xlsx e .xls: ')
    print(path)

    arquivos = os.listdir(path)
    file_pistas = []
    codigo_pistas = []

    for filename in os.listdir(path): # filtrando arquivos excel
        if filename.endswith('.xlsx'):
            file_pistas.append(os.path.join(path, filename))
            codigo_pistas.append(filename)
        if filename.endswith('.xls'):
            file_pistas.append(os.path.join(path, filename))
            codigo_pistas.append(filename)

    pista_simples = []
    pista_dupla_cres = []
    pista_dupla_decres = []
    pista_adc_cres = []
    pista_adc_decres = []
    pista_ramo = []

    print('\nAgrupando trechos')
    for pista in codigo_pistas:
        if pista.split('_')[4][:3] == 'ADC':
            if pista.split('_')[3] == 'C':
                pista_adc_cres.append(pista)
            elif pista.split('_')[3] == 'D':
                pista_adc_decres.append(pista)
        elif pista.split('_')[4][:4] == 'RAMO':
            pista_ramo.append(pista)
        else:
            if pista.split('_')[2] == 'S':
                pista_simples.append(pista)
            elif pista.split('_')[2] == 'D':
                if pista.split('_')[3] == 'C':
                    pista_dupla_cres.append(pista)
                elif pista.split('_')[3] == 'D':
                    pista_dupla_decres.append(pista)
    
    trechos = []
    for x in range(len(codigo_pistas)): # descobrindo nome dos trechos
        if codigo_pistas[x].split('_')[1] in trechos:
            pass
        else:
            trechos.append(codigo_pistas[x].split('_')[1])

    dc_pista_simples = {}
    dc_pista_dupla_cresc = {}
    dc_pista_dupla_decresc = {}
    dc_pista_adc_cres = {}
    dc_pista_adc_decres = {}
    dc_pista_ramo = {}

    atr_pista_simples = {}
    atr_pista_dupla_cresc = {}
    atr_pista_dupla_decresc = {}
    atr_pista_adc_cres = {}
    atr_pista_adc_decres = {}
    atr_pista_ramo = {}

    i = 0
    while i < len(trechos):
        temp = []
        temp_atr = []
        for teste in pista_simples:
            if not teste.split('_')[-1][:3] == 'ATR' and trechos[i] == teste.split('_')[1]:
                temp.append(teste)
            if temp != []:
                dc_pista_simples[trechos[i]] = temp
            if teste.split('_')[-1][:3] == 'ATR' and trechos[i] == teste.split('_')[1]:
                temp_atr.append(teste)
            if temp != []:
                atr_pista_simples[trechos[i]] = temp_atr
            
        temp = []
        temp_atr = []
        for teste in pista_dupla_cres:
            if not teste.split('_')[-1][:3] == 'ATR' and trechos[i] == teste.split('_')[1]:
                temp.append(teste)
            if temp != []:
                dc_pista_dupla_cresc[trechos[i]] = temp
            if teste.split('_')[-1][:3] == 'ATR' and trechos[i] == teste.split('_')[1]:
                temp_atr.append(teste)
            if temp != []:
                atr_pista_dupla_cresc[trechos[i]] = temp_atr
    
        temp = []
        temp_atr = []
        for teste in pista_dupla_decres:
            if not teste.split('_')[-1][:3] == 'ATR' and trechos[i] == teste.split('_')[1]:
                temp.append(teste)
            if temp != []:
                dc_pista_dupla_decresc[trechos[i]] = temp
            if teste.split('_')[-1][:3] == 'ATR' and trechos[i] == teste.split('_')[1]:
                temp_atr.append(teste)
            if temp != []:
                atr_pista_dupla_decresc[trechos[i]] = temp_atr
    
        temp = []
        temp_atr = []
        for teste in pista_adc_cres:
            if not teste.split('_')[-1][:3] == 'ATR' and trechos[i] == teste.split('_')[1]:
                temp.append(teste)
            if temp != []:
                dc_pista_adc_cres[trechos[i]] = temp
            if teste.split('_')[-1][:3] == 'ATR' and trechos[i] == teste.split('_')[1]:
                temp_atr.append(teste)
            if temp != []:
                atr_pista_adc_cres[trechos[i]] = temp_atr
    
        temp = []
        temp_atr = []
        for teste in pista_adc_decres:
            if not teste.split('_')[-1][:3] == 'ATR' and trechos[i] == teste.split('_')[1]:
                temp.append(teste)
            if temp != []:
                dc_pista_adc_decres[trechos[i]] = temp
            if teste.split('_')[-1][:3] == 'ATR' and trechos[i] == teste.split('_')[1]:
                temp_atr.append(teste)
            if temp != []:
                atr_pista_adc_decres[trechos[i]] = temp_atr

        temp = []
        temp_atr = []
        for teste in pista_ramo:
            if not teste.split('_')[-1][:3] == 'ATR' and trechos[i] == teste.split('_')[1]:
                temp.append(teste)
            if temp != []:
                dc_pista_ramo[trechos[i]] = temp
            if teste.split('_')[-1][:3] == 'ATR' and trechos[i] == teste.split('_')[1]:
                temp_atr.append(teste)
            if temp != []:
                atr_pista_ramo[trechos[i]] = temp_atr
    
        i = i + 1

    print('Finalizado! \n')
    print('------------ Trechos que serão processados ------------')
    print('-> Pista simples: ')
    for chave in dc_pista_simples:
        print(chave)
        print(dc_pista_simples[chave])
        print(atr_pista_simples[chave])
        print('\n')
    print('-> Pista dupla crescente: ')
    for chave in dc_pista_dupla_cresc:
        print(chave)
        print(dc_pista_dupla_cresc[chave])
        print(atr_pista_dupla_cresc[chave])
        print('\n')
    print('-> Pista dupla decrescente: ')
    for chave in dc_pista_dupla_decresc:
        print(chave)
        print(dc_pista_dupla_decresc[chave])
        print(atr_pista_dupla_decresc[chave])
        print('\n')
    print('-> Faixas adicionais crescentes: ')
    for chave in dc_pista_adc_cres:
        print(chave)
        print(dc_pista_adc_cres[chave])
        print(atr_pista_adc_cres[chave])
        print('\n')
    print('-> Faixas adicionais decrescentes: ')
    for chave in dc_pista_adc_decres:
        print(chave)
        print(dc_pista_adc_decres[chave])
        print(atr_pista_adc_decres[chave])
        print('\n')
    print('-> Ramos de dispositivos: ')
    for chave in dc_pista_ramo:
        print(chave)
        print(dc_pista_ramo[chave])
        print(atr_pista_ramo[chave])
        print('\n')
    print('-------------------------------------------------------')

    return dc_pista_simples, dc_pista_dupla_cresc, dc_pista_dupla_decresc, dc_pista_adc_cres, dc_pista_adc_decres, dc_pista_ramo, atr_pista_simples, atr_pista_dupla_cresc, atr_pista_dupla_decresc, atr_pista_adc_cres, atr_pista_adc_decres, atr_pista_ramo


# dicionario_arquivos(os.getcwd())

