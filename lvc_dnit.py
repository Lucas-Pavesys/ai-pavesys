import pandas as pd
import numpy as np
from tabelaslvc import lvc_conceitos as lc
from icpf import icpf

class lvc_dnit():

    def __init__(self,src_path,intervalo,tipo) -> None:

        self.src_path = src_path
        self.intervalo = intervalo # em metros
        self.tipo = tipo # 1 -> Padrão | 0 -> Detalhado (TECLAS)

        self.modelo_list = ['Início','Fim',
                            'ICPF','IGGE',
                            'P','TR','J','JE','TB','R','AF','O','D','EX','E',
                            'IES','Código','Conceito',
                            'Observação']
        
        self.geral = ['TRR', 'O', 'P', 'E', 'EX', 'D', 'R',
                        'FI', 'J', 'JE', 'TB', 'TBE', 
                        'TTC', 'TTL', 'TLC', 'TLL', 
                        'ALP', 'ALC', 'ATP', 'ATC']
        
        self.solos = ['P','J','JE','R','O','D','EX','E']

        self.TR = ['TRR','TTC', 'TTL', 'TLC', 'TLL']
        self.TB = ['TB','TBE']
        self.AF = ['ALP', 'ALC', 'ATP', 'ATC']

        pass
    
    def run_command(self):

        result_df = []
        sub_list, main_df = self.read_lvc()

        for i in range(len(sub_list)):

            ini_index = sub_list["ini"].iloc[i]
            fim_index = sub_list["fim"].iloc[i]
            single_df = main_df.iloc[ini_index:fim_index]

            ini_value = single_df["Início"].iloc[0]
            fim_value = single_df["Fim"].iloc[-1]

            empty_row = pd.DataFrame(np.nan, index=[0], columns=single_df.columns)

            single_df["Início"] = single_df["Início"].astype(float)
            single_df["Fim"] = single_df["Fim"].astype(float)

            single_df["Início"] = round(single_df["Início"],3)
            single_df["Fim"] = round(single_df["Fim"],3)
            
            if single_df["Início"].iloc[0] < single_df["Fim"].iloc[0]:
                self.sentido_var = 1
            else:
                self.sentido_var = -1

            self.xl_df = single_df

            self.get_lists()
            self.one2twenty()
            self.df2calc()
            self.freq_defeito()

            self.df_lvc["Início"].iloc[0] = ini_value
            self.df_lvc["Fim"].iloc[-1] = fim_value
            
            result_df.append(self.df_lvc)
            result_df.append(empty_row)

        final_result = pd.concat(result_df, ignore_index=True)

        return final_result
    
    def read_lvc(self):
        """
        Lê a planilha excel e retorna os conjuntos de DataFrame.
        """

        src_df = pd.read_excel(
            self.src_path,
            na_filter=False,
            header=7,
        )

        src_df.rename(columns={"Ex": "EX"}, inplace=True)
        mask = np.where(
            (src_df["Início"] == "")
            | (src_df["Início"] == np.nan)
            | (src_df["Início"] == " ")
        )[0]
        values_to_add = [-1, int(len(src_df))]
        mask = sorted(np.append(mask, values_to_add))
        km_values = []
        for i in range(len(mask) - 1):
            ini = mask[i]
            fim = mask[i + 1]
            km_values.append([ini, fim])
        km_values = pd.DataFrame(km_values, columns=["ini", "fim"])
        km_values["fim"] = km_values["fim"].values
        km_values["ini"] = km_values["ini"].values + 1

        return km_values, src_df

    def get_lists(self):
        '''
        Denomina o tipo de tecla
        '''
        if self.tipo == 1:
            pass
        else:
            self.detalhada2padrao()

    def ajuste_teclas(self,src_df:pd.DataFrame,base_df:pd.DateOffset,src_list:list,dst:str) -> None:

        for i in range(len(src_df)):
            for col in src_list:
                if src_df[col].iloc[i] == 'x':
                    base_df[dst].at[i] = 'x'
                    break
                else:
                    base_df[dst].at[i] = ''

    def detalhada2padrao(self):

        self.base_df = pd.DataFrame(columns=['Início', 'Fim', 
                                    'TRR', 'O', 'P', 'E', 'EX', 'D', 'R',
                                    'FI', 'J', 'JE', 'TB', 'TBE', 
                                    'TTC', 'TTL', 'TLC', 'TLL', 
                                    'ALP', 'ALC', 'ATP', 'ATC', 
                                    'DG', 
                                    'Observação', 
                                    'Latitude', 'Longitude', 'Altitude',
                                    'Data','Hora'],index=self.xl_df.index)
        

        # Solos
        self.base_df[['TRR','D','TB','TBE','E']] = self.xl_df[['TRR','D','TB','TBE','E']]

        # O
        o_list = ['OND.BE', 'OND.ATRE', 'OND.F', 'OND.ATRD', 'OND.BD']
        self.ajuste_teclas(self.xl_df,self.base_df,o_list,'O')

        # P
        p_list = ['Panela.BE.A', 'Panela.BE.M', 'Panela.BE.B', 'Panela.ATRE.A', 'Panela.ATRE.M', 
                  'Panela.ATRBE.B', 'Panela.F.A', 'Panela.F.M', 'Panela.F.B', 'Panela.ATRD.A', 
                  'Panela.ATRD.M', 'Panela.ATRD.B', 'Panela.BD.A', 'Panela.BD.M', 'Panela.BD.B']
        self.ajuste_teclas(self.xl_df,self.base_df,p_list,'P')

        # EX
        ex_list =  ['Exsudação.BE', 'Exsudação.ATRE', 'Exsudação.F', 'Exsudação.ATRD', 'Exsudação.BD']
        self.ajuste_teclas(self.xl_df,self.base_df,ex_list,'Ex')

        # R
        r_list =  ['Remendo.BE', 'Remendo.ATRE', 'Remendo.F', 'Remendo.ATRD', 'Remendo.BD']
        self.ajuste_teclas(self.xl_df,self.base_df,r_list,'R')

        # FI
        fi_list = ['Fi.FC-1.BE', 'Fi.FC-1.ATRE', 'Fi.FC-1.F', 'Fi.FC-1.ATRD', 'Fi.FC-1.BD']
        self.ajuste_teclas(self.xl_df,self.base_df,fi_list,'FI')

        # J
        j_list =  ['J.FC-2.BE', 'J.FC-2.ATRE', 'J.FC-2.F', 'J.FC-2.ATRD', 'J.FC-2.BD']
        self.ajuste_teclas(self.xl_df,self.base_df,j_list,'J')

        # JE
        je_list = ['JE.FC-3.BE', 'JE.FC-3.ATRE', 'JE.FC-3.F', 'JE.FC-3.ATRD', 'JE.FC-3.BD']
        self.ajuste_teclas(self.xl_df,self.base_df,je_list,'JE')

        # TTC
        ttc_list = ['TTC.FC-23.BE', 'TTC.FC-23.ATRE', 'TTC.FC-23.F', 'TTC.FC-23.ATRD', 'TTC.FC-23.BD']
        self.ajuste_teclas(self.xl_df,self.base_df,ttc_list,'TTC')

        # TTL
        ttl_list = ['TTL.FC-23.BE', 'TTL.FC-23.ATRE', 'TTL.FC-23.F', 'TTL.FC-23.ATRD', 'TTL.FC-23.BD']
        self.ajuste_teclas(self.xl_df,self.base_df,ttl_list,'TTL')

        # TLC
        tlc_list = ['TLC.FC-23.BE', 'TLC.FC-23.ATRE', 'TLC.FC-23.F', 'TLC.FC-23.ATRD', 'TLC.FC-23.BD']
        self.ajuste_teclas(self.xl_df,self.base_df,tlc_list,'TLC')

        # TLL
        tll_list =  ['TLL.FC-23.BE', 'TLL.FC-23.ATRE', 'TLL.FC-23.F', 'TLL.FC-23.ATRD', 'TLL.FC-23.BD']
        self.ajuste_teclas(self.xl_df,self.base_df,tll_list,'TLL')

        # ATP
        atp_list = ['ATP-23.BE', 'ATP-23.ATRE', 'ATP-23.F', 'ATP-23.ATRD', 'ATP-23.BD']
        self.ajuste_teclas(self.xl_df,self.base_df,atp_list,'ATP')

        # ATC
        atc_list =  ['ATC-23.BE', 'ATC-23.ATRE', 'ATC-23.F', 'ATC-23.ATRD', 'ATC-23.BD']
        self.ajuste_teclas(self.xl_df,self.base_df,atc_list,'ATC')

        # ALP
        alp_list =  ['ALP-23.BE', 'ALP-23.ATRE', 'ALP-23.F', 'ALP-23.ATRD', 'ALP-23.BD']
        self.ajuste_teclas(self.xl_df,self.base_df,alp_list,'ALP')

        # ALC
        alc_list =  ['ALC-23.BE', 'ALC-23.ATRE', 'ALC-23.F', 'ALC-23.ATRD', 'ALC-23.BD']
        self.ajuste_teclas(self.xl_df,self.base_df,alc_list,'ALC')

        self.xl_df = self.base_df

    def segmentos(self, step) -> np.ndarray:
        '''
        Função que pega o dataframe da planilha e array de índices de acordo com estaqueamento.

        Parameters:
            step (float64): intervalo entre estacas
            xl_df (pd.dataFrame): DataFrame inicial (planilha)
        Returns:
            mask (np.array)       
        '''
        self.ini_int = int(self.xl_df['Início'].iloc[0])
        self.fim_int = int(self.xl_df['Início'].iloc[-1])
        self.ini = self.xl_df['Início'].iloc[0]
        self.fim = self.xl_df['Início'].iloc[-1]

        if self.sentido_var == 1:
            print("Crescente")
            seg_temp = np.arange(self.ini_int, self.fim_int + 1, step)
        else:
            print("Decrescente")
            seg_temp = np.arange(self.fim_int, self.ini_int, -step)

            sub_seg = np.arange(self.ini_int, self.ini, -step)
            seg_temp = np.concatenate((seg_temp, sub_seg))
            seg_temp = np.sort(seg_temp)[::-1]

        # Handle both ascending and descending dataframes
        if self.sentido_var == 1:
            # For ascending dataframes, segmentos within the range of [self.ini, self.fim]
            segmentos = list(seg_temp[(seg_temp >= self.ini) & (seg_temp <= self.fim)])
        else:
            # For descending dataframes, segmentos within the range of [self.fim, self.ini]
            segmentos = list(seg_temp[(seg_temp >= self.fim) & (seg_temp <= self.ini)])

        segmentos = [round(value, 3) for value in segmentos]

        if self.ini not in segmentos:
            segmentos.append(self.ini)
        if self.fim not in segmentos:
            segmentos.append(self.fim)

        segmentos_df = pd.DataFrame(segmentos, columns=['Segmentos'])
        
        mask = np.where(self.xl_df['Início'].isin(segmentos_df['Segmentos']))[0]

        return mask

    def km_obs(self,mask:np.arange,col_names:list) -> pd.DataFrame:
        '''
        Calcula Estacas e Observações

        Parameters:
            mask: (np.arrange): índices de acordo com estaqueamento
            col_names: (list): nomes das colunas a serem consideradas

        Returns:
            df (pd.DataFrame): dataframe calculado
        '''

        df = pd.DataFrame(columns=col_names, index=range(len(mask)-1))

        for i in range(len(mask) - 1):

            df['Início'].at[i] = self.xl_df['Início'].iloc[mask[i]]
            df['Fim'].at[i] = self.xl_df['Início'].iloc[mask[i+1]]

        x = 0

        for i in range(len(mask) - 1):
            pres = mask[i]
            pos = mask[i + 1]

            selected_rows = self.xl_df.loc[pres:(pos - 1)]
            concatenated_values = ' | '.join(value for value in selected_rows['Observação'] if value != '')    

            df['Observação'].at[x] = str(concatenated_values)

            x += 1

        return df

    def one2twenty(self):
        '''
        Função que concatena as informações iniciais (de 1 metro para 20 metros).
        '''
        mask20 = self.segmentos(0.020 * self.sentido_var)
        col_names = self.xl_df.columns
        df20m = self.km_obs(mask20,col_names)

        for col in self.geral:

            x = 0
            for i in range(len(mask20) - 1):
                pres = mask20[i]
                pos = mask20[i + 1]

                selected_rows = self.xl_df.iloc[pres:(pos - 1)] # TESTAR DEPOIS

                concatenated_values = ''.join(selected_rows[col])

                if concatenated_values == '':
                    value = ''
                else:
                    value = 'x'

                df20m[col].at[x] = str(value)

                x += 1

        self.df20m = df20m

    def df2calc(self):
        '''
        Função que conta de acordo com lvc dnit.

        Parameters:
            df20m (pd.DataFrame): dataframe para cálculo

        Returns:
            df_lvc (pd.DataFrame): dataframe lvc calculado
        '''
        self.xl_df = self.df20m.reset_index(drop=True)
        mask = self.segmentos(self.intervalo / 1000 * self.sentido_var)
        df_lvc = self.km_obs(mask,self.modelo_list)
        
        # contagem solos:
        for col in self.solos:
            x = 0
            for i in range(len(mask) - 1):
                pres = mask[i]
                pos = mask[i + 1]

                print(self.df20m["Início"].iloc[pres], self.df20m["Fim"].iloc[(pos - 1)])
                
                prop_ext = round((abs(round(self.df20m["Início"].iloc[pres] - self.df20m["Fim"].iloc[(pos - 1)],3))),3)
                

                selected_rows = self.df20m.loc[pres:(pos - 1)]

                count = 0
                for value in selected_rows[col]:
                    if value == 'x':
                        count += 1

                if col == 'P' or col == 'R':
                    df_lvc[col].at[x] = lc.cod_pnl_rmd((count / prop_ext))
                else:
                    df_lvc[col].at[x] = lc.cod_demais((count / prop_ext))

                x += 1

        # contagem demais:
        # AF
        x = 0
        for i in range(len(mask) - 1):
            count = 0
            pres = mask[i]
            pos = mask[i + 1]
            for col in self.AF:
                selected_rows = self.df20m.loc[pres:(pos - 1)]
                for value in selected_rows[col]:
                    if value == 'x':
                        count += 1
            df_lvc['AF'].at[x] = lc.cod_demais((count  / prop_ext))
            x += 1

        # TR
        x = 0
        for i in range(len(mask) - 1):
            count = 0
            pres = mask[i]
            pos = mask[i + 1]
            for col in self.TR:
                selected_rows = self.df20m.loc[pres:(pos - 1)]
                for value in selected_rows[col]:
                    if value == 'x':
                        count += 1
            df_lvc['TR'].at[x] = lc.cod_demais((count  / prop_ext))
            x += 1

        # TB
        x = 0
        for i in range(len(mask) - 1):
            count = 0
            pres = mask[i]
            pos = mask[i + 1]
            for col in self.TB:
                selected_rows = self.df20m.loc[pres:(pos - 1)]
                for value in selected_rows[col]:
                    if value == 'x':
                        count += 1
            df_lvc['TB'].at[x] = lc.cod_demais((count  / prop_ext))
            x += 1

        self.df_lvc = df_lvc

    def count_cod(self,col_list:list,i:int,df:pd.DataFrame):

        A = 0
        M = 0
        B = 0

        for col in col_list:
            if df[col].iloc[i] == 'A':
                A = A + 1
            elif df[col].iloc[i] == 'M':
                M = M + 1
            elif df[col].iloc[i] == 'B':
                B = B + 1

        return A, M, B
    
    def freq_defeito(self):
        '''
        Calcula frequencia para cada defeito e calcula IGGE
        '''

        for i in range(len(self.df_lvc)):

            col_trincas = ['TR','J','JE','TB']
            col_def = ['AF','O']
            col_pr = ['P','R']
            col_todos = ['P','TR','J','JE','TB','R','AF','O','D','EX','E']
            col_pesados = ['J','JE','TB','P','R','AF']

            Atrinca, Mtrinca, Btrinca = self.count_cod(col_trincas,i,self.df_lvc)
            Adef, Mdef, Bdef = self.count_cod(col_def,i,self.df_lvc)
            Apr, Mpr, Bpr = self.count_cod(col_pr,i,self.df_lvc)
            Atodos, Mtodos, Btodos = self.count_cod(col_todos,i,self.df_lvc)
            Apesado, Mpesado, Bpesado = self.count_cod(col_pesados,i,self.df_lvc)

            #print(f"{self.df_lvc['Início'].iloc[i]} >> {Atrinca} >> {Mtrinca} >> {Btrinca}")

            total = (Atodos + Mtodos + Btodos)

            freq_trinca = lc.freq_trincas(Atrinca, Mtrinca, Btrinca)
            grav_trinca = lc.grav_trincas(Atrinca, Mtrinca, Btrinca)

            freq_def = lc.freq_def(Adef, Mdef, Bdef)
            grav_def = lc.grav_def(Adef, Mdef, Bdef)

            freq_panrem = lc.freq_panrem(Apr, Mpr, Bpr)
            grav_panrem = lc.grav_panrem(Apr, Mpr, Bpr)

            igge = (freq_trinca * grav_trinca) + (freq_def + grav_def) + (freq_panrem + grav_panrem)
            
            JE = self.df_lvc['JE'].at[i]
            J = self.df_lvc['J'].at[i]
            TB = self.df_lvc['TB'].at[i]
            AF = self.df_lvc['AF'].at[i]
            P = self.df_lvc['P'].at[i]
            R = self.df_lvc['R'].at[i]

            icpf_value = icpf.run_icpf(igge,JE,P,AF,TB,J,R,Apesado,Mpesado,Bpesado,total)
            ies,cod,conceito = lc.ies_conceito(igge,icpf_value)

            self.df_lvc['ICPF'].at[i] = icpf_value
            self.df_lvc['IGGE'].at[i] = igge
            self.df_lvc['IES'].at[i] = ies
            self.df_lvc['Código'].at[i] = cod
            self.df_lvc['Conceito'].at[i] = conceito