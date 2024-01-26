import pandas as pd
import numpy as np
import os

class pre_processo:
    def __init__(self, src_path, xl_list, ini, fim, tipo) -> None:
        """
        Classe que faz o pre processamento dos dados.

        Parameters:

        xl_list (list[path]): lista de planilhas a integrar o processo
        ini (float64): km de início da planilha
        fim (float64): km de fim da planilhas
        nome (str): nome da planilha a ser exportada

        return:

        pp_xl (excel sheet): planilha exportada
        """

        self.src_path = src_path
        self.xl_list = xl_list

        self.ini = ini
        self.fim = fim
        self.step = 0.001

        self.tipo = tipo

        pass

    def create_base(self):
        """
        Cria Dataframe para servir como base para os dados

        """

        if self.ini < self.fim:
            self.step = self.step
            self.sentido = True
        elif self.fim < self.ini:
            self.step = -(self.step)
            self.sentido = False

        if self.tipo == 0:
            print("Detalhada")
            self.base_df = pd.DataFrame(
                columns=[
                    "Início",
                    "Fim",
                    "TRR",
                    "E",
                    "D",
                    "Fi.FC-1.BE",
                    "Fi.FC-1.ATRE",
                    "Fi.FC-1.F",
                    "Fi.FC-1.ATRD",
                    "Fi.FC-1.BD",
                    "J1.FC-1.BE",
                    "J1.FC-1.ATRE",
                    "J1.FC-1.F",
                    "J1.FC-1.ATRD",
                    "J1.FC-1.BD",
                    "J.FC-2.BE",
                    "J.FC-2.ATRE",
                    "J.FC-2.F",
                    "J.FC-2.ATRD",
                    "J.FC-2.BD",
                    "JE.FC-3.BE",
                    "JE.FC-3.ATRE",
                    "JE.FC-3.F",
                    "JE.FC-3.ATRD",
                    "JE.FC-3.BD",
                    "TB",
                    "TBE",
                    "TTC.FC-23.BE",
                    "TTC.FC-23.ATRE",
                    "TTC.FC-23.F",
                    "TTC.FC-23.ATRD",
                    "TTC.FC-23.BD",
                    "TTL.FC-23.BE",
                    "TTL.FC-23.ATRE",
                    "TTL.FC-23.F",
                    "TTL.FC-23.ATRD",
                    "TTL.FC-23.BD",
                    "TLC.FC-23.BE",
                    "TLC.FC-23.ATRE",
                    "TLC.FC-23.F",
                    "TLC.FC-23.ATRD",
                    "TLC.FC-23.BD",
                    "TLL.FC-23.BE",
                    "TLL.FC-23.ATRE",
                    "TLL.FC-23.F",
                    "TLL.FC-23.ATRD",
                    "TLL.FC-23.BD",
                    "ALP-23.BE",
                    "ALP-23.ATRE",
                    "ALP-23.F",
                    "ALP-23.ATRD",
                    "ALP-23.BD",
                    "ALC-23.BE",
                    "ALC-23.ATRE",
                    "ALC-23.F",
                    "ALC-23.ATRD",
                    "ALC-23.BD",
                    "ATP-23.BE",
                    "ATP-23.ATRE",
                    "ATP-23.F",
                    "ATP-23.ATRD",
                    "ATP-23.BD",
                    "ATC-23.BE",
                    "ATC-23.ATRE",
                    "ATC-23.F",
                    "ATC-23.ATRD",
                    "ATC-23.BD",
                    "OND.BE",
                    "OND.ATRE",
                    "OND.F",
                    "OND.ATRD",
                    "OND.BD",
                    "Panela.BE.A",
                    "Panela.BE.M",
                    "Panela.BE.B",
                    "Panela.ATRE.A",
                    "Panela.ATRE.M",
                    "Panela.ATRBE.B",
                    "Panela.F.A",
                    "Panela.F.M",
                    "Panela.F.B",
                    "Panela.ATRD.A",
                    "Panela.ATRD.M",
                    "Panela.ATRD.B",
                    "Panela.BD.A",
                    "Panela.BD.M",
                    "Panela.BD.B",
                    "Exsudação.BE",
                    "Exsudação.ATRE",
                    "Exsudação.F",
                    "Exsudação.ATRD",
                    "Exsudação.BD",
                    "Remendo.BE",
                    "Remendo.ATRE",
                    "Remendo.F",
                    "Remendo.ATRD",
                    "Remendo.BD",
                    "DG",
                    "Observação",
                    "Latitude",
                    "Longitude",
                    "Altitude",
                    "Data",
                    "Hora",
                ]
            )
        elif self.tipo == 1:
            print("Padrão")
            self.base_df = pd.DataFrame(
                columns=[
                    "Início",
                    "Fim",
                    "TRR",
                    "O",
                    "P",
                    "E",
                    "Ex",
                    "D",
                    "R",
                    "FI",
                    "J",
                    "JE",
                    "TB",
                    "TBE",
                    "TTC",
                    "TTL",
                    "TLC",
                    "TLL",
                    "ALP",
                    "ALC",
                    "ATP",
                    "ATC",
                    "DG",
                    "Observação",
                    "Latitude",
                    "Longitude",
                    "Altitude",
                    "Data",
                    "Hora",
                ]
            )
        num_elements = int((self.fim - self.ini) / self.step)

        self.base_df["Início"] = np.arange(
            self.ini, self.ini + num_elements * self.step, self.step
        )

        self.base_df["Início"] = round(self.base_df["Início"], 3)

    def read_excel(self, file_path: str):
        """
        Lê a planilha excel e retorna o DataFrame.
        """

        if file_path.endswith(".xls"):
            xl_df = pd.read_excel(
                file_path,
                na_filter=False,
                engine="xlrd",
                na_values=["", "NA", "N/A", "NaN", "None"],
                header=7,
            )

        elif file_path.endswith(".xlsx"):
            xl_df = pd.read_excel(
                file_path,
                na_filter=False,
                engine="openpyxl",
                na_values=["", "NA", "N/A", "NaN", "None"],
                header=7,
            )
        else:
            pass

        return xl_df

    def remove_blanks(self, df:pd.DataFrame) -> pd.DataFrame:

        end = (df.columns.get_loc("Observação"))
        order = (True if df["Início"].iloc[0] < df["Fim"].iloc[0] else False)

        # removing blank and doubling rows
        test = df[df.columns[2:end]].apply(
            lambda x: "".join(x.dropna().astype(str)),
            axis=1
        )
        mask = test != ""
        df = df[mask].reset_index(drop=True)

        return df

    def list2base(self):
        """
        Adiciona os dataframes na base.

        """

        # loop para ler e adicionar dataframe na base
        for file in self.xl_list:
            temp_path = os.path.join(self.src_path, file)
            
            # Read the Excel file into a DataFrame
            m_temp_df = self.read_excel(temp_path)
            
            temp_df = self.remove_blanks(m_temp_df)

            # Round the "Início" and "Fim" columns to 3 decimal places
            temp_df["Início"] = round(temp_df["Início"], 3)
            temp_df["Fim"] = round(temp_df["Fim"], 3)
            
            # Filter rows in 'temp_df' where "Início" is in 'self.base_df'
            mask = temp_df["Início"].isin(self.base_df["Início"])
            filtered_temp_df = temp_df[mask]
            
            # Set "Início" as the index for both DataFrames
            self.base_df.set_index("Início", inplace=True)
            filtered_temp_df.set_index("Início", inplace=True)
            
            # Update values in 'self.base_df' based on 'filtered_temp_df'
            self.base_df.update(filtered_temp_df)
            
            # Reset the index of 'self.base_df'
            self.base_df.reset_index(inplace=True)
            
            # Calculate "Fim" based on "Início" and 'self.step'
            self.base_df["Fim"] = self.base_df["Início"] + self.step

        return self.base_df

    def run_pp(self):
        print("Início do Processo")
        self.create_base()
        print("Início Concatenação")
        dataframe = self.list2base()
        print("Retornando DataFrame")
        return dataframe
