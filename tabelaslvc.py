import numpy as np
import random

class lvc_conceitos():

    def freq_trincas(A, M, B):
        if A == 4:
            return random.randint(90, 100)
        elif A == 3:
            if M == 1:
                return random.randint(80, 90)
            elif B == 1:
                return random.randint(75, 80)
            elif M == 0 and B == 0:
                return random.randint(65, 75)
        elif A == 2:
            return random.randint(55, 65)
        elif A == 1:
            return random.randint(50, 55)
        elif M == 4:
            return random.randint(40, 50)
        elif M == 3:
            return random.randint(30, 40)
        elif M == 2:
            return random.randint(20, 30)
        elif M == 1:
            return random.randint(10, 20)
        elif B == 4:
            return random.randint(9, 10)
        elif B == 3:
            return random.randint(8, 9)
        elif B == 2:
            return random.randint(6, 7)
        elif B == 1:
            return random.randint(3, 6)
        else:
            return 0

    def freq_def(A, M, B):
        if A == 2:
            return random.randint(65, 75)
        elif A == 1:
            if M == 1:
                return random.randint(60, 65)
            elif B == 1:
                return random.randint(55, 60)
            elif M == 0 and B == 0:
                return random.randint(50, 55)
        elif M == 2:
            return random.randint(30, 50)
        elif M == 1:
            return random.randint(10, 30)
        elif B == 2:
            return random.randint(7, 10)
        elif B == 1:
            return random.randint(2, 7)
        elif A + M + B == 0:
            return 0
        
    def freq_panrem(A, M, B):
        if A == 2:
            return random.randint(40, 70)
        elif A == 1:
            return random.randint(30, 40)
        elif A + M == 2:
            return random.randint(20, 30)
        elif A + M == 1:
            return random.randint(10, 20)
        elif B == 2:
            return random.randint(7, 10)
        elif B == 1:
            return random.randint(2, 7)
        elif A + M + B == 0:
            return 0

    def cod_demais(defeito):
        if defeito != 0:
            if defeito <= (5):
                cod = 'B'
            elif defeito < (25):
                cod = 'M'
            else:
                cod = 'A'
        else:
            cod = np.nan
        return cod

    def cod_pnl_rmd(defeito):
        if defeito != 0:
            if defeito <= 2:
                cod = 'B'
            elif defeito < 5:
                cod = 'M'
            else:
                cod = 'A'
        else:
            cod = np.nan
        return cod

    def grav_trincas(A, M, B):
        if A == 0 and M == 0 and B == 0:
            return 0
        elif A > 0:
            return 0.65
        elif M > 0:
            return 0.45
        else:
            return 0.3

    def grav_def(A, M, B):
        if A == 0 and M == 0 and B == 0:
            return 0
        elif A > 0:
            return 1
        elif M > 0:
            return 0.7
        else:
            return 0.6

    def grav_panrem(A, M, B):
        if A == 0 and M == 0 and B == 0:
            return 0
        elif A > 0:
            return 1
        elif M > 0:
            return 0.8
        else:
            return 0.7


    def ies_conceito(igge,icpf):
        if igge <= 20 and icpf > 3.5:
            ies = 0
            cod = "A"
            conceito = "Ótimo"
        elif igge <=20 and icpf <= 3.5:
            ies = 1
            cod = "B"
            conceito = "Bom"
        elif 20 < igge <= 40 and icpf > 3.5:
            ies = 2
            cod = "B"
            conceito = "Bom" 
        elif 20 < igge <= 40 and icpf <= 3.5:
            ies = 3
            cod = "C"
            conceito = "Regular"
        elif 40 < igge <= 60 and icpf > 2.5:
            ies = 4
            cod = "C"
            conceito = "Regular"
        elif 40 < igge <= 60 and icpf <= 2.5:
            ies = 5
            cod = "D"
            conceito = "Ruim"
        elif 60 < igge <= 90 and icpf > 2.5:
            ies = 7
            cod = "D"
            conceito = "Ruim"
        elif 60 < igge <= 90 and icpf <= 2.5:
            ies = 8
            cod = "E"
            conceito = "Péssimo"
        else:
            ies = 10
            cod = "E"
            conceito = "Péssimo"
        return ies,cod,conceito