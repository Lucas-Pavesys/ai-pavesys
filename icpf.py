import random

class icpf():

    def first_class(igge, JE, P, AF, Mpesado, Apesado, TB, J, R):
        if igge > 100 and JE == "A" and P == "A" and AF == "A":
            return "1 ou 2"
        elif igge < 15 and Mpesado < 1 and Apesado < 1:
            return "9 ou 10"
        elif igge < 30 and TB != "A" and JE != "A" and J != "A" and JE != "M":
            return "7 ou 8"
        elif igge < 70 and P != "A" and R != "A":
            return "5 ou 6"
        elif igge < 100:
            return "3 ou 4"
        else:
            return "3 ou 4"
        
    var = first_class(41,"B","","",1,1,"M","A","")

    def zero_um(fc_var, P, JE):
        if fc_var == "1 ou 2" and P == "A" and JE == "A":
            return 1
        elif fc_var == "1 ou 2":
            return 2
        else:
            return 0 

    def um_dois(fc_var, var_zero_um, Apesado):
        if fc_var == "3 ou 4" and Apesado >= 3:
            return 3
        elif fc_var == "3 ou 4":
            return 4
        else:
            return 0

    def dois_tres(fc_var, var_um_dois, Apesado, Mpesado):
        if fc_var == "5 ou 6" and Apesado >= 2 and Mpesado >= 2:
            return 5
        elif fc_var == "5 ou 6":
            return 6
        else:
            return 0 

    def tres_quatro(fc_var, var_dois_tres, Mpesado, Bpesado):
        if fc_var == "7 ou 8" and (Mpesado >= 2 or Bpesado >= 2):
            return 7
        elif fc_var == "7 ou 8":
            return 8
        else:
            return 0

    def quatr_cinco(fc_var, var_tres_quatro, total):
        if fc_var == "9 ou 10" and total == 0:
            return 10
        elif fc_var == "9 ou 10":
            return 9
        else:
            return 0
        
    def icpf_final(var_zero_um, var_um_dois, var_dois_tres, var_tres_quatro, var_quatro_cinco):
        sum_range = var_zero_um + var_um_dois + var_dois_tres +var_tres_quatro + var_quatro_cinco
        if sum_range == 1:
            return random.choice([0.0, 0.5])
        elif sum_range == 2:
            return random.choice([0.5, 1.0])
        elif sum_range == 3:
            return random.choice([1.0, 1.5])
        elif sum_range == 4:
            return random.choice([1.5, 2.0])
        elif sum_range == 5:
            return random.choice([2.0, 2.5])
        elif sum_range == 6:
            return random.choice([2.5, 3.0])
        elif sum_range == 7:
            return random.choice([3.0, 3.5])
        elif sum_range == 8:
            return random.choice([3.5, 4.0])
        elif sum_range == 9:
            return random.choice([4.0, 4.5])
        else:
            return random.choice([4.5, 5.0])
    
    def run_icpf(igge,JE,P,AF,TB,J,R,Apesado,Mpesado,Bpesado,total):

        fc_var = icpf.first_class(igge,JE,P,AF,Mpesado,Apesado,TB,J,R)
        
        var_zero_um = icpf.zero_um(fc_var,P,JE)
        var_um_dois = icpf.um_dois(fc_var,var_zero_um,Apesado)
        var_dois_tres = icpf.dois_tres(fc_var,var_um_dois,Apesado,Mpesado)
        var_tres_quatro = icpf.tres_quatro(fc_var,var_dois_tres,Mpesado,Bpesado)
        var_quatro_cinco = icpf.quatr_cinco(fc_var,var_tres_quatro,total)

        icpf_value = icpf.icpf_final(var_zero_um, var_um_dois, var_dois_tres, var_tres_quatro, var_quatro_cinco)

        return icpf_value
