import re
from tkinter import *
from file_creator import *




#verifie si une ligne de code ne contient pas de caractere interdit
def check_for_specials(sentence):
        specials = ",;!?.<>[]@"
        for i in range(len(specials)):
                if (specials[i] in sentence):
                        return True
        return False

#verifie la syntaxe des lignes de code
def check_for_mistakes(filename):
        test_list = []
        newfile = open(filename,"r",encoding="utf-8")
        lis = newfile.readlines()
        err_msg = ""
        boo = False
        for i in range(len(lis)):
                if (lis[i][0:4] == "%cod"):
                    test_list = lis[i].split(":")
                    if (len(test_list) != 17):
                        err_msg = err_msg + "\n" + "Nombre de champs incorrects dans la ligne" + str(i)
                        boo = True
                    if (check_for_specials(lis[i])):
                        err_msg = err_msg + "\n"+ "Caractère interdit au niveau de la ligne" + str(i)
                        boo = True
                    j = i+1
                    while (lis[j][0] == "\t"):
                        test_list = lis[j].split(":")
                        if (len(test_list) != 16):
                                err_msg = err_msg + "\n" + "Nombre de champs incorrect dans la ligne" + str(j)
                                boo = True
                        if (check_for_specials(lis[j])):
                                err_msg = err_msg + "\n"+"Caractère interdit au niveau de la ligne" + str(j)
                                boo = True
                        j = j+1
                        if (j == len(lis)):
                                break
        if (boo):
            messagebox.showerror("Liste d'erreurs", "Nom du fichier: " + removepath(filename) + "\n" + err_msg)
        return False
def check_subd(filename):
	new_file = open(filename,"r",encoding="utf-8")
	lis = detectcodelines(new_file.readlines())
	ret_val = True
	err_msg = ""
	for i in range(len(lis)):
		divided = re.split(r":|&|#|\|",lis[i])
		if (is_sole_stim(divided)):
			stim = divided[0]
			continue
		elif (not (count_specials(lis[i]) and lis[i][0:3] == "\t$:")):
			ret_val = False
			err_msg = err_msg + "Erreur dans le stimuli: " + stim + "dans la ligne de code suivante\n" + lis[i] + "\n" + "la ligne n'est pas bien partitionnée\n"
		elif (" " in lis[i]):
			ret_val = False
			err_msg = err_msg + "Erreur dans le stimuli: " + stim + "dans la ligne de code suivante\n" + lis[i] + "\n" + "il y a un ou plusieurs espaces\n"
		elif (not(lis[i][-2] == "}" or lis[i][-1] == "}")):
			err_msg = err_msg + "Erreur dans le stimuli: " + stim + "dans la ligne de code suivante\n" + lis[i] + "\n" + "Il manque le dernier caractère { \n"
		else:
			continue
	if (not(ret_val)):
                messagebox.showerror("Liste d'erreurs", "Nom du fichier: " + removepath(filename) + "\n" + err_msg)

def is_sole_stim(ste_list):
        if (len(ste_list) == 1):
                return True
        else:
                return False

def count_specials(ste):
        num_and = ste.count("&")
        num_pip = ste.count("|")
        num_hash = ste.count("#")
        if (num_and == 5 and num_hash == 1 and num_pip >=6):
                return True
        else:
                return False