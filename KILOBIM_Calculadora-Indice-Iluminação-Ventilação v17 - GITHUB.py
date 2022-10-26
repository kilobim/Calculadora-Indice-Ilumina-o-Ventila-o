import tkinter as tk
from tkinter import *
from tkinter import ttk
from idlelib.tooltip import Hovertip
from tkinter import messagebox
from PIL import Image, ImageTk
import os, sys 
import pandas as pd
import pathlib
from pathlib import Path
import webbrowser

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

###janela
janela= Tk()
janela.geometry("350x550")
janela.title("Calculadora | Áreas de iluminação e ventilação")
janela.minsize(350,600)
janela.maxsize(350,600)

# data_folder_ico = resource_path('favicon-novalogo.ico')
# print(data_folder_ico)
# janela.iconbitmap(data_folder_ico)

# ### logotipo

# # Read the Image
# data_folder_img = resource_path('kilobim_v2_003_transparent.png')
# image = Image.open(data_folder_img)

# # Resize the image using resize() method
# resize_image = image.resize((50,15))

# img = ImageTk.PhotoImage(resize_image)

# # create label and add resize image
# label1 = Label(image=img)
# label1.image = img
# label1.place(x=290,y=570)

label_kilo = tk.Label(janela, text = "Desenvolvido por:")
label_kilo.place(x=210,y=572)
label_kilo.config(font=("Arial",'7',"italic"))


###opções

def numero_lado(lx):
	try:
		float(lx)
		# print('Ok')

	except ValueError:
		messagebox.showwarning('Erro','Utilize números para definir área.')


## lado 1
label_area_amb = tk.Label(janela, text = "Área do ambiente:")
label_area_amb.place(x=20,y=30)

# entrada

area_amb = tk.Entry(janela,width=20)
area_amb.place(x=150,y=30)
area_amb.insert(0,float(0))
numero_lado(area_amb.get())


### lista cidades

###criando dataframe

file_location=Path(__file__).parent.resolve()

if Path('ListaCidadesIndicesEdit.xlsx').is_file():
	print('lista EXISTE')
	data_folder_df = ('ListaCidadesIndicesEdit.xlsx')
	cidades_df = pd.read_excel(data_folder_df)
	print(cidades_df)
else:
	try:
		print('lista NAO EXISTE')
		data_folder_df = resource_path('ListaCidadesIndices.xlsx')
		cidades_df = pd.read_excel(data_folder_df)
		
		export_dados=pd.ExcelWriter('ListaCidadesIndicesEdit.xlsx')
		cidades_df.to_excel(export_dados,'Planilha1',index=False)
		export_dados.save()

		data_folder_df = ('ListaCidadesIndicesEdit.xlsx')
		cidades_df = pd.read_excel(data_folder_df)
		print(cidades_df)
	except PermissionError:
		messagebox.showwarning('Erro','O programa não pôde ser aberto pois ele está localizado no diretório "C:\". \n Para solucionar erro, por favor coloque o programa em outra pasta.',parent=janela)
		sys.exit()

# valid_df=cidades_df.sort_values('Cidade',ascending=True)	

# print(valid_df)

###conversão cidades dataframe >> UI (tkinter)
cidades_list=[]
for i in cidades_df.loc[:,'Cidade']:
	cidades_list.append(i)

def updt():
	listacidade.delete(0,END)
	u=cidades_list.sort()
	for u in cidades_list:
		listacidade.insert(END,u)

def save_df():
	cidades_df.sort_values(['Cidade'],inplace=True)
	cidades_df.reset_index(drop=True, inplace=True)
	cidades_list.sort()
	export_dados=pd.ExcelWriter('ListaCidadesIndicesEdit.xlsx')
	cidades_df.to_excel(export_dados,'Planilha1',index=False)
	export_dados.save()

print(cidades_list)

# print(cidades_list)
cidades_var = tk.StringVar(value=cidades_list)

## frame 
frame = Frame(janela)
frame.place(x=150,y=75)

##lista a ser vista na UI

listacidade = tk.Listbox(frame,listvariable=cidades_var)
listacidade.pack(side='left', fill='y')

## configurações scrollbar
scrollbar=Scrollbar(frame, orient="vertical")
scrollbar.config(command=listacidade.yview)
scrollbar.pack(side='right',fill='y')

listacidade.config(yscrollcommand=scrollbar.set)

label_cidades = tk.Label(janela, text = "Cidade:")
label_cidades.place(x=20,y=80)
label_cidades.config(font=("Times New Roman",'12'))


def janela3():
	
	# Toplevel object which will
	# be treated as a new window
	novajanela = tk.Toplevel(janela)

	# sets the title of the
	# Toplevel widget
	novajanela.title("Adicionar Cidade")
	novajanela.iconbitmap(data_folder_ico)

	# sets the geometry of toplevel
	novajanela.geometry("300x120")
	novajanela.minsize(300,120)
	novajanela.maxsize(300,120)

	disty = 5
	distx = 5
	distx2 = 140

	novacidade = tk.Entry(novajanela,width=20)
	novacidade.place(x=distx2,y=disty)
	

	novacidade_label = tk.Label(novajanela, text = "Cidade:")
	novacidade_label.place(x=distx,y=disty)

	disty = disty + 20

	novo_indperm = tk.Entry(novajanela,width=20)
	novo_indperm.place(x=distx2,y=disty)
	

	novo_indperm_label = tk.Label(novajanela, text = "Índice prolongado:")
	novo_indperm_label.place(x=distx,y=disty)

	disty = disty + 20

	novo_indtransi = tk.Entry(novajanela,width=20)
	novo_indtransi.place(x=distx2,y=disty)
	

	novo_indtransi_label = tk.Label(novajanela, text = "Índice transitório:")
	novo_indtransi_label.place(x=distx,y=disty)

	disty = disty + 20

	novo_indoutro = tk.Entry(novajanela,width=20)
	novo_indoutro.place(x=distx2,y=disty)
	

	novo_indoutro_label = tk.Label(novajanela, text = "Outro:")
	novo_indoutro_label.place(x=distx,y=disty)

	def addlista():
		try:
			if novacidade.get() == "":
				messagebox.showwarning('Erro','Nome da cidade não definido.',parent=novajanela)
			if float(novo_indperm.get()) < 1 or float(novo_indtransi.get()) < 1 or float(novo_indoutro.get()) < 1:
				messagebox.showwarning('Erro','Atribua índices maiores que zero.',parent=novajanela)
			else:				
				if novacidade.get() in cidades_df.values:
					messagebox.showwarning('Erro','Cidade já existente. Caso queira atualizar, clique em "Remover cidade", depois, adicione-a novamente com os índices requisitados.',parent=novajanela)
				else:
					listacidade.insert(END,novacidade.get()) ### adicionar cidade na lista da UI
					cidades_list.append(novacidade.get()) ### adicionar cidade na lista de conversão dataframe >> UI
					cidades_df.loc[len(cidades_df)]=[novacidade.get(),float(novo_indperm.get()),float(novo_indtransi.get()),float(novo_indoutro.get())] ### adicionar cidade no dataframe
					novacidade.delete(0,END)			
					novo_indperm.delete(0,END)
					novo_indtransi.delete(0,END)
					novo_indoutro.delete(0,END)
					cidades_list.sort()
					updt()
					print('ok')
					save_df()
					print(cidades_df)
					novajanela.destroy()
		except ValueError:
 			messagebox.showwarning('Erro','Valores Índices prolongado e transitório em falta.Atribua a eles apenas números maiores que zero.')

	disty = disty + 20

	addbutton = tk.Button(novajanela, text='Adicionar',command=addlista)
	addbutton.place(x=50,y=disty)

	# sairbutton = tk.Button(novajanela, text='Sair',command=Close)
	# sairbutton.place(x=200,y=disty)


def removelista():
	MsgBox = tk.messagebox.askquestion ('Remover.','Gostaria de remover cidade e seus índices da lista?',icon = 'warning')
	if MsgBox == 'yes':
		global item
		global tupla
		item=listacidade.get(listacidade.curselection())
		tupla=listacidade.curselection()		### tupla para identificar o indice da cidade clicadas
		cidades_df.drop(tupla[0],inplace=True)				### remover a cidade do data frame
		cidades_list.remove(item) 				###remover cidade na lista de conversão dataframe >> UI
		listacidade.delete(listacidade.curselection()) ###remover cidade da lista da UI
		print(item)
		cidades_list.sort()
		updt()
		save_df()

disty = 120

disty = disty + 30

addbutton = tk.Button(janela, text='Adicionar cidade',width = 15,command=janela3)
addbutton.place(x=20,y=disty)

disty = disty + 30

removebutton = tk.Button(janela, text='Remover cidade',width = 15,command=removelista)
removebutton.place(x=20,y=disty)



def prova():
	# save_df()
	for i in cidades_df.loc[:,'Cidade']:
		x = i
		print (x)
	print(numero_lado(area_amb.get()))

# def provaentry():
# 	print(area_amb)
# 	print(lado2)

## Caixa de diálogo para tipos de ambientes

frame = tk.LabelFrame(janela, text = "Tipo de área de permanência:", padx=10,pady=20)
frame.place (x=20,y=410)	

# r = IntVar()
# r.set = 0

listacidade.select_set(0)

perm_type_combobox = ttk.Combobox(frame, values=["Prolongado","Transitório","Outro"])
perm_type_combobox.grid(column=0, row=1)
perm_type_combobox.current(1)

label_ind1=StringVar()
label_ind2=IntVar()
label_ind3=IntVar()


def obter_indices():
	try:
		global r_indperm
		global r_indtransi
		global r_outro

		r_tupla=listacidade.curselection()
		r_indperm=cidades_df.loc[r_tupla[0],'Índice permanente']
		r_indtransi=cidades_df.loc[r_tupla[0],'Índice transitório']
		r_outro=cidades_df.loc[r_tupla[0],'Outro']
		print(r_tupla)
		print(r_indperm)
		print(r_indtransi)
		print(r_outro)
		
		label_ind1.configure(text= 'Cidade escolhida: ' + cidades_df.loc[r_tupla[0],'Cidade'])

		label_ind2.configure(text= 'Índice permanente: ' + str(r_indperm))

		label_ind3.configure(text= 'Índice transitório: ' + str(r_indtransi))

		label_ind4.configure(text= 'Outro: ' + str(r_outro))

		label_alert.destroy()


	except IndexError:
		messagebox.showwarning('Erro','Escolha uma cidade.')
	except UnboundLocalError:
		messagebox.showwarning('Erro','Escolha uma cidade.')

	return(r_indperm)
	return(r_indtransi)
	return(r_outro)
	return(validation)


label_alert = tk.Label(janela, text = "Nenhuma cidade escolhida. \n Escolha cidade e clique em 'Obter índices' para obter dados.")
label_alert.config(font=("Times New Roman",'8'))
label_alert.config(fg="Red")
label_alert.place(x=25,y=245)



yind=325

label_ind1 = Label(janela)
label_ind1.place(x=150,y=yind)
label_ind1.config(font=("Times New Roman",'10'))

yind=yind+20

label_ind2 = Label(janela)
label_ind2.place(x=150,y=yind)
label_ind2.config(font=("Times New Roman",'10'))

yind=yind+20

label_ind3 = Label(janela)
label_ind3.place(x=150,y=yind)
label_ind3.config(font=("Times New Roman",'10'))

yind=yind+20

label_ind4 = Label(janela)
label_ind4.place(x=150,y=yind)
label_ind4.config(font=("Times New Roman",'10'))

botao_obter = tk.Button(janela,text="Obter índices", width = 18, pady = 5,command =obter_indices)
botao_obter.place(x=150,y=285)

#hovertip
tip = Hovertip(frame,'Prolongada: ambientes de uso contínuo. Ex.: sala, cozinha, quartos, escritórios, etc.\nTransitória: ambientes de uso por curto período de tempo. Ex.: banheiros, garagem, corredores, etc. \nOutros: pode ser usada para definir índice de área de ventilação apenas, ou ainda, para outros tipos de classificação de ambientes\nnão descritos aqui, conforme código de obras da cidade indicada.')
# tip.bind_widget(frame,balloonmsg='Prolongada: ambientes de uso contínuo. Ex.: sala, cozinha, quartos, escritórios, etc.\nTransitória: ambientes de uso por curto período de tempo. Ex.: banheiros, garagem, corredores, etc.')

# calcular



def calculo():
	try:
		global l1
		global valor_total
		global factor

		if perm_type_combobox.get() == "Prolongado":
			r=r_indperm
		elif perm_type_combobox.get() == "Transitório":
			r=r_indtransi
		else:
			r=r_outro

		l1 = float(area_amb.get().replace(",","."))
		factor = float(r)

		if factor < 1:
			messagebox.showwarning('Erro','Escolha um tipo de ambiente.Ou verfique se há índice com valor zero e faça correção no botão "Editar Cidades".')
			valor_total=0
		else:
			valor_calculado = l1*1/factor
			valor_total = '{:.2f}'.format(valor_calculado) + ' m²'	#*(1/factor)

		label_valor_total.configure(text= valor_total)

		print(valor_total)

	except ValueError:
		messagebox.showwarning('Erro','Utilize números para definir área.')
	except NameError:
		messagebox.showwarning('Erro','Utilize números para definir área.')
	return (valor_total)

label_valor_total = Label(janela)
label_valor_total.place(x=180,y=537)
label_valor_total.config(font=("Times New Roman",'15'))



botao_print2 = tk.Button(janela,text="Calcular", width = 8, pady = 10,command =calculo)
botao_print2.place(x=220,y=440)



## resultado


label_calc = tk.Label(janela, text = "Área mínima para janelas:")
label_calc.place(x=20,y=540)

## ajuda

def LinkAjuda():
	webbrowser.open("https://kilobim.com/?v=cf46f4701eab")

helpbutton = tk.Button(janela, text='Ajuda', fg='blue', command=LinkAjuda)
helpbutton.place(x=5,y=570)

###manter a janela aberta para ver
janela.mainloop()