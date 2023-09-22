import os, sys
import tkinter as tk
from matplotlib.backends.backend_tkagg import FigureCanvasAgg
from tkinter import ttk, filedialog, messagebox
from tkinter.filedialog import asksaveasfile
import pandas as pd
import matplotlib.pyplot as plt
from PIL import Image


def restart ():
  os.execl(sys.executable, os.path.abspath(__file__), *sys.argv) 

def start():

  # Abrindo o file dialog
  filetypes = (
      ('Excel files', '.xlsx .xls'),
      ('text files', '*.txt')
  )

  patch_file = filedialog.askopenfilename(
      title='Abrir Arquivo',
      initialdir='/',
      filetypes=filetypes)
  
  try:
    df = pd.read_excel(patch_file)
  except FileNotFoundError:
    restart()

  # Limpando sessões para novos widgets
  for widgets in left_frame.winfo_children():
    widgets.destroy()

  for widgets in right_frame.winfo_children():
    widgets.destroy()
  #Gerando coluna com pontuacao
  rating = []
  try:
    for row in df['RESPOSTAS']:
            if row == 'Sim, porém ainda existem falhas a serem resolvidas.' : rating.append(50)
            elif row == 'Sim, de forma completa.' : rating.append(100)
            elif row == 'Sim.' : rating.append(100)
            elif row == 'Não temos.' : rating.append(0)
            elif row == 'Não.' : rating.append(0)
            elif row == 'Não temos essa prática.' : rating.append(0)
            else:
              rating.append(None)
  except KeyError: 
    messagebox.showerror(title='Formato inválido', message='Selecione uma planilha com\n'
      'os segunintes headers:\n'
      'QUESTÕES| RESPOSSTAS | CATEGORIA ')
    restart()
  df['PONTUACAO'] = rating

  #limpando dados do DataFrame
  df_limpo = df.loc[df['PONTUACAO'] != None]

  #Criando o valor max esperado nas colunas (Mas não sei se é útil para o gráfico)
  df_limpo.loc[:, 'MAX'] = 100
  
  #AGRUPANDO POR CATEGORIA
  calc = df_limpo.groupby('CATEGORIA',as_index=False)[['PONTUACAO','MAX']].sum()
  calc['%'] = round(calc['PONTUACAO'] / calc['MAX'] *100)
  final = calc[['CATEGORIA', '%']]
  categorias = final['CATEGORIA'].values.tolist()


  #jogando no treeview
  final_rset= final.to_numpy().tolist()
  final_list= list(final.columns.values)
  df_tree= ttk.Treeview(left_frame, columns=final_list)
  df_tree['show'] = 'headings'
  df_tree.grid(row=0, column=0)

  for i in final:
    df_tree.column(i,width=150,anchor='c')
    df_tree.heading(i,text=i)

  for dt in final_rset:
    v=[r for r in dt]
    df_tree.insert('','end', values=v)

  label_1 = ttk.Label(right_frame, text='Selecione as \nCategorias:', font=("Arial", 20), )
  label_1.grid(row=0, column=0,sticky="w")

  #Criando filtros
  # categorias = ["Python", "Perl", "Java", "C++", "C"]
  checkboxes = {}

  placement=1
  for categoria in range(len(categorias)):
    name = categorias[categoria]
    current_var = tk.IntVar()
    filtros = ttk.Checkbutton(
              right_frame, text=name, variable=current_var)
    filtros.var = current_var
    filtros.grid(row=placement, column=0, sticky="nsew")
    checkboxes[filtros] = name
    placement+=1

  #Capturar seleções
  def get_checked_boxes():
    output_categories = []
    for cat in checkboxes:
      if cat.var.get() == 1:
          output_categories.append(checkboxes[cat])
    # print(output_categories)

    #Tentando atualizar o treeview depois do filtro, mas não vi
    # df_tree=ttk.Treeview(left_frame,columns=final_list)
    # df_tree['show'] = 'headings'
    # df_tree.grid(row=0, column=0)

    #CRIANDO O GRAFICO
    df_filtrado = calc[calc.CATEGORIA.isin(output_categories)]
    df_plot = df_filtrado[['CATEGORIA', '%']]
    #Salvar aquivo
    patch_save = filedialog.askdirectory()

    # display(df_to_plot)
    graphic = df_plot.plot.barh(x='CATEGORIA', y='%', color='green',)
    graphic.axis(xmin=0, xmax=100)
    graphic.bar_label(graphic.containers[0], fmt='%.0f%%')
    plt.subplots_adjust(left=0.35)

    plt.savefig(patch_save + '/Graphic.png', format='png', dpi=300, transparent=True)
    plt.close('all')
    tk.messagebox.showinfo(title='Salvo com sucesso!', message=f'Gráfico gerado com sucesso em:\n{patch_save}')
    # plt.show()


  button_filter= ttk.Button(right_frame, text='Gerar', style="Accent.TButton", command=get_checked_boxes)
  button_filter.grid(row=+placement, column=0,sticky="nesw")

  button_filter= ttk.Button(right_frame, text='Gerar novo', style="Accent.TButton", command=restart)
  button_filter.grid(row=+placement+1, column=0,sticky="nesw", pady=5)





root = tk.Tk()
style = ttk.Style(root)
dir_path = os.path.dirname(os.path.realpath(__file__))
root.tk.call('source', 'azure.tcl')
root.tk.call("set_theme", "dark")
root.title("Raio-X Empresarial")
root.iconbitmap('icone.ico')

#Create frame
frame = ttk.Frame(root,)
frame.pack(pady=20, padx=20)

#Frame Esquerdo
left_frame = ttk.Labelframe(frame, text='Raio X', )
left_frame.grid(row=0, column=0,)


button= ttk.Button(left_frame, text='Abrir Arquivo', style="Accent.TButton", command=start)
button.grid(row=0, column=0, pady=100, padx=100,)

#Frame direito
right_frame = ttk.Frame(frame )
right_frame.grid(row=0, column=1, padx=20)


label_1 = ttk.Label(right_frame, text='O que é o Raio X \nEmpresarial?', font=("Arial", 20), )
label_1.grid(row=0, column=0,sticky="w")

label_2 = ttk.Label(right_frame, text='Gere pontuação de uma planilha com \ngráficos automáticos, além de filtros personalizados.', )
label_2.grid(row=1, column=0,sticky="w")

image1 = tk.PhotoImage(file='/Users/dominiquemaik/Documents/PROJETOS TKINTER/Azure-ttk-theme-main/Imagem_01.png')
image_label = ttk.Label(right_frame, image=image1)
image_label.grid(row=2, column=0,sticky="w")

# Adicionando padding em todos os elementos do frame direito
for widgets in right_frame.winfo_children():
  widgets.grid_configure(padx=10, pady=10)

#Criando o loop
root.mainloop()