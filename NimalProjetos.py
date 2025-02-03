import tkinter as tk
from PIL import Image, ImageTk
import time
import tkinter as tk
from PIL import Image, ImageTk
import time
def splash_screen():
    global splash, gif, label

    # Criar a Splash Screen rapidamente
    splash = tk.Tk()
    splash.geometry("300x70")
    splash.title("Carregando...")
    splash.configure(bg="#d9d9d9")
    splash.overrideredirect(True)

    # Centralizar na tela
    screen_width = splash.winfo_screenwidth()
    screen_height = splash.winfo_screenheight()
    x_position = (screen_width - 300) // 2
    y_position = (screen_height - 70) // 2
    splash.geometry(f"+{x_position}+{y_position}")

    # Adicionar um Label de carregamento (GIF)
    gif_path = "barra.gif"  # Certifique-se de que o caminho do GIF está correto
    gif = Image.open(gif_path)

    # Criar um Label para exibir o GIF
    label = tk.Label(splash, bg="#d9d9d9")  # Fundo igual ao da splash
    label.pack(expand=True)

    def atualizar_gif(frame=0):
        try:
            gif.seek(frame)  # Seleciona o frame atual
            frame_atual = ImageTk.PhotoImage(gif.copy())  # Cria uma cópia para evitar erros
            label.config(image=frame_atual)
            label.image = frame_atual  # Manter referência para evitar coleta de lixo
            # Agendar o próximo frame
            splash.after(100, atualizar_gif, (frame + 1) % gif.n_frames)
        except Exception as e:
            print("Erro ao carregar o GIF:", e)

    atualizar_gif()

    splash.after(7000, splash.destroy)
    splash.mainloop()
splash_screen()
import mysql.connector
from customtkinter import CTkImage, CTkLabel
from PIL import Image
import pandas as pd
import customtkinter as ctk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from tkcalendar import DateEntry

conexao = mysql.connector.connect(
    host="192.168.0.101",
    user="seu_usuario",
    password="sua_senha",
    database="nimalnotas"
)

cursor = conexao.cursor()
def mostrar_visao_geral():
        global tree, frame2

        frame2.configure(width=1000, height=400)
        visao_frame = ctk.CTkFrame(frame2, fg_color="gray")
        visao_frame.place(relwidth=1, relheight=1)

        def carregar_dados(valor=None,coluna=None):
            for row in tree.get_children():
                tree.delete(row)

            # Conecta ao banco de dados
            conexao = mysql.connector.connect(
                host="192.168.0.101",
                user="seu_usuario",
                password="sua_senha",
                database="nimalnotas"
            )
            cursor = conexao.cursor()

            if conexao.is_connected():
                print("Conexão bem-sucedida ao MySQL")

            # SQL básico com filtro opcional
            sql_query = """
                SELECT id, servico, cliente, data_inicio, data_fim, status, participantes, horas, assunto  
                FROM projetos
            """
            if valor:  # Se um valor for passado, filtra pela data
                sql_query += f" WHERE {coluna} LIKE %s "
                cursor.execute(sql_query, (valor,))
            else:  # Se não houver filtro, executa o SELECT sem condições
                cursor.execute(sql_query)

            # Carregar os resultados e exibir na Treeview
            resultados = cursor.fetchall()
            for i, linha in enumerate(resultados):
                tag = "odd" if i % 2 == 0 else "even"  # Alternando entre 'odd' e 'even'
                tree.insert("", tk.END, values=linha, tags=(tag,))

            # Fechar o cursor e a conexão
            conexao.commit()
            cursor.close()
            conexao.close()

        def aplicar_filtro():
            coluna_selecionada = combobox_colunas.get()  # Obtém o texto exibido na ComboBox
            coluna_real = opcoes.get(coluna_selecionada)  # Converte para o valor interno (nome real da coluna)

            valor_filtrado = entry_filtro.get().strip() + '%'

            if coluna_real:  # Verifica se a coluna foi selecionada corretamente
                # Ajusta a consulta para usar a coluna selecionada
                for row in tree.get_children():
                    tree.delete(row)

                conexao = mysql.connector.connect(
                    host="192.168.0.101",
                    user="seu_usuario",
                    password="sua_senha",
                    database="nimalnotas"
                )
                cursor = conexao.cursor()

                # Adicionando aspas ao redor dos nomes das colunas para evitar erro de sintaxe SQL
                sql_query = f"SELECT id, servico, cliente, data_inicio, data_fim, status, participantes, horas, assunto FROM projetos WHERE `{coluna_real}` LIKE %s"
                cursor.execute(sql_query, (valor_filtrado,))
                resultados = cursor.fetchall()

                for i, linha in enumerate(resultados):
                    tag = "odd" if i % 2 == 0 else "even"
                    tree.insert("", tk.END, values=linha, tags=(tag,))

                cursor.close()
                conexao.close()
            else:
                messagebox.showerror("Erro", "Por favor, selecione uma coluna válida para filtrar.")

                # Estilo do cabeçalho do Treeview

        style = ttk.Style()
        style.theme_use("alt")

        # Configurações de estilo para o cabeçalho
        style.configure("Treeview.Heading",
                        font=("Arial",10),
                        background="#4682B4",
                        foreground="white",
                        padding=(10, 5))

        # Estilo do corpo do Treeview
        style.configure("Treeview",
                        fieldbackground="#EEEEEE",
                        foreground="black",
                        rowheight=50,
                        borderwidth=100,
                        relief="flat")

        # Criando a tabela com Treeview
        colunas = (
            "ID do Projeto", "Serviço", "Cliente","Data Início", "Data Fim", "Status", "Participantes", "Horas",
            "Assunto")

        # Criando o Treeview
        tree = ttk.Treeview(visao_frame, columns=colunas, show="headings")
        tree.pack(fill=tk.BOTH, expand=True)


        # Configurar as tags para as linhas
        tree.tag_configure("odd", background="#E6E6E6")  # Cor para linhas ímpares
        tree.tag_configure("even", background="#EEEEEE")  # Cor para linhas pares

        for coluna in colunas:
            tree.heading(coluna, text=coluna)
            tree.column(coluna, width=50)

        # Carregar dados ao inicializar
        carregar_dados()

        # Botão exportar
        def exportar_para_excel():
            caminho_arquivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if caminho_arquivo:
                # Extraindo os dados da Treeview para um DataFrame
                dados = [tree.item(item)["values"] for item in tree.get_children()]
                df = pd.DataFrame(dados, columns=colunas)

                # Usando ExcelWriter para definir a linha inicial de escrita
                with pd.ExcelWriter(caminho_arquivo, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, startcol=1)

                messagebox.showinfo("Exportação", f"Dados exportados com sucesso para {caminho_arquivo}")

        def adicionar_dados():
            janela_edicao = ctk.CTkToplevel()
            janela_edicao.title("Adicionar Dados")
            janela.iconbitmap("nimal.ico")
            janela_edicao.geometry("900x800")
            janela_edicao.resizable(width=False, height=False)
            janela_edicao.grab_set()  # Isso impede interações com a janela principal enquanto a janela de adição estiver aberta
            janela_edicao.focus_force()

            # frames
            campos = [
                "Serviço", "Cliente", "Data de Início", "Data de Término", "Status", "Participantes",
                "Horas", "Assunto","Descrição"
            ]

            entradas = {}

            frameA = ctk.CTkFrame(janela_edicao, fg_color="#EEEEEE", width=900, height=700, corner_radius=10)
            frameA.place(relx=0.5, rely=0.43, anchor='center')

            frameC = ctk.CTkFrame(frameA, width=900, height=300, corner_radius=10, fg_color="#4682B4")
            frameC.place(relx=0.5, rely=0.1, anchor='center')

            frameB = ctk.CTkFrame(frameA, fg_color="#EEEEEE", width=900, height=900, corner_radius=10)
            frameB.place(relx=0.5, rely=0.8, anchor='center')

            # labels e entries

            logo_img_data = Image.open("nimall.png")
            logo_img = CTkImage(dark_image=logo_img_data, light_image=logo_img_data, size=(90, 90))
            img = ctk.CTkLabel(master=frameC, text="", image=logo_img)
            img.place(relx=0.8, rely=0.45, anchor='center')

            titulo = ctk.CTkLabel(frameC, text="Adicionar Projeto", font=("Arial Black", 24), text_color="white")
            titulo.place(relx=0.17, rely=0.4, anchor='w')

            CTkLabel(frameA, text="Serviços", font=("Arial", 12), text_color="black").place(relx=0.17, rely=0.21,
                                                                                            anchor='w')
            CTkLabel(frameA, text="Status", font=("Arial", 12), text_color="black").place(relx=0.57, rely=0.21,
                                                                                          anchor='w')
            CTkLabel(frameA, text="Cliente", font=("Arial", 12), text_color="black").place(relx=0.17, rely=0.31,
                                                                                           anchor='w')
            CTkLabel(frameA, text="Data de Início", font=("Arial", 12), text_color="black").place(relx=0.17, rely=0.41,
                                                                                                  anchor='w')
            CTkLabel(frameA, text="Data do Fim", font=("Arial", 12), text_color="black").place(relx=0.57, rely=0.41,
                                                                                               anchor='w')
            CTkLabel(frameA, text="Participantes", font=("Arial", 12), text_color="black").place(relx=0.17, rely=0.51,
                                                                                                 anchor='w')
            CTkLabel(frameA, text="Horas", font=("Arial", 12), text_color="black").place(relx=0.57, rely=0.51,
                                                                                         anchor='w')
            CTkLabel(frameA, text="Assunto", font=("Arial", 12), text_color="black").place(relx=0.17, rely=0.61,
                                                                                           anchor='w')
            CTkLabel(frameA, text="Descrição", font=("Arial", 12), text_color="black").place(relx=0.17, rely=0.71,
                                                                                             anchor='w')

            # Entrys
            servicos_opcoes = ["Call", "Visita", "Suporte Interno", "Suporte Externo"]
            entrada_servico = ctk.CTkComboBox(frameA, values=servicos_opcoes,button_color="#4682B4",font=("Arial", 12),border_color="#4682B4" ,width=250, height=30)
            entrada_servico.place(relx=0.305, rely=0.25, anchor='center')
            entrada_servico.set("Selecione uma opção")
            entradas["Serviço"] = entrada_servico  # Adiciona "Serviço" ao dicionário entradas

            status_opcoes = ["Concluído", "Em aberto", "Cancelado"]
            entrada_status = ctk.CTkComboBox(frameA, values=status_opcoes,button_color="#4682B4", font=("Arial", 12),border_color="#4682B4" ,width=250, height=30)
            entrada_status.place(relx=0.705, rely=0.25, anchor='center')
            entrada_status.set("Selecione uma opção")
            entradas["Status"] = entrada_status

            entrada_cliente = ctk.CTkEntry(frameA, font=("Arial", 12),border_color="#4682B4", width=600)
            entrada_cliente.place(relx=0.5, rely=0.35, anchor='center')
            entradas["Cliente"] = entrada_cliente

            entrada_inicio = DateEntry(frameA, width=18, font=("Arial", 12),border_color="#4682B4",date_pattern="dd/mm/yyyy")
            entrada_inicio.place(relx=0.27, rely=0.45, anchor='center')
            entradas["Data de Início"] = entrada_inicio

            entrada_fim = DateEntry(frameA, width=18, font=("Arial", 12),date_pattern="dd/mm/yyyy")
            entrada_fim.place(relx=0.67, rely=0.45, anchor='center')
            entradas["Data de Término"] = entrada_fim  # Adiciona "Data de Término" ao dicionário entradas

            participantes_opcoes = ["Ana", "Dowglas", "Jennifer", "Rui", "Tarso", "Giovanna"]
            entrada_participantes = ctk.CTkComboBox(frameA, values=participantes_opcoes,button_color="#4682B4",border_color="#4682B4" ,font=("Arial", 12), width=250,
                                                    height=30)
            entrada_participantes.place(relx=0.305, rely=0.55, anchor='center')
            entrada_participantes.set("Selecione uma opção")
            entradas["Participantes"] = entrada_participantes

            entrada_horas = ctk.CTkEntry(frameA, font=("Arial", 12),border_color="#4682B4",width=250)
            entrada_horas.place(relx=0.705, rely=0.55, anchor='center')
            entrada_horas.insert(0, "1 hora")
            entradas["Horas"] = entrada_horas

            entrada_assunto = ctk.CTkEntry(frameA, font=("Arial", 12),border_color="#4682B4", width=600)
            entrada_assunto.place(relx=0.5, rely=0.65, anchor='center')
            entradas["Assunto"] = entrada_assunto

            entrada_descricao = ctk.CTkTextbox(frameA, font=("Arial", 12), width=550, height=100,
                                               border_color="#4682B4")
            entrada_descricao.place(relx=0.5, rely=0.8, anchor='center')
            entradas["Descrição"] = entrada_descricao


            # Função para salvar as edições
            def confirmar_adicoes():
                novos_valores = []
                for campo in campos:
                    entrada = entradas[campo]
                    if campo == "Descrição":
                        novos_valores.append(entradas["Descrição"].get("1.0", "end-1c").strip())
                    else:
                        valor = entrada.get()  # Para entradas normais (Entry, Combobox, etc.)
                    # Verifica se o valor é uma string (para usar .strip())
                        if isinstance(valor, str):
                            novos_valores.append(valor.strip())  # Para strings, aplica .strip()
                        else:
                            novos_valores.append(valor)  # Para inteiros, adiciona diretamente

                # Verifica se todos os campos foram preenchidos
                if any(valores == "" for valores in novos_valores if isinstance(valores, str)):
                    messagebox.showwarning("Aviso", "Por favor, preencha todos os campos.")
                    return

                # Adiciona os dados ao Treeview
                tree.insert("", "end", values=novos_valores)

                # Conexão ao banco de dados
                conexao = mysql.connector.connect(
                    host="192.168.0.101",
                    user="seu_usuario",
                    password="sua_senha",
                    database="nimalnotas"
                )
                cursor = conexao.cursor()

                try:
                    # SQL para inserir os dados no banco de dados (sem o campo "ID")
                    sql_insert = """
                                                    INSERT INTO projetos (servico, cliente, data_inicio, data_fim, 
                                                                          status, participantes, horas, assunto,descricao)
                                                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s,%s)
                                                """
                    cursor.execute(sql_insert, tuple(novos_valores))
                    conexao.commit()
                    messagebox.showinfo("Sucesso", "Dados adicionados com sucesso!")
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao adicionar dados: {e}")
                finally:
                    cursor.close()
                    conexao.close()
                    carregar_dados()
                    atualizar_contagem()

                janela_edicao.destroy()

            # Botão para confirmar as alterações
            ctk.CTkButton(
                janela_edicao, text="Confirmar", command=confirmar_adicoes,
                fg_color="#4682B4", hover_color="#3d719d", width=200, height=40, font=("Arial", 14)
            ).place(relx=0.3, rely=0.85, anchor=tk.CENTER)

            # Botão para cancelar a edição
            ctk.CTkButton(
                janela_edicao, text="Cancelar", command=janela_edicao.destroy,
                fg_color="#4682B4", hover_color="#3d719d", width=200, height=40, font=("Arial", 14)
            ).place(relx=0.7, rely=0.85, anchor=tk.CENTER)

        def editar_dados():
            # Verifica se há uma linha selecionada
            item_selecionado = tree.selection()
            if not item_selecionado:
                messagebox.showwarning("Aviso", "Por favor, selecione uma linha para editar.")
                return

            # Recupera os valores da linha selecionada
            valores_atuais = tree.item(item_selecionado, "values")

            janela_edicao = ctk.CTkToplevel()
            janela_edicao.title("Editar Dados")
            janela.iconbitmap("nimal.ico")
            janela_edicao.geometry("900x800")
            janela_edicao.resizable(width=False, height=False)
            janela_edicao.grab_set()  # Isso impede interações com a janela principal enquanto a janela de adição estiver aberta
            janela_edicao.focus_force()

            # frames
            campos = [
                "ID", "Serviço", "Cliente", "Data de Início", "Data de Término", "Status", "Participantes",
                "Horas", "Assunto","Descrição"
            ]

            entradas = {}

            frameA = ctk.CTkFrame(janela_edicao, fg_color="#EEEEEE", width=900, height=700, corner_radius=10)
            frameA.place(relx=0.5, rely=0.43, anchor='center')

            frameC = ctk.CTkFrame(frameA, width=900, height=300, corner_radius=10, fg_color="#4682B4")
            frameC.place(relx=0.5, rely=0.1, anchor='center')

            frameB = ctk.CTkFrame(frameA, fg_color="#EEEEEE", width=900, height=900, corner_radius=10)
            frameB.place(relx=0.5, rely=0.8, anchor='center')

            # labels e entries

            logo_img_data = Image.open("nimall.png")
            logo_img = CTkImage(dark_image=logo_img_data, light_image=logo_img_data, size=(90, 90))
            img = ctk.CTkLabel(master=frameC, text="", image=logo_img)
            img.place(relx=0.8, rely=0.45, anchor='center')

            titulo = ctk.CTkLabel(frameC, text="Editar Projeto", font=("Arial Black", 24), text_color="white")
            titulo.place(relx=0.1, rely=0.4, anchor='w')

            idlabel = ctk.CTkLabel(frameC, text=f"ID: {valores_atuais[0]}", font=("Arial Black", 18), text_color="white")
            idlabel.place(relx=0.1, rely=0.5, anchor='w')

            CTkLabel(frameA, text="Serviços", font=("Arial", 12), text_color="black").place(relx=0.17, rely=0.21,
                                                                                            anchor='w')
            CTkLabel(frameA, text="Status", font=("Arial", 12), text_color="black").place(relx=0.57, rely=0.21,
                                                                                          anchor='w')
            CTkLabel(frameA, text="Cliente", font=("Arial", 12), text_color="black").place(relx=0.17, rely=0.31,
                                                                                           anchor='w')
            CTkLabel(frameA, text="Data de Início", font=("Arial", 12), text_color="black").place(relx=0.17, rely=0.41,
                                                                                                  anchor='w')
            CTkLabel(frameA, text="Data do Fim", font=("Arial", 12), text_color="black").place(relx=0.57, rely=0.41,
                                                                                               anchor='w')
            CTkLabel(frameA, text="Participantes", font=("Arial", 12), text_color="black").place(relx=0.17, rely=0.51,
                                                                                                 anchor='w')
            CTkLabel(frameA, text="Horas", font=("Arial", 12), text_color="black").place(relx=0.57, rely=0.51,
                                                                                         anchor='w')
            CTkLabel(frameA, text="Assunto", font=("Arial", 12), text_color="black").place(relx=0.17, rely=0.61,
                                                                                           anchor='w')
            CTkLabel(frameA, text="Descrição", font=("Arial", 12), text_color="black").place(relx=0.17, rely=0.71,
                                                                                           anchor='w')
            # Entrys
            servicos_opcoes = ["Call", "Visita", "Suporte Interno", "Suporte Externo"]
            entrada_servico = ctk.CTkComboBox(frameA, values=servicos_opcoes,button_color="#4682B4",border_color="#4682B4", font=("Arial", 12), width=250, height=30)
            entrada_servico.place(relx=0.305, rely=0.25, anchor='center')
            entrada_servico.set(valores_atuais[1])  # Valor inicial do campo
            entradas["Serviço"] = entrada_servico  # Adiciona "Serviço" ao dicionário entradas

            status_opcoes = ["Concluído", "Em aberto", "Cancelado"]
            entrada_status = ctk.CTkComboBox(frameA, values=status_opcoes,button_color="#4682B4",border_color="#4682B4", font=("Arial", 12), width=250, height=30)
            entrada_status.place(relx=0.705, rely=0.25, anchor='center')
            entrada_status.set(valores_atuais[5])  # Valor inicial do campo
            entradas["Status"] = entrada_status

            entrada_cliente = ctk.CTkEntry(frameA, font=("Arial", 12), width=600,border_color="#4682B4")
            entrada_cliente.place(relx=0.5, rely=0.35, anchor='center')
            entrada_cliente.insert(0, valores_atuais[2])  # Valor inicial do campo
            entradas["Cliente"] = entrada_cliente

            entrada_inicio = DateEntry(frameA, width=18, font=("Arial", 12), date_pattern="dd/mm/yyyy")
            entrada_inicio.place(relx=0.27, rely=0.45, anchor='center')
            entrada_inicio.set_date(valores_atuais[3])  # Valor inicial do campo
            entradas["Data de Início"] = entrada_inicio

            entrada_fim = DateEntry(frameA, width=18, font=("Arial", 12), date_pattern="dd/mm/yyyy")
            entrada_fim.place(relx=0.67, rely=0.45, anchor='center')
            entrada_fim.set_date(valores_atuais[4])  # Valor inicial do campo
            entradas["Data de Término"] = entrada_fim  # Adiciona "Data de Término" ao dicionário entradas

            participantes_opcoes = ["Ana", "Dowglas", "Jennifer", "Rui", "Tarso", "Giovanna"]
            entrada_participantes = ctk.CTkComboBox(frameA, values=participantes_opcoes,button_color="#4682B4",border_color="#4682B4", font=("Arial", 12), width=250,
                                                    height=30)
            entrada_participantes.place(relx=0.305, rely=0.55, anchor='center')
            entrada_participantes.set(valores_atuais[6])  # Valor inicial do campo
            entradas["Participantes"] = entrada_participantes

            entrada_horas = ctk.CTkEntry(frameA, font=("Arial", 12), width=250,border_color="#4682B4")
            entrada_horas.place(relx=0.705, rely=0.55, anchor='center')
            entrada_horas.insert(0, valores_atuais[7])  # Valor inicial do campo
            entradas["Horas"] = entrada_horas

            entrada_assunto = ctk.CTkEntry(frameA, font=("Arial", 12), width=600,border_color="#4682B4")
            entrada_assunto.place(relx=0.5, rely=0.65, anchor='center')
            entrada_assunto.insert(0, valores_atuais[8])  # Valor inicial do campo
            entradas["Assunto"] = entrada_assunto

            entrada_descricao = ctk.CTkTextbox(frameA, font=("Arial", 12), width=550,height=100, border_color="#4682B4")
            entrada_descricao.place(relx=0.5, rely=0.8, anchor='center')
            entradas["Descrição"] = entrada_descricao

            conexao = mysql.connector.connect(
                host="192.168.0.101",
                user="seu_usuario",
                password="sua_senha",
                database="nimalnotas"
            )
            cursor = conexao.cursor()

            # SQL SELECT para pegar a descrição
            sql_select = "SELECT descricao FROM projetos WHERE id = %s"
            cursor.execute(sql_select, (valores_atuais[0],))

            # Pega o resultado do SELECT
            descricao = cursor.fetchone()

            # Se a descrição for encontrada, insira na CTkTextbox
            if descricao:
                entrada_descricao.delete("1.0", "end")  # Apaga o conteúdo atual
                entrada_descricao.insert("1.0", str(descricao[0]))

            # Feche o cursor e a conexão após o uso
            cursor.close()
            conexao.close()

            # Função para salvar as edições
            def confirmar_edicoes():
                novos_valores = []

                # O ID já vem de valores_atuais[0], então não é necessário ter uma entrada para ele.
                novos_valores.append(valores_atuais[0])  # O ID original

                # Adiciona os valores das outras entradas
                for campo in campos:
                    if campo != "ID":  # Ignora o campo "ID", pois ele já foi adicionado
                        # Verifica se o campo é "Descrição" e usa o método get() de forma correta
                        if campo == "Descrição":
                            novos_valores.append(entradas["Descrição"].get("1.0", "end-1c").strip())
                        else:
                            novos_valores.append(entradas[campo].get().strip())

                # Atualiza os valores na Treeview
                tree.item(item_selecionado, values=novos_valores)

                # Atualiza o banco de dados
                conexao = mysql.connector.connect(
                    host="192.168.0.101",
                    user="seu_usuario",
                    password="sua_senha",
                    database="nimalnotas"
                )
                cursor = conexao.cursor()

                try:
                    # Corrigir a ordem dos valores, garantindo que o ID seja o último
                    sql_update = """
                                            UPDATE projetos
                                            SET servico = %s, cliente = %s, data_inicio = %s, data_fim = %s, 
                                                status = %s, participantes = %s, horas = %s, assunto = %s, descricao = %s
                                            WHERE id = %s
                                        """
                    cursor.execute(sql_update, (*novos_valores[1:], novos_valores[0]))  # ID vai como último parâmetro
                    conexao.commit()
                    messagebox.showinfo("Sucesso", "Dados atualizados com sucesso!")
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao atualizar dados: {e}")
                finally:
                    cursor.close()
                    conexao.close()

                atualizar_contagem()
                janela_edicao.destroy()

            # Botão para confirmar as alterações
            ctk.CTkButton(
                janela_edicao, text="Confirmar", command=confirmar_edicoes,
                fg_color="#4682B4", hover_color="#3d719d", width=200, height=40, font=("Arial", 14)
            ).place(relx=0.3, rely=0.85, anchor=tk.CENTER)

            # Botão para cancelar a edição
            ctk.CTkButton(
                janela_edicao, text="Cancelar", command=janela_edicao.destroy,
                fg_color="#4682B4", hover_color="#3d719d", width=200, height=40, font=("Arial", 14)
            ).place(relx=0.7, rely=0.85, anchor=tk.CENTER)

        def remover_dados():
            item_selecionado = tree.selection()
            if not item_selecionado:
                messagebox.showwarning("Aviso", "Por favor, selecione uma linha para remover.")
                return
            selecionado = tree.item(tree.selection())["values"][0]
            resposta = messagebox.askyesno("Confirmação",
                                           f"Tem certeza que deseja remover o pedido {selecionado} do banco de dados?")
            if resposta:
                try:
                    conexao = mysql.connector.connect(
                        host="192.168.0.101",
                        user="seu_usuario",
                        password="sua_senha",
                        database="nimalnotas"
                    )
                    cursor = conexao.cursor()

                    # Remover a linha selecionada
                    sql_delete = "DELETE FROM projetos WHERE id = %s"
                    cursor.execute(sql_delete, (selecionado,))

                    conexao.commit()
                    cursor.close()
                    conexao.close()

                    # Recarregar a visão geral após a remoção
                    carregar_dados()
                    atualizar_contagem()

                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao remover dados: {e}")

        def contar_elementos():
            # Conexão com o banco de dados
            conexao = mysql.connector.connect(
                host="192.168.0.101",  # Substitua pelo IP do seu banco de dados
                user="seu_usuario",  # Substitua pelo seu usuário do banco
                password="sua_senha",  # Substitua pela sua senha
                database="nimalnotas"  # Nome do banco de dados
            )
            cursor = conexao.cursor()

            # Executando o comando SQL para contar os elementos
            query = "SELECT COUNT(*) FROM projetos"  # Substitua 'projetos' pelo nome da sua tabela
            cursor.execute(query)
            resultado = cursor.fetchone()[0]

            # Fechando a conexão
            cursor.close()
            conexao.close()


            return resultado

        def contar_concluidos():
            # Conexão com o banco de dados
            conexao = mysql.connector.connect(
                host="192.168.0.101",  # Substitua pelo IP do seu banco de dados
                user="seu_usuario",  # Substitua pelo seu usuário do banco
                password="sua_senha",  # Substitua pela sua senha
                database="nimalnotas"  # Nome do banco de dados
            )
            cursor = conexao.cursor()

            # Executando o comando SQL para contar os elementos com "Status = Concluído"
            query = "SELECT COUNT(*) FROM projetos WHERE status = 'Concluído'"  # Ajuste 'projetos' e 'status' se necessário
            cursor.execute(query)
            resultado = cursor.fetchone()[0]

            # Fechando a conexão
            cursor.close()
            conexao.close()


            return resultado

        def contar_em_aberto():
            # Conexão com o banco de dados
            conexao = mysql.connector.connect(
                host="192.168.0.101",  # Substitua pelo IP do seu banco de dados
                user="seu_usuario",  # Substitua pelo seu usuário do banco
                password="sua_senha",  # Substitua pela sua senha
                database="nimalnotas"  # Nome do banco de dados
            )
            cursor = conexao.cursor()

            # Executando o comando SQL para contar os elementos com "Status = Em Aberto"
            query = "SELECT COUNT(*) FROM projetos WHERE status = 'Em Aberto'"  # Ajuste 'projetos' e 'status' se necessário
            cursor.execute(query)
            resultado = cursor.fetchone()[0]

            # Fechando a conexão
            cursor.close()
            conexao.close()


            return resultado

        def atualizar_contagem():
            try:
                conexao = mysql.connector.connect(
                    host="192.168.0.101",
                    user="seu_usuario",
                    password="sua_senha",
                    database="nimalnotas"
                )
                cursor = conexao.cursor()

                cursor.execute("SELECT COUNT(*) FROM projetos WHERE status = 'Concluído'")
                total_concluidos = cursor.fetchone()[0]

                cursor.execute("SELECT COUNT(*) FROM projetos WHERE status = 'Em Aberto'")
                total_em_aberto = cursor.fetchone()[0]

                cursor.execute("SELECT COUNT(*) FROM projetos;")
                total_elementos = cursor.fetchone()[0]


                label1.configure(text=f"Total: {total_elementos}")
                label2.configure(text=f"Concluídos: {total_concluidos}")
                label3.configure(text=f"Em Aberto: {total_em_aberto}")

            except mysql.connector.Error as e:
                print(f"Erro ao atualizar contagem: {e}")
            finally:
                if cursor:
                    cursor.close()
                if conexao:
                    conexao.close()

        def ver_detalhes():
            # Verifica se há uma linha selecionada
            item_selecionado = tree.selection()
            if not item_selecionado:
                messagebox.showwarning("Aviso", "Por favor, selecione uma linha para visualizar.")
                return

            # Recupera o ID da linha selecionada
            valores_atuais = tree.item(item_selecionado, "values")
            id_projeto = valores_atuais[0]

            # Conectar ao banco de dados para pegar os detalhes
            conexao = mysql.connector.connect(
                host="192.168.0.101",
                user="seu_usuario",
                password="sua_senha",
                database="nimalnotas"
            )
            cursor = conexao.cursor()

            sql_select = """SELECT servico, cliente, data_inicio, data_fim, status, participantes, horas, assunto 
                            FROM projetos WHERE id = %s"""
            cursor.execute(sql_select, (id_projeto,))
            detalhes = cursor.fetchone()

            cursor.close()
            conexao.close()

            if not detalhes:
                messagebox.showerror("Erro", "Não foi possível encontrar os detalhes do projeto.")
                return

            # Criar a janela de detalhes
            janela_detalhes = ctk.CTkToplevel()
            janela_detalhes.title("Detalhes do Projeto")
            janela_detalhes.geometry("900x800")
            janela_detalhes.resizable(width=False, height=False)
            janela_detalhes.configure(fg_color="#4682B4")
            janela_detalhes.grab_set()
            janela_detalhes.focus_force()

            frameA = ctk.CTkFrame(janela_detalhes, fg_color="#EEEEEE", width=900, height=700, corner_radius=10)
            frameA.place(relx=0.5, rely=0.43, anchor='center')

            frameC = ctk.CTkFrame(frameA, width=900, height=300, corner_radius=10, fg_color="#4682B4")
            frameC.place(relx=0.5, rely=0.1, anchor='center')

            frameB = ctk.CTkFrame(frameA, fg_color="#EEEEEE", width=900, height=900, corner_radius=10)
            frameB.place(relx=0.5, rely=0.8, anchor='center')

            logo_img_data = Image.open("nimall.png")
            logo_img = CTkImage(dark_image=logo_img_data, light_image=logo_img_data, size=(90, 90))
            img = ctk.CTkLabel(master=frameC, text="", image=logo_img)
            img.place(relx=0.8, rely=0.45, anchor='center')

            titulo = ctk.CTkLabel(frameC, text="Detalhes do Projeto", font=("Arial Black", 24), text_color="white")
            titulo.place(relx=0.3, rely=0.4, anchor='w')


            # Criar labels para exibir os detalhes
            campos = ["Serviço", "Cliente", "Data de Início", "Data de Término", "Status", "Participantes", "Horas",
                      "Assunto", "Descrição"]

            label_servico = ctk.CTkLabel(frameA, text="Serviço:", font=("Arial", 14, "bold"),
                                         text_color="black")
            label_servico.place(relx=0.3, rely=0.2, anchor='w')

            valor_servico = ctk.CTkLabel(frameA, text=detalhes[0], font=("Arial", 14), text_color="black",
                                         wraplength=500, anchor="w", justify="left")
            valor_servico.place(relx=0.6, rely=0.2, anchor='w')

            label_cliente = ctk.CTkLabel(frameA, text="Cliente:", font=("Arial", 14, "bold"),
                                         text_color="black")
            label_cliente.place(relx=0.3, rely=0.25, anchor='w')

            valor_cliente = ctk.CTkLabel(frameA, text=detalhes[1], font=("Arial", 14), text_color="black",
                                         wraplength=500, anchor="w", justify="left")
            valor_cliente.place(relx=0.6, rely=0.25, anchor='w')

            label_data_inicio = ctk.CTkLabel(frameA, text="Data de Início:", font=("Arial", 14, "bold"),
                                             text_color="black")
            label_data_inicio.place(relx=0.3, rely=0.3, anchor='w')

            valor_data_inicio = ctk.CTkLabel(frameA, text=detalhes[2], font=("Arial", 14), text_color="black",
                                             wraplength=500, anchor="w", justify="left")
            valor_data_inicio.place(relx=0.6, rely=0.3, anchor='w')

            label_data_fim = ctk.CTkLabel(frameA, text="Data de Término:", font=("Arial", 14, "bold"),
                                          text_color="black")
            label_data_fim.place(relx=0.3, rely=0.35, anchor='w')

            valor_data_fim = ctk.CTkLabel(frameA, text=detalhes[3], font=("Arial", 14), text_color="black",
                                          wraplength=500, anchor="w", justify="left")
            valor_data_fim.place(relx=0.6, rely=0.35, anchor='w')

            label_status = ctk.CTkLabel(frameA, text="Status:", font=("Arial", 14, "bold"), text_color="black")
            label_status.place(relx=0.3, rely=0.4, anchor='w')

            valor_status = ctk.CTkLabel(frameA, text=detalhes[4], font=("Arial", 14), text_color="black",
                                        wraplength=500, anchor="w", justify="left")
            valor_status.place(relx=0.6, rely=0.4, anchor='w')

            label_participantes = ctk.CTkLabel(frameA, text="Participantes:", font=("Arial", 14, "bold"),
                                               text_color="black")
            label_participantes.place(relx=0.3, rely=0.45, anchor='w')

            valor_participantes = ctk.CTkLabel(frameA, text=detalhes[5], font=("Arial", 14), text_color="black",
                                               wraplength=500, anchor="w", justify="left")
            valor_participantes.place(relx=0.6, rely=0.45, anchor='w')

            label_horas = ctk.CTkLabel(frameA, text="Horas:", font=("Arial", 14, "bold"), text_color="black")
            label_horas.place(relx=0.3, rely=0.5, anchor='w')

            valor_horas = ctk.CTkLabel(frameA, text=detalhes[6], font=("Arial", 14), text_color="black",
                                       wraplength=500, anchor="w", justify="left")
            valor_horas.place(relx=0.6, rely=0.5, anchor='w')

            label_assunto = ctk.CTkLabel(frameA, text="Assunto:", font=("Arial", 14, "bold"),
                                         text_color="black")
            label_assunto.place(relx=0.3, rely=0.55, anchor='w')

            valor_assunto = ctk.CTkLabel(frameA, text=detalhes[7], font=("Arial", 14), text_color="black",
                                         wraplength=500, anchor="w", justify="left")
            valor_assunto.place(relx=0.6, rely=0.55, anchor='w')

            label_descricao = ctk.CTkLabel(frameA, text="Descrição:", font=("Arial", 14, "bold"),
                                           text_color="black")
            label_descricao.place(relx=0.3, rely=0.6, anchor='w')

            valor_descricao = ctk.CTkTextbox(frameA,font=("Arial", 14),width=450,height=160, border_color="#4682B4")
            valor_descricao.place(relx=0.25, rely=0.75, anchor='w')

            conexao = mysql.connector.connect(
                host="192.168.0.101",
                user="seu_usuario",
                password="sua_senha",
                database="nimalnotas"
            )
            cursor = conexao.cursor()

            # SQL SELECT para pegar a descrição
            sql_select = "SELECT descricao FROM projetos WHERE id = %s"
            cursor.execute(sql_select, (valores_atuais[0],))

            # Pega o resultado do SELECT
            descricao = cursor.fetchone()

            # Se a descrição for encontrada, insira na CTkTextbox

            valor_descricao.insert("1.0", str(descricao[0]))

            # Feche o cursor e a conexão após o uso
            cursor.close()
            conexao.close()


            # Botão para fechar a janela
            botao_fechar = ctk.CTkButton(frameA, text="Fechar", command=janela_detalhes.destroy,
                                         fg_color="#4682B4",width=200, height=50)
            botao_fechar.place(relx=0.5, rely=0.93, anchor='center')


        #Imagens

        logo_img_data = Image.open("nimall.png")
        logo_img = CTkImage(dark_image=logo_img_data, light_image=logo_img_data, size=(130,130))
        CTkLabel(master=frame1, text="", image=logo_img).pack(pady=(100, 120), anchor="center")

        img1 = Image.open("logistics_icon.png")
        img1 = CTkImage(dark_image=img1, light_image=img1, size=(45,45))
        CTkLabel(master=frameA, text="", image=img1).place(x=20,rely=0.2)

        img1 = Image.open("delivered_icon.png")
        img1 = CTkImage(dark_image=img1, light_image=img1, size=(45, 45))
        CTkLabel(master=frameB, text="", image=img1).place(x=20, rely=0.2)

        img1 = Image.open("shipping_icon.png")
        img1 = CTkImage(dark_image=img1, light_image=img1, size=(45, 45))
        CTkLabel(master=frameC, text="", image=img1).place(x=20, rely=0.2)

        total_elementos = contar_elementos()
        total_concluidos = contar_concluidos()
        total_em_aberto = contar_em_aberto()


        # Exibindo o resultado
        label1 = tk.Label(frameA, text=f"Total: {total_elementos}",
                         font=("Arial Black", 14), fg="white",bg="#4682B4")
        label1.place(x=80, rely=0.25)

        label2 = tk.Label(frameB, text=f"Concluídos: {total_concluidos}",
                         font=("Arial Black", 14), fg="white", bg="#4682B4")
        label2.place(x=80, rely=0.25)

        label3 = tk.Label(frameC, text=f"Em Aberto: {total_em_aberto}",
                          font=("Arial Black", 14), fg="white", bg="#4682B4")
        label3.place(x=80, rely=0.25)

        entry_filtro = ctk.CTkEntry(frameD, width=400, font=("Arial", 12),border_color="#4682B4")
        entry_filtro.place(x=40, rely=0.3)


        #botões

        botao_editar = ctk.CTkButton(frame1, text="Editar", command=editar_dados, fg_color="#4682B4", hover_color="#3d719d",
                                     width=200, height=50, font=("Arial Bold", 16),image=imgb1,anchor="w")
        botao_editar.pack(pady=5, padx=5)


        botao_remover = ctk.CTkButton(frame1, text="Remover", command=remover_dados, fg_color="#4682B4",
                                      hover_color="#3d719d", width=200, height=50, font=("Arial Bold", 16), image=imgb2,anchor="w")
        botao_remover.pack(pady=5, padx=5)


        botao_adicionar = ctk.CTkButton(frame1, text="Adicionar", command=adicionar_dados, fg_color="#4682B4",
                                        hover_color="#3d719d", width=200, height=50, font=("Arial Bold", 16),image=imgb3,anchor="w")
        botao_adicionar.pack(pady=5, padx=5)


        botao_exportar = ctk.CTkButton(frame1, text="Gerar Relatório", command=exportar_para_excel, fg_color="#4682B4",
                                       hover_color="#3d719d", width=200, height=50, font=("Arial Bold", 16),image=imgb4,anchor="w")
        botao_exportar.pack(pady=5, padx=5)

        botao_ver_detalhes  = ctk.CTkButton(frame1, text="Ver Detalhes", command=ver_detalhes, fg_color="#4682B4",
                                       hover_color="#3d719d", width=200, height=50, font=("Arial Bold", 16),
                                       image=imgb6, anchor="w")
        botao_ver_detalhes .pack(pady=5, padx=(20, 20))

        botao_filtrar = ctk.CTkButton(frameD, text="Pesquisar", command=aplicar_filtro, fg_color="#4682B4",
                                      hover_color="#3d719d", width=200, height=30, font=("Arial Bold", 16),image=imgb5,anchor="w")
        botao_filtrar.place(x=800, rely=0.5, anchor="center")


        opcoes = {
            "Id": "id",
            "Serviço": "servico",
            "Cliente": "cliente",
            "Data de Início": "data_inicio",
            "Data de Término": "data_fim",
            "Status": "status",
            "Participantes": "participantes",
            "Horas": "horas",
            "Assunto": "assunto"
        }


        combobox_colunas = ctk.CTkComboBox(
            frameD,
            values=list(opcoes.keys()),  # Exibe as chaves do dicionário
            fg_color="#EEEEEE",
            border_color="#4682B4",
            width=200,
            text_color="black",
            button_color="#4682B4",
            button_hover_color="#3d719d",
            dropdown_hover_color="#EEEEEE"
        )
        combobox_colunas.place(x=470, rely=0.3)
        combobox_colunas.set("Selecione uma opção")
def get_screen_size():
    screen_width = janela.winfo_screenwidth()
    screen_height = janela.winfo_screenheight()
    return screen_width, screen_height
def centralizar_janela(largura, altura):
    # Obter dimensões da tela
    largura_tela = janela.winfo_screenwidth()
    altura_tela = janela.winfo_screenheight()

    # Calcular posição x e y para centralizar
    pos_x = (largura_tela - largura) // 2
    pos_y = (altura_tela - altura) // 2

    # Retorna a posição como string
    return f"{largura}x{altura}+{pos_x}+{pos_y}"

janela = ctk.CTk()
screen_width, screen_height = get_screen_size()
dimensoes = centralizar_janela(screen_width, screen_height)
janela.geometry(dimensoes)
janela.resizable(True, True)
janela.title("NimalResearch")
janela.state('zoomed')
janela.iconbitmap("nimal.ico")
ctk.set_appearance_mode("light")

imagem_pdf = ctk.CTkImage(Image.open("pdf.png"), size=(30, 30))
imagem_logo = ctk.CTkImage(Image.open("nimal.ico"), size=(80, 80))
imgb1 = ctk.CTkImage(Image.open("icons8-editar-50.png"), size=(30, 30))
imgb2 = ctk.CTkImage(Image.open("icons8-apagar-para-sempre-24.png"), size=(30, 30))
imgb3 = ctk.CTkImage(Image.open("icons8-adicionar-50.png"), size=(30, 30))
imgb4 = ctk.CTkImage(Image.open("icons8-lista-de-arquivo-de-peças-30.png"), size=(30, 30))
imgb5 = ctk.CTkImage(Image.open("icons8-pesquisar-64.png"), size=(20, 20))
imgb6 = ctk.CTkImage(Image.open("icons8-lista-64.png"), size=(30, 30))

frame1 = ctk.CTkFrame(janela, fg_color="#4682B4", width=100, height=650,corner_radius=0)
frame1.pack(fill="y", anchor="w", side="left")

frame3 = ctk.CTkFrame(janela,fg_color="#E6E6E6", width=1000, height=220, corner_radius=10)
frame3.place(relx=0.6, rely=0.18, anchor='center')

frameA = ctk.CTkFrame(frame3,fg_color="#4682B4", width=300, height=70, corner_radius=10)
frameA.place(relx=0.18, rely=0.4, anchor='center')

frameB = ctk.CTkFrame(frame3,fg_color="#4682B4", width=300, height=70, corner_radius=10)
frameB.place(relx=0.50, rely=0.4, anchor='center')

frameC = ctk.CTkFrame(frame3,fg_color="#4682B4", width=300, height=70, corner_radius=10)
frameC.place(relx=0.82, rely=0.4, anchor='center')

frameD = ctk.CTkFrame(frame3,fg_color="#EEEEEE", width=950, height=70, corner_radius=10)
frameD.place(relx=0.5, rely=0.8, anchor='center')

frame2 = ctk.CTkFrame(janela, fg_color="#33415c", width=500, height=280, corner_radius=10)
frame2.place(relx=0.6, rely=0.68, anchor='center')

(CTkLabel(master=frame3, text="Projetos", font=("Arial Black", 30), text_color="#4682B4")
.place(x = 30 , y = 10))

mostrar_visao_geral()

janela.mainloop()
