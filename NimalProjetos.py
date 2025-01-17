import mysql.connector
from customtkinter import CTkImage, CTkLabel
from openpyxl import load_workbook
import os
from PIL import Image
import pandas as pd
import tkinter as tk
from tkinter import ttk
import customtkinter as ctk
from tkinter import messagebox
from tkinter import filedialog
from tkcalendar import DateEntry

conexao = mysql.connector.connect(
    host="",
    user="seu_usuario",
    password="sua_senha",
    database="nimalnotas"
)

cursor = conexao.cursor()

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


def mostrar_visao_geral():
    global tree, frame2

    frame2.configure(width=1000, height=500)
    frame2.pack_propagate(False)
    # Limpa o conteúdo anterior no frame de menu
    for widget in frame2.winfo_children():
        widget.destroy()

    # Frame para a visão geral
    visao_frame = ctk.CTkFrame(frame2, fg_color="gray")
    visao_frame.place(relwidth=1, relheight=1)

    def carregar_dados(valor=None,coluna=None):
        for row in tree.get_children():
            tree.delete(row)

        # Conecta ao banco de dados
        conexao = mysql.connector.connect(
            host="",
            user="seu_usuario",
            password="sua_senha",
            database="nimalnotas"
        )
        cursor = conexao.cursor()

        if conexao.is_connected():
            print("Conexão bem-sucedida ao MySQL")

        # SQL básico com filtro opcional
        sql_query = """
            SELECT id, servico, cliente, visita, data_inicio, data_fim, status, participantes, horas, assunto  
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
                host="",
                user="seu_usuario",
                password="sua_senha",
                database="nimalnotas"
            )
            cursor = conexao.cursor()

            # Adicionando aspas ao redor dos nomes das colunas para evitar erro de sintaxe SQL
            sql_query = f"SELECT id, servico, cliente, visita, data_inicio, data_fim, status, participantes, horas, assunto FROM projetos WHERE `{coluna_real}` LIKE %s"
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
        "Id do Projeto", "Serviço", "Cliente", "Visita", "Data Início", "Data Fim", "Status", "Participantes", "Horas",
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
        janela_adicao = ctk.CTkToplevel()
        janela_adicao.title("Adicionar Dados")
        janela_adicao.geometry("500x700")

        janela_adicao.grab_set()  # Isso impede interações com a janela principal enquanto a janela de adição estiver aberta
        janela_adicao.focus_force()

        # Labels e entradas para os novos dados
        campos = [
            "Serviço", "Cliente", "Visita", "Data de Início", "Data de Término", "Status", "Participantes",
            "Horas", "Assunto"
        ]

        entradas = {}
        for i, campo in enumerate(campos):
            frame_linha = ctk.CTkFrame(janela_adicao)
            frame_linha.pack(fill=tk.X, padx=10, pady=5)

            ctk.CTkLabel(frame_linha, text=campo.capitalize(), font=("Arial", 12), width=15, anchor="w").pack(
                side=tk.LEFT, padx=5)

            if campo == "Serviço":
                # Combobox para "Serviço"
                servicos_opcoes = ["Call", "Visita", "Suporte Interno", "Suporte Externo"]
                entrada = ctk.CTkComboBox(frame_linha, values=servicos_opcoes, font=("Arial", 12))
                entrada.set(servicos_opcoes[0])  # Define o valor padrão como "Call"

            elif campo == "Participantes":
                # Combobox para participantes
                participantes_opcoes = ["Ana", "Dowglas", "Jennifer", "Rui", "Tarso", "Giovanna"]
                entrada = ctk.CTkComboBox(frame_linha, values=participantes_opcoes, font=("Arial", 12))
                entrada.set(participantes_opcoes[0])  # Define o valor padrão como "Ana"
            elif campo == "Horas":
                # Entrada de horas com valor padrão "1 hora"
                entrada = ctk.CTkEntry(frame_linha, font=("Arial", 12))
                entrada.insert(0, "1 hora")  # Valor padrão para horas
            elif campo == "Data de Início" or campo == "Data de Término":
                # Usando o DateEntry do tkcalendar com estilo customizado
                entrada = DateEntry(frame_linha, width=18, font=("Arial", 12), date_pattern="dd/mm/yyyy")
                # Personalizando a aparência para combinar com o CustomTkinter

                entrada.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=5)
            else:
                # Entrada padrão para os outros campos
                entrada = ctk.CTkEntry(frame_linha, font=("Arial", 12))

            entrada.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=5)
            entradas[campo] = entrada

        # Função para confirmar a adição dos dados
        def confirmar_adicao():
            novos_valores = []

            for campo in campos:
                entrada = entradas[campo]

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
                host="",
                user="seu_usuario",
                password="sua_senha",
                database="nimalnotas"
            )
            cursor = conexao.cursor()

            try:
                # SQL para inserir os dados no banco de dados (sem o campo "ID")
                sql_insert = """
                    INSERT INTO projetos (servico, cliente, visita, data_inicio, data_fim, 
                                          status, participantes, horas, assunto)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
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

            janela_adicao.destroy()

        # Botão para confirmar a adição
        botao_confirmar = ctk.CTkButton(janela_adicao, text="Adicionar", command=confirmar_adicao, fg_color="#4CAF50",
                                        hover_color="#45a049", width=150, height=40, font=("Arial", 12))
        botao_confirmar.pack(pady=10)

        # Botão para fechar a janela
        botao_fechar = ctk.CTkButton(janela_adicao, text="Fechar", command=janela_adicao.destroy, fg_color="#f44336",
                                     hover_color="#e53935", width=150, height=40, font=("Arial", 12))
        botao_fechar.pack(pady=5)

    def editar_dados():
        # Verifica se há uma linha selecionada
        item_selecionado = tree.selection()
        if not item_selecionado:
            messagebox.showwarning("Aviso", "Por favor, selecione uma linha para editar.")
            return

        # Recupera os valores da linha selecionada
        valores_atuais = tree.item(item_selecionado, "values")

        # Cria uma nova janela para edição
        janela_edicao = ctk.CTkToplevel()
        janela_edicao.title("Editar Dados")
        janela_edicao.geometry("700x700")

        # Labels e entradas para edição dos valores
        campos = [
            "ID","Serviço", "Cliente", "Visita", "Data de Início", "Data de Término", "Status", "Participantes",
            "Horas", "Assunto"
        ]

        entradas = {}
        for i, campo in enumerate(campos):
            frame_linha = ctk.CTkFrame(janela_edicao)
            frame_linha.pack(fill=tk.X, padx=10, pady=5)

            ctk.CTkLabel(frame_linha, text=campo.capitalize(), font=("Arial", 12), width=15, anchor="w").pack(
                side=tk.LEFT, padx=5)
            entrada = ctk.CTkEntry(frame_linha, font=("Arial", 12))
            entrada.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=5)
            entrada.insert(0, valores_atuais[i])  # Preenche com o valor atual
            entradas[campo] = entrada

        # Função para salvar as edições
        def confirmar_edicoes():
            novos_valores = [entradas[campo].get().strip() for campo in campos]

            # Atualiza os valores na Treeview
            tree.item(item_selecionado, values=novos_valores)

            # Atualiza o banco de dados
            conexao = mysql.connector.connect(
                host="",
                user="seu_usuario",
                password="sua_senha",
                database="nimalnotas"
            )
            cursor = conexao.cursor()

            try:
                sql_update = """
                    UPDATE projetos
                    SET id = %s, servico = %s, cliente = %s, visita = %s, data_inicio = %s, data_fim = %s, 
                        status = %s, participantes = %s, horas = %s, assunto = %s
                    WHERE id = %s
                """
                cursor.execute(sql_update,
                               (*novos_valores, valores_atuais[1]))  # Usa o orçamento como identificador único
                conexao.commit()
                messagebox.showinfo("Sucesso", "Dados atualizados com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao atualizar dados: {e}")
            finally:
                cursor.close()
                conexao.close()

            janela_edicao.destroy()

        # Botão para confirmar as alterações
        ctk.CTkButton(
            janela_edicao, text="Confirmar", command=confirmar_edicoes,
            fg_color="#001427", hover_color="#4361ee", width=200, height=40, font=("Impact", 14)
        ).pack(pady=20)

        # Botão para cancelar a edição
        ctk.CTkButton(
            janela_edicao, text="Cancelar", command=janela_edicao.destroy,
            fg_color="#001427", hover_color="#e63946", width=200, height=40, font=("Impact", 14)
        ).pack(pady=10)

    def remover_dados():
        global selecionado

        selecionado = tree.item(tree.selection())["values"][0]

        resposta = messagebox.askyesno("Confirmação",
                                       f"Tem certeza que deseja remover o pedido {selecionado} do banco de dados?")
        if resposta:
            try:
                conexao = mysql.connector.connect(
                    host="",
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
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao remover dados: {e}")

    def contar_elementos():
        # Conexão com o banco de dados
        conexao = mysql.connector.connect(
            host="",  # Substitua pelo IP do seu banco de dados
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
        carregar_dados()

        return resultado

    def contar_concluidos():
        # Conexão com o banco de dados
        conexao = mysql.connector.connect(
            host="",  # Substitua pelo IP do seu banco de dados
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
        carregar_dados()

        return resultado

    def contar_em_aberto():
        # Conexão com o banco de dados
        conexao = mysql.connector.connect(
            host="",  # Substitua pelo IP do seu banco de dados
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
        carregar_dados()

        return resultado



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

    entry_filtro = ctk.CTkEntry(frameD, width=400, font=("Arial", 12))
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
    botao_exportar.pack(pady=5, padx=(20,20))

    botao_filtrar = ctk.CTkButton(frameD, text="Pesquisar", command=aplicar_filtro, fg_color="#4682B4",
                                  hover_color="#3d719d", width=200, height=30, font=("Arial Bold", 16),image=imgb5,anchor="w")
    botao_filtrar.place(x=800, rely=0.5, anchor="center")

    opcoes = {
        "Id": "id",
        "Serviço": "servico",
        "Cliente": "cliente",
        "Visita": "visita",
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


janela = ctk.CTk()
screen_width, screen_height = get_screen_size()
janela.geometry(f"{screen_width}x{screen_height}")
janela.resizable(True, True)
janela.title("NimalResearch")
janela.state('zoomed')
janela.iconbitmap("nimal.ico")
ctk.set_appearance_mode("light")
bg_label = ctk.CTkLabel(janela)
bg_label.place(relwidth=1, relheight=1)

imagem_pdf = ctk.CTkImage(Image.open("pdf.png"), size=(30, 30))
imagem_logo = ctk.CTkImage(Image.open("nimal.ico"), size=(80, 80))
imgb1 = ctk.CTkImage(Image.open("icons8-editar-50.png"), size=(30, 30))
imgb2 = ctk.CTkImage(Image.open("icons8-apagar-para-sempre-24.png"), size=(30, 30))
imgb3 = ctk.CTkImage(Image.open("icons8-adicionar-50.png"), size=(30, 30))
imgb4 = ctk.CTkImage(Image.open("icons8-lista-de-arquivo-de-peças-30.png"), size=(30, 30))
imgb5 = ctk.CTkImage(Image.open("icons8-pesquisar-64.png"), size=(20, 20))


frame1 = ctk.CTkFrame(janela, fg_color="#4682B4", width=100, height=650,corner_radius=0)
frame1.pack(fill="y", anchor="w", side="left")

frame2 = ctk.CTkFrame(janela, fg_color="#33415c", width=500, height=280, corner_radius=10)
frame2.place(relx=0.6, rely=0.65, anchor='center')

frame3 = ctk.CTkFrame(janela,fg_color="#E6E6E6", width=1000, height=300, corner_radius=10)
frame3.place(relx=0.6, rely=0.2, anchor='center')

frameA = ctk.CTkFrame(frame3,fg_color="#4682B4", width=300, height=70, corner_radius=10)
frameA.place(relx=0.18, rely=0.4, anchor='center')

frameB = ctk.CTkFrame(frame3,fg_color="#4682B4", width=300, height=70, corner_radius=10)
frameB.place(relx=0.50, rely=0.4, anchor='center')

frameC = ctk.CTkFrame(frame3,fg_color="#4682B4", width=300, height=70, corner_radius=10)
frameC.place(relx=0.82, rely=0.4, anchor='center')

frameD = ctk.CTkFrame(frame3,fg_color="#EEEEEE", width=950, height=70, corner_radius=10)
frameD.place(relx=0.5, rely=0.8, anchor='center')


CTkLabel(master=frame3, text="Projetos", font=("Arial Black", 30), text_color="#4682B4").place(x = 30 , y = 10)

mostrar_visao_geral()

janela.mainloop()
