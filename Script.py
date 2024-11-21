import html  # Para converter as entidades HTML de volta para caracteres
import os
import re  # Usado para remover as tags HTML
from tkinter import Tk, filedialog, Label, Button
import pandas as pd
import xlwings as xw
from sqlalchemy import create_engine


# Função para conectar ao banco de dados usando SQLAlchemy e pymysql
def conectar_banco():
    try:
        # Criação da URL de conexão com SQLAlchemy
        engine = create_engine('mysql+pymysql://user:senha@endereco-do-database:port/name-database?charset=utf8mb4')
        conn = engine.connect()
        return conn
    except Exception as err:
        print(f"Erro ao conectar: {err}")
        return None

# Função para remover todas as tags HTML
def remover_tags_html(texto):
    if isinstance(texto, str):
        texto = html.unescape(texto)  # Decodificar entidades HTML
        texto = re.sub(r'<.*?>', '', texto)  # Remover as tags HTML
    return texto

# Função para exportar dados para uma planilha Excel
def exportar_para_modelo_excel(status_label, root):
    conn = conectar_banco()
    if conn is None:
        return

    try:
        # Exibir "Gerando Relatório"
        status_label.config(text="Gerando Relatório...")
        status_label.update()

        # Consulta SQL
        query = """
            SELECT 
            gu.name AS Usuário, 
            gt.name AS Titulo_do_Chamado, 
            gt.date AS Data_da_Abertura, 
            gt.closedate AS Data_do_Fechamento,
            gt.content AS Comentario,
            gi.content AS Solução
            FROM glpi_tickets gt
            JOIN glpi_users gu ON gt.users_id_recipient = gu.id
            JOIN glpi_itilsolutions gi ON gt.id = gi.items_id
            WHERE gi.itemtype = 'Ticket';
            """

        df = pd.read_sql(query, conn)
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].apply(remover_tags_html)

        #Ordenar os dados pela Data_de_abertura
        df['Data_da_Abertura'] = pd.to_datetime(df['Data_da_Abertura'], errors='coerce')  # Converte para datetime
        df = df.sort_values(by='Data_da_Abertura', ascending=True)  # Ordena pela Data_da_Abertura

        # Carregar o modelo e iniciar xlwings
        caminho_modelo = 'Modelo.xlsx'  # Certifique-se de que este arquivo está no mesmo diretório
        if not os.path.exists(caminho_modelo):
            print(f"Erro: o arquivo {caminho_modelo} não foi encontrado.")
            status_label.config(text=f"Erro: o arquivo {caminho_modelo} não foi encontrado.")
            status_label.update()
            root.after(10000, root.quit)
            return

        app = xw.App(visible=False)
        wb_xlwings = app.books.open(caminho_modelo)
        ws_xlwings = wb_xlwings.sheets[0]

        # Desabilitar a atualização de tela e o cálculo automático
        app.screen_updating = False
        app.calculation = 'manual'

        # Colar os dados em massa
        ws_xlwings.range('A3').value = df.values.tolist()

        # Ativar quebra de texto automática nas colunas "Solução," "Comentário," e "Titulo_do_Chamado"
        for col_name in ["Solução", "Comentario", "Titulo_do_Chamado"]:
            try:
                col_idx = df.columns.get_loc(col_name) + 1  # Índice da coluna (1 baseado)
                intervalo_coluna = f"{chr(64 + col_idx)}3:{chr(64 + col_idx)}{len(df) + 2}"  # A3:F(len + 2)
                ws_xlwings.range(intervalo_coluna).api.WrapText = True  # Ativar quebra de texto
            except Exception as e:
                print(f"Erro ao ativar quebra de texto na coluna {col_name}: {e}")

        # Ajustar largura de todas as colunas, exceto "Solução," e alinhar "Titulo_do_Chamado" à esquerda
        for j, col_name in enumerate(df.columns):
            try:
                if col_name != "Solução":  # Ignorar ajuste de largura na coluna "Solução"
                    ws_xlwings.range(1, j + 1).api.EntireColumn.AutoFit()
                if col_name == "Titulo_do_Chamado":
                    for i in range(1, len(df) + 3):
                        ws_xlwings.range(i, j + 1).api.HorizontalAlignment = -4131  # Alinhado à esquerda
            except Exception as e:
                print(f"Erro ao ajustar coluna {col_name}: {e}")

        # Ajustar altura das linhas automaticamente
        for i in range(3, len(df) + 3):
            try:
                ws_xlwings.range(f"A{i}:F{i}").api.EntireRow.AutoFit()
            except Exception as e:
                print(f"Erro ao ajustar altura da linha {i}: {e}")

        # Centralizar "Usuário," "Data_da_Abertura," e "Data_do_Fechamento"
        for i in range(1, len(df) + 3):
            for j, column_name in enumerate(df.columns):
                if column_name in ["Usuário", "Data_da_Abertura", "Data_do_Fechamento"]:
                    ws_xlwings.range(i, j + 1).api.HorizontalAlignment = -4108

        # Adicionar bordas em todas as células com dados de forma eficiente
        intervalo = f"A3:F{len(df) + 2}"
        ws_xlwings.range(intervalo).api.Borders(7).LineStyle = 1  # Borda superior
        ws_xlwings.range(intervalo).api.Borders(8).LineStyle = 1  # Borda inferior
        ws_xlwings.range(intervalo).api.Borders(9).LineStyle = 1  # Borda à esquerda
        ws_xlwings.range(intervalo).api.Borders(10).LineStyle = 1  # Borda à direita

        # Abrir janela para selecionar onde salvar o arquivo
        caminho_saida = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Escolha onde salvar a planilha"
        )

        if caminho_saida:  # Verificar se o usuário escolheu um local para salvar
            wb_xlwings.save(caminho_saida)
            print(f"Dados exportados com sucesso para '{caminho_saida}'")
        else:
            print("Nenhum local de salvamento escolhido. A operação foi cancelada.")

        wb_xlwings.close()

        # Reabilitar cálculo automático e atualização de tela
        app.calculation = 'automatic'
        app.screen_updating = True

        # Exibir a mensagem de sucesso e fechar a janela após 10 segundos
        status_label.config(text="Relatório gerado com sucesso!")
        status_label.update()

        # Fechar a janela após 10 segundos
        root.after(5000, root.quit)

    except Exception as e:
        print(f"Ocorreu um erro ao exportar os dados: {e}")
        status_label.config(text="Ocorreu um erro ao exportar.")
        status_label.update()
        root.after(10000, root.quit)
    finally:
        conn.close()

# Função para criar a interface gráfica
def criar_interface():
    root = Tk()
    root.title("Gerador de Relatório")

    # Configurar tamanho e layout
    root.geometry("400x200")

    # Adicionar um título
    label = Label(root, text="Clique no botão para gerar o relatório:", font=("Arial", 12))
    label.pack(pady=20)

    # Adicionar um botão para gerar o relatório
    status_label = Label(root, text="", font=("Arial", 10))
    status_label.pack(pady=10)

    gerar_button = Button(root, text="Gerar Relatório", command=lambda: exportar_para_modelo_excel(status_label, root))
    gerar_button.pack(pady=20)

    # Iniciar o loop da interface gráfica
    root.mainloop()

# Executar a interface gráfica
criar_interface()