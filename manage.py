import sys
import json
import os
import zipfile
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QPushButton, QLabel, QLineEdit, QTableWidget,
                             QTableWidgetItem, QTabWidget, QListWidget, QListWidgetItem, QMessageBox, QFileDialog, QInputDialog, QHBoxLayout)
from PyQt6.QtCore import Qt, QItemSelectionModel, QItemSelection, QItemSelectionRange, QEvent, QCoreApplication, QUrl
from PyQt6.QtGui import QDesktopServices
from datetime import datetime
from fpdf import FPDF
import shutil

DATA_FILE = "projetos.json"
DOCUMENTOS_DIR = "documentos"

# Funções para carregar e salvar dados JSON
def carregar_dados():
    if not os.path.exists(DATA_FILE):
        # Cria um arquivo vazio com a estrutura padrão
        with open(DATA_FILE, 'w') as f:
            json.dump({"projetos": []}, f)
        return {"projetos": []}
    with open(DATA_FILE, 'r') as f:
        content = f.read()
        if not content.strip():
            return {"projetos": []}
        return json.loads(content)

def salvar_dados(dados):
    with open(DATA_FILE, 'w') as f:
        json.dump(dados, f, indent=4)

# Função para compactar um projeto em um arquivo ZIP
def compactar_projeto(projeto, nome_arquivo_zip):
    with zipfile.ZipFile(nome_arquivo_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # Adiciona o arquivo JSON
        zipf.writestr('projeto.json', json.dumps(projeto, indent=4))

        # Adiciona os documentos associados
        for comprovante in projeto["comprovantes"]:
            caminho_comprovante = os.path.join(DOCUMENTOS_DIR, comprovante)
            if os.path.exists(caminho_comprovante):
                zipf.write(caminho_comprovante, os.path.basename(caminho_comprovante))

        for nfe in projeto["nfe"]:
            caminho_nfe = os.path.join(DOCUMENTOS_DIR, nfe)
            if os.path.exists(caminho_nfe):
                zipf.write(caminho_nfe, os.path.basename(caminho_nfe))

        for orcamento in projeto["orcamentos"]:
            caminho_orcamento = os.path.join(DOCUMENTOS_DIR, orcamento)
            if os.path.exists(caminho_orcamento):
                zipf.write(caminho_orcamento, os.path.basename(caminho_orcamento))

        for arquivo_adicional in projeto["arquivos_adicionais"]:
            caminho_arquivo_adicional = os.path.join(DOCUMENTOS_DIR, arquivo_adicional)
            if os.path.exists(caminho_arquivo_adicional):
                zipf.write(caminho_arquivo_adicional, os.path.basename(caminho_arquivo_adicional))

# Função para descompactar um projeto de um arquivo ZIP
def descompactar_projeto(nome_arquivo_zip):
    with zipfile.ZipFile(nome_arquivo_zip, 'r') as zipf:
        zipf.extractall(DOCUMENTOS_DIR)
        with zipf.open('projeto.json') as f:
            projeto = json.load(f)
            return projeto

# Classe de Relatório PDF
class PDFReport:
    def __init__(self, projetos):
        self.projetos = projetos

    def gerar_relatorio(self, nome_arquivo):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, txt="Relatório de Projetos de Iniciação Científica", ln=True, align='C')

        for projeto in self.projetos:
            pdf.cell(200, 10, txt=f"Nome: {projeto['nome']}", ln=True)
            pdf.cell(200, 10, txt=f"Responsável: {projeto['responsavel']}", ln=True)
            pdf.cell(200, 10, txt=f"Valor Financiamento: R$ {projeto['valor_financiamento']:.2f}", ln=True)
            pdf.cell(200, 10, txt=f"Data de Cadastro: {projeto['data_cadastro']}", ln=True)

            # Adiciona despesas
            pdf.cell(200, 10, txt="Despesas:", ln=True)
            for despesa in projeto["despesas"]:
                pdf.cell(200, 10, txt=f"  Nome: {despesa['nome']}, Descrição: {despesa['descricao']}, Valor: R$ {despesa['valor']:.2f}, NF-e: {despesa['nfe']}", ln=True)

            # Adiciona orçamentos
            pdf.cell(200, 10, txt="Orçamentos:", ln=True)
            for orcamento in projeto["orcamentos"]:
                pdf.cell(200, 10, txt=f"  {orcamento}", ln=True)

            # Adiciona NF-e
            pdf.cell(200, 10, txt="Notas Fiscais (NF-e):", ln=True)
            for nfe in projeto["nfe"]:
                pdf.cell(200, 10, txt=f"  {nfe}", ln=True)

            # Adiciona comprovantes
            pdf.cell(200, 10, txt="Comprovantes de Pagamento:", ln=True)
            for comprovante in projeto["comprovantes"]:
                pdf.cell(200, 10, txt=f"  {comprovante}", ln=True)

            # Adiciona arquivos adicionais
            pdf.cell(200, 10, txt="Arquivos Adicionais:", ln=True)
            for arquivo_adicional in projeto["arquivos_adicionais"]:
                pdf.cell(200, 10, txt=f"  {arquivo_adicional}", ln=True)

            pdf.cell(200, 10, txt=" ", ln=True)  # Espaço entre projetos
        pdf.output(nome_arquivo)

# Classe principal de gerenciamento de projetos
class ProjetoManager(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gerenciamento de Projetos de IC")
        self.setGeometry(100, 100, 800, 600)

        # Cria o diretório para documentos se não existir
        if not os.path.exists(DOCUMENTOS_DIR):
            os.makedirs(DOCUMENTOS_DIR)

        # Tab principal para cadastrar e visualizar projetos
        self.tab_widget = QTabWidget()
        self.setCentralWidget(self.tab_widget)

        # Aba para cadastrar novos projetos
        self.tab_cadastrar = QWidget()
        self.tab_widget.addTab(self.tab_cadastrar, "Cadastrar Projeto")
        self.layout_cadastrar = QVBoxLayout(self.tab_cadastrar)

        # Campos de entrada para cadastro de projeto
        self.nome_input = QLineEdit()
        self.responsavel_input = QLineEdit()
        self.valor_input = QLineEdit()

        self.layout_cadastrar.addWidget(QLabel("Nome do Projeto:"))
        self.layout_cadastrar.addWidget(self.nome_input)
        self.layout_cadastrar.addWidget(QLabel("Responsável:"))
        self.layout_cadastrar.addWidget(self.responsavel_input)
        self.layout_cadastrar.addWidget(QLabel("Valor Financiamento:"))
        self.layout_cadastrar.addWidget(self.valor_input)

        cadastrar_button = QPushButton("Salvar Projeto")
        cadastrar_button.clicked.connect(self.salvar_projeto)
        self.layout_cadastrar.addWidget(cadastrar_button)

        # Botão para exportar projeto
        exportar_button = QPushButton("Exportar Projeto")
        exportar_button.clicked.connect(self.exportar_projeto)
        self.layout_cadastrar.addWidget(exportar_button)

        # Botão para importar projeto
        importar_button = QPushButton("Importar Projeto")
        importar_button.clicked.connect(self.importar_projeto)
        self.layout_cadastrar.addWidget(importar_button)

        # Aba para visualizar projetos
        self.tab_visualizar = QWidget()
        self.tab_widget.addTab(self.tab_visualizar, "Projetos Cadastrados")
        self.layout_visualizar = QVBoxLayout(self.tab_visualizar)
        self.tabela = QTableWidget()
        self.layout_visualizar.addWidget(self.tabela)

        # Botão para gerar relatório de todos os projetos
        gerar_relatorio_todos_button = QPushButton("Gerar Relatório de Todos os Projetos")
        gerar_relatorio_todos_button.clicked.connect(self.gerar_relatorio_todos_projetos)
        self.layout_visualizar.addWidget(gerar_relatorio_todos_button)

        self.atualizar_tabela()

    def salvar_projeto(self):
        dados = carregar_dados()
        try:
            novo_projeto = {
                "nome": self.nome_input.text(),
                "responsavel": self.responsavel_input.text(),
                "valor_financiamento": float(self.valor_input.text()),
                "despesas": [],  # Inicialmente, a lista de despesas é vazia
                "data_cadastro": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "orcamentos": [],  # Inicialização da lista de orçamentos
                "nfe": [],         # Inicialização da lista de NF-e
                "comprovantes": [],  # Inicialização da lista de comprovantes
                "arquivos_adicionais": []  # Inicialização da lista de arquivos adicionais
            }
            dados["projetos"].append(novo_projeto)
            salvar_dados(dados)
            QMessageBox.information(self, "Sucesso", "Projeto cadastrado com sucesso!")
            self.nome_input.clear()
            self.responsavel_input.clear()
            self.valor_input.clear()
            self.atualizar_tabela()
        except ValueError:
            QMessageBox.warning(self, "Erro", "Valor de financiamento inválido.")

    def atualizar_tabela(self):
        dados = carregar_dados()
        self.tabela.setColumnCount(3)
        self.tabela.setHorizontalHeaderLabels(["Nome", "Responsável", "Valor Financiamento"])
        self.tabela.setRowCount(len(dados["projetos"]))

        for i, projeto in enumerate(dados["projetos"]):
            self.tabela.setItem(i, 0, QTableWidgetItem(projeto["nome"]))
            self.tabela.setItem(i, 1, QTableWidgetItem(projeto["responsavel"]))
            self.tabela.setItem(i, 2, QTableWidgetItem(f"R$ {projeto['valor_financiamento']:.2f}"))

        # Conexão do sinal para clicar na célula
        self.tabela.cellDoubleClicked.connect(self.on_cell_double_clicked)

    def on_cell_double_clicked(self, row):
        dados = carregar_dados()  # Recarregar dados para garantir que temos a versão mais atual
        if 0 <= row < len(dados["projetos"]):
            self.abrir_pagina_projeto(dados["projetos"][row])
        else:
            print(f"Invalid row index: {row}")

    def abrir_pagina_projeto(self, projeto):
        projeto_tab = QWidget()
        self.tab_widget.addTab(projeto_tab, projeto["nome"])
        layout = QVBoxLayout(projeto_tab)

        layout.addWidget(QLabel(f"Nome: {projeto['nome']}"))
        layout.addWidget(QLabel(f"Responsável: {projeto['responsavel']}"))
        layout.addWidget(QLabel(f"Valor Financiamento: R$ {projeto['valor_financiamento']:.2f}"))

        # Botão para excluir projeto
        excluir_projeto_button = QPushButton("Excluir Projeto")
        excluir_projeto_button.clicked.connect(lambda: self.excluir_projeto(projeto))
        layout.addWidget(excluir_projeto_button)

        # Botão para editar projeto
        editar_projeto_button = QPushButton("Editar Projeto")
        editar_projeto_button.clicked.connect(lambda: self.editar_projeto(projeto))
        layout.addWidget(editar_projeto_button)

        # Botão para gerar relatório do projeto
        gerar_relatorio_projeto_button = QPushButton("Gerar Relatório do Projeto")
        gerar_relatorio_projeto_button.clicked.connect(lambda: self.gerar_relatorio_projeto(projeto))
        layout.addWidget(gerar_relatorio_projeto_button)

        # Tabela de despesas
        despesas_label = QLabel("Despesas:")
        layout.addWidget(despesas_label)

        despesas_table = QTableWidget()
        despesas_table.setColumnCount(4)
        despesas_table.setHorizontalHeaderLabels(["Nome", "Descrição", "Valor", "NF-e"])
        layout.addWidget(despesas_table)

        self.carregar_despesas(projeto, despesas_table)

        # Botão para adicionar nova despesa
        adicionar_despesa_button = QPushButton("Adicionar Despesa")
        adicionar_despesa_button.clicked.connect(lambda: self.adicionar_despesa(projeto))
        layout.addWidget(adicionar_despesa_button)

        # Adiciona orçamentos
        orcamentos_label = QLabel("Orçamentos:")
        layout.addWidget(orcamentos_label)

        orcamentos_layout = QVBoxLayout()
        layout.addLayout(orcamentos_layout)

        for orcamento in projeto["orcamentos"]:
            orcamento_item = QListWidgetItem(orcamento)
            orcamento_item.setData(Qt.ItemDataRole.UserRole, os.path.join(DOCUMENTOS_DIR, orcamento))
            self.add_item_with_delete_button(
                orcamentos_layout, 
                orcamento, 
                orcamento_item, 
                lambda: self.excluir_arquivo(projeto, os.path.join(DOCUMENTOS_DIR, orcamento), "orcamentos"), 
                lambda: self.abrir_documento(orcamento_item)
            )

        # Botão para adicionar orçamento
        adicionar_orcamento_button = QPushButton("Adicionar Orçamento")
        adicionar_orcamento_button.clicked.connect(lambda: self.adicionar_orcamento(projeto, orcamentos_layout))
        layout.addWidget(adicionar_orcamento_button)

        # Adiciona NF-e
        nfe_label = QLabel("Notas Fiscais (NF-e):")
        layout.addWidget(nfe_label)

        nfe_layout = QVBoxLayout()
        layout.addLayout(nfe_layout)

        for nfe in projeto["nfe"]:
            nfe_item = QListWidgetItem(nfe)
            nfe_item.setData(Qt.ItemDataRole.UserRole, os.path.join(DOCUMENTOS_DIR, nfe))
            self.add_item_with_delete_button(
                nfe_layout, 
                nfe, 
                nfe_item, 
                lambda: self.excluir_arquivo(projeto, os.path.join(DOCUMENTOS_DIR, nfe), "nfe"), 
                lambda: self.abrir_documento(nfe_item)
            )

        # Botão para adicionar NF-e
        adicionar_nfe_button = QPushButton("Adicionar NF-e")
        adicionar_nfe_button.clicked.connect(lambda: self.adicionar_nfe(projeto, nfe_layout))
        layout.addWidget(adicionar_nfe_button)

        # Adiciona comprovantes
        comprovantes_label = QLabel("Comprovantes de Pagamento:")
        layout.addWidget(comprovantes_label)

        comprovantes_layout = QVBoxLayout()
        layout.addLayout(comprovantes_layout)

        for comprovante in projeto["comprovantes"]:
            comprovante_item = QListWidgetItem(comprovante)
            comprovante_item.setData(Qt.ItemDataRole.UserRole, os.path.join(DOCUMENTOS_DIR, comprovante))
            self.add_item_with_delete_button(
                comprovantes_layout, 
                comprovante, 
                comprovante_item, 
                lambda: self.excluir_arquivo(projeto, os.path.join(DOCUMENTOS_DIR, comprovante), "comprovantes"), 
                lambda: self.abrir_documento(comprovante_item)
            )

        # Botão para adicionar comprovante
        adicionar_comprovante_button = QPushButton("Adicionar Comprovante")
        adicionar_comprovante_button.clicked.connect(lambda: self.adicionar_comprovante(projeto, comprovantes_layout))
        layout.addWidget(adicionar_comprovante_button)

        # Adiciona arquivos adicionais
        arquivos_adicionais_label = QLabel("Arquivos Adicionais:")
        layout.addWidget(arquivos_adicionais_label)

        arquivos_adicionais_layout = QVBoxLayout()
        layout.addLayout(arquivos_adicionais_layout)

        for arquivo_adicional in projeto["arquivos_adicionais"]:
            arquivo_adicional_item = QListWidgetItem(arquivo_adicional)
            arquivo_adicional_item.setData(Qt.ItemDataRole.UserRole, os.path.join(DOCUMENTOS_DIR, arquivo_adicional))
            self.add_item_with_delete_button(
                arquivos_adicionais_layout, 
                arquivo_adicional, 
                arquivo_adicional_item, 
                lambda: self.excluir_arquivo(projeto, os.path.join(DOCUMENTOS_DIR, arquivo_adicional), "arquivos_adicionais"), 
                lambda: self.abrir_documento(arquivo_adicional_item)
            )

        # Botão para adicionar arquivo adicional
        adicionar_arquivo_adicional_button = QPushButton("Adicionar Arquivo Adicional")
        adicionar_arquivo_adicional_button.clicked.connect(lambda: self.adicionar_arquivo_adicional(projeto, arquivos_adicionais_layout))
        layout.addWidget(adicionar_arquivo_adicional_button)

    def add_item_with_delete_button(self, layout, item_text, item, delete_callback, open_callback):
        item_layout = QHBoxLayout()
        item_label = QLabel(item_text)
        delete_button = QPushButton("Excluir")
        delete_button.clicked.connect(delete_callback)
        open_button = QPushButton("Abrir")
        open_button.clicked.connect(open_callback)
        item_layout.addWidget(item_label)
        item_layout.addWidget(open_button)
        item_layout.addWidget(delete_button)
        layout.addLayout(item_layout)

    def excluir_arquivo(self, projeto, arquivo, tipo):
        if os.path.exists(arquivo):
            os.remove(arquivo)
        projeto[tipo].remove(os.path.basename(arquivo))
        dados = carregar_dados()
        for p in dados["projetos"]:
            if p["nome"] == projeto["nome"]:
                p[tipo] = projeto[tipo]
        salvar_dados(dados)
        QMessageBox.information(self, "Sucesso", f"Arquivo {os.path.basename(arquivo)} excluído com sucesso!")
        self.atualizar_tabela()

    def carregar_despesas(self, projeto, table):
        table.setRowCount(len(projeto["despesas"]))
        for i, despesa in enumerate(projeto["despesas"]):
            table.setItem(i, 0, QTableWidgetItem(despesa["nome"]))
            table.setItem(i, 1, QTableWidgetItem(despesa["descricao"]))
            table.setItem(i, 2, QTableWidgetItem(f"R$ {despesa['valor']:.2f}"))
            table.setItem(i, 3, QTableWidgetItem(despesa["nfe"]))

    def excluir_projeto(self, projeto):
        dados = carregar_dados()
        dados["projetos"].remove(projeto)
        salvar_dados(dados)
        self.atualizar_tabela()
        QMessageBox.information(self, "Sucesso", "Projeto excluído com sucesso!")
        self.tab_widget.removeTab(self.tab_widget.currentIndex())

    def editar_projeto(self, projeto):
        nome, ok1 = QInputDialog.getText(self, "Editar Nome do Projeto", "Insira o novo nome do projeto:", text=projeto["nome"])
        if ok1:
            responsavel, ok2 = QInputDialog.getText(self, "Editar Responsável", "Insira o novo responsável:", text=projeto["responsavel"])
            if ok2:
                valor, ok3 = QInputDialog.getDouble(self, "Editar Valor Financiamento", "Insira o novo valor de financiamento:", projeto["valor_financiamento"], 0.0, 0.0)
                if ok3:
                    projeto["nome"] = nome
                    projeto["responsavel"] = responsavel
                    projeto["valor_financiamento"] = valor
                    dados = carregar_dados()
                    for p in dados["projetos"]:
                        if p["nome"] == projeto["nome"]:
                            p.update(projeto)
                    salvar_dados(dados)
                    QMessageBox.information(self, "Sucesso", "Projeto editado com sucesso!")
                    self.atualizar_tabela()

    def adicionar_despesa(self, projeto):
        nome, ok1 = QInputDialog.getText(self, "Nome da Despesa", "Insira o nome da despesa:")
        if ok1:
            descricao, ok2 = QInputDialog.getText(self, "Descrição da Despesa", "Insira a descrição da despesa:")
            if ok2:
                valor, ok3 = QInputDialog.getDouble(self, "Valor da Despesa", "Insira o valor da despesa:", 0.0, 0.0)
                if ok3:
                    nfe, ok4 = QInputDialog.getText(self, "NF-e", "Insira a NF-e:")
                    if ok4:
                        projeto["despesas"].append({
                            "nome": nome,
                            "descricao": descricao,
                            "valor": valor,
                            "nfe": nfe
                        })
                        dados = carregar_dados()
                        for p in dados["projetos"]:
                            if p["nome"] == projeto["nome"]:
                                p["despesas"] = projeto["despesas"]
                        salvar_dados(dados)
                        QMessageBox.information(self, "Sucesso", "Despesa adicionada com sucesso!")
                        self.atualizar_tabela()

    def adicionar_orcamento(self, projeto, orcamentos_layout):
        orcamento, ok = QFileDialog.getOpenFileName(self, "Adicionar Orçamento", "", "Arquivos PDF (*.pdf);;Todos os Arquivos (*)")
        if ok and orcamento:
            nome_orcamento = os.path.basename(orcamento)
            destino_orcamento = os.path.join(DOCUMENTOS_DIR, nome_orcamento)
            if os.path.exists(destino_orcamento):
                nome_orcamento = self.adicionar_sufixo(nome_orcamento)
                destino_orcamento = os.path.join(DOCUMENTOS_DIR, nome_orcamento)
            shutil.copy(orcamento, destino_orcamento)
            projeto["orcamentos"].append(nome_orcamento)
            dados = carregar_dados()
            for p in dados["projetos"]:
                if p["nome"] == projeto["nome"]:
                    p["orcamentos"] = projeto["orcamentos"]
            salvar_dados(dados)
            QMessageBox.information(self, "Sucesso", "Orçamento adicionado com sucesso!")
            orcamento_item = QListWidgetItem(nome_orcamento)
            orcamento_item.setData(Qt.ItemDataRole.UserRole, destino_orcamento)
            self.add_item_with_delete_button(
                orcamentos_layout, 
                nome_orcamento, 
                orcamento_item, 
                lambda: self.excluir_arquivo(projeto, destino_orcamento, "orcamentos"), 
                lambda: self.abrir_documento(orcamento_item)
            )

    def adicionar_nfe(self, projeto, nfe_layout):
        nfe, ok = QFileDialog.getOpenFileName(self, "Adicionar NF-e", "", "Arquivos PDF (*.pdf);;Todos os Arquivos (*)")
        if ok and nfe:
            nome_nfe = os.path.basename(nfe)
            destino_nfe = os.path.join(DOCUMENTOS_DIR, nome_nfe)
            if os.path.exists(destino_nfe):
                nome_nfe = self.adicionar_sufixo(nome_nfe)
                destino_nfe = os.path.join(DOCUMENTOS_DIR, nome_nfe)
            shutil.copy(nfe, destino_nfe)
            projeto["nfe"].append(nome_nfe)
            dados = carregar_dados()
            for p in dados["projetos"]:
                if p["nome"] == projeto["nome"]:
                    p["nfe"] = projeto["nfe"]
            salvar_dados(dados)
            QMessageBox.information(self, "Sucesso", "NF-e adicionada com sucesso!")
            nfe_item = QListWidgetItem(nome_nfe)
            nfe_item.setData(Qt.ItemDataRole.UserRole, destino_nfe)
            self.add_item_with_delete_button(
                nfe_layout, 
                nome_nfe, 
                nfe_item, 
                lambda: self.excluir_arquivo(projeto, destino_nfe, "nfe"), 
                lambda: self.abrir_documento(nfe_item)
            )

    def adicionar_comprovante(self, projeto, comprovantes_layout):
        comprovante, ok = QFileDialog.getOpenFileName(self, "Selecionar Comprovante", "", "Arquivos PDF (*.pdf);;Todos os Arquivos (*)")
        if ok and comprovante:
            nome_comprovante = os.path.basename(comprovante)
            destino = os.path.join(DOCUMENTOS_DIR, nome_comprovante)
            if os.path.exists(destino):
                nome_comprovante = self.adicionar_sufixo(nome_comprovante)
                destino = os.path.join(DOCUMENTOS_DIR, nome_comprovante)
            shutil.copy(comprovante, destino)
            projeto["comprovantes"].append(nome_comprovante)
            dados = carregar_dados()
            for p in dados["projetos"]:
                if p["nome"] == projeto["nome"]:
                    p["comprovantes"] = projeto["comprovantes"]
            salvar_dados(dados)
            QMessageBox.information(self, "Sucesso", "Comprovante adicionado com sucesso!")
            comprovante_item = QListWidgetItem(nome_comprovante)
            comprovante_item.setData(Qt.ItemDataRole.UserRole, destino)
            self.add_item_with_delete_button(
                comprovantes_layout, 
                nome_comprovante, 
                comprovante_item, 
                lambda: self.excluir_arquivo(projeto, destino, "comprovantes"), 
                lambda: self.abrir_documento(comprovante_item)
            )

    def adicionar_arquivo_adicional(self, projeto, arquivos_adicionais_layout):
        arquivo_adicional, ok = QFileDialog.getOpenFileName(self, "Adicionar Arquivo Adicional", "", "Todos os Arquivos (*)")
        if ok and arquivo_adicional:
            nome_arquivo_adicional = os.path.basename(arquivo_adicional)
            destino_arquivo_adicional = os.path.join(DOCUMENTOS_DIR, nome_arquivo_adicional)
            if os.path.exists(destino_arquivo_adicional):
                nome_arquivo_adicional = self.adicionar_sufixo(nome_arquivo_adicional)
                destino_arquivo_adicional = os.path.join(DOCUMENTOS_DIR, nome_arquivo_adicional)
            shutil.copy(arquivo_adicional, destino_arquivo_adicional)
            projeto["arquivos_adicionais"].append(nome_arquivo_adicional)
            dados = carregar_dados()
            for p in dados["projetos"]:
                if p["nome"] == projeto["nome"]:
                    p["arquivos_adicionais"] = projeto["arquivos_adicionais"]
            salvar_dados(dados)
            QMessageBox.information(self, "Sucesso", "Arquivo adicional adicionado com sucesso!")
            arquivo_adicional_item = QListWidgetItem(nome_arquivo_adicional)
            arquivo_adicional_item.setData(Qt.ItemDataRole.UserRole, destino_arquivo_adicional)
            self.add_item_with_delete_button(
                arquivos_adicionais_layout, 
                nome_arquivo_adicional, 
                arquivo_adicional_item, 
                lambda: self.excluir_arquivo(projeto, destino_arquivo_adicional, "arquivos_adicionais"), 
                lambda: self.abrir_documento(arquivo_adicional_item)
            )

    def adicionar_sufixo(self, nome_arquivo):
        nome, ext = os.path.splitext(nome_arquivo)
        sufixo = 1
        novo_nome = f"{nome}_{sufixo}{ext}"
        while os.path.exists(os.path.join(DOCUMENTOS_DIR, novo_nome)):
            sufixo += 1
            novo_nome = f"{nome}_{sufixo}{ext}"
        return novo_nome

    def exportar_projeto(self):
        index = self.tabela.currentRow()
        if index == -1:
            QMessageBox.warning(self, "Erro", "Selecione um projeto para exportar.")
            return
        dados = carregar_dados()
        projeto = dados["projetos"][index]
        nome_arquivo_zip, _ = QFileDialog.getSaveFileName(self, "Salvar Projeto como ZIP", "", "ZIP Files (*.zip)")
        if nome_arquivo_zip:
            compactar_projeto(projeto, nome_arquivo_zip)
            QMessageBox.information(self, "Sucesso", "Projeto exportado com sucesso!")

    def importar_projeto(self):
        nome_arquivo_zip, _ = QFileDialog.getOpenFileName(self, "Importar Projeto", "", "ZIP Files (*.zip)")
        if nome_arquivo_zip:
            projeto = descompactar_projeto(nome_arquivo_zip)
            dados = carregar_dados()
            dados["projetos"].append(projeto)
            salvar_dados(dados)
            QMessageBox.information(self, "Sucesso", "Projetos importados com sucesso!")
            self.atualizar_tabela()

    def abrir_documento(self, item):
        caminho_documento = item.data(Qt.ItemDataRole.UserRole)
        if os.path.exists(caminho_documento):
            QDesktopServices.openUrl(QUrl.fromLocalFile(caminho_documento))
        else:
            QMessageBox.warning(self, "Erro", "O arquivo não foi encontrado.")

    def gerar_relatorio_projeto(self, projeto):
        pdf = PDFReport([projeto])
        nome_arquivo, _ = QFileDialog.getSaveFileName(self, "Salvar Relatório como PDF", "", "PDF Files (*.pdf)")
        if nome_arquivo:
            pdf.gerar_relatorio(nome_arquivo)
            QMessageBox.information(self, "Sucesso", "Relatório gerado com sucesso!")

    def gerar_relatorio_todos_projetos(self):
        dados = carregar_dados()
        pdf = PDFReport(dados["projetos"])
        nome_arquivo, _ = QFileDialog.getSaveFileName(self, "Salvar Relatório como PDF", "", "PDF Files (*.pdf)")
        if nome_arquivo:
            pdf.gerar_relatorio(nome_arquivo)
            QMessageBox.information(self, "Sucesso", "Relatório gerado com sucesso!")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ProjetoManager()
    window.show()
    sys.exit(app.exec())