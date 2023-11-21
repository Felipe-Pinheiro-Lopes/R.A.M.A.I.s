from PyQt5 import uic,QtWidgets, QtCore, QtGui
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from pathlib import Path
from PyQt5.QtCore import QDateTime, Qt, QDate, QTime, QStandardPaths
import os, openpyxl, _sqlite3, getpass
from datetime import datetime

# Conexão com o Banco de Dados
con = _sqlite3.connect('Ramais.db')
cursor = con.cursor()

#Função de Lista 
def funcao_listar():
    cursor.execute("SELECT r.ID, n.Nome, t.Numeros, s.Nome, r.Observacao FROM Ramais r JOIN Nomes n ON r.Nome = n.ID JOIN Telefones t ON r.Telefone = t.ID JOIN Setor s ON r.Setor = s.ID")
    resultados = cursor.fetchall()

    formulario.tableWidget.setRowCount(len(resultados))
    formulario.tableWidget.setColumnCount(5)

    for i, row in enumerate(resultados):
        for j, col in enumerate(row):
            formulario.tableWidget.setItem(i, j, QtWidgets.QTableWidgetItem(str(col)))
            formulario.tableWidget.setColumnWidth(0, 2)
            formulario.tableWidget.setColumnWidth(1, 160)
            formulario.tableWidget.setColumnWidth(2, 50)
            formulario.tableWidget.setColumnWidth(4, 200)

#Funções para obiter os IDs das tabelas extrangeiras 
def obter_id_setor(setor):
    
    cursor.execute(f"SELECT id FROM Setor WHERE Nome = '{setor}'")
    resultado = cursor.fetchone()
    if resultado:
        return resultado[0]
    else:
        cursor.execute(f"INSERT INTO Setor (Nome) VALUES ('{setor}')")
        con.commit()
        return cursor.lastrowid

def obter_id_telefone(telefone):
    
    cursor.execute(f"SELECT id FROM Telefones WHERE Numeros = '{telefone}'")
    resultado = cursor.fetchone()
    if resultado:
        return resultado[0]
    else:
        cursor.execute(f"INSERT INTO Telefones (Numeros) VALUES ('{telefone}')")
        con.commit()
        return cursor.lastrowid

def obter_id_nome(nome):
    
    cursor.execute(f"SELECT id FROM Nomes WHERE Nome = '{nome}'")
    resultado = cursor.fetchone()
    if resultado:
        return resultado[0]
    else:
        cursor.execute(f"INSERT INTO Nomes (Nome) VALUES ('{nome}')")
        con.commit()
        return cursor.lastrowid

#Função de Cadastro 
def funcao_principal():
    nome = formulario.lineEdit.text()
    telefone = formulario.lineEdit_2.text()
    observacao = formulario.lineEdit_3.text()
    setor = formulario.comboBox.currentText()
    id_setor = obter_id_setor(setor)
    id_telefone = obter_id_telefone(telefone)
    id_nome = obter_id_nome(nome)
    
    cursor.execute(f"SELECT * FROM Ramais WHERE Nome='{id_nome}' AND Telefone='{id_telefone}' AND Setor='{id_setor}'")
    resultado = cursor.fetchone()
    if resultado:
        QtWidgets.QMessageBox.warning(formulario, 'Aviso', 'Já existe um ramal com esses valores!')
        return

    cursor.execute(f"INSERT INTO Ramais (Nome, Telefone, Observacao, Setor) VALUES ('{id_nome}', '{id_telefone}', '{observacao}', '{id_setor}')")
    con.commit()
    formulario.lineEdit.setText ("")
    formulario.lineEdit_2.setText("")
    formulario.lineEdit_3.setText("")

#Funções para obiter os valores das tabelas extrangeiras 
def obter_nome_setor(id_setor):
    cursor.execute(f"SELECT Nome FROM Setor WHERE id = {id_setor}")
    resultado = cursor.fetchone()
    if resultado:
        return resultado[0]
    else:
        return None
    
def obter_numero_telefone(id_telefone):
    cursor.execute(f"SELECT Numeros FROM Telefones WHERE id = {id_telefone}")
    resultado = cursor.fetchone()
    if resultado:
        return resultado[0]
    else:
        return None
    
def obter_nome_nome(id_nome):
    cursor.execute(f"SELECT Nome FROM Nomes WHERE id = {id_nome}")
    resultado = cursor.fetchone()
    if resultado:
        return resultado[0]
    else:
        return None

#Função de Filtro para Pesquisa
def funcao_Filtro_pesquisar():
    setor = formulario.comboBox_3.currentText()
    cargo = None

    if formulario.radioButton.isChecked():
        cargo = "Chef"
    elif formulario.radioButton_4.isChecked():
        cargo = "Estag"
    elif formulario.radioButton_2.isChecked():
        cargo = "func"

    filtro_setor = f" AND Setor.Nome = '{setor}'" if setor else ""
    filtro_cargo = ""

    if cargo == "Chef":
        filtro_cargo = " AND Ramais.Observacao LIKE '%Chefe%'"
    elif cargo == "Estag":
        filtro_cargo = " AND (Ramais.Observacao LIKE '%Estagiário%' OR Nomes.Nome LIKE '%Estagiários%')"
    elif cargo == "func":
        filtro_cargo = " AND (Ramais.Observacao NOT LIKE '%Chefe%' AND Ramais.Observacao NOT LIKE '%Estagiário%')"

    ramal = formulario.lineEdit_9.text()
    if not ramal:
        query = f"SELECT Nomes.Nome, Telefones.Numeros, Setor.Nome, Ramais.Observacao FROM Ramais INNER JOIN Nomes ON Ramais.Nome=Nomes.ID INNER JOIN Telefones ON Ramais.Telefone=Telefones.ID INNER JOIN Setor ON Ramais.Setor=Setor.ID WHERE 1{filtro_setor}{filtro_cargo} ORDER BY Nomes.Nome, Telefones.Numeros, Setor.Nome, Ramais.Observacao"
    else:
        query = f"SELECT Nomes.Nome, Telefones.Numeros, Setor.Nome, Ramais.Observacao FROM Ramais INNER JOIN Nomes ON Ramais.Nome=Nomes.ID INNER JOIN Telefones ON Ramais.Telefone=Telefones.ID INNER JOIN Setor ON Ramais.Setor=Setor.ID WHERE (Nomes.Nome LIKE '%{ramal}%' OR Telefones.Numeros LIKE '%{ramal}%'){filtro_setor}{filtro_cargo} ORDER BY Nomes.Nome, Telefones.Numeros, Setor.Nome, Ramais.Observacao"

    cursor.execute(query)
    resultados = cursor.fetchall()

    formulario.radioButton.setAutoExclusive(False)
    formulario.radioButton.setChecked(False)
    formulario.radioButton_2.setAutoExclusive(False)
    formulario.radioButton_2.setChecked(False)
    formulario.radioButton_4.setAutoExclusive(False)
    formulario.radioButton_4.setChecked(False)
    formulario.radioButton.setAutoExclusive(True)
    formulario.radioButton_2.setAutoExclusive(True)
    formulario.radioButton_4.setAutoExclusive(True)

    if resultados:
        formulario.tableWidget_3.setRowCount(len(resultados))  # Definir o número de linhas na tabela
        formulario.tableWidget_3.setColumnCount(4)

        # Preenchendo a tabela com os valores dos resultados
        for i, resultado in enumerate(resultados):
            nome_setor, nome_nome, numero_telefone, observacao = resultado

            # Inserindo os nomes na tabela
            formulario.tableWidget_3.setItem(i, 0, QtWidgets.QTableWidgetItem(nome_setor))
            formulario.tableWidget_3.setItem(i, 1, QtWidgets.QTableWidgetItem(nome_nome))
            formulario.tableWidget_3.setItem(i, 2, QtWidgets.QTableWidgetItem(numero_telefone))
            formulario.tableWidget_3.setItem(i, 3, QtWidgets.QTableWidgetItem(observacao))

        formulario.tableWidget_3.setColumnWidth(0, 160)
        formulario.tableWidget_3.setColumnWidth(1, 50)
        formulario.tableWidget_3.setColumnWidth(2, 50)
        formulario.tableWidget_3.setColumnWidth(3, 200)
    else:
        QtWidgets.QMessageBox.warning(formulario, 'Aviso', 'Nenhum resultado encontrado!')

#Função de Pesquisa
def funcao_pesquisar():
    ramal = formulario.lineEdit_9.text()
    cursor.execute(f"SELECT Nomes.Nome, Telefones.Numeros, Setor.Nome, Ramais.Observacao FROM Ramais INNER JOIN Nomes ON Ramais.Nome=Nomes.ID INNER JOIN Telefones ON Ramais.Telefone=Telefones.ID INNER JOIN Setor ON Ramais.Setor=Setor.ID WHERE Nomes.Nome LIKE '%{ramal}%' OR Telefones.Numeros LIKE '%{ramal}%'")
    resultados = cursor.fetchall()

    if resultados:
        formulario.tableWidget_3.setRowCount(len(resultados))  # Definir o número de linhas na tabela
        formulario.tableWidget_3.setColumnCount(4)

        # Preenchendo a tabela com os valores dos resultados
        for i, resultado in enumerate(resultados):
            nome_setor, nome_nome, numero_telefone, observacao = resultado

            # Inserindo os nomes na tabela
            formulario.tableWidget_3.setItem(i, 0, QtWidgets.QTableWidgetItem(nome_setor))
            formulario.tableWidget_3.setItem(i, 1, QtWidgets.QTableWidgetItem(nome_nome))
            formulario.tableWidget_3.setItem(i, 2, QtWidgets.QTableWidgetItem(numero_telefone))
            formulario.tableWidget_3.setItem(i, 3, QtWidgets.QTableWidgetItem(observacao))

        formulario.tableWidget_3.setColumnWidth(0, 160)
        formulario.tableWidget_3.setColumnWidth(1, 50)
        formulario.tableWidget_3.setColumnWidth(2, 110)
        formulario.tableWidget_3.setColumnWidth(3, 200)
    else:
        QtWidgets.QMessageBox.warning(formulario, 'Aviso', 'Ramal não encontrado!')

#Função de Edição do Banco
def funcao_editar_ramal():
    linha_selecionada = formulario.tableWidget.currentRow()
    if linha_selecionada < 0:
        QtWidgets.QMessageBox.warning(formulario, 'Aviso', 'Selecione um registro para excluir!')
        formulario.Paginas.setCurrentWidget(formulario.Lista)
        return
    
    formulario.lineEdit_10.setText(formulario.tableWidget.item(linha_selecionada, 0).text())
    formulario.lineEdit_11.setText(formulario.tableWidget.item(linha_selecionada, 1).text())
    formulario.lineEdit_12.setText(formulario.tableWidget.item(linha_selecionada, 2).text())
    formulario.lineEdit_14.setText(formulario.tableWidget.item(linha_selecionada, 4).text())
    funcao_comboBox2_setor()

#Função para salvar a alteração de ramal
def salvar_alteracoes():
    id = formulario.lineEdit_10.text()
    nome = formulario.lineEdit_11.text()
    telefone = formulario.lineEdit_12.text()
    setor = formulario.comboBox_2.currentText()
    observacao = formulario.lineEdit_14.text()

    id_setor = obter_id_setor(setor)
    id_telefone = obter_id_telefone(telefone)
    id_nome = obter_id_nome(nome)

    cursor.execute(f"SELECT * FROM Ramais WHERE Nome='{id_nome}' AND Telefone='{id_telefone}' AND Setor='{id_setor}' AND Observacao='{observacao}'")
    resultado = cursor.fetchone()
    if resultado:
        QtWidgets.QMessageBox.warning(formulario, 'Aviso', 'Já existe um ramal com esses valores!')
        return

    cursor.execute(f"UPDATE Ramais SET Nome='{id_nome}', Telefone='{id_telefone}', Observacao='{observacao}', Setor='{id_setor}' WHERE id = {id}")
    con.commit()

    formulario.lineEdit_10.setText("")
    formulario.lineEdit_11.setText("")
    formulario.lineEdit_12.setText("")
    formulario.lineEdit_14.setText("")
    clear_combox()
    funcao_listar()

#Função para Exclusão do Banco
def funcao_excluir_ramal():
    # Obtém a linha selecionada na tabela
    linha_selecionada = formulario.tableWidget.currentRow()

    # Verifica se alguma linha foi selecionada
    if linha_selecionada < 0:
        QtWidgets.QMessageBox.warning(formulario, 'Aviso', 'Selecione um registro para excluir!')
        return

    # Obtém o valor do ID na coluna 0 da linha selecionada
    id_ramal = formulario.tableWidget.item(linha_selecionada, 0).text()

    # Exclui o registro correspondente na tabela Ramais
    cursor.execute(f"DELETE FROM Ramais WHERE Id='{id_ramal}'")
    con.commit()

    funcao_listar()

def funcao_excluir_nota():
    # Obtém a linha selecionada na tabela
    linha_selecionada = formulario.tableWidget.currentRow()

    # Verifica se alguma linha foi selecionada
    if linha_selecionada < 0:
        QtWidgets.QMessageBox.warning(formulario, 'Aviso', 'Selecione um registro para excluir!')
        return

    # Obtém o valor do ID na coluna 0 da linha selecionada
    id_nota = formulario.tableWidget.item(linha_selecionada, 0).text()

    # Exclui o registro correspondente na tabela Ramais
    cursor.execute(f"DELETE FROM Service_Desk WHERE Id='{id_nota}'")
    con.commit()

    TableNotas2()

#Funções dos ComboBox's
def funcao_comboBox_setor():
    formulario.comboBox.clear()
    cursor.execute("SELECT Nome FROM Setor")
    resultados = cursor.fetchall()
    nomes_setor = [resultado[0] for resultado in resultados]
    formulario.comboBox.addItems(nomes_setor)

def funcao_comboBox3_setor():
    formulario.comboBox_3.clear()
    cursor.execute("SELECT Nome FROM Setor")
    resultados = cursor.fetchall()
    nomes_setor = [resultado[0] for resultado in resultados]
    formulario.comboBox_3.addItems(nomes_setor)

def funcao_comboBox2_setor():
    linha_selecionada = formulario.tableWidget.currentRow()
    valor_selecionado = formulario.tableWidget.item(linha_selecionada, 3).text()
    cursor.execute("SELECT Nome FROM Setor")
    resultados = cursor.fetchall()
    nomes_setor = [resultado[0] for resultado in resultados]
    formulario.comboBox_2.addItems(nomes_setor)
    formulario.comboBox_2.setCurrentText(valor_selecionado)
    
def funcao_cadastro_setor():
    setor = formulario.lineEdit_15.text().upper()
    cursor.execute("SELECT * FROM Setor WHERE Nome=?", (setor,))
    resultado = cursor.fetchone()
    if resultado:
        QtWidgets.QMessageBox.warning(formulario, 'Aviso', 'Já existe um Setor com esse nome!')
        return

    cursor.execute("INSERT INTO Setor (Nome) VALUES (?)", (setor,))
    con.commit()
    funcao_comboBox_setor()

def clear_combox():
    formulario.comboBox_2.clear()

#Função de Download em PDF
def get_resultados_pesquisa():
    setor = formulario.comboBox_3.currentText()
    cargo = None

    if formulario.radioButton.isChecked():
        cargo = "Chef"
    elif formulario.radioButton_4.isChecked():
        cargo = "Estag"
    elif formulario.radioButton_2.isChecked():
        cargo = "func"

    filtro_setor = f" AND Setor.Nome = '{setor}'" if setor else ""
    filtro_cargo = ""

    if cargo == "Chef":
        filtro_cargo = " AND (Ramais.Observacao LIKE '%Chefe%' OR Ramais.Observacao LIKE '%Coordenador%')"
    elif cargo == "Estag":
        filtro_cargo = " AND (Ramais.Observacao LIKE '%Estagiário%' OR Ramais.Observacao LIKE '%Estagiários%' OR Ramais.Observacao LIKE '%Estagiária%' OR Ramais.Observacao LIKE '%Estagiárias%')"
    elif cargo == "func":
        filtro_cargo = " AND (Ramais.Observacao NOT LIKE '%Chefe%' AND Ramais.Observacao NOT LIKE '%Estagiá%')"

    ramal = formulario.lineEdit_9.text()

    # Verifica se não há filtros definidos
    if not setor and not cargo and not ramal:
        filtro_setor = ""
        filtro_cargo = ""

    if not ramal:
        query = f"SELECT Setor.Nome, Nomes.Nome, Telefones.Numeros, Ramais.Observacao FROM Ramais INNER JOIN Nomes ON Ramais.Nome=Nomes.ID INNER JOIN Telefones ON Ramais.Telefone=Telefones.ID INNER JOIN Setor ON Ramais.Setor=Setor.ID WHERE 1{filtro_setor}{filtro_cargo} ORDER BY Nomes.Nome, Telefones.Numeros, Setor.Nome, Ramais.Observacao"
    elif not ramal and not cargo:
        query = f"SELECT Nomes.Nome, Telefones.Numeros, Setor.Nome, Ramais.Observacao FROM Ramais INNER JOIN Nomes ON Ramais.Nome=Nomes.ID INNER JOIN Telefones ON Ramais.Telefone=Telefones.ID INNER JOIN Setor ON Ramais.Setor=Setor.ID WHERE Nomes.Nome LIKE '%{ramal}%' OR Telefones.Numeros LIKE '%{ramal}%'"
    else:
        query = f"SELECT Setor.Nome, Nomes.Nome, Telefones.Numeros, Ramais.Observacao FROM Ramais INNER JOIN Nomes ON Ramais.Nome=Nomes.ID INNER JOIN Telefones ON Ramais.Telefone=Telefones.ID INNER JOIN Setor ON Ramais.Setor=Setor.ID WHERE (Nomes.Nome LIKE '%{ramal}%' OR Telefones.Numeros LIKE '%{ramal}%'){filtro_setor}{filtro_cargo} ORDER BY Nomes.Nome, Telefones.Numeros, Setor.Nome, Ramais.Observacao"

    cursor.execute(query)
    resultados = cursor.fetchall()

    return resultados

def pdf_2():
    desktop_path = str(Path.home() / "Desktop")
    pdf_path = os.path.join(desktop_path, "lista_ramais.pdf")
    pdf_canvas = canvas.Canvas(pdf_path, pagesize=letter)
    x = 50
    y = 750
    pdf_canvas.setFont("Helvetica-Bold", 10)
    pdf_canvas.drawString(x, y, "Lista de Ramais")
    x = 50
    y -= 50

    resultados = get_resultados_pesquisa()
    print(resultados)  
    pdf_canvas.setFont("Helvetica", 10)
    for row in resultados:
        nome = f"Nome: {row[1]}"
        telefone = f"Telefone: {row[2]}"
        setor = f"Setor: {row[0]}"
        observacao = f"Observação: {row[3]}"

        if y < 50:
            pdf_canvas.showPage()
            y = 750
            pdf_canvas.setFont("Helvetica", 10)
            x = 50
            y -= 50

        pdf_canvas.drawString(x, y, nome)
        y -= 20
        pdf_canvas.drawString(x, y, telefone)
        y -= 20
        pdf_canvas.drawString(x, y, setor)
        y -= 20
        pdf_canvas.drawString(x, y, observacao)
        y -= 40

    pdf_canvas.save()
    QtWidgets.QMessageBox.information(formulario, 'Sucesso', 'Arquivo PDF salvo na área de trabalho!')

#Função de Download em Ecxel
def excel_2():
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet['A1'] = 'Nome'
    sheet['B1'] = 'Telefone'
    sheet['C1'] = 'Setor'
    sheet['D1'] = 'Observação'

    resultados = get_resultados_pesquisa()
    row_num = 2 
    for row in resultados:
        nome = row[1]
        telefone = row[2]
        setor = row[0]
        observacao = row[3]
        sheet.cell(row=row_num, column=1, value=nome)
        sheet.cell(row=row_num, column=2, value=telefone)
        sheet.cell(row=row_num, column=3, value=setor)
        sheet.cell(row=row_num, column=4, value=observacao)
        row_num += 1

    desktop_path = str(Path.home() / "Desktop")
    excel_path = f"{desktop_path}/lista_ramais.xlsx"
    workbook.save(excel_path)

    QtWidgets.QMessageBox.information(formulario, 'Sucesso', 'Dados salvos em um arquivo Excel na sua área de trabalho!')

#Função Busca de Id's das Notas/Nomes
def obter_id_nome_Service(nomeS):
    cursor.execute(f"SELECT id FROM Service_Desk WHERE Nome = '{nomeS}'")
    resultado = cursor.fetchone()
    if resultado:
        return resultado[0]
    else:
        return None

#Função da Data/Hora
def SetData_1():
    current_datetime = QDateTime.currentDateTime()
    formulario.dateTimeEdit.setDateTime(current_datetime)

def SetData_2():
    current_datetime = QDateTime.currentDateTime()
    formulario.dateTimeEdit_2.setDateTime(current_datetime)
    formulario.dateTimeEdit.dateTimeChanged.connect(validateDateTimeRange)

def Clear_Data():
    data_hora_inicial = QDateTime(QDate(2000, 1, 1), QTime(0, 0))
    formulario.dateTimeEdit.setDateTime(data_hora_inicial)
    data_hora_final = QDateTime(QDate(2000, 1, 1), QTime(0, 0))
    formulario.dateTimeEdit_2.setDateTime(data_hora_final)
    formulario.lineEdit_13.setText("")
    formulario.plainTextEdit.setPlainText("")

def validateDateTimeRange():
    data_hora_inicial = formulario.dateTimeEdit.dateTime()
    data_hora_final = formulario.dateTimeEdit_2.dateTime()

    if data_hora_final < data_hora_inicial:
        # Defina a data e hora final como a mesma que a data e hora inicial
        formulario.dateTimeEdit_2.setDateTime(data_hora_inicial)

def convert_datetime_to_str(dt):
        return dt.toString("dd/MM/yyyy hh:mm:ss")

#Função para Cadastro de Notas 
def Service():
    
    nomeS = formulario.lineEdit_13.text()
    data_ini = formulario.dateTimeEdit.text()
    data_fim = formulario.dateTimeEdit_2.text()
    Nota = formulario.plainTextEdit.toPlainText()
    ct = QDateTime.currentDateTime()
    ct2 = convert_datetime_to_str(ct)
    Username = getpass.getuser()
    id_nome = obter_id_nome_Service(nomeS)

    if not Nota:
        QtWidgets.QMessageBox.warning(formulario, 'Aviso', 'A nota está vazia!')
        return
    
    if not nomeS:
        QtWidgets.QMessageBox.warning(formulario, 'Aviso', 'O Titulo está vazio!')
        return
    
    cursor.execute(f"SELECT * FROM Service_Desk WHERE Nome='{id_nome}'")
    resultado = cursor.fetchone()
    if resultado:
        QtWidgets.QMessageBox.warning(formulario, 'Aviso', 'Já existe uma nota com esse nome!')
        return

    cursor.execute(f"INSERT INTO Service_Desk (Nome, Data_ini, Data_fim, Nota, Responsavel, Ult_Mod) VALUES ('{nomeS}', '{data_ini}', '{data_fim}', '{Nota}', '{Username}', '{ct2}')")
    con.commit()
    Clear_Data()

#Função para Tabela de Notas
def TableNotas():
    cursor.execute("SELECT ID, Nome, Data_ini, Data_fim, Responsavel, Ult_Mod FROM Service_Desk")
    resultados = cursor.fetchall()

    formulario.tableWidget_2.setRowCount(len(resultados))
    formulario.tableWidget_2.setColumnCount(6)

    for i, row in enumerate(resultados):
        for j, col in enumerate(row):
            formulario.tableWidget_2.setItem(i, j, QtWidgets.QTableWidgetItem(str(col)))

def TableNotas2():
    cursor.execute("SELECT ID, Nome, Data_ini, Data_fim, Responsavel, Ult_Mod FROM Service_Desk")
    resultados = cursor.fetchall()

    formulario.tableWidget.setRowCount(len(resultados))
    formulario.tableWidget.setColumnCount(6)

    for i, row in enumerate(resultados):
        for j, col in enumerate(row):
            formulario.tableWidget.setItem(i, j, QtWidgets.QTableWidgetItem(str(col)))
            formulario.tableWidget.setColumnWidth(0, 2)
            formulario.tableWidget.setColumnWidth(1, 160)
            formulario.tableWidget.setColumnWidth(2, 105)
            formulario.tableWidget.setColumnWidth(3, 105)
            formulario.tableWidget.setColumnWidth(4, 70)
            formulario.tableWidget.setColumnWidth(5, 120)


#Função de Download dos Arquivos em Notas
def DowloadNota():
    # Recuperar o conteúdo da nota e o nome da nota do banco de dados
    n1 = "("
    n2 = ")"
    row = formulario.tableWidget_2.currentRow()
    nome_nota = formulario.tableWidget_2.item(row, 1).text()
    nome_user = formulario.tableWidget_2.item(row, 4).text()
    
    cursor.execute("SELECT Nome, Nota FROM Service_Desk WHERE Nome=?", (nome_nota,))
    resultado = cursor.fetchone()
    if resultado is None:
        QtWidgets.QMessageBox.warning(formulario, 'Aviso', 'Nota não encontrada!')
        return

    nome_nota, nota = resultado

    # Obter o diretório da área de trabalho
    diretorio_area_trabalho = QStandardPaths.writableLocation(QStandardPaths.DesktopLocation)

    # Criar o caminho completo para o arquivo de destino na área de trabalho
    arquivo_destino = os.path.join(diretorio_area_trabalho, f'({nome_user}) {nome_nota}.txt')

    # Salvar o conteúdo da nota no arquivo de destino
    with open(arquivo_destino, 'w', encoding='utf-8') as arquivo:
        arquivo.write(nota)

    QtWidgets.QMessageBox.information(formulario, 'Download concluído', f'A nota "{nome_nota}" foi salva na área de trabalho.')

#Função para Edição de Notas
def EditarNotas():
    linha_selecionada = formulario.tableWidget_2.currentRow()
    if linha_selecionada < 0:
        formulario.Paginas.setCurrentWidget(formulario.Notas)
        QtWidgets.QMessageBox.warning(formulario, 'Aviso', 'Selecione uma Nota!!')
        return

    nome_nota = formulario.tableWidget_2.item(linha_selecionada, 1).text()
    data_ini = formulario.tableWidget_2.item(linha_selecionada, 2).text()
    data_fim = formulario.tableWidget_2.item(linha_selecionada, 3).text()
    responsavel = formulario.tableWidget_2.item(linha_selecionada, 4).text()

    cursor.execute("SELECT Nome, Ult_Mod FROM Service_Desk WHERE Nome=? AND Data_ini=? AND Data_fim=? AND Responsavel=?",
                   (nome_nota, data_ini, data_fim, responsavel))
    Ult_Mod_tuple = cursor.fetchone()

    if Ult_Mod_tuple is not None:
        # O valor de "Ult_Mod" está na segunda posição da tupla (índice 1)
        Ult_Mod = Ult_Mod_tuple[1]

    # Agora você pode usar a variável "Ult_Mod" para definir o texto do label diretamente
    formulario.label_37.setText(Ult_Mod)

    formulario.lineEdit_16.setText(nome_nota)
    formulario.label_34.setText(data_ini)
    formulario.label_35.setText(data_fim)
    formulario.label_37.setText(Ult_Mod)

    cursor.execute("SELECT Nome, Nota FROM Service_Desk WHERE Nome=? AND Data_ini=? AND Data_fim=? AND Responsavel=?",
                   (nome_nota, data_ini, data_fim, responsavel))
    resultado = cursor.fetchone()
    if resultado is None:
        QtWidgets.QMessageBox.warning(formulario, 'Aviso', 'Nota não encontrada!')
        return
    
    nome_nota, nota = resultado
    formulario.plainTextEdit_2.setPlainText(nota)

def EditarNotas2():
    linha_selecionada = formulario.tableWidget.currentRow()
    if linha_selecionada < 0:
        formulario.Paginas.setCurrentWidget(formulario.Lista)
        QtWidgets.QMessageBox.warning(formulario, 'Aviso', 'Selecione uma Nota!!')
        return

    nome_nota = formulario.tableWidget.item(linha_selecionada, 1).text()
    data_ini = formulario.tableWidget.item(linha_selecionada, 2).text()
    data_fim = formulario.tableWidget.item(linha_selecionada, 3).text()
    responsavel = formulario.tableWidget.item(linha_selecionada, 4).text()

    cursor.execute("SELECT Nome, Ult_Mod FROM Service_Desk WHERE Nome=? AND Data_ini=? AND Data_fim=? AND Responsavel=?",
                   (nome_nota, data_ini, data_fim, responsavel))
    Ult_Mod_tuple = cursor.fetchone()

    if Ult_Mod_tuple is not None:
        # O valor de "Ult_Mod" está na segunda posição da tupla (índice 1)
        Ult_Mod = Ult_Mod_tuple[1]

    # Agora você pode usar a variável "Ult_Mod" para definir o texto do label diretamente
    formulario.label_37.setText(Ult_Mod)

    formulario.lineEdit_17.setText(nome_nota)
    formulario.label_41.setText(data_ini)
    formulario.label_43.setText(data_fim)
    formulario.label_45.setText(Ult_Mod)

    cursor.execute("SELECT Nome, Nota FROM Service_Desk WHERE Nome=? AND Data_ini=? AND Data_fim=? AND Responsavel=?",
                   (nome_nota, data_ini, data_fim, responsavel))
    resultado = cursor.fetchone()
    if resultado is None:
        QtWidgets.QMessageBox.warning(formulario, 'Aviso', 'Nota não encontrada!')
        return
    
    nome_nota, nota = resultado
    formulario.plainTextEdit_3.setPlainText(nota)

#Função para Salvar Notas
def SalvarEdNota():
    nome = formulario.lineEdit_16.text()
    timei = formulario.label_34.text()
    timef = formulario.label_35.text()
    timeu = formulario.label_37.text()
    nota = formulario.plainTextEdit_2.toPlainText()
    linha_selecionada = formulario.tableWidget_2.currentRow()
    id = formulario.tableWidget_2.item(linha_selecionada, 0).text()
    responsavel = formulario.tableWidget_2.item(linha_selecionada, 4).text()

    n = getpass.getuser()
    if n not in responsavel:
        formulario.Paginas.setCurrentWidget(formulario.Notas)
        QtWidgets.QMessageBox.warning(formulario, 'Aviso', 'Você não tem permissão para editar essa Nota!!')
        return

    cursor.execute(f"SELECT * FROM Service_Desk WHERE Nome='{nome}' AND Data_ini='{timei}' AND Data_fim='{timef}' AND Ult_Mod='{timeu}' AND Nota='{nota}'")
    resultado = cursor.fetchone()
    if resultado:
        QtWidgets.QMessageBox.warning(formulario, 'Aviso', 'Já existe uma nota com esses valores!')
        return
    ct = QDateTime.currentDateTime()
    ct2 = convert_datetime_to_str(ct)
    Username = getpass.getuser()

    cursor.execute(f"UPDATE Service_Desk SET Ult_Mod='{ct2}', Nota='{nota}', Responsavel='{Username}' WHERE id = {id}")
    con.commit()
    formulario.Paginas.setCurrentWidget(formulario.Notas)
    TableNotas()

def SalvarEdNota2():
    nome = formulario.lineEdit_17.text()
    timei = formulario.label_41.text()
    timef = formulario.label_43.text()
    timeu = formulario.label_45.text()
    nota = formulario.plainTextEdit_3.toPlainText()
    linha_selecionada = formulario.tableWidget.currentRow()
    id = formulario.tableWidget.item(linha_selecionada, 0).text()

    cursor.execute(f"SELECT * FROM Service_Desk WHERE Nome='{nome}' AND Data_ini='{timei}' AND Data_fim='{timef}' AND Ult_Mod='{timeu}' AND Nota='{nota}'")
    resultado = cursor.fetchone()
    if resultado:
        QtWidgets.QMessageBox.warning(formulario, 'Aviso', 'Já existe uma nota com esses valores!')
        return
    ct = QDateTime.currentDateTime()
    ct2 = convert_datetime_to_str(ct)
    Username = getpass.getuser()

    cursor.execute(f"UPDATE Service_Desk SET Ult_Mod='{ct2}', Nota='{nota}', Responsavel='{Username}' WHERE id = {id}")
    con.commit()
    formulario.Paginas.setCurrentWidget(formulario.Lista)
    TableNotas2()

#Função Dia/Tarde
def is_manha_ou_tarde():
    # Obter o horário atual
    now = datetime.now().time()

    # Definir os limites de horário para manhã e tarde
    manha_inicio = datetime.strptime('06:00:00', '%H:%M:%S').time()
    manha_fim = datetime.strptime('13:00:00', '%H:%M:%S').time()
    tarde_inicio = datetime.strptime('13:00:00', '%H:%M:%S').time()
    tarde_fim = datetime.strptime('18:00:00', '%H:%M:%S').time()

    # Verificar se é manhã ou tarde
    if manha_inicio <= now <= manha_fim:
        return "Bom dia"
    elif tarde_inicio <= now <= tarde_fim:
        return "Boa tarde"
    else:
        return "Boa noite"

#Admin
def Validar_User():
    us = getpass.getuser()
    if us not in ["t_fplopes", "gmacedo", "itsilva", "asjsantos"]: # Basta substituir/adicionar os nomes de Usuários
        QtWidgets.QMessageBox.warning(formulario, 'Aviso', 'Você não possui nível de acesso!')
        return 
    else:
        formulario.Paginas.setCurrentWidget(formulario.Lista)

app=QtWidgets.QApplication([])
formulario=uic.loadUi("Style.ui") #Carrega o Visual
un2 = getpass.getuser() #Carrega o User
periodo = is_manha_ou_tarde() #Carrega o Periodo do dia
formulario.label_38.setText(f'{periodo}, {un2}')

#Botões Gerais
formulario.pushButton.clicked.connect(funcao_principal)         # Botão para Cadastro
formulario.pushButton_7.clicked.connect(funcao_pesquisar)       # Botão para Pesquisa
formulario.pushButton_buscar.clicked.connect(funcao_comboBox3_setor) 
formulario.pushButton_11.clicked.connect(funcao_listar)          # Botão para Listar
formulario.pushButton_16.clicked.connect(TableNotas2)            # Botão para Listar
formulario.pushButton_9.clicked.connect(lambda: formulario.Paginas.setCurrentWidget(formulario.Editar)) # Botão conectar a Pagina "Editar"
formulario.pushButton_9.clicked.connect(funcao_editar_ramal)
formulario.pushButton_10.clicked.connect(funcao_excluir_ramal)   # Botão para Excluir
formulario.pushButton_19.clicked.connect(funcao_excluir_nota)   # Botão para Excluir
formulario.pushButton_2.clicked.connect(salvar_alteracoes)   # Botão para Salvar Edição
formulario.pushButton_2.clicked.connect(lambda: formulario.Paginas.setCurrentWidget(formulario.Lista)) # Botão conectar a Pagina "Lista"
formulario.downloadpdf_2.clicked.connect(pdf_2)
formulario.downloadexcel_2.clicked.connect(excel_2)
formulario.pushButton_8.clicked.connect(Service)
formulario.pushButton_8.clicked.connect(TableNotas)
formulario.pushButton_8.clicked.connect(lambda: formulario.Paginas.setCurrentWidget(formulario.Notas)) # Botão conectar a Pagina "Notas"
formulario.pushButton_5.clicked.connect(SetData_1)
formulario.pushButton_6.clicked.connect(SetData_2)
formulario.downloadtxt.clicked.connect(DowloadNota)
formulario.pushButton_13.clicked.connect(lambda: formulario.Paginas.setCurrentWidget(formulario.EditarNotas)) # Botão conectar a Pagina "EditarNotas"
formulario.pushButton_13.clicked.connect(EditarNotas)
formulario.pushButton_18.clicked.connect(lambda: formulario.Paginas.setCurrentWidget(formulario.EditarNotas2)) # Botão conectar a Pagina "EditarNotas2"
formulario.pushButton_18.clicked.connect(EditarNotas2)
formulario.pushButton_20.clicked.connect(SalvarEdNota2)
formulario.pushButton_17.clicked.connect(SalvarEdNota)
formulario.pushButton_15.clicked.connect(funcao_Filtro_pesquisar)

#Validadores
validator = QtGui.QIntValidator()
formulario.lineEdit_2.setValidator(validator)
formulario.lineEdit_12.setValidator(validator)

nome_validator = QtGui.QRegularExpressionValidator(QtCore.QRegularExpression("[A-Za-z ]+"))
formulario.lineEdit.setValidator(nome_validator)
formulario.comboBox.setValidator(nome_validator)
formulario.lineEdit_11.setValidator(nome_validator)

class validar_setor(QtGui.QValidator):
    def validate(self, text, pos):
        if text:
            for char in text:
                if not char.isalpha() and char != '/':
                    return (QtGui.QValidator.Invalid, text, pos)
        return (QtGui.QValidator.Acceptable, text, pos)
setor_validator = validar_setor()
formulario.comboBox.setValidator(setor_validator)
formulario.comboBox_2.setValidator(setor_validator)
formulario.lineEdit_15.setValidator(setor_validator)

formulario.lineEdit_10.setReadOnly(True)
formulario.lineEdit_16.setReadOnly(True)
formulario.lineEdit_17.setReadOnly(True)

#Botões do menu lateral
formulario.pushButton_cadastro.clicked.connect(lambda: formulario.Paginas.setCurrentWidget(formulario.Cadastro))
formulario.pushButton_cadastro.clicked.connect(lambda: formulario.CAD.setCurrentWidget(formulario.CAD_1))
formulario.pushButton_cadastro.clicked.connect(funcao_comboBox_setor)    # ComboBox
formulario.pushButton_buscar.clicked.connect(lambda: formulario.Paginas.setCurrentWidget(formulario.Pesquisa))
formulario.pushButton_listar.clicked.connect(Validar_User)
formulario.pushButton_listar.clicked.connect(clear_combox)
formulario.pushButton_updates.clicked.connect(lambda: formulario.Paginas.setCurrentWidget(formulario.Updates))
formulario.pushButton_3.clicked.connect(lambda: formulario.CAD.setCurrentWidget(formulario.CAD_2))
formulario.pushButton_4.clicked.connect(funcao_cadastro_setor)
formulario.pushButton_4.clicked.connect(lambda: formulario.CAD.setCurrentWidget(formulario.CAD_1))
formulario.pushButton_Service.clicked.connect(lambda: formulario.Paginas.setCurrentWidget(formulario.Notas))
formulario.pushButton_Service.clicked.connect(TableNotas)
formulario.pushButton_12.clicked.connect(lambda: formulario.Paginas.setCurrentWidget(formulario.NewNotas))

#Eventos
formulario.skin.clicked.connect(lambda: formulario.Paginas.setCurrentWidget(formulario.secret))
formulario.smile.clicked.connect(lambda: formulario.secret_2.setCurrentWidget(formulario.stackedWidgetPage2))


#Abre o aplicativo
formulario.show()
app.exec()

# Banco de Dados = SQLite3
# Visual = PyQt5 ("Qt Designer")
# Python 3.11.3