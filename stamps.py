# stamps

import os
import re
import fitz
from collections import defaultdict, OrderedDict
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QLabel, QProgressBar, QMessageBox, QRadioButton, QButtonGroup, QVBoxLayout, QWidget, QHBoxLayout, QSpacerItem, QSizePolicy
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon
import sys

class OrganizadorCarimbador(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Organizador e Carimbador de PDFs")
        self.setGeometry(100, 100, 600, 320)

        # Lista de CNPJs e mapeamento para números
        self.cnpjs_desejados = [
           "05.105.998/0001-02", "10.862.556/0001-40", "05.992.240/0001-33", "02.507.257/0001-60", "03.638.042/0001-40", 
    "04.153.473/0001-80", "02.369.525/0001-24", "03.515.169/0001-72", "01.458.172/0001-76", "11.894.682/0001-40", 
    "09.021.992/0001-08", "18.005.070/0001-06", "18.005.075/0001-20", "19.172.445/0001-87", "18.508.678/0001-45", 
    "18.005.071/0001-42", "18.005.073/0001-31", "22.669.145/0001-12", "18.042.471/0001-28", "18.018.793/0001-31", 
    "18.005.068/0001-29", "18.469.909/0001-59", "18.377.381/0001-98", "18.227.914/0001-55", "18.469.910/0001-83", 
    "19.236.564/0001-56", "18.508.681/0001-69", "19.218.484/0001-78", "17.959.931/0001-14", "19.337.501/0001-96", 
    "18.781.305/0001-43", "18.005.065/0001-95", "17.965.332/0001-03", "18.404.660/0001-01", "19.337.499/0001-55", 
    "18.005.062/0001-51", "18.005.061/0001-07", "18.781.306/0001-98", "18.018.791/0001-42", "03.992.934/0001-45", 
    "06.292.993/0001-07", "04.503.753/0001-70", "10.729.476/0001-11", "10.317.955/0001-20", "27.547.527/0001-97", 
    "26.915.636/0001-57", "26.712.478/0001-38", "08.997.252/0001-49", "10.219.115/0001-25", "10.203.184/0001-40", 
    "03.588.270.0001-53", "18.781.304.0001-07", "18.377.380/0001-43", "32.753.965/0001-41", "35.494.840/0001-32", 
    "49.731.012/0001-85", "10.204.389/0001-40", "35.575.350/0001-60", "36.484.216/0001-17", "09.529.695/0001-78", 
    "43.157.833/0001-73", "44.906.774/0001-51", "01.944.121/0001-54", "02.693.579/0001-40", "65.748.147/0001-00", 
    "13.445.954/0001-50", "10.968.502/0001-64", "43.557.693/0001-20", "07.893.536/0001-22", "53.989.109/0001-60", 
    "06.327.564/0001-10", "05.863.083/0001-66", "44.108.629/0001-25", "56.136.807/0001-00", "53.977.773/0001-99", 
    "60.884.566/0001-55", "43.299.650/0001-92", "22.506.756/0001-40", "19.700.192/0001-77", "12.109.019/0001-50", 
    "12.481.045/0001-04", "04.909.888/0001-30", "15.800.103/0001-03", "14.891.983/0001-08", "20.975.527/0001-49", 
    "66.599.887/0001-94", "37.405.361/0001-28", "34.995.268/0001-22", "56.900.129/0001-00", "40.203.038/0001-86", 
    "30.432.615/0001-58", "47.123.500/0001-84", "00.612.412/0001-82", "10.627.882/0001-73", "46.466.279/0001-02", 
    "12.616.427/0001-06", "03.085.922/0001-37", "18.508.679/0001-90", "18.781.303/0001-54", "18.043.212/0001-11", 
    "18.043.211/0001-77", "10.410.220/0001-47", "49.093.156/0001-53", "31.232.818/0001-63", "03.303.976/0001-21", 
    "03.796.292/0001-09", "07.356.133/0001-44", "08.699.770/0001-86", "10.532.173/0001-04", "15.840.053/0001-98", 
    "33.458.899/0001-40", "34.662.094/0001-86", "18.923.270/0001-30", "03.136.100/0001-38", "51.902.725/0001-06", 
    "02.169.559/0001-75", "17.837.137/0001-06", "09.536.396/0001-60", "02.828.214/0001-86", "08.881.961/0001-64", 
    "32.267.669/0001-30", "33.358.113/0001-12", "26.666.665/0001-22", "31.507.732/0001-04", "29.562.682/0001-08", 
    "37.782.063/0001-57", "18.683.354/0001-43", "42.995.751/0001-35", "48.238.285/0001-20", "24.660.979/0001-92", 
    "05.115.935/0001-37", "77.488.005/0001-30", "77.488.005/0002-10", "77.488.005/0008-06", "77.488.005/0009-97", 
    "31.149.457/0001-96", "77.488.005/0004-82", "77.488.005/0006-44", "77.488.005/0010-20", "77.488.005/0007-25", 
    "77.488.005/0003-00", "77.488.005/0011-01", "07.551.516/0001-73", "55.509.384/0001-64", "07.674.988/0001-13"
        ]
        self.numeros_por_cnpj = {
            "05.105.998/0001-02": "084", "10.862.556/0001-40": "141", "05.992.240/0001-33": "150", "02.507.257/0001-60": "156",
    "03.638.042/0001-40": "159", "04.153.473/0001-80": "161", "02.369.525/0001-24": "167", "03.515.169/0001-72": "173", 
    "01.458.172/0001-76": "180", "11.894.682/0001-40": "181", "09.021.992/0001-08": "201", "18.005.070/0001-06": "202", 
    "18.005.075/0001-20": "211", "19.172.445/0001-87": "212", "18.508.678/0001-45": "220", "18.005.071/0001-42": "223", 
    "18.005.073/0001-31": "230", "22.669.145/0001-12": "231", "18.042.471/0001-28": "243", "18.018.793/0001-31": "246", 
    "18.005.068/0001-29": "252", "18.469.909/0001-59": "261", "18.377.381/0001-98": "270", "18.227.914/0001-55": "271", 
    "18.469.910/0001-83": "286", "19.236.564/0001-56": "293", "18.508.681/0001-69": "295", "19.218.484/0001-78": "299", 
    "17.959.931/0001-14": "345", "19.337.501/0001-96": "352", "18.781.305/0001-43": "353", "18.005.065/0001-95": "362", 
    "17.965.332/0001-03": "364", "18.404.660/0001-01": "369", "19.337.499/0001-55": "371", "18.005.062/0001-51": "401", 
    "18.005.061/0001-07": "402", "18.781.306/0001-98": "403", "18.018.791/0001-42": "404", "03.992.934/0001-45": "405", 
    "06.292.993/0001-07": "406", "04.503.753/0001-70": "407", "10.729.476/0001-11": "408", "10.317.955/0001-20": "409", 
    "27.547.527/0001-97": "410", "26.915.636/0001-57": "411", "26.712.478/0001-38": "412", "08.997.252/0001-49": "414", 
    "10.219.115/0001-25": "415", "10.203.184/0001-40": "416", "03.588.270.0001-53": "417", "18.781.304.0001-07": "418", 
    "18.377.380/0001-43": "419", "32.753.965/0001-41": "420", "35.494.840/0001-32": "421", "49.731.012/0001-85": "422", 
    "10.204.389/0001-40": "423", "35.575.350/0001-60": "424", "36.484.216/0001-17": "425", "09.529.695/0001-78": "426", 
    "43.157.833/0001-73": "427", "44.906.774/0001-51": "428", "01.944.121/0001-54": "429", "02.693.579/0001-40": "430", 
    "65.748.147/0001-00": "431", "13.445.954/0001-50": "432", "10.968.502/0001-64": "433", "43.557.693/0001-20": "434", 
    "07.893.536/0001-22": "435", "53.989.109/0001-60": "436", "06.327.564/0001-10": "437", "05.863.083/0001-66": "438", 
    "44.108.629/0001-25": "439", "56.136.807/0001-00": "440", "53.977.773/0001-99": "441", "60.884.566/0001-55": "442", 
    "43.299.650/0001-92": "443", "22.506.756/0001-40": "445", "19.700.192/0001-77": "446", "12.109.019/0001-50": "447", 
    "12.481.045/0001-04": "448", "04.909.888/0001-30": "449", "15.800.103/0001-03": "450", "14.891.983/0001-08": "451", 
    "20.975.527/0001-49": "452", "66.599.887/0001-94": "453", "37.405.361/0001-28": "454", "34.995.268/0001-22": "455", 
    "56.900.129/0001-00": "456", "40.203.038/0001-86": "457", "30.432.615/0001-58": "458", "47.123.500/0001-84": "459", 
    "00.612.412/0001-82": "460", "10.627.882/0001-73": "461", "46.466.279/0001-02": "462", "12.616.427/0001-06": "463", 
    "03.085.922/0001-37": "464", "18.508.679/0001-90": "465", "18.781.303/0001-54": "466", "18.043.212/0001-11": "467", 
    "18.043.211/0001-77": "468", "10.410.220/0001-47": "469", "49.093.156/0001-53": "470", "31.232.818/0001-63": "471", 
    "03.303.976/0001-21": "472", "03.796.292/0001-09": "473", "07.356.133/0001-44": "474", "08.699.770/0001-86": "475", 
    "10.532.173/0001-04": "476", "15.840.053/0001-98": "477", "33.458.899/0001-40": "478", "34.662.094/0001-86": "479", 
    "18.923.270/0001-30": "480", "03.136.100/0001-38": "481", "51.902.725/0001-06": "482", "02.169.559/0001-75": "483", 
    "17.837.137/0001-06": "484", "09.536.396/0001-60": "485", "02.828.214/0001-86": "486", "08.881.961/0001-64": "487", 
    "32.267.669/0001-30": "488", "33.358.113/0001-12": "490", "26.666.665/0001-22": "491", "31.507.732/0001-04": "492", 
    "29.562.682/0001-08": "493", "37.782.063/0001-57": "494", "18.683.354/0001-43": "495", "42.995.751/0001-35": "496", 
    "48.238.285/0001-20": "497", "24.660.979/0001-92": "498", "05.115.935/0001-37": "499", "77.488.005/0001-30": "501", 
    "77.488.005/0002-10": "502", "77.488.005/0008-06": "503", "77.488.005/0009-97": "504", "31.149.457/0001-96": "505", 
    "77.488.005/0004-82": "506", "77.488.005/0006-44": "507", "77.488.005/0010-20": "508", "77.488.005/0007-25": "509", 
    "77.488.005/0003-00": "510", "77.488.005/0011-01": "511", "07.551.516/0001-73": "513", "55.509.384/0001-64": "514", 
    "07.674.988/0001-13": "515"
        }


        # Layout principal (vertical)
        layout_principal = QVBoxLayout()

        # Widget central para organizar os elementos
        widget_central = QWidget()
        layout_central = QVBoxLayout(widget_central)

        # Definir o ícone da janela
        self.setWindowIcon(QIcon('Icone.ico'))  # Substitua pelo caminho correto do seu ícone

        # Fonte
        self.fonte = QFont()
        self.fonte.setPointSize(14)

        # Label da posição do carimbo (centralizado)
        self.label_posicao = QLabel("Selecione A Posição Do Carimbo:", self)
        self.label_posicao.setFont(self.fonte)  # Usar self.fonte
        self.label_posicao.setFont(self.fonte)
        layout_label = QHBoxLayout()  # Layout horizontal para o label
        layout_label.addStretch(1)  # Espaço flexível à esquerda
        layout_label.addWidget(self.label_posicao)
        layout_label.addStretch(1)  # Espaço flexível à direita
        layout_central.addLayout(layout_label)
        
        # Radio buttons (criados na função criar_radio_buttons)
        self.criar_radio_buttons()
        layout_radio_buttons = QHBoxLayout()
        layout_radio_buttons.addStretch(1)
        layout_radio_buttons.addWidget(self.radio_esquerda)
        layout_radio_buttons.addItem(QSpacerItem(20, 0, QSizePolicy.Fixed, QSizePolicy.Minimum))  # Espaçamento
        layout_radio_buttons.addWidget(self.radio_meio)
        layout_radio_buttons.addItem(QSpacerItem(20, 0, QSizePolicy.Fixed, QSizePolicy.Minimum))  # Espaçamento
        layout_radio_buttons.addWidget(self.radio_direita)
        layout_radio_buttons.addStretch(1)
        layout_central.addLayout(layout_radio_buttons)

        # Botão de executar e encerrar (Movidos para cima)
        self.botao_executar = QPushButton("Executar", self)
        self.botao_executar.clicked.connect(self.iniciar_processo)
        self.botao_encerrar = QPushButton("Encerrar", self)
        self.botao_encerrar.clicked.connect(self.close)  # Fechar a janela ao clicar
        self.botao_encerrar.setEnabled(False)  # Desabilitado inicialmente

        # Barra de progresso e label de status (Movidos para cima)
        self.barra_progresso = QProgressBar(self)
        self.label_status = QLabel(self)

        # Botão de organizar sem carimbo (novo)
        self.botao_organizar_sem_txt = QPushButton("Apenas Organizar", self)
        self.botao_organizar_sem_txt.clicked.connect(self.iniciar_processo_sem_txt)
        self.botao_organizar_sem_txt.move(100, 130)  # Posicionar abaixo dos radio buttons

        # Botões (horizontalmente)
        layout_botoes = QHBoxLayout()
        layout_botoes.addWidget(self.botao_organizar_sem_txt)  # Adiciona o novo botão
        layout_botoes.addWidget(self.botao_executar)
        layout_botoes.addWidget(self.botao_encerrar)
        layout_central.addLayout(layout_botoes)

        layout_central.addWidget(self.barra_progresso)
        layout_central.addItem(QSpacerItem(60, 15, QSizePolicy.Minimum, QSizePolicy.Fixed))  # Espaçamento vertical de 10 pixels
        layout_central.addWidget(self.label_status)

        # Ajustar altura da barra de progresso
        altura_botao = self.botao_executar.height()  # Obter a altura do botão
        self.barra_progresso.setFixedHeight(altura_botao)  # Definir a mesma altura para a barra

        # Barra de progresso e label de status
        layout_central.addWidget(self.barra_progresso)
        layout_central.addWidget(self.label_status)

        # Definir o layout central como layout do widget central
        widget_central.setLayout(layout_central)

        # Adicionar o widget central ao layout principal
        layout_principal.addWidget(widget_central)

        # Criar um widget para conter o layout principal
        container = QWidget()
        container.setLayout(layout_principal)

        # Definir o widget container como o widget central da janela
        self.setCentralWidget(container)    
    
        # Posição do carimbo
        self.posicao_carimbo = "Esquerda"  # Valor padrão
        # self.criar_radio_buttons()

        # Mensagem de status
        self.label_status = QLabel(self)
        self.label_status.setGeometry(100, 350, 400, 25)
        self.label_status.setAlignment(Qt.AlignCenter)

        # Thread para o processamento em segundo plano
        self.thread_processamento = None

    def criar_radio_buttons(self):
        self.button_group = QButtonGroup(self)
        self.radio_esquerda = QRadioButton("Esquerda", self)
        self.radio_esquerda.setChecked(True)  # Marcar "Esquerda" como padrão
        self.radio_esquerda.move(120, 130)
        self.radio_esquerda.setFont(self.fonte)  # Aplicar a fonte aos radio buttons
        self.button_group.addButton(self.radio_esquerda)

        self.radio_meio = QRadioButton("Meio", self)
        self.radio_meio.move(250, 130)
        self.radio_meio.setFont(self.fonte)
        self.button_group.addButton(self.radio_meio)

        self.radio_direita = QRadioButton("Direita", self)
        self.radio_direita.move(330, 130)
        self.radio_direita.setFont(self.fonte)
        self.button_group.addButton(self.radio_direita)

        self.button_group.buttonClicked.connect(self.atualizar_posicao_carimbo)

    def atualizar_posicao_carimbo(self, button):
        self.posicao_carimbo = button.text()

    def iniciar_processo(self):
        self.botao_executar.setEnabled(False)
        self.barra_progresso.setValue(0)
        self.label_status.setText("Processando...")

        # Inicia a thread de processamento
        self.thread_processamento = ThreadProcessamento(self)
        self.thread_processamento.progresso.connect(self.atualizar_progresso)
        self.thread_processamento.concluido.connect(self.processamento_concluido)
        self.thread_processamento.start()

    def atualizar_progresso(self, valor):
        self.barra_progresso.setValue(valor)

    def processamento_concluido(self, mensagem):
        self.botao_executar.setEnabled(True)
        self.label_status.setText(mensagem)  # Atualizar a label com a mensagem
        QMessageBox.information(self, "Processo Concluído", mensagem)  # Exibir a mensagem em uma caixa de diálogo
        self.botao_encerrar.setEnabled(True)

        # Habilitar botão de encerrar após a conclusão
        self.botao_encerrar.setEnabled(True)

    # Função para extrair páginas que contêm um CNPJ específico
    def extrair_paginas_com_cnpj(self, pdf_document, cnpj):
        paginas_com_cnpj = []
        for pagina_num in range(len(pdf_document)):
            pagina = pdf_document[pagina_num]
            texto = pagina.get_text()
            if re.search(cnpj, texto):
                paginas_com_cnpj.append(pagina_num)
        return paginas_com_cnpj

    # Função para limpar o CNPJ e torná-lo um nome de arquivo válido
    def limpar_cnpj(self, cnpj):
        return re.sub(r'[^a-zA-Z0-9]', '.', cnpj)

    # Função para adicionar carimbo a todas as páginas do PDF
    def carimbar_pdf(self, pdf_document, texto_carimbo, posicao_carimbo):
        for pagina_num in range(len(pdf_document)):
            pagina = pdf_document[pagina_num]
            if posicao_carimbo == "Esquerda":
                pagina.insert_text((50, 700), texto_carimbo, fontsize=12, color=(0, 0, 0))
            elif posicao_carimbo == "Meio":
                pagina.insert_text((250, 700), texto_carimbo, fontsize=12, color=(0, 0, 0))
            elif posicao_carimbo == "Direita":
                pagina.insert_text((430, 700), texto_carimbo, fontsize=12, color=(0, 0, 0))

    # Função para obter o caminho do recurso
    def obter_caminho_recurso(self, caminho_relativo):
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, caminho_relativo)
        return os.path.join(os.path.abspath("."), caminho_relativo)
    
    def iniciar_processo_sem_txt(self):
        self.botao_executar.setEnabled(False)
        # self.botao_organizar_sem_txt.setEnabled(False)
        self.botao_organizar_sem_txt.setEnabled(False)  # Desabilitar o novo botão também
        self.barra_progresso.setValue(0)
        self.label_status.setText("Processando...")

        # Inicia a thread de processamento sem usar o modelo.txt
        self.thread_processamento = ThreadProcessamento(self, usar_modelo_txt=False)
        self.thread_processamento.progresso.connect(self.atualizar_progresso)
        self.thread_processamento.concluido.connect(self.processamento_concluido)
        self.thread_processamento.start()

class ThreadProcessamento(QThread):
    progresso = pyqtSignal(int)
    concluido = pyqtSignal(str)

    def __init__(self, parent=None, usar_modelo_txt=True):
        super().__init__(parent)
        self.organizador_carimbador = parent
        self.usar_modelo_txt = usar_modelo_txt  # Armazena se deve usar o modelo.txt ou não

    def run(self):
        try:
            # Diretórios
            base_folder = r"C:\Organizar"
            input_folder = os.path.join(base_folder, "Desorganizadas")
            pasta_destino = os.path.join(base_folder, "Organizadas")

            # Criar pastas se não existirem
            os.makedirs(input_folder, exist_ok=True)
            os.makedirs(pasta_destino, exist_ok=True)

            # Criar arquivo modelo.txt se não existir
            carimbo_texto_path = os.path.join(base_folder, "modelo.txt")
            if not os.path.exists(carimbo_texto_path):
                with open(carimbo_texto_path, 'w', encoding='utf-8') as file:
                    file.write("Despesa:   Investimento:   \nO Próprio:   Rateio:   \nClassificação: \nVencimento: \nNome:  \nDep: ")

            # Lendo o texto do carimbo de um arquivo .txt com codificação UTF-8
            with open(carimbo_texto_path, 'r', encoding='utf-8') as file:
                texto_carimbo = file.read().strip()
            
            # Carregar o texto do carimbo (apenas se usar_modelo_txt for True)
            if self.usar_modelo_txt:
                with open(carimbo_texto_path, 'r', encoding='utf-8') as file:
                    texto_carimbo = file.read().strip()
            else:
                texto_carimbo = ""  # Texto vazio se não usar o modelo.txt

            # Dicionário para armazenar os PDFs separados por CNPJ
            pdfs_por_cnpj = defaultdict(list)

            # Processar arquivos PDF na pasta de entrada
            for root, _, files in os.walk(input_folder):
                for file in files:
                    if file.endswith(".pdf"):
                        input_pdf_file = os.path.join(root, file)
                        pdf_document = fitz.open(input_pdf_file)
                        for cnpj in self.organizador_carimbador.cnpjs_desejados:
                            paginas = self.organizador_carimbador.extrair_paginas_com_cnpj(pdf_document, cnpj)
                            if paginas:
                                pdfs_por_cnpj[cnpj].append((input_pdf_file, paginas))
                        pdf_document.close()

            # Ordenar os arquivos e páginas para cada CNPJ
            pdfs_por_cnpj_ordenado = OrderedDict()
            for cnpj in self.organizador_carimbador.cnpjs_desejados:
                if cnpj in pdfs_por_cnpj:
                    pdfs_por_cnpj_ordenado[cnpj] = sorted(pdfs_por_cnpj[cnpj], key=lambda x: x[0])

            # Salvar os PDFs combinados por CNPJ
            pdf_files = []
            total_arquivos = len(pdfs_por_cnpj_ordenado)
            arquivos_processados = 0
            for cnpj, lista_pdf_paginas in pdfs_por_cnpj_ordenado.items():
                cnpj_limpo = self.organizador_carimbador.limpar_cnpj(cnpj)
                numero_cnpj = self.organizador_carimbador.numeros_por_cnpj.get(cnpj, "XX")
                output_pdf_file = os.path.join(pasta_destino, f'{numero_cnpj}_{cnpj_limpo}.pdf')

                pdf_combinado = fitz.open()
                for input_pdf_file, paginas in lista_pdf_paginas:
                    pdf_document = fitz.open(input_pdf_file)
                    for pagina_num in sorted(paginas):
                        pdf_combinado.insert_pdf(pdf_document, from_page=pagina_num, to_page=pagina_num)
                    pdf_document.close()

                self.organizador_carimbador.carimbar_pdf(pdf_combinado, texto_carimbo, self.organizador_carimbador.posicao_carimbo)
                pdf_combinado.save(output_pdf_file)
                pdf_combinado.close()
                pdf_files.append(output_pdf_file)

                arquivos_processados += 1
                progresso = int((arquivos_processados / total_arquivos) * 100)
                self.progresso.emit(progresso)

            # Nome do arquivo PDF final unificado
            output_final_pdf_file = os.path.join(pasta_destino, 'PDF_UNIFICADO.pdf')

            # Verificar se há pelo menos um arquivo PDF na pasta
            if pdf_files:
                pdfs_unificados = fitz.open()
                for pdf_file in pdf_files:
                    pdf_to_merge = fitz.open(pdf_file)
                    pdfs_unificados.insert_pdf(pdf_to_merge)
                    pdf_to_merge.close()
                pdfs_unificados.save(output_final_pdf_file)
                pdfs_unificados.close()
                print("PDF's unificado com sucesso:", output_final_pdf_file)
                self.concluido.emit("Arquivos gerados e carimbados com sucesso!")  # Emitir sinal apenas se houver PDFs
            else:
                self.concluido.emit("Nenhum arquivo PDF válido encontrado na pasta.")  # Emitir sinal com a mensagem
                print("Nenhum arquivo PDF encontrado na pasta.")
        except Exception as e:
            self.concluido.emit(f"Erro: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OrganizadorCarimbador()
    window.show()
    sys.exit(app.exec_())


