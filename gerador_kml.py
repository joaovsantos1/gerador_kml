import os
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd
import sys
import zipfile
import shutil
import speech_recognition as sr

def recognize_speech():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Diga algo...")
        audio = recognizer.listen(source)

        try:
            text = recognizer.recognize_google(audio, language="pt-BR")
            print(f"Voce disse: {text}")
            return text
        except sr.UnknownValueError:
            print("Não foi possível entender o áudio")
        except sr.RequestError as e:
            print(f"Erro ao acessar o serviço de reconhecimento de voz: {e}")
        return None

def sanitize_kml_content(content):
    content = content.replace('&', '&amp;')
    content = content.replace('<', '&lt;')
    content = content.replace('>', '&gt;')
    content = content.replace('"', '&quot;')
    content = content.replace("'", '&#39;')
    return content

def extract_kmz(kmz_path, extract_to):
    """Descompacta o arquivo KMZ e retorna o caminho do arquivo KML extraído"""
    with zipfile.ZipFile(kmz_path, 'r') as kmz:
        kmz.extractall(extract_to)
        for file in os.listdir(extract_to):
            if file.endswith('.kml'):
                return os.path.join(extract_to, file)
    return None

def create_kml(coordinates, names, descriptions, sheet_name="Combined Data", filename="output.kml"):
    kml_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
<Document>
<name>{sanitize_kml_content(sheet_name)}</name>
<description>Gerado KML do arquivo: {sanitize_kml_content(sheet_name)}</description>
"""
    for coord, name, description in zip(coordinates, names, descriptions):
        kml_content += f"""
<Placemark>
    <name>{sanitize_kml_content(name)}</name>
    <description>{sanitize_kml_content(description)}</description>
    <Point>
        <coordinates>{coord}</coordinates>
    </Point>
</Placemark>
"""
    kml_content += """
</Document>
</kml>"""

    with open(filename, "w", encoding="utf-8-sig") as file:
        file.write(kml_content)

def read_excel(file_path, name_column='Etiqueta', description_column='Descrição'):
    df = pd.read_excel(file_path, header=None)

    if 'Longitude' in df.iloc[0].values and 'Latitude' in df.iloc[0].values:
        header = 0
    elif 'Longitude' in df.iloc[1].values and 'Latitude' in df.iloc[1].values:
        header = 1
    else:
        raise ValueError("As colunas 'Longitude' e 'Latitude' não foram encontradas na planilha.")
    
    df = pd.read_excel(file_path, header=header)

    if 'Longitude' in df.columns and 'Latitude' in df.columns:
        df = df.dropna(subset=['Longitude', 'Latitude'])
        coordinates = df['Longitude'].astype(str) + ',' + df['Latitude'].astype(str) + ',0'
        names = df[name_column].fillna("").tolist()

        if description_column in df.columns:
            descriptions = df[description_column].fillna("").tolist()
        else:
            descriptions = [""] * len(names)
    else:
        raise ValueError("As colunas 'Longitude' e 'Latitude' não foram encontradas na planilha.")

    return coordinates.tolist(), names, descriptions

def process_all_excels_in_folder(folder_path, single_kml, name_column):
    if single_kml:
        all_coordinates = []
        all_names = []
        all_descriptions = []
        
        for filename in os.listdir(folder_path):
            if filename.endswith(".xlsx") or filename.endswith(".xls"):
                file_path = os.path.join(folder_path, filename)
                try:
                    coordinates, names = read_excel(file_path, name_column=name_column)
                    all_coordinates.extend(coordinates)
                    all_names.extend(names)
                    all_descriptions.extend(descriptions)
                except Exception as e:
                    print(f"Erro ao processar {filename}: {e}")
        
        if all_coordinates:
            output_kml = os.path.join(folder_path, "all_data.kml")
            create_kml(all_coordinates, all_names, all_descriptions, sheet_name="All Data", filename=output_kml)
            print(f"KML gerado com sucesso para todas as planilhas.")
        else:
            print("Nenhuma coordenada válida encontrada em todas as planilhas.")
    else:
        for filename in os.listdir(folder_path):
            if filename.endswith(".xlsx") or filename.endswith(".xls"):
                file_path = os.path.join(folder_path, filename)
                try:
                    coordinates, names, descriptions = read_excel(file_path, name_column=name_column)
                    sheet_name = os.path.splitext(filename)[0]
                    output_kml = filename.replace(".xlsx", ".kml").replace(".xls", ".kml")
                    create_kml(coordinates, names, descriptions, sheet_name=sheet_name, filename=os.path.join(folder_path, output_kml))
                    print(f"KML gerado com sucesso para: {filename}")
                except Exception as e:
                    print(f"Erro ao processar {filename}: {e}")



def process_kml_or_kmz(file_path, output_excel, name_column):
    """Processa o arquivo KML ou KMZ e converte para Excel"""
    if file_path.endswith('.kmz'):
        # Criar um diretório temporário para extrair o KMZ
        temp_dir = os.path.join(os.path.dirname(file_path), "temp_kmz")
        os.makedirs(temp_dir, exist_ok=True)
        
        # Extrair KMZ
        kml_file = extract_kmz(file_path, temp_dir)
        
        if kml_file:
            print(f"Arquivo KMZ descompactado: {kml_file}")
            kml_to_excel(kml_file, output_excel, name_column)
        else:
            print("Erro: Nenhum arquivo KML encontrado dentro do KMZ")
        
        # Limpar diretório temporário após o processamento
        shutil.rmtree(temp_dir)

    elif file_path.endswith('.kml'):
        # Converter diretamente o KML
        kml_to_excel(file_path, output_excel, name_column)

    else:
        print("Erro: O arquivo não é KML nem KMZ")

def combine_kmls(kml_files, output_file="all_kml.kml"):
    colors = [
        "ff0000ff",  # Azul
        "ff00ff00",  # Verde
        "ffff0000",  # Vermelho
        "ff00ffff",  # Ciano
        "ffff00ff",  # Magenta
        "ffffff00",  # Amarelo
        "ff9900ff",  # Laranja
    ]

    combined_kml = """<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
<Document>
<name>KMLs Combinados</name>
<description>Todos os arquivos em KMLs foram combinados com diferentes cores</description>
"""

    for i, color in enumerate(colors):
        combined_kml += f"""
<Style id="style{i}">
    <IconStyle>
        <color>{color}</color>
        <Icon>
            <href>http://maps.google.com/mapfiles/kml/paddle/1.png</href>
        </Icon>
    </IconStyle>
</Style>
"""

    for i, kml_file in enumerate(kml_files):
        try:
            with open(kml_file, 'r', encoding='utf-8') as file:
                content = file.read()
                placemarks = content.split('<Placemark>')[1:]  # Evita o cabeçalho do arquivo
                file_name = os.path.basename(kml_file).split('.')[0]  # Nome do arquivo sem extensão
                for placemark in placemarks:
                    placemark_content = placemark.split("</Placemark>")[0]

                    if "<description>" not in placemark_content:
                        placemark_content += f"""
<description>Origem: {file_name}</description>"""

                    combined_kml += f"""
<Placemark>
    <styleUrl>#style{i}</styleUrl>
    {placemark_content}
</Placemark>\n"""
        except ET.ParseError as e:
            print(f"Erro de análise XML ao processar {kml_file}: {e}")
        except Exception as e:
            print(f"Erro ao combinar {kml_file}: {e}")

    combined_kml += "</Document>\n</kml>"

    with open(output_file, "w", encoding="utf-8-sig") as f:
        f.write(combined_kml)
    print(f"KML combinado gerado com sucesso: {output_file}")

def kml_to_excel(kml_file, output_excel, name_column):
    tree = ET.parse(kml_file)
    root = tree.getroot()
    
    # Define o namespace do KML
    namespace = {'kml': 'http://www.opengis.net/kml/2.2'}

    # Itera sobre os elementos Placemark
    names = []
    coordinates = []
    descriptions = []
    for placemark in root.findall('.//kml:Placemark', namespace):
        name = placemark.find('kml:name', namespace)
        coord = placemark.find('.//kml:coordinates', namespace)
        description = placemark.find('kml:description', namespace)
        
        if name is not None and coord is not None:
            names.append(name.text.strip())
            coordinates.append(coord.text.strip().split(',')[:2])  # Pega apenas Longitude e Latitude
            if description.text is not None:
                descriptions.append(description.text.strip())
            else:
                descriptions.append("") #adiciona vazio se nao houver nada

    # Cria um DataFrame
    df = pd.DataFrame(coordinates, columns=['Longitude', 'Latitude'])
    df[name_column] = names
    df['Descrição'] = descriptions #adiciona como nova coluna
    
    # Salva o DataFrame em um arquivo Excel
    df.to_excel(output_excel, index=False)

def run_program(folder_path, combine=False, name_column='Etiqueta', convert_kml=False):
    if convert_kml:
        kml_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith('.kml') or f.endswith('.kmz')]
        for kml_file in kml_files:
            output_excel = os.path.splitext(kml_file)[0] + '.xlsx'
            process_kml_or_kmz(kml_file, output_excel, name_column)
            print(f"Planilha Excel gerada com sucesso para: {kml_file}")
    else:
        process_all_excels_in_folder(folder_path, single_kml=combine, name_column=name_column)
        print("Processamento concluído. Todos os arquivos KML foram gerados na pasta selecionada.")

def get_resource_path(relative_path):
    """Retorna o caminho absoluto do recurso, seja executado em modo script ou como um executável."""
    if hasattr(sys, '_MEIPASS'):
        # Quando em modo executável, o ícone estará dentro de uma pasta temporária
        return os.path.join(sys._MEIPASS, relative_path)
    # Em modo script, o caminho relativo será usado
    return os.path.join(os.path.abspath("."), relative_path)

def voice_command_listener(name_column_entry, folder_path_entry, select_folder, start_processing, combine_kmls_action_wrapper, kml_convert_var):
    """Função para ativar o comando de voz e preencher os campos"""
    recognizer = sr.Recognizer()
    mic = sr.Microphone()

    # Função para capturar a fala e reconhecê-la como texto
    def recognize_speech():
        with mic as source:
            recognizer.adjust_for_ambient_noise(source)
            audio = recognizer.listen(source)

        try:
            speech_text = recognizer.recognize_google(audio, language='pt-BR')
            return speech_text.lower()  # Converte o texto para minúsculas para facilitar a comparação
        except sr.UnknownValueError:
            return None
        except sr.RequestError:
            return None

    

def run_gui():
    def start_processing():
        folder_path = folder_path_entry.get().strip()
        if not folder_path:
            messagebox.showerror("Erro", "Por favor, selecione a pasta que contém os arquivos.")
            return
        
        name_column = name_column_entry.get().strip()
        if not name_column:
            messagebox.showerror("Erro", "O nome da coluna não pode estar vazio. Por favor, insira um nome válido.")
            return
        
        convert_kml = kml_convert_var.get()
        run_program(folder_path, combine=combine_var.get(), name_column=name_column, convert_kml=convert_kml)
        
        if convert_kml:
            messagebox.showinfo("Processamento Concluído", "Planilhas Excel geradas para todos os arquivos KML na pasta selecionada.")
        elif combine_var.get():
            messagebox.showinfo("Processamento Concluído", "KML único 'all_data.kml' gerado na pasta selecionada.")
        else:
            messagebox.showinfo("Processamento Concluído", "Arquivos KML individuais gerados na pasta selecionada.")

    def select_folder():
        folder_path = filedialog.askdirectory()
        folder_path_entry.delete(0, tk.END)
        folder_path_entry.insert(0, folder_path)

    def combine_kmls_action_wrapper():
        folder_path = folder_path_entry.get().strip()
        if not folder_path:
            messagebox.showerror("Erro", "Por favor, selecione a pasta que contém os arquivos KML.")
            return

        kml_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith('.kml')]
        
        if kml_files:
            output_file = os.path.join(folder_path, "all_kml.kml")
            combine_kmls(kml_files, output_file)
            messagebox.showinfo("Processamento Concluído", f"Arquivos KML combinados em '{output_file}' gerado com sucesso.")
        else:
            messagebox.showerror("Erro", "Nenhum arquivo KML encontrado para combinar.")
            # Função para preencher os campos com base no que foi falado
    def voice_command_wrapper():
        messagebox.showinfo("Comando de Voz", "Aguardando comando de voz para preencher os campos...")

        # Primeiro comando: Nome da coluna
        messagebox.showinfo("Comando de Voz", "Diga o nome da coluna para os marcadores.")
        column_name = recognize_speech().capitalize()
        if column_name:
            name_column_entry.delete(0, tk.END)
            name_column_entry.insert(0, column_name)
        else:
            messagebox.showerror("Erro", "Não foi possível reconhecer o nome da coluna.")

        # Segundo comando: Caminho da pasta
        messagebox.showinfo("Comando de Voz", "Diga o caminho da pasta ou fale 'selecionar' para escolher manualmente.")
        folder_command = recognize_speech()
        if folder_command == "selecionar":
            select_folder()
        elif folder_command:
            folder_path = folder_command.capitalize()
            if os.path.exists(folder_path):
                folder_path_entry.delete(0, tk.END)
                folder_path_entry.insert(0, folder_command)
            else:
                messagebox.showerror("Caminho Inválido", "O caminho da pasta não é válido. Selecione manualmente.")
                select_folder()
        else:
            messagebox.showerror("Erro","Não foi possível reconhecer o caminho da pasta.")

        # Terceiro comando: Ação a ser realizada
        messagebox.showinfo("Comando de Voz", "Diga a ação que deseja realizar: 'iniciar', 'combinar', 'converter' ou 'unir'.")
        action = recognize_speech()
        if action == "iniciar":
            start_processing()
        elif action == "combinar":
            combine_kmls_action_wrapper()
        elif action == "converter":
            kml_convert_var.set(True)
            start_processing()
        else:
            messagebox.showerror("Erro", "Não foi possível reconhecer a ação ou ação não válida.")

    root = tk.Tk()
    root.title("Gerador de KML")

    # Definindo a geometria da janela
    window_width = 490
    window_height = 220
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")

    # Definindo o ícone da janela com o caminho obtido pela função get_resource_path
    icon_path = get_resource_path('earth.ico')
    root.iconbitmap(icon_path)

    # Desativa o redimensionamento
    root.resizable(False, False)

    frame = tk.Frame(root)
    frame.grid(pady=20, padx=20, row=0, column=0, sticky='nsew')

    tk.Label(frame, text="Nome da coluna para os marcadores:").grid(row=0, column=0, sticky='w')
    name_column_entry = tk.Entry(frame)
    name_column_entry.grid(row=1, column=0, pady=5, sticky='ew')
    

    tk.Label(frame, text="Pasta dos arquivos:").grid(row=2, column=0, sticky='w')
    folder_path_entry = tk.Entry(frame, width=50)
    folder_path_entry.grid(row=3, column=0, pady=5, sticky='ew')
    select_folder_button = tk.Button(frame, text="Selecionar Pasta", command=select_folder)
    select_folder_button.grid(row=3, column=1, padx=5)
    

    start_button = tk.Button(frame, text="Iniciar", command=start_processing)
    start_button.grid(row=4, column=0, padx=10, pady=10, sticky='ew')

    combine_button = tk.Button(frame, text="Combinar KMLs", command=combine_kmls_action_wrapper)
    combine_button.grid(row=4, column=1, padx=10, pady=10, sticky='ew')

    combine_var = tk.BooleanVar()
    tk.Checkbutton(frame, text="Gerar KML único", variable=combine_var).grid(row=5, column=0, pady=10, sticky='w')

    kml_convert_var = tk.BooleanVar()
    tk.Checkbutton(frame, text="Converter KML para Excel", variable=kml_convert_var).grid(row=5, column=1, pady=10, sticky='w')

    # Botão para ativar o comando de voz
    voice_button = tk.Button(frame, text="Comando de Voz", command=voice_command_wrapper)
    voice_button.grid(row=1, column=1, padx=10, pady=10, sticky='ew')

    root.mainloop()


if __name__ == "__main__":
    run_gui()
