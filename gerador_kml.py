import os
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd
import sys

def sanitize_kml_content(content):
    content = content.replace('&', '&amp;')
    content = content.replace('<', '&lt;')
    content = content.replace('>', '&gt;')
    content = content.replace('"', '&quot;')
    content = content.replace("'", '&#39;')
    return content

def create_kml(coordinates, names, sheet_name="Combined Data", filename="output.kml"):
    kml_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
<Document>
<name>{sanitize_kml_content(sheet_name)}</name>
<description>Generated KML from {sanitize_kml_content(sheet_name)}</description>
"""
    for coord, name in zip(coordinates, names):
        kml_content += f"""
<Placemark>
    <name>{sanitize_kml_content(name)}</name>
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

def read_excel(file_path, name_column='Etiqueta'):
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
    else:
        raise ValueError("As colunas 'Longitude' e 'Latitude' não foram encontradas na planilha.")

    return coordinates.tolist(), names

def process_all_excels_in_folder(folder_path, single_kml, name_column):
    if single_kml:
        all_coordinates = []
        all_names = []
        
        for filename in os.listdir(folder_path):
            if filename.endswith(".xlsx") or filename.endswith(".xls"):
                file_path = os.path.join(folder_path, filename)
                try:
                    coordinates, names = read_excel(file_path, name_column=name_column)
                    all_coordinates.extend(coordinates)
                    all_names.extend(names)
                except Exception as e:
                    print(f"Erro ao processar {filename}: {e}")
        
        if all_coordinates:
            output_kml = os.path.join(folder_path, "all_data.kml")
            create_kml(all_coordinates, all_names, sheet_name="All Data", filename=output_kml)
            print(f"KML gerado com sucesso para todas as planilhas.")
        else:
            print("Nenhuma coordenada válida encontrada em todas as planilhas.")
    else:
        for filename in os.listdir(folder_path):
            if filename.endswith(".xlsx") or filename.endswith(".xls"):
                file_path = os.path.join(folder_path, filename)
                try:
                    coordinates, names = read_excel(file_path, name_column=name_column)
                    sheet_name = os.path.splitext(filename)[0]
                    output_kml = filename.replace(".xlsx", ".kml").replace(".xls", ".kml")
                    create_kml(coordinates, names, sheet_name=sheet_name, filename=os.path.join(folder_path, output_kml))
                    print(f"KML gerado com sucesso para: {filename}")
                except Exception as e:
                    print(f"Erro ao processar {filename}: {e}")

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
<name>Combined KML</name>
<description>All KML files combined with different marker colors</description>
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

def parse_kml(kml_file):
    coordinates = []
    names = []

    tree = ET.parse(kml_file)
    root = tree.getroot()
    
    # Define o namespace do KML
    namespace = {'kml': 'http://www.opengis.net/kml/2.2'}

    # Itera sobre os elementos Placemark
    for placemark in root.findall('.//kml:Placemark', namespace):
        name = placemark.find('kml:name', namespace)
        coord = placemark.find('.//kml:coordinates', namespace)
        
        if name is not None and coord is not None:
            names.append(name.text.strip())
            coordinates.append(coord.text.strip().split(',')[:2])  # Pega apenas Longitude e Latitude

    return names, coordinates

def kml_to_excel(kml_file, output_excel, name_column):
    names, coordinates = parse_kml(kml_file)
    
    # Cria um DataFrame
    df = pd.DataFrame(coordinates, columns=['Longitude', 'Latitude'])
    df[name_column] = names
    
    # Salva o DataFrame em um arquivo Excel
    df.to_excel(output_excel, index=False)

def run_program(folder_path, combine=False, name_column='Etiqueta', convert_kml=False):
    if convert_kml:
        kml_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith('.kml')]
        for kml_file in kml_files:
            output_excel = os.path.splitext(kml_file)[0] + '.xlsx'
            kml_to_excel(kml_file, output_excel, name_column)
            print(f"Planilha Excel gerada com sucesso para: {kml_file}")
    else:
        process_all_excels_in_folder(folder_path, single_kml=combine, name_column=name_column)
        print("Processamento concluído. Todos os arquivos KML foram gerados na pasta selecionada.")

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

    # Definindo o ícone da janela
    root.iconbitmap(r'C:\Users\JoaoVictor\Documents\JavaScript Projects\gerador_kml\earth.ico')

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

    root.mainloop()


if __name__ == "__main__":
    run_gui()
