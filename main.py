import tkinter as tk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import os
import pandas as pd
import random
from collections import Counter
from typing import List, Optional, Tuple
from decimal import Decimal
import openpyxl.styles
import openai
import openpyxl.utils
import subprocess
import os 
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENTATION
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, Inches, Mm, RGBColor

from datetime import datetime
from api_key import openai_api_key


# Configurar la API key de OpenAI
openai.api_key = openai_api_key


# Verificar si la API key está configurada
if not openai.api_key:
    messagebox.showerror("Error", "No se encontró la API key de OpenAI. Por favor, verifique su configuración.")
    exit()

class Global:
    MAXL = 200        
    LINELEN = 800    
    MAXS = 2400      
    BASE_UPP = 65    
    BASE_LOW = 97    
    OFFSET = 32      
    NUL = -1         
    LIM = -2         
    DUB = -3         
    TAB = 4          
    PLEN = 202       
    ELEN = 208       
    NSEG = 26        
    
    SAL_COEF = [5, 40, 50, 10, 10, 10, 10, 5, 1, 5, 5, 5, 10]
    FACTOR = 100
    
    @staticmethod
    def conv(f: float) -> int:
        return int(f * Global.FACTOR)
    
    @staticmethod
    def de_conv(k: int) -> float:
        return float(k) / Global.FACTOR
    
    @staticmethod
    def max_score() -> int:
        return Global.conv(35)
    
    @staticmethod
    def skip_cost() -> int:
        return Global.conv(10)
    
    @staticmethod
    def vwl_decr() -> int:
        return Global.conv(10)
    
    @staticmethod
    def score_margin() -> float:
        return float(1)

class Stack:
    def __init__(self):
        self.stack = [[0] * Global.MAXL for _ in range(2)]
        self.top = 0
    
    def push(self, i1: int, i2: int = Global.NUL):
        if 0 <= self.top < Global.MAXL:
            self.stack[0][self.top] = i1
            self.stack[1][self.top] = i2
            self.top += 1
    
    def pop(self, k: int = 1):
        self.top -= k
    
    def clear(self):
        self.top = 0

class Word:
    def __init__(self, text: str):
        self.text = text
        self.length = len(text)

class Wagner:
    def __init__(self):
        self.s = [[0] * (Global.MAXL + 1) for _ in range(Global.MAXL + 1)]
        self.cost = Stack()
        self.trace = Stack()
        self.out = Stack()
        self.fall_thru = False
        self.dp_score = 0
        self.acpt_score = 0

    def similarity(self, w_a: Word, w_b: Word) -> int:
        # Implementar la lógica de similitud aquí
        pass

    def align(self, w_a: Word, w_b: Word):
        len_a = w_a.length
        len_b = w_b.length
        
        self.cost.clear()
        self.trace.clear()
        self.out.clear()
        self.fall_thru = False
        
        sg_max = self.similarity(w_a, w_b)
        self.dp_score = self.s[len_a][len_b]
        self.acpt_score = self.dp_score * Global.score_margin()
        
        i = len_a
        j = len_b
        
        if self.s[i][j] >= self.acpt_score:
            for j1 in range(len_b, j, -1):
                self.out.push(Global.NUL, j1)
            for i1 in range(len_a, i, -1):
                self.out.push(i1, Global.NUL)
                
            self.out.push(Global.LIM, Global.LIM)
            self.alignment(w_a, w_b, i, j, 0)

class Aline:
    def __init__(self, search_word: str = None):
        self.w_search_word = None
        self.w_compare_word = None
        self.wagner = Wagner()
        self.d_max_value: Decimal = Decimal('0')
        self.d_max_factor: Decimal = Decimal('0')
        self.results = {}  # Almacena resultados únicos

        if search_word is not None:
            self.search_word = search_word

    @property
    def search_word(self) -> str:
        return self.w_search_word.text

    @search_word.setter
    def search_word(self, value: str):
        self.w_search_word = Word(value)
        self.set_max_value()

    def verify_search_word(self):
        if self.w_search_word is None:
            raise Exception("A search word must be set before calling this function.")

    def set_max_value(self):
        self.max_value = self.get_phonetic_match(self.w_search_word.text)

    @property
    def max_factor(self) -> float:
        return self.d_max_factor

    @max_factor.setter
    def max_factor(self, value: float):
        self.d_max_factor = value

    @property
    def max_value(self) -> float:
        return self.d_max_value

    @max_value.setter
    def max_value(self, value: float):
        self.d_max_value = value
        if self.d_max_value != 0:
            self.max_factor = (self.d_max_value / 
                              (len(self.w_search_word.text) * self.d_max_value))
        else:
            self.max_factor = 0

    def get_percentage(self, result: int) -> float:
        result_decimal = Decimal(str(result))
        len_search = Decimal(str(len(self.w_search_word.text)))
        len_compare = Decimal(str(len(self.w_compare_word.text)))
        
        max_len = max(len_search, len_compare)
        denominator = max_len * self.d_max_value
        
        if denominator == 0:
            return 0.0
        
        current_row_value = (result_decimal / denominator) / Decimal(str(self.max_factor))
        percentage = current_row_value * Decimal('100')
        
        final_value = int(percentage) + 1
        
        return float(final_value)

    def get_phonetic_match(self, compare_word: str) -> int:
        self.w_compare_word = Word(compare_word)
        return self._get_phonetic_match()

    def _get_phonetic_match(self) -> int:
        self.wagner.align(self.w_search_word, self.w_compare_word)
        return self.wagner.dp_score

    def add_result(self, compare_word: str):
        match_score = self.get_phonetic_match(compare_word)
        percentage = self.get_percentage(match_score)
        self.results[compare_word] = percentage

    def get_sorted_results(self, limit: int) -> List[Tuple[str, float]]:
        # Ordenar resultados en orden descendente y limitar la cantidad
        sorted_results = sorted(self.results.items(), key=lambda x: x[1], reverse=True)
        return sorted_results[:limit]

def comparar_nombres(nombre1: str, nombre2: str, limit: int) -> List[Tuple[str, float]]:
    aline = Aline(nombre1.lower())
    aline.add_result(nombre2.lower())
    return aline.get_sorted_results(limit)

def main():
    print("\n=== Comparador de Similitud Fonética ===\n")
    
    while True:
        nombre1 = input("Ingrese el primer nombre (o 'salir' para terminar): ")
        if nombre1.lower() == 'salir':
            break
            
        nombre2 = input("Ingrese el segundo nombre: ")
        limit = int(input("Ingrese el número de resultados a mostrar: "))
        
        try:
            resultados = comparar_nombres(nombre1, nombre2, limit)
            print("\nResultados de similitud fonética:")
            for nombre, similitud in resultados:
                print(f"{nombre}: {similitud}%")
            print("\n")
            
        except Exception as e:
            print(f"\nError: {str(e)}\n")
        
        print("-" * 40 + "\n")

class OrthoGraphic:
    def __init__(self):
        self.h = {
            "oc": "or",
            "re": "le",
            "#c": "#l",
            "ra": "la",
            "re": "ro",
            "za": "ra",
            "ar": "os",
            "os": "or",
            "ri": "vi",
            "#s": "#p",
            "#a": "si",
            "in": "iv",
            "ra": "ir",
            "#l": "#t",
            "or": "ir",
            "at": "al",
            "iv": "re",
            "ne": "se",
            "on": "er",
            "xa": "na",
            "an": "av",
            "re": "se",
            "iv": "in",
            "on": "en",
            "xa": "na",
            "an": "av",
            "re": "se",
            "iv": "in",
            "on": "en",     
            "xa": "na",
        }
        self.c_score = 0.75
        self.p_score = 0.0
        self.debug = True
        self.init_bi_list()

    def init_bi_list(self):
        # Crear una lista de las claves actuales de self.h
        keys = list(self.h.keys())
        for bi in keys:  # Iterar sobre la lista de claves
            if bi[:2] + " " + bi[3:5] not in self.h:
                self.h[bi[:2] + " " + bi[3:5]] = 1
            if bi[3:5] + " " + bi[:2] not in self.h:
                self.h[bi[3:5] + " " + bi[:2]] = 1

    def score(self, word1, word2):
        len1 = len(word1)
        len2 = len(word2)
        minL = min(len1, len2)
        avgL = (len1 + len2) / 2.0
        Bscore = min(self.init_match(word1, word2), 3)
        Escore = min(self.init_match(self.reverse(word1), self.reverse(word2)), 1)
        Ascore = self.adaptive(word1, word2)

        wAscore = (Bscore + Escore + Ascore) / (avgL + 3)
        Uscore = (self.match_uni(word1, word2) + self.match_uni(word2, word1)) / 2
        Lscore = self.lcsq(word1, word2)
        wLUscore = (Lscore + Uscore) / (len1 + len2)
        Sscore = self.lcsb(word1, word2)
        wSscore = (Sscore + Sscore) / (len1 + len2 + 2)

        wFscore = max(wAscore, max(wLUscore, wSscore))

        if (wFscore == 1) and (word1 != word2):
            wAscore = (Bscore + Escore + Ascore) / (minL + 4)
            wFscore = max(wAscore, max(wLUscore, wSscore))

        return wFscore * 100

    def init_match(self, word1, word2):
        lenE = len(word1)
        lenF = len(word2)
        i = 0

        while i < min(lenE, lenF):
            if word1[:i + 1] != word2[:i + 1]:
                break
            i += 1

        return i

    def lcsq(self, word1, word2):
        lenE = len(word1)
        lenF = len(word2)
        MaxLen = max(lenE, lenF)
        S = [lenF] * MaxLen

        for l in range(lenE):
            for j in range(lenF - 1, -1, -1):
                if word1[l] == word2[j]:
                    l = 0
                    while j > S[l]:
                        l += 1
                    S[l] = j

        for l in range(lenE):
            if S[l] == lenF:
                return l

        return lenE

    def match_uni(self, word1, word2):
        lenE = len(word1)
        lenF = len(word2)
        memE = {}
        matchF = 0

        for i in range(lenE):
            uni = word1[i]
            if uni in memE:
                memE[uni] += 1
            else:
                memE[uni] = 1

        for j in range(lenF):
            uni = word2[j]
            if uni in memE:
                matchF += 1

        return matchF

    def lcsb(self, word1, word2):
        lenE = len(word1)
        lenF = len(word2)
        S1 = [0] * (lenF + 2)
        z = 0
        prev = 1

        for j in range(lenF):
            S1[j] = 0

        for i in range(lenE):
            for j in range(lenF):
                curr = S1[j]
                if word1[i] == word2[j]:
                    S1[j] = prev + 1
                    if S1[j] > z:
                        z = S1[j]
                else:
                    S1[j] = 0
                prev = curr

            prev = 0

        return z

    @staticmethod
    def reverse(s):
        return s[::-1]

    def simil(self, biE, biF):
        if biE == biF:
            return 1
        if biE + " " + biF in self.h:
            return self.c_score

        a1 = biE[0]
        a2 = biE[1]
        b1 = biF[0]
        b2 = biF[1]
        if a1 == b1 or a2 == b2:
            return self.p_score
        else:
            return 0

    def adaptive(self, searchword, compareword):
        lenE = len(searchword)
        lenF = len(compareword)
        S = [0] * (lenF + 2)

        for j in range(lenF):
            S[j] = 0

        for i in range(1, lenE + 1):
            prev = 0
            biE = searchword[i - 1:i + 1] if i < lenE else searchword[i - 1] + "#"

            for j in range(1, lenF + 1):
                curr = S[j]
                biF = compareword[j - 1:j + 1] if j < lenF else compareword[j - 1] + "#"

                ident = self.simil(biE, biF)
                diag = prev + ident
                if diag > curr:
                    if diag > S[j - 1]:
                        S[j] = diag
                    else:
                        S[j] = S[j - 1]
                else:
                    if curr > S[j - 1]:
                        S[j] = curr
                    else:
                        S[j] = S[j - 1]
                prev = curr

        return S[lenF]

def buscar_archivo(nombre_archivo, directorio_inicio='.'):
    """Busca un archivo por nombre en el directorio actual y subdirectorios."""
    for root, dirs, files in os.walk(directorio_inicio):
        if nombre_archivo in files:
            return os.path.join(root, nombre_archivo)
    return None

class ComparisonForm:
    def __init__(self, root):
        self.root = root
        self.root.title("Comparador de Nombres")
        
        # Establecer tamaño mínimo de la ventana
        self.root.minsize(800, 600)
        
        # Configurar expansión del frame principal
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=1)
        
        # Frame principal con más padding
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.main_frame.grid_columnconfigure(1, weight=1)  # Hacer que la segunda columna se expanda
        

            
        # Inicializar el algoritmo de comparación
        try:
            self.ortho = OrthoGraphic()
        except Exception as e:
            messagebox.showerror("Error", f"Error al inicializar el algoritmo: {str(e)}")
            self.root.destroy()
            return
        
        # Buscar archivos en el directorio actual y subdirectorios
        self.archivo_ema = buscar_archivo('EMA.xlsx')
        self.archivo_inn_stem = buscar_archivo('INN stem.xlsx')
        self.archivo_inn = buscar_archivo('INN.xlsx')
        self.archivo_ema_check = buscar_archivo('EMA_check.py', 'factcheck')
        self.archivo_inn_check = buscar_archivo('inn_check.py', 'factcheck')
        self.archivo_factcheck = buscar_archivo('factcheck.py', 'factcheck')
        
        # Verificar si se encontraron los archivos necesarios
        if not self.archivo_ema:
            messagebox.showwarning("Advertencia", "No se encontró el archivo EMA.xlsx")
        if not self.archivo_inn:
            messagebox.showwarning("Advertencia", "No se encontró el archivo INN.xlsx")
        if not self.archivo_inn_stem:
            messagebox.showwarning("Advertencia", "No se encontró el archivo INN stem.xlsx")

        
        # Agregar atributos para las rutas de archivos
        self.ruta_ema = tk.StringVar(value="EMA.xlsx")  
        self.ruta_inn = tk.StringVar(value="INN.xlsx")  
        self.ruta_innstem = tk.StringVar(value="INN stem.xlsx")
        self.results_window = None
        
        # Frame para archivos con más espacio
        self.file_frame = ttk.LabelFrame(self.main_frame, text="Archivos de datos", padding="10")
        self.file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        self.file_frame.grid_columnconfigure(1, weight=1)  
        
        # Botones y etiquetas para archivos
        ttk.Label(self.file_frame, text="EMA:", width=20).grid(row=0, column=0, sticky=tk.W, padx=5)
        ttk.Label(self.file_frame, textvariable=self.ruta_ema, width=40).grid(row=0, column=1, sticky=tk.W)
        ttk.Button(self.file_frame, text="Subir archivo EMA", width=20, command=self.subir_ema).grid(row=0, column=2, padx=10)
        
        ttk.Label(self.file_frame, text="INN:", width=20).grid(row=1, column=0, sticky=tk.W, padx=5)
        ttk.Label(self.file_frame, textvariable=self.ruta_inn, width=40).grid(row=1, column=1, sticky=tk.W)
        ttk.Button(self.file_frame, text="Subir archivo INN", width=20, command=self.subir_inn).grid(row=1, column=2, padx=10)
        
        ttk.Label(self.file_frame, text="INN stem:", width=20).grid(row=2, column=0, sticky=tk.W, padx=5)
        ttk.Label(self.file_frame, textvariable=self.ruta_innstem, width=40).grid(row=2, column=1, sticky=tk.W)
        ttk.Button(self.file_frame, text="Subir archivo INN stem", width=21, command=self.subir_innstem).grid(row=2, column=2, padx=10)
        
        # Campos de entrada
        ttk.Label(self.main_frame, text="Primer nombre:").grid(row=1, column=0, sticky=tk.W)
        self.nombre1 = ttk.Entry(self.main_frame, width=30)
        self.nombre1.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(self.main_frame, text="Segundo nombre:").grid(row=2, column=0, sticky=tk.W)
        self.nombre2 = ttk.Entry(self.main_frame, width=30)
        self.nombre2.grid(row=2, column=1, padx=5, pady=5)
        
        # Frame para opciones de búsqueda
        self.options_frame = ttk.LabelFrame(self.main_frame, text="Opciones de búsqueda", padding="5")
        self.options_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Spinbox para número de resultados
        ttk.Label(self.options_frame, text="Número de resultados:").grid(row=0, column=0, padx=5)
        self.num_resultados = ttk.Spinbox(
            self.options_frame,
            from_=1,
            to=100,
            width=5,
            validate='key',
            validatecommand=(self.root.register(self.validar_numero), '%P')
        )
        self.num_resultados.set("5")  # Valor por defecto
        self.num_resultados.grid(row=0, column=1, padx=5)
        
        # Checkbuttons en el frame de opciones con valores por defecto True
        self.inn_var = tk.BooleanVar(value=True)  # Preseleccionado
        self.ema_var = tk.BooleanVar(value=True)  # Preseleccionado
        self.innstem_var = tk.BooleanVar(value=True)  # Preseleccionado
        
        ttk.Checkbutton(self.options_frame, text="INN", variable=self.inn_var).grid(row=1, column=0, sticky=tk.W)
        ttk.Checkbutton(self.options_frame, text="EMA", variable=self.ema_var).grid(row=1, column=1, sticky=tk.W)
        ttk.Checkbutton(self.options_frame, text="INN stem", variable=self.innstem_var).grid(row=1, column=2, sticky=tk.W)
        
        # Frame para opciones de generación
        self.gen_frame = ttk.LabelFrame(self.main_frame, text="Opciones de generación", padding="5")
        self.gen_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Agregar campo para letras iniciales
        ttk.Label(self.gen_frame, text="Iniciar con:").grid(row=0, column=0, padx=5, pady=5)
        self.inicio_nombre = ttk.Entry(self.gen_frame, width=10)
        self.inicio_nombre.grid(row=0, column=1, padx=5, pady=5)
        self.inicio_nombre.insert(0, "aba")  # Valor predeterminado
        
        # Spinbox para número de nombres a generar
        ttk.Label(self.gen_frame, text="Cantidad de nombres:").grid(row=0, column=2, padx=5)
        self.num_nombres_gen = ttk.Spinbox(
            self.gen_frame,
            from_=1,
            to=99999,
            width=7,
            validate='key',
            validatecommand=(self.root.register(self.validar_numero), '%P')
        )
        self.num_nombres_gen.set("50")  # Valor por defecto
        self.num_nombres_gen.grid(row=0, column=3, padx=5)
        
        
        # Botones
        ttk.Button(self.main_frame, text="Comparar", command=self.comparar).grid(row=5, column=0, pady=10)
        ttk.Button(self.main_frame, text="Generar", command=self.generar).grid(row=5, column=1, pady=10)
        
        # Frame para botones inferiores
        self.bottom_frame = ttk.Frame(self.main_frame)
        self.bottom_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.E), pady=5)
        
        # Botón Generar Excel (inicialmente deshabilitado)
        self.btn_excel = ttk.Button(self.bottom_frame, text="Generar Excel", command=self.generar_excel, state='disabled')
        self.btn_excel.pack(side=tk.RIGHT, padx=5)
        
        # Botón Informe
        self.btn_informe = ttk.Button(self.bottom_frame, text="Informe", command=self.generar_informe)
        self.btn_informe.pack(side=tk.RIGHT, padx=5)
        
        self.btn_siguiente = ttk.Button(self.bottom_frame, text="Siguiente", command=self.siguiente_funcion)
        self.btn_siguiente.pack(side=tk.RIGHT, padx=5)
        
        # Aumentar espaciado entre secciones
        for widget in [self.file_frame, self.options_frame, self.gen_frame]:
            widget.grid(pady=15)
        
        # Ajustar el ancho de los campos de entrada
        self.nombre1.configure(width=50)
        self.nombre2.configure(width=50)
        
        # Aumentar el espaciado entre elementos en los frames de opciones
        for frame in [self.options_frame, self.gen_frame]:
            for child in frame.winfo_children():
                child.grid(padx=10, pady=5)
        
        # Ajustar el espaciado de los botones
        self.bottom_frame.grid(pady=20)


    
    def subir_ema(self):
        filename = filedialog.askopenfilename(
            title="Seleccionar archivo EMA",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.ruta_ema.set(os.path.basename(filename))
            self.archivo_ema = filename

    def subir_inn(self):
        filename = filedialog.askopenfilename(
            title="Seleccionar archivo INN",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.ruta_inn.set(os.path.basename(filename))
            self.archivo_inn = filename

    def subir_innstem(self):
        filename = filedialog.askopenfilename(
            title="Seleccionar archivo INN stem",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.ruta_innstem.set(os.path.basename(filename))
            self.archivo_innstem = filename

    def comparar(self):
        nombre1 = self.nombre1.get().strip().lower()
        
        if not nombre1:
            messagebox.showwarning("Advertencia", "Por favor, ingrese el primer nombre")
            return
        
        try:
            # Obtener el número de resultados deseados
            try:
                num_resultados = int(self.num_resultados.get())
            except ValueError:
                num_resultados = 5  # valor por defecto si hay error
            
            # Si EMA está seleccionado, buscar en el archivo EMA
            if self.ema_var.get():
                try:
                    if not os.path.exists(self.archivo_ema):
                        messagebox.showerror("Error", f"No se encuentra el archivo: {self.archivo_ema}")
                        return
                    
                    df_ema = pd.read_excel(self.archivo_ema)
                    nombres_ema = df_ema['EMA'].str.lower().tolist()
                    
                    # Lista para almacenar todos los nombres similares
                    similitudes = []
                    for nombre_ema in nombres_ema:
                        similitud_actual = self.ortho.score(nombre1, nombre_ema)
                        similitudes.append((nombre_ema, similitud_actual))
                    
                    # Ordenar por similitud y obtener los N más similares
                    similitudes.sort(key=lambda x: x[1], reverse=True)
                    top_n = similitudes[:num_resultados]
                    
                    # Cargar nombres INN
                    INN_list = self.cargar_nombres_inn()
                    
                    # Crear resultados para mostrar
                    resultados = []
                    nombres_usados = set()  # Para evitar duplicados
                    
                    for nombre_ema, similitud_ema in top_n:
                        # Calcular similitudes para todos los nombres INN
                        similitudes_inn = []
                        for nombre_inn in INN_list:
                            if isinstance(nombre_inn, str) and nombre_inn not in nombres_usados:
                                caracteres_presentes = sum(1 for c in nombre_inn.lower() if c in nombre1.lower())
                                porcentaje = (caracteres_presentes / len(nombre_inn)) * 100
                                similitudes_inn.append((nombre_inn, caracteres_presentes, len(nombre_inn), porcentaje))
                        
                        # Ordenar por porcentaje y tomar los N mejores diferentes
                        similitudes_inn.sort(key=lambda x: x[3], reverse=True)
                        
                        # Agregar resultados hasta alcanzar num_resultados
                        for inn_data in similitudes_inn:
                            if len(resultados) >= num_resultados:
                                break
                                
                            nombre_inn, chars_presentes, total_chars, porcentaje = inn_data
                            if nombre_inn not in nombres_usados:
                                resultados.append((
                                    nombre1,
                                    nombre_inn,
                                    f"{chars_presentes}/{total_chars}",
                                    f"{porcentaje:.2f}%",
                                    "-",
                                    nombre_ema,
                                    similitud_ema
                                ))
                                nombres_usados.add(nombre_inn)
                
                except Exception as e:
                    messagebox.showerror("Error", f"Error al procesar archivo EMA: {str(e)}")
                    return
            else:
                # Comparación directa entre los dos nombres
                nombre2 = self.nombre2.get().strip().lower()
                if not nombre2:
                    messagebox.showwarning("Advertencia", "Por favor, ingrese el segundo nombre")
                    return
                
                # Calcular similitud EMA
                similitud_ema = self.ortho.score(nombre1, nombre2)
                
                # Calcular porcentaje INN
                caracteres_presentes = sum(1 for c in nombre2.lower() if c in nombre1.lower())
                porcentaje_inn = (caracteres_presentes / len(nombre2)) * 100
                
                resultados = [(
                    nombre1,
                    nombre2,
                    f"{caracteres_presentes}/{len(nombre2)}",
                    f"{porcentaje_inn:.2f}%",
                    "-",
                    nombre2,
                    similitud_ema
                )]
            
            self.mostrar_resultados(resultados)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al comparar nombres: {str(e)}")

    def mostrar_resultados(self, resultados):
        if self.results_window is not None and self.results_window.winfo_exists():
            self.results_window.destroy()
        
        self.results_window = tk.Toplevel(self.root)
        self.results_window.title("Resultados")
        
        # Establecer un tamaño inicial del 50% de la pantalla
        screen_width = self.results_window.winfo_screenwidth()
        screen_height = self.results_window.winfo_screenheight()
        window_width = int(screen_width * 0.8)  # 80% del ancho de la pantalla
        window_height = int(screen_height * 0.8)
        self.results_window.geometry(f"{window_width}x{window_height}")
        
        # Frame principal para dividir la ventana
        main_frame = ttk.Frame(self.results_window)
        main_frame.pack(fill='both', expand=True)
        
        # Frame izquierdo para información (1/3 de la pantalla)
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side='left', fill='both', padx=5, pady=5, expand=False)
        
        # Frame derecho para la tabla (2/3 de la pantalla)
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        
        # Crear un LabelFrame con título para la sección de información
        info_frame = ttk.LabelFrame(left_frame, text="Información", padding="10")
        info_frame.pack(fill='both', expand=True)
        
        # Texto informativo con mejor diseño
        info_text = tk.Text(
            info_frame,
            wrap=tk.NONE,  # Cambiado a NONE para evitar el wrap automático
            padx=15,
            pady=15,
            width=75,  # Ancho fijo más grande
            height=30,
            font=('Courier New', 10),  # Font monoespaciado
            background='#F5F5F5',
            relief='flat',
            borderwidth=0
        )
        info_text.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Mensaje inicial informativo con formato
        mensaje_inicial = """
╔══════════════════════════════════════════╗
║           SECCIÓN DE INFORMACIÓN         ║
╚══════════════════════════════════════════╝

Esta sección muestra información detallada sobre cada columna de la tabla.

• Haga clic en cualquier encabezado de columna para ver su descripción
• La información mostrada le ayudará a entender los datos presentados
• Cada columna tiene un propósito específico en el análisis

▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ ANÁLISIS INN ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓

► PASO 1: Cálculo del Porcentaje INN

┌──────────────────────────────────────────────────────────────────────┐
│                                                                      │
│        caracteres del nombre INN en nombre ingresado                 │
│ %INN = ────────────────────────────────────────────────    x 100     │
│           total de caracteres del nombre INN                         │
│                                                                      │
└──────────────────────────────────────────────────────────────────────┘

Interpretación del %INN:
• Si la diferencia entre caracteres totales y compartidos es
 1 o menos → NO APTO
• Si tienen cadena común y está en la misma posición → NO APTO
• Si tienen cadena común pero en diferente posición → APTO
• Si no tienen cadena común → APTO

► PASO 2: Análisis de Stem

Definición: Un stem es la raíz de una palabra
Ejemplo: 
• paracetamol → stem: paracetam
• acetaminofen → stem: acetaminof

Evaluación de stem:
✓ No comparten stem → APTO
✗ Comparten stem → NO APTO

► PASO 3: Resultado Final INN

• INN APTO + Stem APTO → RESULTADO APTO
• INN NO APTO o Stem NO APTO → RESULTADO NO APTO

▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ ANÁLISIS EMA ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓

Evaluación según POCA:

┌─────────────────────────────────────────────────────────┐
│ Riesgo Alto: 70-100%    → NO APTO                       │
│ Riesgo Moderado: 55-69% → REQUIERE EVALUACIÓN ADICONAL  │
│ Riesgo Bajo: <55%       → APTO                          │
└─────────────────────────────────────────────────────────┘

Para Riesgo Moderado:
• Se aplica sistema de evaluación EMA
• Se consideran atenuantes y agravantes
• Objetivo: Evitar riesgo alto
• Posibilidad de ser APTO con evaluación favorable

───────────────────────────────────────────────────────────────────────────

» Haga clic en un encabezado para ver información específica «
"""
        info_text.insert('1.0', mensaje_inicial)
        
        # Hacer el texto de solo lectura
        info_text.configure(state='disabled')
        
        # Modificar la función lambda que actualiza el texto
        def mostrar_info(columna):
            info_text.configure(state='normal')
            info_text.delete('1.0', tk.END)
            info = f"""
        Información: {column_headings[columna]}
        ────────────────────────────────
        
        {column_info[columna]}
        
        ───────────────────────────────
        """
            info_text.insert('1.0', info)
            info_text.configure(state='disabled')
        
        # Definir las columnas y configuraciones primero
        columns = (
            'nombre_ingresado', 'nombre_inn', 'porcentaje_inn', 'caracteres_inn',
            'cadena_comun', 'misma_cadena', 'tam_cadena_comun', 'peso_cadena', 'misma_posicion', 'resultado_inn',
            'inn_stem', 'resultado_inn_stem', 'resultado_final',
            'nombre_ema', 'similitud_ema'
        )
        
        # Configurar anchos de columna iniciales
        column_widths = {
            'nombre_ingresado': 120,  # Nombres
            'nombre_inn': 120,
            'porcentaje_inn': 80,     # Porcentajes
            'caracteres_inn': 100,    # Fracciones
            'cadena_comun': 100,      # Cadenas
            'misma_cadena': 100,      # SÍ/NO
            'tam_cadena_comun': 100,  # Fracciones
            'peso_cadena': 100,       # Nueva columna para el porcentaje
            'misma_posicion': 100,    # SÍ/NO
            'resultado_inn': 100,     # APTO/NO APTO
            'inn_stem': 120,          # Nueva columna
            'resultado_inn_stem': 120, # Nueva columna
            'resultado_final': 180,    # Aumentado de 120 a 180 para texto completo
            'nombre_ema': 120,        # Nombres
            'similitud_ema': 100      # Porcentajes
        }
        
        # Textos de los encabezados
        column_headings = {
            'nombre_ingresado': 'Nombre Ingresado',
            'nombre_inn': 'Nombre INN',
            'porcentaje_inn': '% INN',
            'caracteres_inn': 'Caracteres INN',
            'cadena_comun': 'Cadena Común',
            'misma_cadena': '¿Misma Cadena?',
            'tam_cadena_comun': 'Tamaño Cadena',
            'peso_cadena': 'Peso de Cadena',  # Nueva columna
            'misma_posicion': '¿Misma Posición?',
            'resultado_inn': 'Resultado INN',
            'inn_stem': 'INN stem',
            'resultado_inn_stem': 'Resultado INN stem',
            'resultado_final': 'Resultado INN e INN stem',
            'nombre_ema': 'Nombre EMA',
            'similitud_ema': 'Similitud EMA'
        }
        
        # Diccionario con la información para cada columna
        column_info = {
            'nombre_ingresado': 'Nombre que se está evaluando, ingresado por el usuario.',
            'nombre_inn': 'Nombre INN encontrado en la base de datos, ordenado por similitud.',
            'porcentaje_inn': 'Porcentaje calculado a partir de la fracción de caracteres INN.',
            'caracteres_inn': 'Fracción que representa cuántos caracteres del nombre INN están presentes en el nombre ingresado.',
            'cadena_comun': 'Subcadena más larga que comparten ambos nombres.',
            'misma_cadena': 'Indica si existe una cadena común entre los nombres.',
            'tam_cadena_comun': 'Tamaño de la cadena común respecto al nombre INN.',
            'peso_cadena': 'Porcentaje de la cadena común respecto al nombre INN.',
            'misma_posicion': 'Indica si la cadena común está en la misma posición en ambos nombres.',
            'resultado_inn': 'Resultado final del análisis: APTO o NO APTO.',
            'inn_stem': 'Stem del nombre INN encontrado.',
            'resultado_inn_stem': 'Resultado del análisis del stem: APTO o NO APTO.',
            'resultado_final': 'Resultado combinado del análisis INN y stem.',
            'nombre_ema': 'Nombre EMA encontrado en la base de datos.',
            'similitud_ema': 'Porcentaje de similitud orthographic entre los nombres.'
        }
        
        # Control frame para el número de resultados visibles (ahora en el frame derecho)
        control_frame = ttk.Frame(right_frame, padding="5")
        control_frame.pack(fill='x')
        
        # Spinbox para ajustar resultados visibles
        ttk.Label(control_frame, text="Mostrar top:").pack(side=tk.LEFT, padx=5)
        num_visible = ttk.Spinbox(
            control_frame,
            from_=1,
            to=999999,  # Valor muy alto para que sea prácticamente ilimitado
            width=5,
            validate='key',
            validatecommand=(self.root.register(self.validar_numero), '%P')
        )
        num_visible.set(min(5, len(resultados)))  # Valor por defecto
        num_visible.pack(side=tk.LEFT, padx=5)
        
        # Botón para actualizar vista
        ttk.Button(
            control_frame, 
            text="Actualizar", 
            command=lambda: self.actualizar_resultados_visibles(tree, resultados, int(num_visible.get()))
        ).pack(side=tk.LEFT, padx=5)
        
        # Checkboxes para filtros (nuevo)
        self.mostrar_solo_aptos = tk.BooleanVar(value=False)
        self.mostrar_solo_inn_aptos = tk.BooleanVar(value=False)  # Nuevo checkbox para INN

        def filtrar_aptos():
            # Limpiar la tabla
            for item in tree.get_children():
                tree.delete(item)
            
            # Si el checkbox está desmarcado, mostrar todos los resultados
            if not self.mostrar_solo_aptos.get() and not self.mostrar_solo_inn_aptos.get():
                for resultado in resultados:  # Use resultados instead of todos_resultados
                    tree.insert('', tk.END, values=resultado)
            elif self.mostrar_solo_inn_aptos.get():
                # Mostrar solo los resultados donde INN es APTO
                for resultado in resultados:  # Use resultados instead of todos_resultados
                    if len(resultado) >= 9:
                        resultado_inn = resultado[8]
                        if resultado_inn == "APTO":
                            tree.insert('', tk.END, values=resultado)
            else:
                # Mostrar solo los resultados donde INN e INN stem son APTOS
                for resultado in resultados:  # Use resultados instead of todos_resultados
                    if len(resultado) >= 12:
                        resultado_final = resultado[11]
                        similitud_ema = float(resultado[13].strip('%'))
                        if resultado_final == "APTO" and similitud_ema < 70:
                            tree.insert('', tk.END, values=resultado)

      
        
        # Crear Treeview en el frame derecho
        tree = ttk.Treeview(right_frame, columns=columns, show='headings')
        
        # Agregar scrollbar horizontal
        h_scrollbar = ttk.Scrollbar(right_frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(xscrollcommand=h_scrollbar.set)
        
        # Agregar scrollbar vertical
        v_scrollbar = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=v_scrollbar.set)
        
        # Función para manejar el scroll horizontal con la rueda del ratón
        def on_mousewheel(event):
            # Scroll horizontal directo con la rueda del ratón
            if event.delta > 0:
                tree.xview_scroll(-80, "units")  # Aumentado de -80 a 80
            else:
                tree.xview_scroll(80, "units")   # Aumentado de -80 a 80
        
        # Vincular el evento de la rueda del ratón
        tree.bind("<MouseWheel>", on_mousewheel)  # Windows
        
        # Posicionar elementos
        tree.pack(fill='both', expand=True)
        h_scrollbar.pack(fill='x')
        v_scrollbar.pack(side='right', fill='y')
        
        # Configurar encabezados y anchos de columna
        for col in columns:
            tree.heading(col, text=column_headings[col],
                        command=lambda c=col: mostrar_info(c))
            tree.column(col, width=column_widths[col], minwidth=column_widths[col], anchor='center')
        
        # Mostrar resultados iniciales
        self.actualizar_resultados_visibles(tree, resultados, int(num_visible.get()))
        
        # Mostrar información inicial
        info_text.insert('1.0', 'Haga clic en los encabezados de las columnas para ver información sobre cada campo.')
        
        # Crear clase para el tooltip
        class ToolTip(object):
            def __init__(self, widget, text=''):
                self.widget = widget
                self.text = text
                self.tooltip = None

            def show_tooltip(self, event=None):
                x, y, _, _ = self.widget.bbox("insert")
                x += self.widget.winfo_rootx() + 25
                y += self.widget.winfo_rooty() + 20

                # Crear ventana tooltip
                self.tooltip = tk.Toplevel(self.widget)
                self.tooltip.wm_overrideredirect(True)
                self.tooltip.wm_geometry(f"+{x}+{y}")

                label = ttk.Label(self.tooltip, text=self.text, justify='left',
                                background="#ffffe0", relief='solid', borderwidth=1,
                                wraplength=300, padding=(5, 5))
                label.pack()

            def hide_tooltip(self, event=None):
                if self.tooltip:
                    self.tooltip.destroy()
                    self.tooltip = None

        def crear_tooltip(widget, text):
            tooltip = ToolTip(widget, text)
            widget.bind('<Enter>', tooltip.show_tooltip)
            widget.bind('<Leave>', tooltip.hide_tooltip)

        # Función para obtener explicación del resultado
        def get_explicacion_resultado(valores):
            nombre_ingresado = valores[0]
            nombre_inn = valores[1]
            porcentaje = float(valores[2].rstrip('%'))
            cadena_comun = valores[4]
            misma_posicion = valores[7]
            resultado = valores[8]

            if resultado == "NO APTO":
                razones = []
                if cadena_comun and misma_posicion == "SÍ" and porcentaje > 50:
                    razones.append(f"Comparte la cadena '{cadena_comun}' en la misma posición y el %INN es > 50%")
                return "NO APTO porque:\n" + "\n".join(f"• {r}" for r in razones)
            else:
                razones = []
                if porcentaje < 50:
                    razones.append(f"El porcentaje INN ({porcentaje:.1f}%) es inferior al 50%")
                elif cadena_comun and misma_posicion == "NO":
                    razones.append("La cadena común está en diferente posición")
                return "APTO porque:\n" + "\n".join(f"• {r}" for r in razones)

        # Modificar la función on_double_click para incluir el tooltip
        def on_item_click(event):
            item = tree.identify_row(event.y)
            column = tree.identify_column(event.x)
            
            if column == "#9":  # Columna de resultado_inn
                valores = tree.item(item)['values']
                explicacion = get_explicacion_resultado(valores)
                
                # Crear tooltip en la posición del cursor
                x = event.x_root + 10
                y = event.y_root + 10
                
                tooltip = tk.Toplevel(tree)
                tooltip.wm_overrideredirect(True)
                tooltip.wm_geometry(f"+{x}+{y}")
                
                label = ttk.Label(tooltip, text=explicacion, justify='left',
                                background="#ffffe0", relief='solid', borderwidth=1,
                                wraplength=300, padding=(5, 5))
                label.pack()
                
                def on_leave(e):
                    tooltip.destroy()
                
                tooltip.bind('<Leave>', on_leave)
                tree.bind('<Motion>', lambda e: tooltip.destroy() if e.x < x-10 or e.x > x+200 or e.y < y-10 or e.y > y+100 else None)

        # Vincular el evento de movimiento del ratón
        tree.bind('<Motion>', on_item_click)

        # Función para mostrar explicación en el panel de texto
        def mostrar_explicacion_en_panel(event):
            item = tree.identify_row(event.y)
            column = tree.identify_column(event.x)
            
            if column == "#9":  # Columna de resultado_inn
                valores = tree.item(item)['values']
                nombre_ingresado = valores[0]
                nombre_inn = valores[1]
                porcentaje = float(valores[2].rstrip('%'))
                fraccion = valores[3]
                cadena_comun = valores[4]
                tiene_cadena = valores[5]
                tam_cadena = valores[6]
                misma_posicion = valores[7]
                resultado = valores[8]

                # Construir explicación detallada
                explicacion = f"""
╔══════════════════════════════════════════╗
║           ANÁLISIS DE RESULTADO          ║
╚══════════════════════════════════════════╝

Comparación entre:
• Nombre ingresado: {nombre_ingresado}
• Nombre INN: {nombre_inn}

Análisis de caracteres:
• Fracción: {fraccion}
• Porcentaje INN: {porcentaje:.1f}%

Análisis de cadena común:
• Cadena encontrada: {cadena_comun if cadena_comun != "-" else "Ninguna"}
• Tamaño de cadena: {tam_cadena}
• Misma posición: {misma_posicion}

RESULTADO: {resultado}
────────────────────────────────────────────
Explicación:
"""
                if resultado == "NO APTO":
                    razones = []
                    if cadena_comun != "-" and misma_posicion == "SÍ" and porcentaje > 50:
                        razones.append(f"• Comparte la cadena '{cadena_comun}' en la misma posición")
                    explicacion += "\nMotivos de NO APTO:\n" + "\n".join(razones)
                else:
                    razones = []
                    if porcentaje < 50:
                        razones.append(f"• El porcentaje INN ({porcentaje:.1f}%) es inferior al 50%")
                    elif cadena_comun != "-" and misma_posicion == "NO":
                        razones.append("• Aunque hay cadena común, está en diferente posición")
                    explicacion += "\nMotivos de APTO:\n" + "\n".join(razones)

                # Actualizar el texto en el panel izquierdo
                info_text.configure(state='normal')
                info_text.delete('1.0', tk.END)
                info_text.insert('1.0', explicacion)
                info_text.configure(state='disabled')

        # Vincular el clic izquierdo en la columna resultado_inn
        tree.bind('<ButtonRelease-1>', mostrar_explicacion_en_panel)

        # Función para mostrar explicación en el panel
        def mostrar_explicacion_resultado_inn(event):
            item = tree.identify_row(event.y)
            if item:
                valores = tree.item(item)['values']
                if valores:
                    nombre_gen = valores[0]
                    nombre_inn = valores[1]
                    caracteres = valores[3]
                    cadena_comun = valores[4]
                    misma_posicion = valores[7]
                    resultado_inn = valores[8]
                    
                    # Construir explicación
                    explicacion = f"""
╔══════════════════════════════════════════╗
║           ANÁLISIS DE RESULTADO INN      ║
╚══════════════════════════════════════════╝

Comparación entre:
• Nombre generado: {nombre_gen}
• Nombre INN: {nombre_inn}

"""
                    
                    # Analizar caracteres
                    caracteres_presentes, total_caracteres = map(int, caracteres.split('/'))
                    diferencia_caracteres = total_caracteres - caracteres_presentes
                    
                    if resultado_inn == "NO APTO":
                        if diferencia_caracteres <= 1:
                            explicacion += f"""
RESULTADO: NO APTO

Motivo: La diferencia entre caracteres totales ({total_caracteres}) 
y caracteres compartidos ({caracteres_presentes}) es {diferencia_caracteres}, 
que es menor o igual a 1."""
                        elif cadena_comun != "-" and misma_posicion == "SÍ":
                            explicacion += f"""
RESULTADO: NO APTO

Motivo: Comparten la cadena común '{cadena_comun}' 
y está en la misma posición en ambos nombres."""
                    else:
                        explicacion += f"""
RESULTADO: APTO

No se encontraron motivos de NO APTO:
• La diferencia de caracteres es {diferencia_caracteres} (mayor a 1)"""
                        if cadena_comun != "-":
                            explicacion += f"""
• Aunque comparten la cadena '{cadena_comun}', 
  está en diferente posición"""
                        else:
                            explicacion += """
• No comparten cadena común"""
                    
                    # Mostrar la explicación en el panel de información
                    info_text.configure(state='normal')
                    info_text.delete('1.0', tk.END)
                    info_text.insert('1.0', explicacion)
                    info_text.configure(state='disabled')
        
        # Vincular el evento de clic en una fila
        tree.bind('<ButtonRelease-1>', mostrar_explicacion_resultado_inn)

    def actualizar_resultados_visibles(self, tree, resultados, num_visible):
        # Limpiar resultados actuales
        for item in tree.get_children():
            tree.delete(item)
        
        # Cargar INN stems
        stems_dict = self.cargar_inn_stems()
        
        # Obtener el nombre ingresado del primer resultado
        nombre_ingresado = resultados[0][0]
        
        # Cargar nombres EMA y calcular similitudes
        try:
            df_ema = pd.read_excel(self.archivo_ema)
            nombres_ema = df_ema['EMA'].str.lower().tolist()
            similitudes_ema = self.calcular_similitudes_ema(nombre_ingresado, nombres_ema)
        except Exception as e:
            similitudes_ema = [("-", 0)]
        
        # Crear diccionario para rastrear mejores resultados por nombre INN
        mejores_resultados = {}
        
        # Procesar todos los resultados para encontrar la mejor similitud para cada nombre INN
        for resultado in resultados:
            nombre_ingresado, nombre_inn, fraccion, porcentaje, cadena_comun, _, _ = resultado
            if nombre_inn != "-":
                # Calcular porcentaje real
                caracteres_comunes = sum(1 for c in nombre_inn.lower() if c in nombre_ingresado.lower())
                total_caracteres = len(nombre_inn)
                porcentaje_num = (caracteres_comunes/total_caracteres) * 100
                
                # Guardar el mejor resultado para cada nombre INN
                nombre_inn_lower = nombre_inn.lower()
                if nombre_inn_lower not in mejores_resultados or porcentaje_num > mejores_resultados[nombre_inn_lower][1]:
                    mejores_resultados[nombre_inn_lower] = (resultado, porcentaje_num)
        
        # Ordenar resultados únicos por porcentaje de similitud
        resultados_ordenados = sorted(
            mejores_resultados.values(),
            key=lambda x: x[1],
            reverse=True
        )
        
        # Mostrar solo los primeros num_visible resultados únicos
        for i, (resultado, porcentaje_num) in enumerate(resultados_ordenados[:num_visible]):
            nombre_ingresado, nombre_inn = resultado[0], resultado[1]
            
            # Calcular caracteres comunes y porcentajes
            caracteres_comunes = sum(1 for c in nombre_inn.lower() if c in nombre_ingresado.lower())
            total_caracteres = len(nombre_inn)
            fraccion = f"{caracteres_comunes}/{total_caracteres}"
            porcentaje = f"{porcentaje_num:.2f}%"
            
            # Obtener y procesar nombre EMA
            nombre_ema, similitud_ema = similitudes_ema[i] if i < len(similitudes_ema) else ("-", 0)
            _, nombre_ema_resaltado = self.resaltar_caracteres_comunes(nombre_ingresado, nombre_ema) if nombre_ema != "-" else (None, "-")
            
            # Resaltar caracteres comunes (tall man lettering)
            nombre_ingresado_resaltado, nombre_inn_resaltado = self.resaltar_caracteres_comunes(nombre_ingresado, nombre_inn)
            
            # Encontrar la cadena común más larga y calcular tamaños
            cadena_comun = self.encontrar_cadena_comun_mas_larga(nombre_ingresado, nombre_inn)
            tiene_misma_cadena = "SÍ" if cadena_comun else "NO"
            tam_cadena = len(cadena_comun) if cadena_comun else 0
            tam_cadena_fraccion = f"{tam_cadena}/{len(nombre_inn)}"
            
            # Calcular relevancia de la cadena
            relevancia_porcentaje = (tam_cadena / len(nombre_inn)) * 100 if cadena_comun else 0
            relevancia = f"{relevancia_porcentaje:.1f}% (low)" if relevancia_porcentaje <= 40 else f"{relevancia_porcentaje:.1f}% (high)"

            # Verificar posición
            if cadena_comun:
                pos_nombre1 = nombre_ingresado.lower().find(cadena_comun.lower())
                pos_nombre2 = nombre_inn.lower().find(cadena_comun.lower())
                misma_posicion = "SÍ" if pos_nombre1 == pos_nombre2 else "NO"
            else:
                misma_posicion = "NO"

            # Nueva lógica para determinar si es NO APTO
            es_no_apto = False
            if relevancia_porcentaje > 40:  # Solo considerar posición si la relevancia es high
                if misma_posicion == "SÍ":
                    es_no_apto = True
            # Si la relevancia es low, siempre es APTO respecto a la posición

            # Verificar error cognitivo
            if caracteres_comunes >= total_caracteres - 2:
                es_no_apto = True

            resultado_inn = "NO APTO" if es_no_apto else "APTO"

            # Verificar INN stem y resultado final
            resultado_inn_stem, stem_encontrado = self.verificar_inn_stem(nombre_ingresado, stems_dict)
            resultado_final = "APTO" if resultado_inn == "APTO" and resultado_inn_stem == "APTO" else "NO APTO"

            # Crear valores para insertar en el tree
            valores = (
                nombre_ingresado_resaltado,
                nombre_inn_resaltado,
                porcentaje,
                fraccion,
                cadena_comun if cadena_comun else "-",
                tiene_misma_cadena,
                tam_cadena_fraccion,
                relevancia,  # Nueva columna
                misma_posicion,
                resultado_inn,
                stem_encontrado,        # INN stem encontrado
                resultado_inn_stem,     # Resultado del análisis del stem
                resultado_final,        # Resultado combinado
                nombre_ema_resaltado,
                f"{similitud_ema:.2f}%"
            )
            
            # Insertar en el tree
            item_id = tree.insert('', tk.END, values=valores)

            # Verificar condiciones para resaltar la fila
            if resultado_inn == "APTO" and resultado_inn_stem == "APTO" and similitud_ema < 70:
                tree.item(item_id, tags=('resaltado',))

        # Configurar el estilo para el tag 'resaltado'
        tree.tag_configure('resaltado', background='lightgreen')

    def calcular_similitudes_ema(self, nombre_ingresado, nombres_ema):
        # Calcular similitudes con todos los nombres EMA
        similitudes_ema = []
        for nombre_ema in nombres_ema:
            # Si los nombres son idénticos, asignar 100%
            if nombre_ingresado.lower() == nombre_ema.lower():
                similitud = 100.0
            else:
                # Calcular similitud y asegurar que no exceda 100%
                similitud = min(self.ortho.score(nombre_ingresado, nombre_ema), 100.0)
            similitudes_ema.append((nombre_ema, similitud))
        
        # Ordenar por similitud de mayor a menor
        similitudes_ema.sort(key=lambda x: x[1], reverse=True)
        
        return similitudes_ema

    def encontrar_cadena_comun_mas_larga(self, str1, str2):
        """
        Encuentra la subcadena común más larga entre dos strings, ignorando mayúsculas/minúsculas
        """
        # Convertir ambas cadenas a minúsculas para la comparación
        str1 = str1.lower()
        str2 = str2.lower()
        
        len1, len2 = len(str1), len(str2)
        matriz = [[0] * (len2 + 1) for _ in range(len1 + 1)]
        max_length = 0
        end_pos = 0
        
        for i in range(1, len1 + 1):
            for j in range(1, len2 + 1):
                if str1[i-1] == str2[j-1]:
                    matriz[i][j] = matriz[i-1][j-1] + 1
                    if matriz[i][j] > max_length:
                        max_length = matriz[i][j]
                        end_pos = i
        
        return str1[end_pos - max_length:end_pos] if max_length > 0 else ""

    def encontrar_mejor_nombre_inn(self, nombre_ingresado, INN_list):
        """
        Calcula la similitud basada en cuántos caracteres del nombre INN están en el nombre ingresado
        """
        longitud_ingresado = len(nombre_ingresado)
        nombres_similares = [n for n in INN_list if abs(len(n) - longitud_ingresado) <= 2]
        
        mejor_similitud = 0
        mejor_nombre = None
        
        for nombre_inn in nombres_similares:
            if isinstance(nombre_inn, str):
                # Contar cuántos caracteres del nombre INN están en el nombre ingresado
                caracteres_presentes = sum(1 for c in nombre_inn.lower() if c in nombre_ingresado.lower())
                # Calcular el porcentaje basado en la longitud del nombre INN
                porcentaje = (caracteres_presentes / len(nombre_inn)) * 100
                
                if porcentaje > mejor_similitud:
                    mejor_similitud = porcentaje
                    mejor_nombre = nombre_inn
        
        return mejor_nombre, mejor_similitud

    def cargar_nombres_inn(self):
        """Carga los nombres INN desde un archivo Excel."""
        try:
            if not self.archivo_inn:
                messagebox.showerror("Error", "No se encontró el archivo INN.xlsx")
                return []
            df = pd.read_excel(self.archivo_inn)
            return df['Nombre'].str.lower().tolist()
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al cargar INN.xlsx:\n{e}")
            return []

    def mostrar_resultados_multiples(self, todos_resultados):
        if self.results_window is not None and self.results_window.winfo_exists():
            self.results_window.destroy()
        
        # Inicializar la lista de resultados guardados si no existe
        if not hasattr(self, 'saved_results'):
            self.saved_results = []
        
        # Guardar los resultados originales para poder restaurarlos
        self.resultados_originales = todos_resultados.copy()
        
        self.results_window = tk.Toplevel(self.root)
        self.results_window.title("Resultados")
        
        # Establecer un tamaño inicial del 80% de la pantalla
        screen_width = self.results_window.winfo_screenwidth()
        screen_height = self.results_window.winfo_screenheight()
        window_width = int(screen_width * 0.8)
        window_height = int(screen_height * 0.8)
        self.results_window.geometry(f"{window_width}x{window_height}")
        
        # Frame principal para dividir la ventana
        main_frame = ttk.Frame(self.results_window)
        main_frame.pack(fill='both', expand=True)
        
        # Frame superior para controles
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill='x', padx=5, pady=5)
        
        # Botón GUARDAR con estilo destacado
        style = ttk.Style()
        style.configure('Guardar.TButton', 
                       font=('Arial', 10, 'bold'),
                       padding=5,
                       width=15)
        
        btn_guardar = ttk.Button(
            control_frame,
            text="GUARDAR",
            style='Guardar.TButton',
            command=lambda: self.guardar_resultados_temporales(tree)
        )
        btn_guardar.pack(side='left', padx=10)
        
        # Label para mostrar cantidad de resultados guardados
        self.lbl_guardados = ttk.Label(
            control_frame, 
            text=f"Resultados guardados: {len(self.saved_results)}",
            font=('Arial', 10)
        )
        self.lbl_guardados.pack(side='left', padx=20)
        
        # Botón FactCheck
        btn_factcheck = ttk.Button(
            control_frame,
            text="FactCheck",
            command=self.ejecutar_factcheck
        )
        btn_factcheck.pack(side='left', padx=20)
        
        # Botón Generar Informe
        btn_generar_informe = ttk.Button(
            control_frame,
            text="Generar Informe",
            command=self.generar_informe
        )
        btn_generar_informe.pack(side='left', padx=10)
        
        # Botón Siguiente
        btn_siguiente = ttk.Button(
            control_frame,
            text="Siguiente",
            command=self.siguiente_funcion
        )
        btn_siguiente.pack(side='left', padx=10)
        
        # Botón Generar Excel
        btn_excel = ttk.Button(
            control_frame,
            text="Generar Excel",
            command=self.generar_excel
        )
        btn_excel.pack(side='left', padx=10)
        
      

        # Crear Treeview
        columns = (
            'nombre_generado', 'nombre_inn', 'porcentaje_inn', 'caracteres_inn',
            'cadena_comun', 'misma_cadena', 'tam_cadena_comun', 'relevancia', 
            'misma_posicion', 'resultado_inn', 'inn_stem', 'resultado_inn_stem', 'resultado_final', 'nombre_ema', 'similitud_ema'
        )
        
        tree = ttk.Treeview(main_frame, columns=columns, show='headings')
        
        # Variable para rastrear el estado de ordenamiento
        self.orden_similitud_ema = False  # False = orden original, True = ordenado
        
        # Función para ordenar por similitud EMA
        def ordenar_por_similitud_ema():
            self.orden_similitud_ema = not self.orden_similitud_ema
            
            # Limpiar el árbol
            for item in tree.get_children():
                tree.delete(item)
            
            if self.orden_similitud_ema:
                # Ordenar por similitud EMA (de mayor a menor)
                resultados_ordenados = []
                for nombre, res in todos_resultados.items():
                    for r in res:
                        resultados_ordenados.append((nombre, r, float(r[5])))  # r[5] es similitud_ema
                resultados_ordenados.sort(key=lambda x: x[2], reverse=True)
                
                # Insertar resultados ordenados
                for _, resultado, _ in resultados_ordenados:
                    nombre_gen, nombre_inn, caracteres, porcentaje_inn, nombre_ema, similitud_ema = resultado
                    
                    # [Resto del código de inserción igual que antes]
                    # Encontrar la cadena común más larga
                    cadena_comun = self.encontrar_cadena_comun_mas_larga(nombre_gen, nombre_inn)
                    tiene_misma_cadena = "SÍ" if cadena_comun else "NO"
                    tam_cadena = len(cadena_comun) if cadena_comun else 0
                    tam_cadena_fraccion = f"{tam_cadena}/{len(nombre_inn)}"
                    
                    # Calcular relevancia de la cadena
                    relevancia_porcentaje = (tam_cadena / len(nombre_inn)) * 100 if cadena_comun else 0
                    relevancia = f"{relevancia_porcentaje:.1f}% (low)" if relevancia_porcentaje <= 40 else f"{relevancia_porcentaje:.1f}% (high)"
                    
                    # Verificar posición
                    if cadena_comun:
                        pos_nombre1 = nombre_gen.lower().find(cadena_comun.lower())
                        pos_nombre2 = nombre_inn.lower().find(cadena_comun.lower())
                        misma_posicion = "SÍ" if pos_nombre1 == pos_nombre2 else "NO"
                    else:
                        misma_posicion = "NO"
                    
                    # Nueva lógica para determinar si es NO APTO
                    es_no_apto = False
                    if relevancia_porcentaje > 40:  # Solo considerar posición si la relevancia es high
                        if misma_posicion == "SÍ":
                            es_no_apto = True
                    
                    # Verificar error cognitivo
                    caracteres_presentes, total_caracteres = map(int, caracteres.split('/'))
                    if caracteres_presentes >= total_caracteres - 1:
                        es_no_apto = True
                    
                    resultado_inn = "NO APTO" if es_no_apto else "APTO"
                    
                    # Verificar INN stem
                    resultado_inn_stem, stem_encontrado = self.verificar_inn_stem(nombre_gen, self.cargar_inn_stems())
                    resultado_final = "APTO" if resultado_inn == "APTO" and resultado_inn_stem == "APTO" else "NO APTO"
                    
                    valores = (
                        nombre_gen,
                        nombre_inn,
                        porcentaje_inn,
                        caracteres,
                        cadena_comun if cadena_comun else "-",
                        tiene_misma_cadena,
                        tam_cadena_fraccion,
                        relevancia,
                        misma_posicion,
                        resultado_inn,
                        stem_encontrado,        # Antes era inn_stem
                        resultado_inn_stem,     # Antes era resultado_stem
                        resultado_final,
                        nombre_ema,
                        f"{similitud_ema:.2f}%"
                    )
                    
                    item_id = tree.insert('', tk.END, values=valores)
                    
                    # Resaltar en verde si cumple las condiciones
                    if resultado_inn == "APTO" and resultado_inn_stem == "APTO" and float(similitud_ema) < 70:
                        tree.tag_configure('resaltado', background='lightgreen')
                        tree.item(item_id, tags=('resaltado',))
            else:
                # Restaurar orden original
                self.mostrar_resultados_tree(tree, self.resultados_originales)
        
        # Configurar encabezados y anchos de columna
        column_widths = {
            'nombre_generado': 120,
            'nombre_inn': 120,
            'porcentaje_inn': 80,
            'caracteres_inn': 120,
            'cadena_comun': 120,
            'misma_cadena': 120,
            'tam_cadena_comun': 120,
            'relevancia': 120,
            'misma_posicion': 100,
            'resultado_inn': 100,
            'inn_stem': 120,
            'resultado_inn_stem': 120,
            'resultado_final': 180,
            'nombre_ema': 120,
            'similitud_ema': 120
        }
        
        column_headings = {
            'nombre_generado': 'Nombre Generado',
            'nombre_inn': 'Nombre INN',
            'porcentaje_inn': '% INN',
            'caracteres_inn': 'Caracteres INN',
            'cadena_comun': 'Cadena Común',
            'misma_cadena': '¿Misma Cadena?',
            'tam_cadena_comun': 'Tamaño Cadena',
            'relevancia': 'Relevancia',
            'misma_posicion': '¿Misma Posición?',
            'resultado_inn': 'Resultado INN',
            'inn_stem': 'INN stem',
            'resultado_inn_stem': 'Resultado INN stem',
            'resultado_final': 'Resultado INN e INN stem',
            'nombre_ema': 'Nombre EMA',
            'similitud_ema': 'Similitud EMA'
        }
        
        for col in columns:
            if col == 'similitud_ema':
                tree.heading(col, text=column_headings[col], command=ordenar_por_similitud_ema)
            else:
                tree.heading(col, text=column_headings[col])
            tree.column(col, width=column_widths[col], minwidth=column_widths[col], anchor='center')
        
        # Agregar scrollbars
        v_scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=tree.yview)
        h_scrollbar = ttk.Scrollbar(main_frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Posicionar elementos
        tree.pack(side='left', fill='both', expand=True)
        v_scrollbar.pack(side='right', fill='y')
        h_scrollbar.pack(side='bottom', fill='x')
        
        # Mostrar resultados iniciales
        self.mostrar_resultados_tree(tree, todos_resultados)
        
        # Configurar el estilo para el tag 'resaltado'
        tree.tag_configure('resaltado', background='lightgreen')

    def mostrar_resultados_tree(self, tree, resultados):
        # Limpiar el árbol
        for item in tree.get_children():
            tree.delete(item)
            
        # Insertar resultados
        for nombre_generado, resultados_nombre in resultados.items():
            for resultado in resultados_nombre:
                # Desempaquetar todos los valores
                nombre_gen, nombre_inn, caracteres, porcentaje_inn, nombre_ema, similitud_ema, inn_stem, resultado_inn_stem, resultado_final = resultado
                
                # Encontrar la cadena común más larga
                cadena_comun = self.encontrar_cadena_comun_mas_larga(nombre_gen, nombre_inn)
                tiene_misma_cadena = "SÍ" if cadena_comun else "NO"
                tam_cadena = len(cadena_comun) if cadena_comun else 0
                tam_cadena_fraccion = f"{tam_cadena}/{len(nombre_inn)}"
                
                # Calcular relevancia de la cadena
                relevancia_porcentaje = (tam_cadena / len(nombre_inn)) * 100 if cadena_comun else 0
                relevancia = f"{relevancia_porcentaje:.1f}% (low)" if relevancia_porcentaje <= 40 else f"{relevancia_porcentaje:.1f}% (high)"
                
                # Verificar posición
                if cadena_comun:
                    pos_nombre1 = nombre_gen.lower().find(cadena_comun.lower())
                    pos_nombre2 = nombre_inn.lower().find(cadena_comun.lower())
                    misma_posicion = "SÍ" if pos_nombre1 == pos_nombre2 else "NO"
                else:
                    misma_posicion = "NO"
                
                # Nueva lógica para determinar si es NO APTO
                es_no_apto = False
                if relevancia_porcentaje > 40:  # Solo considerar posición si la relevancia es high
                    if misma_posicion == "SÍ":
                        es_no_apto = True
                
                # Verificar error cognitivo
                caracteres_presentes, total_caracteres = map(int, caracteres.split('/'))
                if caracteres_presentes >= total_caracteres - 1:
                    es_no_apto = True
                
                resultado_inn = "NO APTO" if es_no_apto else "APTO"
                
                valores = (
                    nombre_gen,
                    nombre_inn,
                    porcentaje_inn,
                    caracteres,
                    cadena_comun if cadena_comun else "-",
                    tiene_misma_cadena,
                    tam_cadena_fraccion,
                    relevancia,
                    misma_posicion,
                    resultado_inn,
                    inn_stem,        # Ahora incluimos estos valores
                    resultado_inn_stem,
                    resultado_final,
                    nombre_ema,
                    f"{similitud_ema:.2f}%"
                )
                
                item_id = tree.insert('', tk.END, values=valores)
                
                # Resaltar en verde si cumple las condiciones
                if resultado_inn == "APTO" and resultado_inn_stem == "APTO" and float(similitud_ema) < 70:
                    tree.tag_configure('resaltado', background='lightgreen')
                    tree.item(item_id, tags=('resaltado',))

    def generar_nombres_chatgpt(self):
        """Genera nombres utilizando la API de ChatGPT y analiza su similitud"""
        try:
            if self.ema_var.get():
                # Verificar archivo EMA
                if not os.path.exists(self.archivo_ema):
                    messagebox.showerror("Error", f"No se encuentra el archivo: {self.archivo_ema}")
                    return

                # Obtener la cantidad exacta de nombres a generar
                cantidad_nombres = int(self.num_nombres_gen.get())
                if cantidad_nombres <= 0:
                    messagebox.showwarning("Advertencia", "La cantidad de nombres debe ser mayor que 0")
                    return

                nombres_generados = []
                intentos = 0
                max_intentos = 5  # Máximo número de intentos para obtener todos los nombres

                while len(nombres_generados) < cantidad_nombres and intentos < max_intentos:
                    try:
                        # Obtener el prefijo deseado
                        prefijo = self.inicio_nombre.get().lower() or 'aba'
                        
                        # Calcular cuántos nombres faltan
                        nombres_faltantes = cantidad_nombres - len(nombres_generados)
                        
                        # Solicitar más nombres de los necesarios para tener margen
                        nombres_a_solicitar = min(nombres_faltantes * 2, 100)  # Limitamos a 100 por solicitud
                        
                        # Construir el prompt
                        prompt = (
                            f"Genera exactamente {nombres_a_solicitar} nombres de medicamentos únicos y diferentes que:"
                            f"\n- Comiencen con '{prefijo}'"
                            f"\n- Tengan entre {len(prefijo)} y 8 caracteres"
                            f"\n- Suenen como nombres de medicamentos reales"
                            f"\n- No incluyan números ni caracteres especiales"
                            f"\n- Sean diferentes a estos nombres ya generados: {', '.join(nombres_generados)}"
                            f"\nResponde solo con la lista de nombres, uno por línea, sin numeración ni explicaciones."
                        )

                        response = openai.ChatCompletion.create(
                            model="gpt-3.5-turbo",
                            messages=[{"role": "user", "content": prompt}],
                            max_tokens=1000,
                            temperature=0.8
                        )

                        # Procesar la respuesta
                        texto_respuesta = response.choices[0].message.content
                        for linea in texto_respuesta.split('\n'):
                            nombre = linea.strip()
                            if (nombre and 
                                nombre.lower().startswith(prefijo) and 
                                len(nombre) < 9 and 
                                nombre.capitalize() not in nombres_generados):
                                nombres_generados.append(nombre.capitalize())
                                if len(nombres_generados) >= cantidad_nombres:
                                    break

                        intentos += 1

                    except Exception as e:
                        messagebox.showerror("Error ChatGPT", f"Error en intento {intentos + 1}: {str(e)}")
                        intentos += 1
                        continue

                if len(nombres_generados) < cantidad_nombres:
                    messagebox.showwarning(
                        "Advertencia", 
                        f"Solo se pudieron generar {len(nombres_generados)} nombres de los {cantidad_nombres} solicitados"
                    )

                # Procesar los nombres generados
                todos_resultados = {}
                
                # Cargar nombres de EMA y INN
                df_ema = pd.read_excel(self.archivo_ema)
                nombres_ema = df_ema['EMA'].str.lower().tolist()
                nombres_ema = [n for n in nombres_ema if ' ' not in n]
                INN_list = self.cargar_nombres_inn()

                # Procesar cada nombre generado
                for nombre_generado in nombres_generados:
                    # Calcular similitudes EMA
                    similitudes_ema = []
                    for nombre_ema in nombres_ema:
                        if isinstance(nombre_ema, str):
                            similitud = self.ortho.score(nombre_generado.lower(), nombre_ema.lower())
                            similitudes_ema.append((nombre_ema, similitud))

                    # Ordenar por similitud y tomar el mejor
                    similitudes_ema.sort(key=lambda x: x[1], reverse=True)
                    mejor_ema = similitudes_ema[0] if similitudes_ema else ("-", 0)

                    # Calcular similitudes INN
                    similitudes_inn = []
                    for nombre_inn in INN_list:
                        if isinstance(nombre_inn, str):
                            caracteres_presentes = sum(1 for c in nombre_inn.lower() if c in nombre_generado.lower())
                            total_caracteres = len(nombre_inn)
                            porcentaje = (caracteres_presentes / total_caracteres) * 100
                            similitudes_inn.append((
                                nombre_inn,
                                caracteres_presentes,
                                total_caracteres,
                                porcentaje
                            ))

                    # Ordenar por porcentaje y tomar el mejor
                    similitudes_inn.sort(key=lambda x: x[3], reverse=True)
                    mejor_inn = similitudes_inn[0] if similitudes_inn else ("-", 0, 0, 0)

                    resultados = [(
                        nombre_generado,
                        mejor_inn[0],  # nombre_inn
                        f"{mejor_inn[1]}/{mejor_inn[2]}",  # caracteres
                        f"{mejor_inn[3]:.2f}%",  # porcentaje_inn
                        mejor_ema[0],  # nombre_ema
                        mejor_ema[1]   # similitud_ema
                    )]

                    todos_resultados[nombre_generado] = resultados

                # Mostrar resultados
                if todos_resultados:
                    self.ultimos_resultados = todos_resultados
                    self.btn_excel['state'] = 'normal'
                    self.mostrar_resultados_multiples(todos_resultados)
                    self.root.update_idletasks()
                

            else:
                messagebox.showwarning("Advertencia", "Active la opción EMA para generar nombres")

        except Exception as e:
            messagebox.showerror("Error", f"Error al generar nombres: {str(e)}")

    def generar(self):
        try:
            if self.ema_var.get():
                # Verificar archivo EMA
                if not os.path.exists(self.archivo_ema):
                    messagebox.showerror("Error", f"No se encuentra el archivo: {self.archivo_ema}")
                    return

                # Obtener la cantidad exacta de nombres a generar
                cantidad_nombres = int(self.num_nombres_gen.get())
                if cantidad_nombres <= 0:
                    messagebox.showwarning("Advertencia", "La cantidad de nombres debe ser mayor que 0")
                    return

                # Generar nombres usando ChatGPT
                try:
                    # Obtener el prefijo deseado
                    prefijo = self.inicio_nombre.get().lower() or 'aba'
                    
                    
                    # Solicitar nombres a ChatGPT
                    response = openai.ChatCompletion.create(
                        model="gpt-3.5-turbo",
                        messages=[
                            {"role": "user", "content": f"Genera exactamente {cantidad_nombres} nombres de medicamentos que comiencen con '{prefijo}' y tengan menos de 9 caracteres. Responde solo con la lista de nombres, uno por línea, sin numeración ni explicaciones. Asegúrate de generar exactamente {cantidad_nombres} nombres válidos."}
                        ],
                        max_tokens=1000,
                        temperature=0.7
                    )

                    # Extraer los nombres generados
                    nombres_generados = []
                    texto_respuesta = response.choices[0].message.content
                    for linea in texto_respuesta.split('\n'):
                        nombre = linea.strip()
                        if nombre and nombre.lower().startswith(prefijo) and len(nombre) < 9:
                            nombres_generados.append(nombre.capitalize())

                    # Verificar si obtuvimos la cantidad correcta de nombres
                    if len(nombres_generados) < cantidad_nombres:
                        # Si faltan nombres, hacer otra solicitud
                        nombres_faltantes = cantidad_nombres - len(nombres_generados)
                        response_adicional = openai.ChatCompletion.create(
                            model="gpt-3.5-turbo",
                            messages=[
                                {"role": "user", "content": f"Genera exactamente {nombres_faltantes} nombres adicionales de medicamentos que comiencen con '{prefijo}' y tengan menos de 9 caracteres, diferentes a los siguientes nombres: {', '.join(nombres_generados)}. Responde solo con la lista de nombres, uno por línea."}
                            ],
                            max_tokens=1000,
                            temperature=0.8
                        )
                        
                        # Agregar los nombres adicionales
                        texto_adicional = response_adicional.choices[0].message.content
                        for linea in texto_adicional.split('\n'):
                            nombre = linea.strip()
                            if nombre and nombre.lower().startswith(prefijo) and len(nombre) < 9:
                                if len(nombres_generados) < cantidad_nombres:
                                    nombres_generados.append(nombre.capitalize())

                except Exception as e:
                    messagebox.showerror("Error ChatGPT", f"Error al generar nombres con ChatGPT: {str(e)}")
                    return

                # Cargar nombres de EMA y INN
                df_ema = pd.read_excel(self.archivo_ema)
                nombres_ema = df_ema['EMA'].str.lower().tolist()
                nombres_ema = [n for n in nombres_ema if ' ' not in n]
                INN_list = self.cargar_nombres_inn()

                todos_resultados = {}

                # Procesar cada nombre generado
                for nombre_generado in nombres_generados:
                    # Calcular similitudes EMA
                    similitudes_ema = []
                    for nombre_ema in nombres_ema:
                        if isinstance(nombre_ema, str):
                            similitud = self.ortho.score(nombre_generado.lower(), nombre_ema.lower())
                            similitudes_ema.append((nombre_ema, similitud))

                    # Ordenar por similitud y tomar el mejor
                    similitudes_ema.sort(key=lambda x: x[1], reverse=True)
                    mejor_ema = similitudes_ema[0] if similitudes_ema else ("-", 0)

                    # Calcular similitudes INN
                    similitudes_inn = []
                    for nombre_inn in INN_list:
                        if isinstance(nombre_inn, str):
                            caracteres_presentes = sum(1 for c in nombre_inn.lower() if c in nombre_generado.lower())
                            total_caracteres = len(nombre_inn)
                            porcentaje = (caracteres_presentes / total_caracteres) * 100
                            similitudes_inn.append((
                                nombre_inn,
                                caracteres_presentes,
                                total_caracteres,
                                porcentaje
                            ))

                    # Ordenar por porcentaje y tomar el mejor
                    similitudes_inn.sort(key=lambda x: x[3], reverse=True)
                    mejor_inn = similitudes_inn[0] if similitudes_inn else ("-", 0, 0, 0)

                    # Verificar INN stem y resultado final
                    resultado_inn_stem, stem_encontrado = self.verificar_inn_stem(nombre_generado, self.cargar_inn_stems())
                    resultado_final = "APTO" if mejor_inn[3] >= 70 and resultado_inn_stem == "APTO" else "NO APTO"

                    resultados = [(
                        nombre_generado,
                        mejor_inn[0],  # nombre_inn
                        f"{mejor_inn[1]}/{mejor_inn[2]}",  # caracteres
                        f"{mejor_inn[3]:.2f}%",  # porcentaje_inn
                        mejor_ema[0],  # nombre_ema
                        mejor_ema[1],  # similitud_ema
                        stem_encontrado,  # INN stem encontrado
                        resultado_inn_stem,  # Resultado del análisis del stem
                        resultado_final  # Resultado combinado
                    )]

                    todos_resultados[nombre_generado] = resultados

                # Mostrar resultados
                if todos_resultados:
                    self.ultimos_resultados = todos_resultados
                    self.btn_excel['state'] = 'normal'
                    self.mostrar_resultados_multiples(todos_resultados)
                    self.root.update_idletasks()
                

            else:
                messagebox.showwarning("Advertencia", "Active la opción EMA para generar nombres")

        except Exception as e:
            messagebox.showerror("Error", f"Error al generar nombres: {str(e)}")

    def setup_results_table(self):
        # Configurar la tabla de resultados
        self.results_window.geometry("800x600")
        main_frame = ttk.Frame(self.results_window)
        main_frame.pack(fill='both', expand=True)
        
        # Frame para los botones
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', padx=5, pady=5)

        # Botón GUARDAR con estilo destacado
        btn_guardar = ttk.Button(
            button_frame,
            text="GUARDAR",
            style='Guardar.TButton',
            command=lambda: self.guardar_resultados_temporales(self.tree)
        )
        btn_guardar.pack(side='left', padx=10)

        # Etiqueta para el botón FactCheck
        label_factcheck = ttk.Label(button_frame, text="Ejecutar FactCheck:")
        label_factcheck.pack(side='left', padx=5)

        # Botón FACTCHECK
        btn_factcheck = ttk.Button(
            button_frame,
            text="FactCheck",
            style='FactCheck.TButton',
            command=self.ejecutar_factcheck  # Cambia a un método que ejecuta main.py
        )
        btn_factcheck.pack(side='left', padx=10)

        # Inicializar el Treeview
        columns = (
            'nombre_generado', 'nombre_inn', 'porcentaje_inn', 'caracteres_inn',
            'cadena_comun', 'misma_cadena', 'tam_cadena_comun', 'relevancia', 
            'misma_posicion', 'resultado_inn', 'inn_stem', 'resultado_inn_stem', 'resultado_final', 'nombre_ema', 'similitud_ema'
        )
        
        self.tree = ttk.Treeview(main_frame, columns=columns, show='headings')  # Asegúrate de usar self.tree
        
        # Configurar encabezados y anchos de columna
        column_widths = {
            'check': 50,  # Ancho para la columna de verificación
            'nombre_generado': 120,
            'nombre_inn': 120,
            'porcentaje_inn': 80,
            'caracteres_inn': 100,
            'cadena_comun': 100,
            'misma_cadena': 100,
            'tam_cadena_comun': 100,
            'relevancia': 100,
            'misma_posicion': 100,
            'resultado_inn': 100,
            'inn_stem': 120,
            'resultado_inn_stem': 120,
            'resultado_final': 180,
            'nombre_ema': 120,
            'similitud_ema': 100
        }
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths[col], minwidth=column_widths[col], anchor='center')

        self.tree.pack(side='top', fill='both', expand=True)

        # Agregar algunos datos de ejemplo al Treeview
        self.agregar_datos_ejemplo()

    def agregar_datos_ejemplo(self):
        # Agregar datos de ejemplo al Treeview
        datos = [
            ("", "Nombre1", "INN1", "80%", "5/6", "Cadena1", "SÍ", "5", "80%", "SÍ", "APTO", "EMA1", "90%"),
            ("", "Nombre2", "INN2", "70%", "4/6", "Cadena2", "NO", "4", "70%", "NO", "NO APTO", "EMA2", "85%"),
            ("", "Nombre3", "INN3", "60%", "3/6", "Cadena3", "SÍ", "3", "60%", "SÍ", "APTO", "EMA3", "75%")
        ]
        
        for item in datos:
            self.tree.insert('', 'end', values=item)

    def generar_informe(self):
        # Verificar si hay resultados guardados
        if not hasattr(self, 'saved_results') or not self.saved_results:
            messagebox.showwarning("Advertencia", "No hay resultados guardados para generar un informe.")
            return

        try:
            # Crear un nuevo documento
            doc = Document()
            
            # Configurar el documento para orientación vertical (por defecto)
            section = doc.sections[0]
            section.page_width = Mm(210)  # A4 width
            section.page_height = Mm(297)  # A4 height
            
            # Estilos personalizados
            styles = doc.styles
            
            # Estilo para título principal
            style_title = styles.add_style('CustomTitle', WD_STYLE_TYPE.PARAGRAPH)
            style_title.font.name = 'Arial'
            style_title.font.size = Pt(24)
            style_title.font.bold = True
            style_title.font.color.rgb = RGBColor(0x00, 0x32, 0x96)  # Azul corporativo
            
            # Estilo para subtítulos
            style_subtitle = styles.add_style('CustomSubtitle', WD_STYLE_TYPE.PARAGRAPH)
            style_subtitle.font.name = 'Arial'
            style_subtitle.font.size = Pt(16)
            style_subtitle.font.bold = True
            style_subtitle.font.color.rgb = RGBColor(0x00, 0x32, 0x96)
            
            # Estilo para texto normal
            style_normal = styles.add_style('CustomNormal', WD_STYLE_TYPE.PARAGRAPH)
            style_normal.font.name = 'Arial'
            style_normal.font.size = Pt(11)
            
            # Título principal
            title = doc.add_paragraph('INFORME EJECUTIVO DE ANÁLISIS DE SIMILITUD FARMACÉUTICA', style='CustomTitle')
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Agregar fecha y hora
            fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
            p_fecha = doc.add_paragraph(f'Fecha de generación: {fecha_actual}', style='CustomNormal')
            p_fecha.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
            # Línea separadora
            doc.add_paragraph('_' * 100, style='CustomNormal').alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # 1. Resumen Ejecutivo
            heading = doc.add_heading('1. RESUMEN EJECUTIVO', level=1)
            heading.style = styles['CustomSubtitle']
            
            total_analizados = len(self.saved_results)
            aptos = sum(1 for r in self.saved_results if r['Resultado Final'] == 'APTO')
            no_aptos = total_analizados - aptos
            
            resumen = doc.add_paragraph(style='CustomNormal')
            resumen.add_run('Se ha realizado un análisis exhaustivo de similitud farmacéutica para ').bold = False
            resumen.add_run(f'{total_analizados} nombres propuestos').bold = True
            resumen.add_run(', evaluando su viabilidad según criterios INN, INN stem y EMA. Del total analizado: \n\n')
            
            # Tabla de resumen
            table = doc.add_table(rows=2, cols=3)
            table.style = 'Table Grid'
            table.autofit = False
            
            # Encabezados
            headers = table.rows[0].cells
            headers[0].text = 'Total Analizados'
            headers[1].text = 'APTOS'
            headers[2].text = 'NO APTOS'
            
            # Datos
            data_row = table.rows[1].cells
            data_row[0].text = str(total_analizados)
            data_row[1].text = str(aptos)
            data_row[2].text = str(no_aptos)
            
            # Formato de tabla
            for row in table.rows:
                for cell in row.cells:
                    cell.width = Inches(2)
                    paragraphs = cell.paragraphs
                    for paragraph in paragraphs:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for run in paragraph.runs:
                            run.font.name = 'Arial'
                            run.font.size = Pt(10)
            
            doc.add_paragraph()  # Espacio después de la tabla
            
            # 2. Análisis Detallado
            heading = doc.add_heading('2. ANÁLISIS DETALLADO DE NOMBRES PROPUESTOS', level=1)
            heading.style = styles['CustomSubtitle']
            
            # Crear tabla para el análisis detallado
            detailed_table = doc.add_table(rows=1, cols=7)
            detailed_table.style = 'Table Grid'
            
            # Encabezados de la tabla detallada
            headers = detailed_table.rows[0].cells
            header_texts = ['Nombre Propuesto', 'Análisis INN', 'Análisis Stem', 'Similitud EMA', 'Resultado Final', 'Observaciones', 'Recomendación']
            for i, text in enumerate(header_texts):
                headers[i].text = text
                headers[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = headers[i].paragraphs[0].runs[0]
                run.font.bold = True
                run.font.name = 'Arial'
                run.font.size = Pt(10)
            
            # Agregar datos a la tabla
            for resultado in self.saved_results:
                row_cells = detailed_table.add_row().cells
                
                # Nombre Propuesto
                row_cells[0].text = resultado['Nombre Generado']
                
                # Análisis INN
                inn_analisis = f"Coincidencia: {resultado['Porcentaje INN']}\n"
                inn_analisis += f"Resultado: {resultado['Resultado INN']}"
                row_cells[1].text = inn_analisis
                
                # Análisis Stem
                stem_analisis = f"Stem encontrado: {resultado['Stem INN']}\n"
                stem_analisis += f"Resultado: {resultado['Resultado INN Stem']}"
                row_cells[2].text = stem_analisis
                
                # Similitud EMA
                similitud_ema = float(resultado['Similitud EMA'].strip('%'))
                row_cells[3].text = f"{similitud_ema:.1f}%"
                
                # Resultado Final
                row_cells[4].text = resultado['Resultado Final']
                
                # Observaciones
                observaciones = []
                if resultado['Cadena Común'] != "-":
                    observaciones.append(f"Cadena común encontrada: {resultado['Cadena Común']}")
                if resultado['Misma Posición'] == "SÍ":
                    observaciones.append("Cadena en misma posición")
                row_cells[5].text = "\n".join(observaciones) if observaciones else "Sin observaciones relevantes"
                
                # Recomendación
                if resultado['Resultado Final'] == 'APTO':
                    recomendacion = "VIABLE para uso"
                    if similitud_ema > 55:
                        recomendacion += "\nRequiere monitoreo"
                else:
                    recomendacion = "NO RECOMENDADO"
                row_cells[6].text = recomendacion
                
                # Aplicar formato a las celdas
                for cell in row_cells:
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.name = 'Arial'
                            run.font.size = Pt(9)
            
            # Ajustar anchos de columna para orientación vertical
            widths = [1.5, 1.5, 1.5, 1, 1, 1.5, 1]  # Reducidos para ajustarse al ancho vertical
            for i, width in enumerate(widths):
                for row in detailed_table.rows:
                    row.cells[i].width = Inches(width)
            
            # 3. Conclusiones y Recomendaciones
            heading = doc.add_heading('3. CONCLUSIONES Y RECOMENDACIONES', level=1)
            heading.style = styles['CustomSubtitle']
            
            # Generar conclusiones con ChatGPT
            try:
                openai.api_key 
                
                # Preparar datos para el prompt
                nombres_aptos = [r['Nombre Generado'] for r in self.saved_results if r['Resultado Final'] == 'APTO']
                nombres_no_aptos = [r['Nombre Generado'] for r in self.saved_results if r['Resultado Final'] == 'NO APTO']
                
                prompt = f"""
                Genera conclusiones y recomendaciones ejecutivas para un informe farmacéutico-legal basado en los siguientes datos:

                - Total de nombres analizados: {total_analizados}
                - Nombres APTOS: {len(nombres_aptos)} ({', '.join(nombres_aptos)})
                - Nombres NO APTOS: {len(nombres_no_aptos)} ({', '.join(nombres_no_aptos)})

                Las conclusiones deben:
                1. Ser profesionales y objetivas
                2. Incluir aspectos legales y regulatorios
                3. Mencionar riesgos y medidas de mitigación
                4. Proporcionar recomendaciones específicas
                5. Usar lenguaje técnico pero comprensible

                Formato deseado:
                - 2-3 párrafos de conclusiones
                - 3-4 recomendaciones específicas
                """

                response = openai.ChatCompletion.create(
                    model="gpt-3.5-turbo",
                    messages=[{"role": "user", "content": prompt}],
                    max_tokens=1000,
                    temperature=0.7
                )

                conclusiones = response.choices[0].message.content
                
                # Agregar conclusiones al documento
                p_conclusiones = doc.add_paragraph(conclusiones, style='CustomNormal')
                
            except Exception as e:
                # Si falla la generación con ChatGPT, usar conclusiones genéricas
                p_conclusiones = doc.add_paragraph(style='CustomNormal')
                p_conclusiones.add_run('Conclusiones del Análisis:\n').bold = True
                p_conclusiones.add_run(f'Del total de {total_analizados} nombres analizados, {aptos} cumplieron con todos los criterios de aptitud establecidos. ')
                p_conclusiones.add_run('Se recomienda proceder con los nombres marcados como APTOS, prestando especial atención a aquellos que requieren monitoreo adicional.')
            
            # 4. Pie de página
            section = doc.sections[0]
            footer = section.footer
            footer_para = footer.paragraphs[0]
            footer_para.text = "Documento generado automáticamente por el sistema de análisis de similitud farmacéutica"
            footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Guardar el documento
            ruta_archivo = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word Documents", "*.docx")],
                title="Guardar informe como"
            )
            
            if ruta_archivo:
                doc.save(ruta_archivo)
                messagebox.showinfo("Éxito", "Informe ejecutivo generado correctamente")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al generar el informe: {str(e)}")

    # Método para ejecutar main.py
    def ejecutar_factcheck(self):
        try:
            if not self.archivo_factcheck:
                messagebox.showerror("Error", "No se encontró el archivo factcheck.py")
                return
            subprocess.run(['python', self.archivo_factcheck], check=True)
        except subprocess.CalledProcessError as e:
            print(f"Error al ejecutar factcheck.py: {e}")
        except Exception as e:
            print(f"Error inesperado: {e}")

    def guardar_resultados_temporales(self, tree):
        """Guarda temporalmente los resultados actuales que están en verde"""
        try:
            # Obtener todos los items del tree
            for item_id in tree.get_children():
                valores = tree.item(item_id)['values']
                tags = tree.item(item_id)['tags']
                
                # Solo guardar los items resaltados en verde
                if 'resaltado' in tags:
                    # Crear un diccionario con los valores
                    resultado = {
                        'Nombre Generado': valores[0],
                        'Nombre INN': valores[1],
                        'Porcentaje INN': valores[2],
                        'Caracteres INN': valores[3],
                        'Cadena Común': valores[4],
                        'Misma Cadena': valores[5],
                        'Tamaño Cadena': valores[6],
                        'Relevancia': valores[7],
                        'Misma Posición': valores[8],
                        'Resultado INN': valores[9],
                        'Stem INN': valores[10],
                        'Resultado INN Stem': valores[11],
                        'Resultado Final': valores[12],
                        'Nombre EMA': valores[13],
                        'Similitud EMA': valores[14]
                    }
                    
                    # Verificar si el resultado ya está guardado
                    if not any(r['Nombre Generado'] == resultado['Nombre Generado'] for r in self.saved_results):
                        self.saved_results.append(resultado)
            
            # Actualizar el label con la cantidad de resultados guardados
            self.lbl_guardados.config(text=f"Resultados guardados: {len(self.saved_results)}")
            messagebox.showinfo("Éxito", f"Se han guardado {len(self.saved_results)} resultados APTOS")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar resultados: {str(e)}")

    def generar_excel(self):
        """Genera un archivo Excel con los resultados guardados"""
        try:
            if not hasattr(self, 'saved_results') or not self.saved_results:
                messagebox.showwarning("Advertencia", "No hay resultados guardados para exportar")
                return
            
            # Solicitar ubicación para guardar el archivo
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Guardar resultados como"
            )
            
            if filename:
                # Crear un DataFrame con los resultados guardados
                df = pd.DataFrame(self.saved_results)
                
                # Crear un escritor de Excel con formato
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    # Escribir el DataFrame empezando desde la fila 2 para dejar espacio para el título
                    df.to_excel(writer, sheet_name='Resultados APTOS', index=False, startrow=2)
                    
                    # Obtener la hoja de trabajo
                    workbook = writer.book
                    worksheet = writer.sheets['Resultados APTOS']
                    
                    # Obtener la última columna
                    last_column = openpyxl.utils.get_column_letter(worksheet.max_column)
                    
                    # Dar formato al título
                    titulo = "RESULTADOS DE ANÁLISIS DE NOMBRES"
                    worksheet.merge_cells(f'A1:{last_column}1')
                    cell_titulo = worksheet.cell(row=1, column=1)
                    cell_titulo.value = titulo
                    cell_titulo.font = openpyxl.styles.Font(size=16, bold=True, color='FFFFFF')
                    cell_titulo.fill = openpyxl.styles.PatternFill(
                        start_color='4F81BD',
                        end_color='4F81BD',
                        fill_type='solid'
                    )
                    cell_titulo.alignment = openpyxl.styles.Alignment(
                        horizontal='center',
                        vertical='center'
                    )
                    
                    # Dar formato a los encabezados
                    for col in range(1, worksheet.max_column + 1):
                        column_letter = openpyxl.utils.get_column_letter(col)
                        cell = worksheet.cell(row=2, column=col)
                        cell.fill = openpyxl.styles.PatternFill(
                            start_color='8DB4E2',
                            end_color='8DB4E2',
                            fill_type='solid'
                        )
                        cell.font = openpyxl.styles.Font(bold=True, color='000000')
                        cell.alignment = openpyxl.styles.Alignment(
                            horizontal='center',
                            vertical='center',
                            wrap_text=True
                        )
                        cell.border = openpyxl.styles.Border(
                            left=openpyxl.styles.Side(style='thin'),
                            right=openpyxl.styles.Side(style='thin'),
                            top=openpyxl.styles.Side(style='thin'),
                            bottom=openpyxl.styles.Side(style='thin')
                        )
                    
                    # Dar formato a las celdas de datos
                    for row in range(3, worksheet.max_row + 1):
                        for col in range(1, worksheet.max_column + 1):
                            cell = worksheet.cell(row=row, column=col)
                            cell.alignment = openpyxl.styles.Alignment(
                                horizontal='center',
                                vertical='center',
                                wrap_text=True
                            )
                            cell.border = openpyxl.styles.Border(
                                left=openpyxl.styles.Side(style='thin'),
                                right=openpyxl.styles.Side(style='thin'),
                                top=openpyxl.styles.Side(style='thin'),
                                bottom=openpyxl.styles.Side(style='thin')
                            )
                            
                            # Colorear filas alternadas
                            if row % 2 == 0:
                                cell.fill = openpyxl.styles.PatternFill(
                                    start_color='F2F2F2',
                                    end_color='F2F2F2',
                                    fill_type='solid'
                                )
                    
                    # Ajustar el ancho de las columnas
                    for col in range(1, worksheet.max_column + 1):
                        column_letter = openpyxl.utils.get_column_letter(col)
                        max_length = 0
                        for row in range(1, worksheet.max_row + 1):
                            cell = worksheet.cell(row=row, column=col)
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = (max_length + 2)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
                    
                    # Ajustar altura de las filas
                    worksheet.row_dimensions[1].height = 30  # Título
                    worksheet.row_dimensions[2].height = 40  # Encabezados
                    for row in range(3, worksheet.max_row + 1):
                        worksheet.row_dimensions[row].height = 25
                
                messagebox.showinfo("Éxito", "Archivo Excel generado correctamente")
                
                # Limpiar resultados guardados después de exportar
                self.saved_results = []
                if hasattr(self, 'lbl_guardados'):
                    self.lbl_guardados.config(text="Resultados guardados: 0")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al generar Excel: {str(e)}")

    def validar_numero(self, valor):
        if valor == "":
            return True
        try:
            num = int(valor)
            return num > 0 and num <= 100
        except ValueError:
            return False

    def resaltar_caracteres_comunes(self, nombre1, nombre2):
        """Devuelve ambos nombres con los caracteres comunes en mayúsculas"""
        nombre1 = nombre1.lower()
        nombre2 = nombre2.lower()
        resultado1 = []
        resultado2 = []
        
        # Crear una lista de pares de posiciones donde hay caracteres iguales
        pares_comunes = []
        for i in range(len(nombre1)):
            for j in range(len(nombre2)):
                if nombre1[i] == nombre2[j]:
                    # Verificar si esta posición ya fue usada
                    if not any(i in par1 or j in par2 for par1, par2 in pares_comunes):
                        pares_comunes.append(([i], [j]))
                        break
        
        # Convertir a sets para fácil búsqueda
        pos_comunes1 = {pos for par in pares_comunes for pos in par[0]}
        pos_comunes2 = {pos for par in pares_comunes for pos in par[1]}
        
        # Construir los nombres con los caracteres comunes en mayúsculas
        for i, char in enumerate(nombre1):
            if i in pos_comunes1:
                resultado1.append(char.upper())
            else:
                resultado1.append(char.lower())
            
        for i, char in enumerate(nombre2):
            if i in pos_comunes2:
                resultado2.append(char.upper())
            else:
                resultado2.append(char.lower())
        
        return ''.join(resultado1), ''.join(resultado2)

    def generar_nombre_aleatorio(self):
        """Genera un nombre farmacéutico pronunciable en español/inglés"""
        # Obtener el inicio especificado
        inicio = self.inicio_nombre.get().lower()
        if not inicio:
            inicio = "aba"  # Valor por defecto si está vacío
        
        # Componentes comunes en nombres farmacéuticos
        silabas_medias = [
            'ce', 'ci', 'co', 'de', 'di', 'do', 'fe', 'fi', 'fo',
            'le', 'li', 'lo', 'me', 'mi', 'mo', 'ne', 'ni', 'no',
            'pe', 'pi', 'po', 're', 'ri', 'ro', 'se', 'si', 'so',
            've', 'vi', 'vo', 'te', 'ti', 'to'
        ]
        
        # Sufijos comunes en medicamentos reales
        sufijos_cortos = [
            'ex',   # Zantex
            'ax',   # Zithromax
            'ox',   # Wilcox
            'al',   # Xenical
            'ol',   # Tylenol
            'il',   # Tensil
            'an',   # Voltaren
            'en',   # Motrin
            'on'    # Naproxen
        ]
        
        sufijos_largos = [
            'ina',  # Aspirina
            'ida',  # Ranitidina
            'ene',  # Tramadene
            'one',  # Cortisone
            'ium',  # Calcium
            'ate',  # Folate
            'ine',  # Morphine
            'ide',  # Acetamide
            'oid',  # Corticoid
            'amol', # Paracetamol
            'icin', # Vancomicin
            'oxin'  # Digoxin
        ]
        
        # Construir el nombre empezando con las letras especificadas
        nombre = list(inicio)
        
        # Agregar 1-2 sílabas que sean fáciles de pronunciar
        num_silabas = random.randint(1, 2)
        ultima_silaba = ''.join(nombre[-2:]) if len(nombre) >= 2 else ''
        
        for _ in range(num_silabas):
            # Elegir una sílaba que no forme combinaciones difciles
            silaba = random.choice(silabas_medias)
            while (ultima_silaba and 
                  (silaba[0] == ultima_silaba[-1] or  # Evitar consonantes duplicadas
                   ultima_silaba[-1] + silaba[0] in ['dt', 'tk', 'pk', 'bt', 'fg', 'kg'])):  # Evitar combinaciones difíciles
                silaba = random.choice(silabas_medias)
            nombre.extend(list(silaba))
            ultima_silaba = silaba
        
        # Elegir un sufijo apropiado
        if random.random() < 0.6:  # 60% de probabilidad de sufijo corto
            sufijo = random.choice(sufijos_cortos)
        else:
            sufijo = random.choice(sufijos_largos)
        
        # Si el nombre termina en vocal y el sufijo empieza con vocal,
        # eliminar la última vocal para evitar combinaciones extrañas
        if nombre[-1] in 'aeiou' and sufijo[0] in 'aeiou':
            nombre.pop()
        
        nombre.extend(list(sufijo))
        
        # Convertir a string y capitalizar
        nombre_final = ''.join(nombre).capitalize()
        
        # Si el nombre es muy largo o muy corto, generar uno nuevo
        if len(nombre_final) > 12 or len(nombre_final) < 5:
            return self.generar_nombre_aleatorio()
        
        return nombre_final

    def fact_check(self):
        """Alterna entre los nombres más similares y los valores originales"""
        try:
            # Obtener todos los items de la tabla
            for item_id in self.tree.get_children():
                values = self.tree.item(item_id, 'values')
                nombre_generado = values[0]

                # Si no tenemos valores originales guardados, estamos en modo normal
                if item_id not in self.valores_originales:
                    # Guardar los valores originales antes de modificar
                    self.valores_originales[item_id] = values

                    # Cargar nombres INN y EMA
                    INN_list = self.cargar_nombres_inn()
                    df_ema = pd.read_excel(self.archivo_ema)
                    nombres_ema = df_ema['EMA'].str.lower().tolist()

                    # Calcular mejor coincidencia INN
                    similitudes_inn = []
                    for nombre_inn in INN_list:
                        if isinstance(nombre_inn, str):
                            caracteres_presentes = sum(1 for c in nombre_inn.lower() if c in nombre_generado.lower())
                            total_caracteres = len(nombre_inn)
                            porcentaje = (caracteres_presentes / total_caracteres) * 100
                            similitudes_inn.append((nombre_inn, caracteres_presentes, total_caracteres, porcentaje))
                    
                    # Ordenar por porcentaje y tomar el mejor
                    similitudes_inn.sort(key=lambda x: x[3], reverse=True)
                    mejor_inn = similitudes_inn[0] if similitudes_inn else ("-", 0, 0, 0)

                    # Calcular mejor coincidencia EMA
                    similitudes_ema = []
                    for nombre_ema in nombres_ema:
                        if isinstance(nombre_ema, str):
                            similitud = self.ortho.score(nombre_generado.lower(), nombre_ema.lower())
                            similitudes_ema.append((nombre_ema, similitud))
                    
                    # Ordenar por similitud y tomar el mejor
                    similitudes_ema.sort(key=lambda x: x[1], reverse=True)
                    mejor_ema = similitudes_ema[0] if similitudes_ema else ("-", 0)

                    # Calcular cadena común y otros valores
                    cadena_comun = self.encontrar_cadena_comun_mas_larga(nombre_generado, mejor_inn[0])
                    tiene_misma_cadena = "SÍ" if cadena_comun else "NO"
                    tam_cadena = len(cadena_comun) if cadena_comun else 0
                    tam_cadena_fraccion = f"{tam_cadena}/{len(mejor_inn[0])}"

                    # Calcular relevancia de la cadena
                    relevancia_porcentaje = (tam_cadena / len(mejor_inn[0])) * 100 if cadena_comun else 0
                    relevancia = f"{relevancia_porcentaje:.1f}% (low)" if relevancia_porcentaje <= 40 else f"{relevancia_porcentaje:.1f}% (high)"

                    # Verificar posición de la cadena común
                    if cadena_comun:
                        pos_nombre1 = nombre_generado.lower().find(cadena_comun.lower())
                        pos_nombre2 = mejor_inn[0].lower().find(cadena_comun.lower())
                        misma_posicion = "SÍ" if pos_nombre1 == pos_nombre2 else "NO"
                    else:
                        misma_posicion = "NO"

                    # Verificar INN stem
                    resultado_inn_stem, stem_encontrado = self.verificar_inn_stem(nombre_generado, self.cargar_inn_stems())
                    
                    # Nueva lógica para determinar si es NO APTO
                    es_no_apto = False
                    if relevancia_porcentaje > 40:  # Solo considerar posición si la relevancia es high
                        if misma_posicion == "SÍ":
                            es_no_apto = True
                    
                    # Verificar error cognitivo
                    if mejor_inn[1] >= mejor_inn[2] - 1:
                        es_no_apto = True

                    resultado_inn = "NO APTO" if es_no_apto else "APTO"
                    resultado_final = "APTO" if resultado_inn == "APTO" and resultado_inn_stem == "APTO" else "NO APTO"

                    # Actualizar la fila con los nuevos valores
                    nuevos_valores = (
                        nombre_generado,
                        mejor_inn[0],
                        f"{mejor_inn[3]:.2f}%",
                        f"{mejor_inn[1]}/{mejor_inn[2]}",
                        cadena_comun if cadena_comun else "-",
                        tiene_misma_cadena,
                        tam_cadena_fraccion,
                        relevancia,
                        misma_posicion,
                        resultado_inn,
                        stem_encontrado,
                        resultado_inn_stem,
                        resultado_final,
                        mejor_ema[0],
                        f"{mejor_ema[1]:.2f}%"
                    )
                    self.tree.item(item_id, values=nuevos_valores)

                    # Actualizar el estilo si cumple las condiciones
                    if resultado_inn == "APTO" and resultado_inn_stem == "APTO" and mejor_ema[1] < 70:
                        self.tree.item(item_id, tags=('resaltado',))
                    else:
                        self.tree.item(item_id, tags=())

                else:
                    # Restaurar los valores originales
                    self.tree.item(item_id, values=self.valores_originales[item_id])
                    del self.valores_originales[item_id]

        except Exception as e:
            messagebox.showerror("Error", f"Error al realizar fact check: {str(e)}")

    def cargar_inn_stems(self):
        try:
            if not self.archivo_inn_stem:
                messagebox.showerror("Error", "No se encontró el archivo INN stem.xlsx")
                return []
            df = pd.read_excel(self.archivo_inn_stem, header=None)
            return df[0].str.lower().tolist()
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar INN stem.xlsx: {str(e)}")
            return []

    def verificar_inn_stem(self, nombre_generado, stems_dict):
        """
        Verifica si el nombre generado tiene un stem en la lista de stems.
        """
        for stem in stems_dict:
            if stem in nombre_generado.lower():
                return "NO APTO", stem  # Retorna NO APTO y el stem encontrado
        return "APTO", "-"  # Si no se encuentra, retorna APTO

    def mostrar_resultados_comparar(self, resultados):
        # Limpiar el Treeview antes de mostrar nuevos resultados
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Insertar resultados en el Treeview
        for resultado in resultados:
            # Asegúrate de que 'resultado' tenga el mismo número de elementos que columnas
            self.tree.insert('', tk.END, values=resultado)

    def siguiente_funcion(self):
        """Abre la ventana de evaluación de riesgos"""
        try:
            # Crear una nueva ventana
            second_window = tk.Toplevel(self.root)
            second_window.title("Evaluación de Riesgos")
            
            # Configurar el tamaño de la ventana
            window_width = 1200
            window_height = 800
            screen_width = second_window.winfo_screenwidth()
            screen_height = second_window.winfo_screenheight()
            x = (screen_width - window_width) // 2
            y = (screen_height - window_height) // 2
            second_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
            
            # Inicializar el GUI de evaluación de riesgos
            risk_gui = RiskAssessmentGUI(second_window)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al abrir la ventana de evaluación de riesgos: {str(e)}")

class SecondGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Segunda Ventana")
        
        # Frame principal
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título
        ttk.Label(self.main_frame, text="Segunda Ventana", font=('Helvetica', 16)).grid(row=0, column=0, pady=10)
        
        # Botón para volver
        ttk.Button(self.main_frame, text="Volver", command=self.volver).grid(row=1, column=0, pady=10)
    
    def volver(self):
        self.root.destroy()

class RiskAssessment:
    def __init__(self):
        # Inicializar los diccionarios de criterios
        self.orthographic_criteria = {
            
            'identical_prefix': {'Print': 0, 'Speech': 0, 'Handwriting': 0},
            'identical_infix': {'Print': 0, 'Speech': 0, 'Handwriting': 0},
            'identical_suffix': {'Print': 0, 'Speech': 0, 'Handwriting': 0},
            'similar_length_of_name': {'Print': 0, 'Speech': 0, 'Handwriting': 0},
            'similar_spelling': {'Print': 0, 'Speech': 0, 'Handwriting': 0},
            'Upstrokes (capital and lower case e.g. P, d) in similar locations': {'Print': 0, 'Speech': 0, 'Handwriting': 0},
            'Downstrokes (e.g. q, y) in similar locations': {'Print': 0, 'Speech': 0, 'Handwriting': 0},
            'Cross-strokes (e.g. x, t) in similar locations': {'Print': 0, 'Speech': 0, 'Handwriting': 0},
            'Dotted letters (e.g. i) in similar locations': {'Print': 0, 'Speech': 0, 'Handwriting': 0},
            'Ambiguity introduced when scripting letters (e.g., P may appear as B, D, or R; lower case r; may appear as e, v or I; lower case a may appear as any vowel; lower case x may appear as lower case t, f or y etc.)': {'Print': 0, 'Speech': 0, 'Handwriting': 0},
            'Similar number of words/groups of characters in a name (A "word" is considered as any group of characters separated by a space)': {'Print': 0, 'Speech': 0, 'Handwriting': 0},
            'Similar number of syllables': {'Print': 0, 'Speech': 0, 'Handwriting': 0},
            'Similar stresses (e.g., Trycel and Triafil have similar stresses: TRY-cel and TRIA-fil; try-CEL and tria-FIL)': {'Print': 0, 'Speech': 0, 'Handwriting': 0},
            'Placement of vowel sounds is similar (e.g., e may sound like a or i; i may sound like a or e; a may sound like e or i etc.)': {'Print': 0, 'Speech': 0, 'Handwriting': 0},
            'Placement of consonant sound is similar (e.g., n may sound like m, dn, gn, kn, mn, pn; t may sound like d, b or pt etc.)': {'Print': 0, 'Speech': 0, 'Handwriting': 0},
            'First letter and/or sound (but made with the same letter) is identical': {'Print': 0, 'Speech': 0, 'Handwriting': 0},
            'Last letter is identical': {'Print': 0, 'Speech': 0, 'Handwriting': 0},
            'Same letters or same letters but for 1 letter but in different order (e.g., Termix and Trevisc - the "er" and "re" can be interpreted as the same and do not provide protection from name confusion)': {'Print': 0, 'Speech': 0, 'Handwriting': 0},
            'Some letters are written but not pronounced (silent letters)': {'Print': 0, 'Speech': 0, 'Handwriting': 0},
        }
        
        self.prescription_risk = {
            'Same therapeutic area or indication': {'YES': 0, 'NO': 0, 'N/A': 0, 'UNCLEAR': 0},
            'Same prescriber or intended H.C.P': {'YES': 0, 'NO': 0, 'N/A': 0, 'UNCLEAR': 0},
            'Close on electronic prescribing lists': {'YES': 0, 'NO': 0, 'N/A': 0, 'UNCLEAR': 0},
            'Handwritten prescriptions appear the same': {'YES': 0, 'NO': 0, 'N/A': 0, 'UNCLEAR': 0},
            'Same intended population': {'YES': 0, 'NO': 0, 'N/A': 0, 'UNCLEAR': 0},
            'Emergency situations': {'YES': 0, 'NO': 0, 'N/A': 0, 'UNCLEAR': 0}
        }
        
        self.dispensing_risk = {
            'Same storage conditions and proximity (e.g. shelf, fridge, controlled drugs locked cupboard, etc.)': {'YES': 0, 'NO': 0, 'N/A': 0, 'UNCLEAR': 0},
            'Close on electronic dispensing lists ': {'YES': 0, 'NO': 0, 'N/A': 0, 'UNCLEAR': 0},
            'Same dispensing facility (hospital pharmacy,'
            'community pharmacy, aseptic department, directly from ward stock, directly shipped' 
            'by manufacturer on patient named basis, etc.)': {'YES': 0, 'NO': 0, 'N/A': 0, 'UNCLEAR': 0},
            'Specialized controls the same packaging': {'YES': 0, 'NO': 0, 'N/A': 0, 'UNCLEAR': 0},
            'Emergency situations': {'YES': 0, 'NO': 0, 'N/A': 0, 'UNCLEAR': 0}
        }
        
        self.preparation_risk = {
            'Both to be mixed together prior to administration (e.g. error of dosing)?': {'YES': 0, 'NO': 0, 'N/A': 0, 'UNCLEAR': 0},
            'Can they both be put in a Monitored Dosage System (MDS)/Individualized dosing system?': {'YES': 0, 'NO': 0, 'N/A': 0, 'UNCLEAR': 0},
            'Complexity of handling the same or similar': {'YES': 0, 'NO': 0, 'N/A': 0, 'UNCLEAR': 0}
        }
        
        self.administration_risk = {
            'Self-administration in same patient population? (patient may confuse both products at home)': {'YES': 0, 'NO': 0, 'N/A': 0, 'UNCLEAR': 0},
            'Both Administered by H.C.P': {'YES': 0, 'NO': 0, 'N/A': 0, 'UNCLEAR': 0},
            'Emergency situations': {'YES': 0, 'NO': 0, 'N/A': 0, 'UNCLEAR': 0}
        }

class OrthographicAnalyzer:
    def __init__(self):
        self.nombre_generado = ""
        self.nombre_riesgo = ""
        
    def set_nombres(self, nombre_generado, nombre_riesgo):
        self.nombre_generado = nombre_generado.lower()
        self.nombre_riesgo = nombre_riesgo.lower()
    
    def check_tall_man_lettering(self):
        """Verifica si hay letras que podrían confundirse en mayúsculas"""
        confusable_pairs = [
            ('l', 'I'), ('O', '0'), ('Z', '2'), ('S', '5'),
            ('B', '8'), ('G', '6'), ('q', 'g'), ('n', 'h')
        ]
        for c1, c2 in confusable_pairs:
            if (c1.lower() in self.nombre_generado and c2.lower() in self.nombre_riesgo) or \
               (c2.lower() in self.nombre_generado and c1.lower() in self.nombre_riesgo):
                return True
        return False
    
    def check_similar_length(self):
        """Verifica si los nombres tienen longitud similar (diferencia ≤ 2)"""
        return abs(len(self.nombre_generado) - len(self.nombre_riesgo)) <= 2
    
    def check_similar_prefix(self):
        """Verifica si comparten prefijo (primeras 2-3 letras)"""
        min_len = min(len(self.nombre_generado), len(self.nombre_riesgo))
        prefix_length = min(3, min_len)
        return self.nombre_generado[:prefix_length] == self.nombre_riesgo[:prefix_length]
    
    def check_similar_suffix(self):
        """Verifica si comparten sufijo (al menos los últimos 2 caracteres en el mismo orden)"""
        # Obtener los últimos 2 caracteres de cada nombre
        min_suffix_length = 2
        gen_suffix = self.nombre_generado[-min_suffix_length:]
        risk_suffix = self.nombre_riesgo[-min_suffix_length:]
        
        # Verificar si los últimos 2 caracteres son iguales
        return gen_suffix == risk_suffix
    
    def check_similar_infix(self):
        """Verifica si comparten una secuencia interna de caracteres"""
        def get_substrings(text, min_length=3):
            return [text[i:j] for i in range(len(text)) 
                    for j in range(i + min_length, len(text) + 1)]
        
        gen_substrings = set(get_substrings(self.nombre_generado))
        risk_substrings = set(get_substrings(self.nombre_riesgo))
        return bool(gen_substrings & risk_substrings)
    
    def check_similar_sounds(self):
        """Verifica si tienen sonidos similares usando reglas fonéticas"""
        def get_phonetic(text):
            # Simplificación fonética básica
            replacements = [
                ('ph', 'f'), ('ae', 'e'), ('oe', 'e'),
                ('th', 't'), ('ch', 'k'), ('sh', 's'),
                ('ck', 'k'), ('cc', 'k'), ('qu', 'k'),
                ('x', 'ks'), ('v', 'b'), ('z', 's')
            ]
            result = text
            for old, new in replacements:
                result = result.replace(old, new)
            return result
        
        return get_phonetic(self.nombre_generado) == get_phonetic(self.nombre_riesgo)
    
    def check_similar_spelling(self):
        """Calcula la distancia de Levenshtein para similitud ortográfica"""
        def levenshtein_distance(s1, s2):
            if len(s1) < len(s2):
                return levenshtein_distance(s2, s1)
            if len(s2) == 0:
                return len(s1)
            previous_row = range(len(s2) + 1)
            for i, c1 in enumerate(s1):
                current_row = [i + 1]
                for j, c2 in enumerate(s2):
                    insertions = previous_row[j + 1] + 1
                    deletions = current_row[j] + 1
                    substitutions = previous_row[j] + (c1 != c2)
                    current_row.append(min(insertions, deletions, substitutions))
                previous_row = current_row
            return previous_row[-1]
        
        distance = levenshtein_distance(self.nombre_generado, self.nombre_riesgo)
        max_length = max(len(self.nombre_generado), len(self.nombre_riesgo))
        similarity = 1 - (distance / max_length)
        return similarity > 0.7  # 70% de similitud
    
    def check_similar_appearance(self):
        """Verifica similitud visual entre caracteres"""
        visual_similar = {
            'o': '0', 'l': '1', 'i': '1', 'z': '2', 
            'e': '3', 'a': '4', 's': '5', 'b': '8',
            'g': '9', 'q': 'g', 'n': 'h', 'u': 'v'
        }
        count = 0
        for c1, c2 in zip(self.nombre_generado, self.nombre_riesgo):
            if c1 == c2 or c2 == visual_similar.get(c1, '') or c1 == visual_similar.get(c2, ''):
                count += 1
        return count >= min(len(self.nombre_generado), len(self.nombre_riesgo)) * 0.7
    
    def check_similar_syllables(self):
        """Verifica si comparten sílabas similares"""
        def get_syllables(word):
            vowels = 'aeiou'
            syllables = []
            current = ''
            for i, char in enumerate(word):
                current += char
                if char in vowels and (i == len(word)-1 or word[i+1] not in vowels):
                    syllables.append(current)
                    current = ''
            return syllables
        
        gen_syllables = set(get_syllables(self.nombre_generado))
        risk_syllables = set(get_syllables(self.nombre_riesgo))
        return len(gen_syllables & risk_syllables) >= 1
    
    def check_similar_stress_pattern(self):
        """Verifica si tienen patrones de acentuación similares"""
        def get_stress_pattern(word):
            vowels = 'aeiou'
            pattern = ''
            for char in word:
                if char in vowels:
                    pattern += 'V'
                else:
                    pattern += 'C'
            return pattern
        
        return get_stress_pattern(self.nombre_generado) == get_stress_pattern(self.nombre_riesgo)
    
    def check_identical_prefix(self):
        """Verifica si comparten prefijo (primeras 2-3 letras)"""
        min_len = min(len(self.nombre_generado), len(self.nombre_riesgo))
        prefix_length = min(3, min_len)
        return self.nombre_generado[:prefix_length] == self.nombre_riesgo[:prefix_length]
    
    def check_identical_infix(self):
        """Verifica si comparten una secuencia interna de caracteres"""
        def get_substrings(text, min_length=3):
            return [text[i:j] for i in range(len(text)) 
                    for j in range(i + min_length, len(text) + 1)]
        
        gen_substrings = set(get_substrings(self.nombre_generado))
        risk_substrings = set(get_substrings(self.nombre_riesgo))
        return bool(gen_substrings & risk_substrings)
    
    def check_identical_suffix(self):
        """Verifica si comparten sufijo (al menos los últimos 2 caracteres en el mismo orden)"""
        min_suffix_length = 2
        gen_suffix = self.nombre_generado[-min_suffix_length:]
        risk_suffix = self.nombre_riesgo[-min_suffix_length:]
        return gen_suffix == risk_suffix
    
    def check_similar_spelling(self):
        """Calcula la distancia de Levenshtein para similitud ortográfica"""
        def levenshtein_distance(s1, s2):
            if len(s1) < len(s2):
                return levenshtein_distance(s2, s1)
            if len(s2) == 0:
                return len(s1)
            previous_row = range(len(s2) + 1)
            for i, c1 in enumerate(s1):
                current_row = [i + 1]
                for j, c2 in enumerate(s2):
                    insertions = previous_row[j + 1] + 1
                    deletions = current_row[j] + 1
                    substitutions = previous_row[j] + (c1 != c2)
                    current_row.append(min(insertions, deletions, substitutions))
                previous_row = current_row
            return previous_row[-1]
        
        distance = levenshtein_distance(self.nombre_generado, self.nombre_riesgo)
        max_length = max(len(self.nombre_generado), len(self.nombre_riesgo))
        similarity = 1 - (distance / max_length)
        return similarity > 0.7  # 70% de similitud

class RiskAssessmentGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Evaluación de Riesgos")
        
        # Establecer tamaño inicial de la ventana
        window_width = 1200
        window_height = 800
        self.root.geometry(f"{window_width}x{window_height}")
        
        self.risk_assessment = RiskAssessment()
        
        # Crear estilo para checkboxes
        self.style = ttk.Style()
        self.style.configure("Custom.TCheckbutton",
                           foreground="green",     # Color del tick
                           background="white",     # Color de fondo
                           indicatorcolor="green", # Color del indicador (tick)
                           indicatorrelief="solid", # Borde sólido para el box
                           borderwidth=2,         # Ancho del borde
                           relief="solid")        # Tipo de borde
        
        # Crear notebook para pestañas
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.analyzer = OrthographicAnalyzer()
        
        # Crear solo la pestaña TOTAL
        self.create_total_tab()

    def create_checkbox(self, parent, row, col, initial_value=False):
        var = tk.BooleanVar(value=False)  # Siempre empezar deseleccionado
        checkbox = ttk.Checkbutton(parent, variable=var, style="Custom.TCheckbutton")
        
        def on_left_click(event):
            # Toggle: si está vacío, seleccionar; si está seleccionado, deseleccionar
            var.set(not var.get())
            return "break"  # Prevenir el comportamiento por defecto
            
        # Vincular click izquierdo
        checkbox.bind('<Button-1>', on_left_click)
        
        # Deshabilitar el comportamiento normal del checkbox
        checkbox.config(takefocus=False)  # Evitar el foco del teclado
        checkbox.grid(row=row, column=col, padx=5, pady=2)
        return var

    def create_total_tab(self):
        # Crear frame con padding
        padding_frame = ttk.Frame(self.notebook)
        self.notebook.add(padding_frame, text="TOTAL")
        
        # Crear frame con scroll
        canvas = tk.Canvas(padding_frame)
        scrollbar = ttk.Scrollbar(padding_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        # Configurar scroll
        def on_mouse_wheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", on_mouse_wheel)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Grid para el canvas y scrollbar
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Configurar el grid del padding_frame
        padding_frame.grid_columnconfigure(0, weight=1)
        padding_frame.grid_rowconfigure(0, weight=1)
        
        current_row = 0
        
        # Crear los campos de entrada para los nombres
        ttk.Label(scrollable_frame, text="Nombre Generado:").grid(row=current_row, column=0, sticky='w', padx=5, pady=2)
        self.nombre_generado = tk.Text(scrollable_frame, height=1, width=20)
        self.nombre_generado.grid(row=current_row, column=1, sticky='w', padx=5, pady=2)
        
        current_row += 1
        ttk.Label(scrollable_frame, text="Nombre de Riesgo:").grid(row=current_row, column=0, sticky='w', padx=5, pady=2)
        self.nombre_riesgo = tk.Text(scrollable_frame, height=1, width=20)
        self.nombre_riesgo.grid(row=current_row, column=1, sticky='w', padx=5, pady=2)
        
        # Añadir botón EJECUTAR al lado de los campos de nombre
        self.ejecutar_button = ttk.Button(
            scrollable_frame,
            text="EJECUTAR",
            command=self.run_automatic_analysis,
            width=15
        )
        self.ejecutar_button.grid(
            row=current_row - 1,  # Misma fila que el primer nombre
            column=2,
            rowspan=2,  # Abarcar ambas filas de nombres
            padx=5,
            pady=5,
            sticky='ns'  # Centrado verticalmente
        )
        
        # Añadir botón de análisis automático justo después de los campos de nombre
        current_row += 1
        self.auto_analyze_button = ttk.Button(
            scrollable_frame,
            text="Analizar Automáticamente",
            command=self.run_automatic_analysis
        )
        self.auto_analyze_button.grid(
            row=current_row,
            column=0,
            columnspan=2,
            padx=5,
            pady=5,
            sticky='ew'
        )
        
        current_row += 1
        
        # Crear un canvas con scrollbar con altura específica
        canvas = tk.Canvas(padding_frame, height=700)
        scrollbar = ttk.Scrollbar(padding_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        # Guardar referencia al scrollable_frame
        self.scrollable_frame = scrollable_frame
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        # Crear ventana en el canvas con padding
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Configurar el grid del tab
        padding_frame.grid_rowconfigure(0, weight=1)
        padding_frame.grid_columnconfigure(0, weight=1)
        
        # Configurar el canvas y scrollbar usando grid
        canvas.grid(row=0, column=0, sticky='nsew')
        scrollbar.grid(row=0, column=1, sticky='ns')
        
        # Mejorar el comportamiento del scroll
        def on_mouse_wheel(event):
            # Hacer el scroll más suave
            canvas.yview_scroll(int(-1*(event.delta/60)), "units")
        
        # Bind the mouse wheel event
        canvas.bind_all("<MouseWheel>", on_mouse_wheel)
        
        # Asegurarse de que el canvas tenga suficiente espacio
        def configure_scroll_region(event):
            # Configurar la región de scroll para incluir todo el contenido más espacio extra
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Añadir espacio extra al final
            current_region = canvas.bbox("all")
            if current_region:
                canvas.configure(scrollregion=(current_region[0], current_region[1], 
                                             current_region[2], current_region[3] + 300))
        
        scrollable_frame.bind('<Configure>', configure_scroll_region)
        
        # Cuando se llegue al final del scroll, añadir más espacio
        def check_scroll_position(event):
            if canvas.yview()[1] >= 0.9:  # Si estamos cerca del final
                current_region = canvas.bbox("all")
                if current_region:
                    canvas.configure(scrollregion=(current_region[0], current_region[1], 
                                                 current_region[2], current_region[3] + 200))
        
        canvas.bind('<MouseWheel>', check_scroll_position)
        
        # Crear un frame contenedor con padding usando grid
        padding_frame = ttk.Frame(scrollable_frame)
        padding_frame.grid(row=0, column=0, padx=20, pady=(10, 200), sticky='nsew')
        
        # Configurar el grid del padding_frame
        padding_frame.grid_columnconfigure(0, weight=1)
        
        # A partir de aquí, usar padding_frame como el padre para todos los widgets
        current_row = 0
        
        # Añadir sección para los nombres antes de Orthographic and Phonetic
        ttk.Label(padding_frame, text="Nombres a evaluar", font=('Arial', 11, 'bold')).grid(
            row=current_row, column=0, columnspan=8, pady=10, sticky='w'
        )
        current_row += 1
        ttk.Button(padding_frame, text="EJECUTAR", command=self.run_automatic_analysis).grid(
            row=current_row, column=0, columnspan=8, pady=10, sticky='w'
        )
        current_row += 1

        # Nombre generado
        ttk.Label(padding_frame, text="Nombre generado:", font=('Arial', 10)).grid(
            row=current_row, column=0, columnspan=2, padx=5, pady=5, sticky='w'
        )
        self.nombre_generado = tk.Text(padding_frame, height=1, width=40)
        self.nombre_generado.grid(row=current_row, column=2, columnspan=4, padx=5, pady=5, sticky='w')
        current_row += 1

        # Nombre de riesgo
        ttk.Label(padding_frame, text="Nombre de riesgo:", font=('Arial', 10)).grid(
            row=current_row, column=0, columnspan=2, padx=5, pady=5, sticky='w'
        )
        self.nombre_riesgo = tk.Text(padding_frame, height=1, width=40)
        self.nombre_riesgo.grid(row=current_row, column=2, columnspan=4, padx=5, pady=5, sticky='w')
        current_row += 2  # Espacio extra antes de la siguiente sección

        # Sección Orthographic and Phonetic
        ttk.Label(padding_frame, text="Orthographic and Phonetic", font=('Arial', 11, 'bold')).grid(
            row=current_row, column=0, columnspan=8, pady=10, sticky='w'
        )
        current_row += 1
        
        # Diccionario para almacenar las variables de los checkboxes
        self.checkbox_vars = {'Print': [], 'Speech': [], 'Handwriting': []}
        
        # Encabezados para la sección ortográfica
        headers = ['#', 'Criterio', 'Print', 'Speech', 'Handwriting']
        for col, header in enumerate(headers):
            ttk.Label(padding_frame, text=header, font=('Arial', 10, 'bold')).grid(
                row=current_row, column=col, padx=5, pady=5
            )
        current_row += 1
        
        # Contenido de la sección ortográfica
        for idx, (criterion, values) in enumerate(self.risk_assessment.orthographic_criteria.items(), 1):
            # Número de fila
            ttk.Label(padding_frame, text=str(idx)).grid(
                row=current_row, column=0, padx=5, pady=2
            )
            # Nombre del criterio
            criterion_name = criterion.replace('_', ' ').title()
            # Ajustar el ancho de la etiqueta para permitir el ajuste de texto
            ttk.Label(padding_frame, text=criterion_name, wraplength=150).grid(  # Ajusta wraplength según sea necesario
                row=current_row, column=1, padx=5, pady=2, sticky='w'
            )
            # Checkboxes
            for col, category in enumerate(['Print', 'Speech', 'Handwriting'], start=2):
                if category == 'Print':
                    if idx in [1,2,3,4,5,6,7,8,9,11,16,17,18]:
                        var = self.create_checkbox(padding_frame, current_row, col, bool(values[category]))
                        self.checkbox_vars[category].append(var)
                    else:
                        ttk.Label(padding_frame, text="").grid(row=current_row, column=col, padx=5, pady=2)
                elif category == 'Speech':
                    if idx in [12, 13, 14, 15, 16, 19]:
                        var = self.create_checkbox(padding_frame, current_row, col, bool(values[category]))
                        self.checkbox_vars[category].append(var)
                    else:
                        ttk.Label(padding_frame, text="").grid(row=current_row, column=col, padx=5, pady=2)
                else:  # Handwriting
                    if idx in [1,2,3,4,5,6,7,8,9,10,11,16,17,18]:
                        var = self.create_checkbox(padding_frame, current_row, col, bool(values[category]))
                        self.checkbox_vars[category].append(var)
                    else:
                        ttk.Label(padding_frame, text="").grid(row=current_row, column=col, padx=5, pady=2)
            current_row += 1
        
        # Añadir fila TOTAL después de Silent Letters
        ttk.Label(padding_frame, text="TOTAL").grid(
            row=current_row, column=1, padx=5, pady=2, sticky='w'
        )
        
        # Labels para los totales y la interpretación
        total_labels = {}
        interpretation_labels = {}
        
        for col, category in enumerate(['Print', 'Speech', 'Handwriting'], start=2):
            total_labels[category] = ttk.Label(padding_frame, text="0")
            total_labels[category].grid(row=current_row, column=col, padx=5, pady=2)
        
        current_row += 1
        
        # Agregar "Interpretation of results"
        ttk.Label(padding_frame, text="Interpretation of results").grid(
            row=current_row, column=1, padx=5, pady=2, sticky='w'
        )
        
        # Labels para las interpretaciones
        for col, category in enumerate(['Print', 'Speech', 'Handwriting'], start=2):
            interpretation_labels[category] = ttk.Label(padding_frame, text="Low")
            interpretation_labels[category].grid(row=current_row, column=col, padx=5, pady=2)
        
        # Función para actualizar los totales y las interpretaciones
        def update_totals_and_interpretation():
            for col, category in enumerate(['Print', 'Speech', 'Handwriting'], start=2):
                total = sum(var.get() for var in self.checkbox_vars[category])
                total_labels[category].config(text=str(total))
                
                # Interpretación para Print
                if category == 'Print':
                    if total >= 9 and total <= 13:
                        interpretation = "High"
                    elif total >= 4 and total <= 8:
                        interpretation = "Medium"
                    else:  # 0-3
                        interpretation = "Low"
                    interpretation_labels[category].config(text=interpretation)
                # Interpretación para Speech
                elif category == 'Speech':
                    if total >= 5 and total <= 6:
                        interpretation = "High"
                    elif total >= 3 and total <= 4:
                        interpretation = "Medium"
                    else:  # 0-2
                        interpretation = "Low"
                    interpretation_labels[category].config(text=interpretation)
                # Interpretación para Handwriting
                elif category == 'Handwriting':
                    if total >= 10 and total <= 14:
                        interpretation = "High"
                    elif total >= 5 and total <= 9:
                        interpretation = "Medium"
                    else:  # 0-4
                        interpretation = "Low"
                    interpretation_labels[category].config(text=interpretation)
            
            padding_frame.after(100, update_totals_and_interpretation)
        
        # Iniciar la actualización
        update_totals_and_interpretation()
        
        current_row += 2
        
        # Sección Risk at Prescription level
        ttk.Label(padding_frame, text="Risk at Prescription level", font=('Arial', 11, 'bold')).grid(
            row=current_row, column=0, columnspan=8, pady=10, sticky='w'
        )
        current_row += 1
        
     
        
        # Diccionarios para almacenar las variables de los checkboxes y totales
        self.risk_vars = {
            'prescription': {'YES': [], 'NO': [], 'N/A': [], 'UNCLEAR': []},
            'dispensing': {'YES': [], 'NO': [], 'N/A': [], 'UNCLEAR': []},
            'preparation': {'YES': [], 'NO': [], 'N/A': [], 'UNCLEAR': []},
            'administration': {'YES': [], 'NO': [], 'N/A': [], 'UNCLEAR': []}
        }
        
        self.risk_total_labels = {
            'prescription': {},
            'dispensing': {},
            'preparation': {},
            'administration': {}
        }
        
        # Diccionario para almacenar los textboxes de interpretación
        self.interpretation_texts = {}

        # Función helper para agregar secciones de riesgo
        def add_risk_section(title, risk_dict):
            nonlocal current_row
            section_name = ""
            
            # Determinar qué sección estamos procesando
            if title == "":
                section_name = 'prescription'
            elif "Dispensing" in title:
                section_name = 'dispensing'
            elif "preparation" in title:
                section_name = 'preparation'
            elif "administration" in title:
                section_name = 'administration'

            if title:
                ttk.Label(padding_frame, text=title, font=('Arial', 11, 'bold')).grid(
                    row=current_row, column=0, columnspan=8, pady=10, sticky='w'
                )
                current_row += 1
                
                # Añadir encabezados con numeración para todas las secciones de riesgo
                headers = ['#', 'Criterio', 'YES', 'NO', 'N/A', 'UNCLEAR']
                for col, header in enumerate(headers):
                    ttk.Label(padding_frame, text=header, font=('Arial', 10, 'bold')).grid(
                        row=current_row, column=col, padx=5, pady=5
                    )
                current_row += 1
                
                # Añadir filas numeradas
                for idx, (criterion, values) in enumerate(risk_dict.items(), 1):
                    # Número de fila
                    ttk.Label(padding_frame, text=str(idx)).grid(
                        row=current_row, column=0, padx=5, pady=2
                    )
                    criterion_name = criterion.replace('_', ' ').title()
                    # Ajustar el ancho de la etiqueta para permitir el ajuste de texto
                    ttk.Label(padding_frame, text=criterion_name, wraplength=300).grid(  # Ajusta wraplength según sea necesario
                        row=current_row, column=1, padx=5, pady=2, sticky='w'
                    )
                    for col, category in enumerate(['YES', 'NO', 'N/A', 'UNCLEAR'], start=2):
                        var = self.create_checkbox(padding_frame, current_row, col, bool(values.get(category, 0)))
                        self.risk_vars[section_name][category].append(var)
                    current_row += 1
                
                # Añadir fila Total
                ttk.Label(padding_frame, text="Total", font=('Arial', 10, 'bold')).grid(
                    row=current_row, column=1, padx=5, pady=2, sticky='w'
                )
                
                # Labels para los totales
                for col, category in enumerate(['YES', 'NO', 'N/A', 'UNCLEAR'], start=2):
                    self.risk_total_labels[section_name][category] = ttk.Label(padding_frame, text="0", font=('Arial', 10))
                    self.risk_total_labels[section_name][category].grid(row=current_row, column=col, padx=5, pady=2)
                current_row += 1
                
                # Añadir fila Interpretación con textbox
                ttk.Label(padding_frame, text="Interpretación", font=('Arial', 10, 'bold')).grid(
                    row=current_row, column=1, padx=5, pady=2, sticky='w'
                )
                
                # Crear y posicionar el textbox
                interpretation_text = tk.Text(padding_frame, height=2, width=50)
                interpretation_text.grid(row=current_row, column=2, columnspan=4, padx=5, pady=2, sticky='w')
                
                def on_return(event):
                    # Desactivar la edición y hacer que el textbox se camufle completamente
                    interpretation_text.configure(
                        state='disabled',  # Hacer no editable
                        bg=padding_frame.cget('background'),  # Color de fondo igual al frame
                        relief='flat',  # Quitar bordes
                        highlightthickness=0,  # Quitar resaltado
                        insertwidth=0,  # Ocultar cursor
                        borderwidth=0,  # Eliminar el borde
                        selectbackground=padding_frame.cget('background'),  # Color de selección igual al fondo
                        selectforeground='black'  # Color del texto seleccionado
                    )
                    return "break"  # Prevenir el salto de línea por defecto
                
                def on_click(event):
                    # Reactivar la edición y restaurar apariencia
                    interpretation_text.configure(
                        state='normal',  # Hacer editable
                        bg='white',  # Fondo blanco
                        relief='sunken',  # Borde hundido
                        highlightthickness=1,  # Resaltado normal
                        insertwidth=2  # Mostrar cursor
                    )
                
                # Vincular eventos
                interpretation_text.bind('<Return>', on_return)  # Enter para confirmar
                interpretation_text.bind('<Button-1>', on_click)  # Click para editar
                
                self.interpretation_texts[section_name] = interpretation_text
                
                current_row += 2  # Espacio extra después de cada sección
            else:
                # Para la sección de prescription (sin título)
                headers = ['#', 'Criterio', 'YES', 'NO', 'N/A', 'UNCLEAR']
                for col, header in enumerate(headers):
                    ttk.Label(padding_frame, text=header, font=('Arial', 10, 'bold')).grid(
                        row=current_row, column=col, padx=5, pady=5
                    )
                current_row += 1
                
                for idx, (criterion, values) in enumerate(risk_dict.items(), 1):
                    # Número de fila
                    ttk.Label(padding_frame, text=str(idx)).grid(
                        row=current_row, column=0, padx=5, pady=2
                    )
                    criterion_name = criterion.replace('_', ' ').title()
                    ttk.Label(padding_frame, text=criterion_name).grid(
                        row=current_row, column=1, padx=5, pady=2, sticky='w'
                    )
                    for col, category in enumerate(['YES', 'NO', 'N/A', 'UNCLEAR'], start=2):
                        var = self.create_checkbox(padding_frame, current_row, col, bool(values.get(category, 0)))
                        self.risk_vars[section_name][category].append(var)
                    current_row += 1
                
                # Añadir fila Total
                ttk.Label(padding_frame, text="Total", font=('Arial', 10, 'bold')).grid(
                    row=current_row, column=1, padx=5, pady=2, sticky='w'
                )
                
                # Labels para los totales
                for col, category in enumerate(['YES', 'NO', 'N/A', 'UNCLEAR'], start=2):
                    self.risk_total_labels[section_name][category] = ttk.Label(padding_frame, text="0", font=('Arial', 10))
                    self.risk_total_labels[section_name][category].grid(row=current_row, column=col, padx=5, pady=2)
                current_row += 1
                
                # Añadir fila Interpretación con textbox
                ttk.Label(padding_frame, text="Interpretación", font=('Arial', 10, 'bold')).grid(
                    row=current_row, column=1, padx=5, pady=2, sticky='w'
                )
                
                # Crear y posicionar el textbox
                interpretation_text = tk.Text(padding_frame, height=2, width=50)
                interpretation_text.grid(row=current_row, column=2, columnspan=4, padx=5, pady=2, sticky='w')
                
                def on_return(event):
                    # Desactivar la edición y hacer que el textbox se camufle completamente
                    interpretation_text.configure(
                        state='disabled',  # Hacer no editable
                        bg=padding_frame.cget('background'),  # Color de fondo igual al frame
                        relief='flat',  # Quitar bordes
                        highlightthickness=0,  # Quitar resaltado
                        insertwidth=0,  # Ocultar cursor
                        borderwidth=0,  # Eliminar el borde
                        selectbackground=padding_frame.cget('background'),  # Color de selección igual al fondo
                        selectforeground='black'  # Color del texto seleccionado
                    )
                    return "break"  # Prevenir el salto de línea por defecto
                
                def on_click(event):
                    # Reactivar la edición y restaurar apariencia
                    interpretation_text.configure(
                        state='normal',  # Hacer editable
                        bg='white',  # Fondo blanco
                        relief='sunken',  # Borde hundido
                        highlightthickness=1,  # Resaltado normal
                        insertwidth=2  # Mostrar cursor
                    )
                
                # Vincular eventos
                interpretation_text.bind('<Return>', on_return)  # Enter para confirmar
                interpretation_text.bind('<Button-1>', on_click)  # Click para editar
                
                self.interpretation_texts[section_name] = interpretation_text
                
                current_row += 2  # Espacio extra después de cada sección
        
        # Agregar todas las secciones de riesgo
        add_risk_section("", self.risk_assessment.prescription_risk)
        add_risk_section("Risk at Dispensing level", self.risk_assessment.dispensing_risk)
        add_risk_section("Risk at preparation", self.risk_assessment.preparation_risk)
        add_risk_section("Risk at administration", self.risk_assessment.administration_risk)
        
        # Función para actualizar todos los totales
        def update_risk_totals():
            for section_name in ['prescription', 'dispensing', 'preparation', 'administration']:
                for category in ['YES', 'NO', 'N/A', 'UNCLEAR']:
                    total = sum(1 for var in self.risk_vars[section_name][category] if var.get())
                    self.risk_total_labels[section_name][category].config(text=str(total))
            padding_frame.after(100, update_risk_totals)
        
        # Iniciar la actualización de totales
        update_risk_totals()
        
        # Añadir sección ELEMENTS
        ttk.Label(padding_frame, text="ELEMENTS that may increase/reduce the risk of confusion", font=('Arial', 11, 'bold')).grid(
            row=current_row, column=0, columnspan=8, pady=10, sticky='w'
        )
        current_row += 1
        
        # Encabezados para la tabla ELEMENTS
        elements_headers = [
            'SAME (Insert an X if this applies)',
            'SIMILAR (Insert an X if this applies)',
            'DIFFERENT (Insert an X if this applies)',
            'N/A (Insert an X if this applies)'
        ]
        
        # Añadir encabezados
        ttk.Label(padding_frame, text="Criterio", font=('Arial', 10, 'bold')).grid(
            row=current_row, column=1, padx=5, pady=5, sticky='w'
        )
        for col, header in enumerate(elements_headers, start=2):
            ttk.Label(padding_frame, text=header, font=('Arial', 10, 'bold')).grid(
                row=current_row, column=col, padx=5, pady=5
            )
        current_row += 1
        
        # Filas de la tabla ELEMENTS
        elements_rows = [
            "Strengths",
            "Pharmaceutical forms",
            "Route of administration",
            "Legal Status",
            "Proposed Labeling",
            "Totals"
        ]
        
        # Diccionario para almacenar las variables de los checkboxes de ELEMENTS
        self.elements_vars = {
            'SAME': [],
            'SIMILAR': [],
            'DIFFERENT': [],
            'N/A': []
        }
        
        # Diccionario para almacenar los labels de totales
        self.elements_total_labels = {}
        
        # Crear las filas con checkboxes
        for idx, row_text in enumerate(elements_rows):
            # Primera columna (criterio)
            ttk.Label(padding_frame, text=row_text, font=('Arial', 10)).grid(
                row=current_row, column=1, padx=5, pady=2, sticky='w'
            )
            
            if row_text == "Totals":
                # Para la fila de totales, crear labels
                for col, category in enumerate(['SAME', 'SIMILAR', 'DIFFERENT', 'N/A'], start=2):
                    total_label = ttk.Label(padding_frame, text="0")
                    total_label.grid(row=current_row, column=col, padx=5, pady=2)
                    self.elements_total_labels[category] = total_label
            else:
                # Para las demás filas, crear checkboxes
                for col, category in enumerate(['SAME', 'SIMILAR', 'DIFFERENT', 'N/A'], start=2):
                    var = self.create_checkbox(padding_frame, current_row, col)
                    self.elements_vars[category].append(var)
            
            current_row += 1
        
        # Función para actualizar los totales de ELEMENTS
        def update_elements_totals():
            for category in ['SAME', 'SIMILAR', 'DIFFERENT', 'N/A']:
                total = sum(1 for var in self.elements_vars[category] if var.get())
                self.elements_total_labels[category].config(text=str(total))
            padding_frame.after(100, update_elements_totals)
        
        # Iniciar la actualización de totales de ELEMENTS
        update_elements_totals()
        
        # Añadir fila de interpretación para ELEMENTS
        ttk.Label(padding_frame, text="Interpretación", font=('Arial', 10, 'bold')).grid(
            row=current_row, column=1, padx=5, pady=2, sticky='w'
        )
        
        # Crear y posicionar el textbox de interpretación
        elements_interpretation_text = tk.Text(padding_frame, height=5, width=60)
        elements_interpretation_text.grid(row=current_row, column=2, columnspan=4, padx=5, pady=2, sticky='w')
        
        # Crear el botón "Exportar"
        export_button = ttk.Button(padding_frame, text="Exportar", command=self.exportar_interpretacion)
        export_button.grid(row=current_row + 1, column=3, padx=5, pady=(2, 10), sticky='e')
        
        current_row += 2  # Espacio extra antes de la siguiente tabla
        
        # Añadir sección Potential for Harm
        ttk.Label(padding_frame, text="Potential for Harm in case of accidental mix-up", font=('Arial', 11, 'bold')).grid(
            row=current_row, column=0, columnspan=8, pady=10, sticky='w'
        )
        current_row += 1
        
        # Crear la tabla de Potential for Harm
        potential_harm_data = [
            ("High", "#ffcccc", "e.g., death or major injury"),  # Rojo claro
            ("Medium", "#ffffcc", "e.g., minor injury"),         # Amarillo claro
            ("Low", "#ccffcc", "e.g., no injury"),              # Verde claro
            ("N/A", "#f0f0f0", "e.g., no risk of confusion identified"),  # Gris muy claro
            ("Unknown", "#f0f0f0", "e.g., when the actual potential for harm is unknown"),  # Gris muy claro
        ]
        
        # Crear las celdas con diferentes colores
        for text, bg_color, description in potential_harm_data:
            # Primera columna (nivel)
            label = tk.Label(padding_frame, text=text, bg=bg_color)
            label.grid(row=current_row, column=1, sticky='w', padx=5, pady=2)
            
            # Segunda columna (descripción)
            desc_label = tk.Label(padding_frame, text=description, bg=bg_color)
            desc_label.grid(row=current_row, column=2, sticky='w', padx=5, pady=2)
            
            # Tercera columna (textbox)
            if text == "Response:":
                # Para la fila de respuesta, crear un Text widget más grande
                text_widget = tk.Text(padding_frame, height=2, width=30)
            else:
                text_widget = tk.Text(padding_frame, height=1, width=30)
            
            text_widget.grid(row=current_row, column=3, sticky='w', padx=5, pady=2)
            text_widget.configure(bg=bg_color)
            
            def on_return_text(event, widget=text_widget, color=bg_color):
                widget.configure(
                    state='disabled',
                    bg=color,
                    relief='flat',
                    highlightthickness=0,
                    insertwidth=0,
                    borderwidth=0,
                    selectbackground=color,
                    selectforeground='black'
                )
                return "break"
            
            def on_click_text(event, widget=text_widget, color=bg_color):
                widget.configure(
                    state='normal',
                    relief='solid',
                    highlightthickness=1,
                    insertwidth=2,
                    borderwidth=1
                )
            
            text_widget.bind('<Return>', on_return_text)
            text_widget.bind('<Button-1>', on_click_text)
            
            current_row += 1
        
        # Después de la sección "Potential for Harm", añadir la fila "Tras revisión"
        current_row += 1
        ttk.Label(padding_frame, text="Tras revisión", font=('Arial', 11, 'bold')).grid(
            row=current_row, column=0, columnspan=8, pady=(30, 10), sticky='w'
        )
        current_row += 1

        # Crear frame contenedor para el textbox y su scrollbar
        review_frame = ttk.Frame(padding_frame)
        review_frame.grid(row=current_row, column=0, columnspan=8, sticky='nsew', padx=5, pady=5)
        review_frame.grid_columnconfigure(0, weight=1)  # Para que el textbox se expanda horizontalmente

        # Crear el textbox con scrollbar
        self.final_review_text = tk.Text(
            review_frame, 
            height=20,          # Aumentado para más contenido visible
            width=120,          # Aumentado para usar más espacio horizontal
            state='normal',     # Cambiar a 'normal' para permitir la edición
            wrap=tk.WORD,      # Para que el texto se ajuste automáticamente
            font=('Arial', 10)  # Fuente más legible
        )
        review_scrollbar = ttk.Scrollbar(review_frame, orient="vertical", command=self.final_review_text.yview)
        self.final_review_text.configure(yscrollcommand=review_scrollbar.set)

        # Colocar el textbox y su scrollbar
        self.final_review_text.grid(row=0, column=0, sticky='nsew')
        review_scrollbar.grid(row=0, column=1, sticky='ns')

        # Función para manejar el scroll del textbox
        def on_mousewheel_text(event):
            self.final_review_text.yview_scroll(int(-1*(event.delta/120)), "units")
            return "break"  # Prevenir que el evento se propague

        # Función para manejar el scroll de la ventana
        def on_mousewheel_canvas(event):
            self.root.yview_scroll(int(-1*(event.delta/120)), "units")
            return "break"  # Prevenir que el evento se propague

        # Vincular el evento de la rueda del ratón directamente al textbox
        self.final_review_text.bind("<MouseWheel>", on_mousewheel_text)

        # Desvincular el evento de la rueda del ratón cuando el ratón sale del textbox
        def on_enter(event):
            self.final_review_text.bind("<MouseWheel>", on_mousewheel_text)  # Vincular el scroll del textbox

        # Revincular el evento de la rueda del ratón cuando el ratón sale del textbox
        def on_leave(event):
            self.final_review_text.unbind("<MouseWheel>")  # Desvincular el scroll del textbox

        # Vincular eventos de entrada/salida del ratón
        self.final_review_text.bind("<Enter>", on_enter)
        self.final_review_text.bind("<Leave>", on_leave)

        # Añadir botón de interpretación
        interpret_button = ttk.Button(
            padding_frame, 
            text="Interpretación de Resultados",
            command=self.generate_interpretation
        )
        interpret_button.grid(
            row=current_row + 1,  # Mover el botón debajo del textbox
            column=0,
            columnspan=8,
            padx=5,
            pady=10,
            sticky='ew'  # Hacer que el botón ocupe todo el ancho
        )
        export_button = ttk.Button(
            padding_frame, 
            text="Exportar", 
            command=self.exportar_interpretacion
        )
        export_button.grid(
            row=current_row + 1, 
            column=0, 
            columnspan=2,
            padx=20, 
            pady=(70),  # Aumentar el valor para bajar el botón
            sticky='e'
        )  # Ajustar la columna a 3 para moverlo a la derecha
        

        # Añadir un frame vacío al final para dar más espacio
        spacing_frame = ttk.Frame(padding_frame, height=100)
        spacing_frame.grid(row=current_row + 2, column=0, columnspan=8, sticky='ew')
        spacing_frame.grid_propagate(False)

    def interpret_risk_level(self, section_name):
        # Obtener totales
        yes_count = int(self.risk_total_labels[section_name]['YES'].cget('text'))
        no_count = int(self.risk_total_labels[section_name]['NO'].cget('text'))
        total_questions = len(self.risk_vars[section_name]['YES'])
        
        # Calcular total_risk_factors
        total_risk_factors = sum(1 for var in self.risk_vars[section_name]['YES'] if var.get())  # Ajusta según tu lógica
        
        # Evaluar riesgo inicial
        risk_level = "Indeterminado"
        
        # Criterios de evaluación:
        # Alto riesgo: >60% YES o SAME>3
        # Bajo riesgo: >60% NO o DIFFERENT>3
        if (yes_count / total_questions) * 100 >= 60 or sum(var.get() for var in self.elements_vars['SAME']) >= 3:
            risk_level = "Alto riesgo de confusión"
            details = [
                "Factores de riesgo detectados:",
                f"- {yes_count}/{total_questions} respuestas YES ({yes_count / total_questions * 100:.1f}%)",
                f"- {sum(var.get() for var in self.elements_vars['SAME'])} elementos SAME vs {sum(var.get() for var in self.elements_vars['DIFFERENT'])} DIFFERENT"
            ]
        elif (no_count / total_questions) * 100 >= 60 or sum(var.get() for var in self.elements_vars['DIFFERENT']) >= 3:
            risk_level = "Bajo riesgo de confusión"
            details = [
                "Factores protectores detectados:",
                f"- {no_count}/{total_questions} respuestas NO ({no_count / total_questions * 100:.1f}%)",
                f"- {sum(var.get() for var in self.elements_vars['DIFFERENT'])} elementos DIFFERENT vs {sum(var.get() for var in self.elements_vars['SAME'])} SAME"
            ]
        else:
            risk_level = "Riesgo medio de confusión"
            details = [
                "Factores mixtos detectados:",
                f"- YES: {yes_count}/{total_questions} ({yes_count / total_questions * 100:.1f}%)",
                f"- NO: {no_count}/{total_questions} ({no_count / total_questions * 100:.1f}%)",
                f"- SAME: {sum(var.get() for var in self.elements_vars['SAME'])}, SIMILAR: {sum(var.get() for var in self.elements_vars['SIMILAR'])}, DIFFERENT: {sum(var.get() for var in self.elements_vars['DIFFERENT'])}"
            ]
        
        # Actualizar el texto de interpretación
        if section_name in self.interpretation_texts:
            text_widget = self.interpretation_texts[section_name]
            text_widget.delete('1.0', tk.END)
            text_widget.insert('1.0', f"{risk_level}\n")
            for detail in details:
                text_widget.insert(tk.END, f"{detail}\n")

    def update_risk_totals(self):
        for section_name in ['prescription', 'dispensing', 'preparation', 'administration']:
            for category in ['YES', 'NO', 'N/A', 'UNCLEAR']:
                total = sum(1 for var in self.risk_vars[section_name][category] if var.get())
                self.risk_total_labels[section_name][category].config(text=str(total))
            # Añadir interpretación automática
            self.interpret_risk_level(section_name)
        self.root.after(100, self.update_risk_totals)

    def generate_interpretation(self):
        self.final_review_text.config(state='normal')
        
        # Recopilar datos
        data = {
            'orthographic': {
                'Print': sum(var.get() for var in self.checkbox_vars['Print']),
                'Speech': sum(var.get() for var in self.checkbox_vars['Speech']),
                'Handwriting': sum(var.get() for var in self.checkbox_vars['Handwriting'])
            },
            'sections': {},
            'elements': {
                'SAME': sum(var.get() for var in self.elements_vars['SAME']),
                'SIMILAR': sum(var.get() for var in self.elements_vars['SIMILAR']),
                'DIFFERENT': sum(var.get() for var in self.elements_vars['DIFFERENT'])
            }
        }
        
        # Recopilar datos de cada sección
        for section in ['prescription', 'dispensing', 'preparation', 'administration']:
            data['sections'][section] = {
                'YES': int(self.risk_total_labels[section]['YES'].cget('text')),
                'NO': int(self.risk_total_labels[section]['NO'].cget('text')),
                'N/A': int(self.risk_total_labels[section]['N/A'].cget('text'))
            }

        # Obtener los nombres ingresados
        nombre_generado = self.nombre_generado.get("1.0", "end-1c").strip()
        nombre_riesgo = self.nombre_riesgo.get("1.0", "end-1c").strip()
        
        # Configurar OpenAI
        openai.api_key = openai_api_key
        
        # Crear el prompt para ChatGPT
        prompt = f"""Actúa como un experto del Comité de Medicamentos de Uso Humano (CHMP) de la EMA.
Analiza los siguientes datos de evaluación de riesgo de confusión entre los nombres de medicamentos y estructura tu respuesta en las siguientes secciones numeradas:

NOMBRES EVALUADOS:
- Nombre propuesto: {nombre_generado if nombre_generado else "[No especificado]"}
- Nombre existente con riesgo de confusión: {nombre_riesgo if nombre_riesgo else "[No especificado]"}

1. ANÁLISIS ORTOGRÁFICO:
- Print: {data['orthographic']['Print']} criterios identificados
- Speech: {data['orthographic']['Speech']} criterios identificados
- Handwriting: {data['orthographic']['Handwriting']} criterios identificados

Reglas de evaluación ortográfica:
Print: 
- High: 9-13 criterios
- Medium: 4-8 criterios
- Low: 0-3 criterios

Speech:
- High: 5-6 criterios
- Medium: 3-4 criterios
- Low: 0-2 criterios

Handwriting:
- High: 10-14 criterios
- Medium: 5-9 criterios
- Low: 0-4 criterios

IMPORTANTE: Si alguno de los criterios (Print, Speech o Handwriting) resulta en Medium o High, debemos continuar con el análisis.

2. ANÁLISIS POR NIVELES (solo si alguno es Medium o High):
Prescripción:
- YES: {data['sections']['prescription']['YES']}
- NO: {data['sections']['prescription']['NO']}
- N/A: {data['sections']['prescription']['N/A']}

Dispensación:
- YES: {data['sections']['dispensing']['YES']}
- NO: {data['sections']['dispensing']['NO']}
- N/A: {data['sections']['dispensing']['N/A']}

Preparación:
- YES: {data['sections']['preparation']['YES']}
- NO: {data['sections']['preparation']['NO']}
- N/A: {data['sections']['preparation']['N/A']}

Administración:
- YES: {data['sections']['administration']['YES']}
- NO: {data['sections']['administration']['NO']}
- N/A: {data['sections']['administration']['N/A']}

3. ELEMENTOS DE SIMILITUD (factor modificador del riesgo ortográfico):
- SAME: {data['elements']['SAME']}
- SIMILAR: {data['elements']['SIMILAR']}
- DIFFERENT: {data['elements']['DIFFERENT']}

Criterios generales de evaluación:
- Un alto número de NO más un mayor número de DIFFERENT vs SAME/SIMILAR más un nivel LOW/N/A de Mix-up sugiere menor probabilidad de confusión.
- Un alto número de YES más un mayor número de SAME vs SIMILAR/DIFFERENT más un nivel MEDIUM/HIGH de Mix-up sugiere mayor probabilidad de confusión.

Por favor, proporciona:
1. Análisis inicial del riesgo ortográfico:
   - Determina si es necesario continuar con más análisis
   - Indica si los ELEMENTS modifican el nivel de riesgo ortográfico

2. Solo si es necesario (si alguno es Medium o High):
   - Análisis detallado de cada sección
   - Identificación de factores que aumentan o reducen el riesgo
   - Evaluación del nivel de riesgo global

3. Conclusión final:
   - Recomendación clara sobre si el nombre puede registrarse
   - Justificación de la decisión
   - Sugerencias de medidas de mitigación si fueran necesarias

Usa un lenguaje claro y accesible, explicando paso a paso el proceso de evaluación.
"""

        try:
            # Llamar a la API de ChatGPT
            response = openai.ChatCompletion.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "Eres un experto del CHMP de la EMA evaluando el riesgo de confusión en nombres de medicamentos."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
                max_tokens=2000
            )
            
            # Obtener la respuesta
            analysis = response.choices[0].message.content
            
        except Exception as e:
            analysis = f"""Error al conectar con el servicio de análisis.
Por favor, revise su conexión a internet o inténtelo más tarde.
Error específico: {str(e)}"""
        
        # Actualizar el textbox
        self.final_review_text.delete('1.0', tk.END)
        self.final_review_text.insert('1.0', analysis)
        self.final_review_text.config(state='disabled')

    def run_automatic_analysis(self):
        # Obtener nombres de los textboxes
        nombre_generado = self.nombre_generado.get("1.0", "end-1c").strip()
        nombre_riesgo = self.nombre_riesgo.get("1.0", "end-1c").strip()
        
        if not nombre_generado or not nombre_riesgo:
            messagebox.showwarning("Advertencia", "Por favor, ingrese ambos nombres para analizar")
            return
        
        # Configurar el analizador
        self.analyzer.set_nombres(nombre_generado, nombre_riesgo)
        
        # Ejecutar análisis y actualizar checkboxes
        checks = {
            'identical_prefix': self.analyzer.check_identical_prefix(),
            'identical_infix': self.analyzer.check_identical_infix(),
            'identical_suffix': self.analyzer.check_identical_suffix(),
            'similar_length_of_name': self.analyzer.check_similar_length(),
            'similar_spelling': self.analyzer.check_similar_spelling(),
            # Agrega más funciones según sea necesario
        }
        
        # Actualizar los checkboxes según los resultados
        for criterion, result in checks.items():
            if criterion in self.risk_assessment.orthographic_criteria:
                idx = list(self.risk_assessment.orthographic_criteria.keys()).index(criterion)
                
                # Print column
                if idx < len(self.checkbox_vars['Print']):
                    self.checkbox_vars['Print'][idx].set(result)
                
                # Speech column - activar si hay similitud fonética
                if idx < len(self.checkbox_vars['Speech']):
                    speech_result = result  # Mantener el resultado original para criterios fonéticos
                    self.checkbox_vars['Speech'][idx].set(speech_result)
                
                # Handwriting column - activar si hay similitud visual
                if idx < len(self.checkbox_vars['Handwriting']):
                    handwriting_result = result  # Mantener el resultado original para criterios visuales
                    self.checkbox_vars['Handwriting'][idx].set(handwriting_result)

    def exportar_interpretacion(self):
        """Método para exportar la interpretación a un archivo Word"""
        nombre_generado = self.nombre_generado.get("1.0", "end-1c").strip()
        interpretation = self.final_review_text.get("1.0", "end-1c")
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Document files", "*.docx")],
            title="Guardar archivo como"
        )
        
        if file_path:
            doc = Document()
            
            # Título principal
            title = doc.add_heading(nombre_generado, level=1)
            title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
            # Secciones numeradas
            doc.add_heading('1. Orthographic and Phonetic Analysis', level=2)
            doc.add_paragraph('Assessment of orthographic, phonetic and handwriting similarities:')
            
            doc.add_heading('2. Risk Assessment by Level', level=2)
            doc.add_paragraph('Analysis of risks at different stages:')
            doc.add_paragraph('• Prescription level\n• Dispensing level\n• Preparation level\n• Administration level', style='List Bullet')
            
            doc.add_heading('3. Elements Analysis', level=2)
            doc.add_paragraph('Evaluation of elements that may increase/reduce confusion risk:')
            
            doc.add_heading('4. Potential for Harm', level=2)
            doc.add_paragraph('Assessment of potential consequences in case of mix-up:')
            
            # Contenido principal con formato
            doc.add_heading('Detailed Assessment', level=2)
            p = doc.add_paragraph(interpretation)
            p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT  # Alineación a la izquierda
            
            # Establecer la fuente y el tamaño
            for run in p.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)
            
            if not p.runs:
                run = p.add_run(interpretation)
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)
            
            # Guardar el documento
            doc.save(file_path)
            messagebox.showinfo("Exportar", f"Interpretación exportada a '{file_path}'.")

def iniciar_aplicacion(root):
    app = ComparisonForm(root)
    return app

if __name__ == "__main__":
    root = tk.Tk()
    app = iniciar_aplicacion(root)
    root.mainloop()

