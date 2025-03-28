import tkinter as tk
from tkinter import ttk, messagebox
from login import Autenticacion
from consultas import ConsultasExcel

class ConsultasApp:
    def __init__(self, root, usuario):
        self.root = root
        self.usuario = usuario
        self.consultas = ConsultasExcel()
        self.columnas_consulta = {}
        self.opciones_filtro = {}
        
        # Configuración de la ventana
        self.root.title(f"Sistema de Consultas - {usuario}")
        self.root.geometry("1100x650")  # Tamaño ligeramente reducido
        self.root.configure(bg='#f0f0f0')
        
        # Estilos mejorados
        self.estilo = ttk.Style()
        self.estilo.theme_use('clam')  # Tema más moderno
        self.estilo.configure('TFrame', background='#f0f0f0')
        self.estilo.configure('TButton', font=('Arial', 10), padding=5, width=20)
        self.estilo.configure('TLabel', background='#f0f0f0', font=('Arial', 10))
        self.estilo.configure('Header.TLabel', font=('Arial', 12, 'bold'))
        self.estilo.configure('Treeview', font=('Arial', 9), rowheight=25)
        self.estilo.configure('Treeview.Heading', font=('Arial', 10, 'bold'))
        self.estilo.map('Treeview', background=[('selected', '#0078d7')])
        
        self.crear_interfaz()
        self.configurar_consultas()
        self.cargar_opciones_filtro()

    def configurar_consultas(self):
        self.columnas_consulta = {
            "1": ["Región", "Ventas Totales", "N° Ventas", "Margen Promedio"],
            "2": ["Producto", "Valor", "Categoría", "Región"],
            "3": ["Categoría", "Margen Promedio"],
            "4": ["Vendedor", "Total Ventas", "N° Ventas", "Productos Vendidos"],
            "5": ["Métrica", "Valor"]
        }

    def cargar_opciones_filtro(self):
        """Carga opciones reales desde los datos"""
        if not self.consultas.df.empty:
            self.opciones_filtro = {
                "1": ["Todas"] + sorted(self.consultas.df['Region'].dropna().unique().tolist()),
                "2": ["Todas"] + sorted(self.consultas.df['Categoria'].dropna().unique().tolist()),
                "3": ["Todas"] + sorted(self.consultas.df['Categoria'].dropna().unique().tolist()),
                "4": ["Todas"] + sorted(self.consultas.df['Vendedor'].dropna().unique().tolist()),
                "5": ["Todas"]  # No necesita filtro
            }

    def crear_interfaz(self):
        # Frame principal con mejor distribución
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Panel izquierdo (controles) - Ancho fijo
        panel_controles = ttk.Frame(main_frame, width=250)
        panel_controles.pack(side=tk.LEFT, fill=tk.Y)
        panel_controles.pack_propagate(False)  # Fija el ancho
        
        # Panel derecho (resultados) - Flexible
        panel_resultados = ttk.Frame(main_frame)
        panel_resultados.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Título y controles
        ttk.Label(panel_controles, text="Consultas Disponibles", style='Header.TLabel').pack(pady=(10, 15))
        
        # Botones de consultas con mejor espaciado
        consultas = [
            ("1. Ventas por Región", "1"),
            ("2. Top Productos", "2"),
            ("3. Margen por Categoría", "3"),
            ("4. Ventas por Vendedor", "4"),
            ("5. Estadísticas", "5")
        ]
        
        for texto, consulta_id in consultas:
            btn = ttk.Button(
                panel_controles, 
                text=texto, 
                command=lambda cid=consulta_id: self.actualizar_filtros(cid),
                style='TButton'
            )
            btn.pack(fill=tk.X, pady=3, padx=5)
        
        # Área de filtros reorganizada
        ttk.Separator(panel_controles).pack(fill=tk.X, pady=10)
        ttk.Label(panel_controles, text="Filtros", style='Header.TLabel').pack()
        
        self.filtro_frame = ttk.Frame(panel_controles)
        self.filtro_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(self.filtro_frame, text="Seleccione:").pack(anchor=tk.W)
        self.filtro_combobox = ttk.Combobox(self.filtro_frame, state="readonly")
        self.filtro_combobox.pack(fill=tk.X, pady=3)
        self.filtro_combobox.bind("<<ComboboxSelected>>", self.aplicar_filtro)
        
        # Área de resultados mejorada
        self.tree_frame = ttk.Frame(panel_resultados)
        self.tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.tree = ttk.Treeview(self.tree_frame)
        self.tree_scroll_y = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.tree_scroll_x = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=self.tree_scroll_y.set, xscrollcommand=self.tree_scroll_x.set)
        
        # Grid layout para mejor control
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree_scroll_y.grid(row=0, column=1, sticky="ns")
        self.tree_scroll_x.grid(row=1, column=0, sticky="ew")
        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)
        
        # Footer centrado
        ttk.Label(panel_controles, text=f"Usuario: {self.usuario}", style='Header.TLabel').pack(side=tk.BOTTOM, pady=10)

    def actualizar_filtros(self, consulta_id):
        """Actualiza las opciones del combobox según la consulta seleccionada"""
        self.filtro_combobox['values'] = self.opciones_filtro.get(consulta_id, ["Todas"])
        self.filtro_combobox.set("Todas")
        self.consulta_actual = consulta_id
        self.aplicar_filtro()

    def aplicar_filtro(self, event=None):
        """Aplica el filtro seleccionado"""
        if hasattr(self, 'consulta_actual'):
            filtro = self.filtro_combobox.get()
            filtro = None if filtro == "Todas" else filtro
            
            # Limpiar tabla
            self.tree.delete(*self.tree.get_children())
            columnas = self.columnas_consulta.get(self.consulta_actual, [])
            self.tree["columns"] = columnas
            
            # Configurar columnas
            for col in columnas:
                self.tree.heading(col, text=col, anchor=tk.CENTER)
                self.tree.column(col, width=120, anchor=tk.CENTER, stretch=True)
            
            # Ejecutar consulta
            datos, titulo = self.consultas.obtener_consultas()[self.consulta_actual](filtro)
            
            # Insertar datos
            for fila in datos:
                self.tree.insert("", tk.END, values=fila)
            
            self.root.title(f"{titulo} - {self.usuario}")

class LoginApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Login - Sistema de Consultas")
        self.root.geometry("320x200")
        self.root.configure(bg='#f0f0f0')
        self.crear_widgets()

    def crear_widgets(self):
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(expand=True, fill=tk.BOTH)
        
        ttk.Label(frame, text="Inicio de Sesión", font=('Arial', 12, 'bold')).grid(row=0, columnspan=2, pady=5)
        
        ttk.Label(frame, text="Usuario:").grid(row=1, column=0, pady=5, sticky=tk.W)
        self.usuario_entry = ttk.Entry(frame, width=22)
        self.usuario_entry.grid(row=1, column=1, pady=5)
        
        ttk.Label(frame, text="Contraseña:").grid(row=2, column=0, pady=5, sticky=tk.W)
        self.password_entry = ttk.Entry(frame, width=22, show="*")
        self.password_entry.grid(row=2, column=1, pady=5)
        
        ttk.Button(frame, text="Ingresar", command=self.verificar, width=15).grid(row=3, columnspan=2, pady=10)
        
        frame.columnconfigure(1, weight=1)

    def verificar(self):
        usuario = self.usuario_entry.get()
        password = self.password_entry.get()
        
        if Autenticacion.verificar_usuario(usuario, password):
            self.root.destroy()
            root_consultas = tk.Tk()
            ConsultasApp(root_consultas, usuario)
            root_consultas.mainloop()
        else:
            messagebox.showerror("Error", "Credenciales incorrectas")

if __name__ == "__main__":
    root = tk.Tk()
    app = LoginApp(root)
    root.mainloop()