import pandas as pd
from tkinter import messagebox

class ConsultasExcel:
    def __init__(self, archivo='datos.xlsx'):
        try:
            self.df = pd.read_excel(archivo)
            # Limpieza y preparación de datos
            self.df['Fecha'] = pd.to_datetime(self.df['Fecha'])
            self.df['Año'] = self.df['Fecha'].dt.year
            self.df['Mes'] = self.df['Fecha'].dt.month_name()
            # Eliminar filas con valores críticos faltantes
            self.df = self.df.dropna(subset=['Categoria', 'Region', 'Vendedor'])
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar datos:\n{str(e)}")
            self.df = pd.DataFrame()

    def consulta_ventas_region(self, filtro=None):
        df_filtrado = self.df[self.df['Tipo'] == 'VENTA']
        if filtro:
            df_filtrado = df_filtrado[df_filtrado['Region'] == filtro]
        
        resumen = df_filtrado.groupby('Region').agg({
            'Valor': ['sum', 'count'],
            'Margen': 'mean'
        }).round(2)
        return resumen.reset_index().values.tolist(), f"Ventas por Región {'- ' + filtro if filtro else ''}"

    def consulta_top_productos(self, filtro=None):
        df_filtrado = self.df[self.df['Tipo'] == 'VENTA']
        if filtro:
            df_filtrado = df_filtrado[df_filtrado['Categoria'] == filtro]
        
        top = df_filtrado.nlargest(10, 'Valor')[['Producto', 'Valor', 'Categoria', 'Region']]
        return top.values.tolist(), f"Top Productos {'- ' + filtro if filtro else ''}"

    def consulta_margen_categoria(self, filtro=None):
        df_filtrado = self.df[self.df['Tipo'] == 'VENTA']
        if filtro:
            df_filtrado = df_filtrado[df_filtrado['Categoria'] == filtro]
        
        margen = df_filtrado.groupby('Categoria')['Margen'].mean().round(3).reset_index()
        return margen.values.tolist(), f"Margen por Categoría {'- ' + filtro if filtro else ''}"

    def consulta_ventas_vendedor(self, filtro=None):
        df_filtrado = self.df[self.df['Tipo'] == 'VENTA']
        if filtro:
            df_filtrado = df_filtrado[df_filtrado['Vendedor'] == filtro]
        
        ventas = df_filtrado.groupby('Vendedor').agg({
            'Valor': ['sum', 'count'],
            'Producto': lambda x: ', '.join(x.unique()[:3]) + ('...' if len(x.unique()) > 3 else '')
        }).round(2)
        return ventas.reset_index().values.tolist(), f"Ventas por Vendedor {'- ' + filtro if filtro else ''}"

    def consulta_estadisticas(self, _):
        stats = []
        if not self.df.empty:
            ventas = self.df[self.df['Tipo'] == 'VENTA']
            gastos = self.df[self.df['Tipo'] == 'GASTO']
            
            stats = [
                ['Total Ventas', f"${ventas['Valor'].sum().round(2)}"],
                ['Total Gastos', f"${abs(gastos['Valor'].sum()).round(2)}"],
                ['Margen Promedio', f"{ventas['Margen'].mean().round(3)*100}%"],
                ['Producto Más Vendido', ventas['Producto'].mode()[0]],
                ['Mejor Vendedor', ventas.groupby('Vendedor')['Valor'].sum().idxmax()],
                ['Región Más Activa', ventas['Region'].mode()[0]]
            ]
        return stats, "Estadísticas Generales"

    def obtener_consultas(self):
        return {
            "1": self.consulta_ventas_region,
            "2": self.consulta_top_productos,
            "3": self.consulta_margen_categoria,
            "4": self.consulta_ventas_vendedor,
            "5": self.consulta_estadisticas
        }