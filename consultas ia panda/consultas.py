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
            self.df['Dia'] = self.df['Fecha'].dt.day_name()
            # Eliminar filas con valores críticos faltantes
            self.df = self.df.dropna(subset=['Categoria', 'Region', 'Vendedor'])
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar datos:\n{str(e)}")
            self.df = pd.DataFrame()

    # --- CONSULTAS ---
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

    def consulta_ventas_mensuales(self, filtro=None):
        df_filtrado = self.df[self.df['Tipo'] == 'VENTA']
        if filtro:
            df_filtrado = df_filtrado[df_filtrado['Region'] == filtro]
        
        ventas_mes = df_filtrado.groupby(['Año', 'Mes']).agg({
            'Valor': 'sum',
            'Producto': 'count'
        }).rename(columns={'Valor': 'Total Ventas', 'Producto': 'N° Ventas'}).round(2)
        return ventas_mes.reset_index().values.tolist(), f"Ventas Mensuales {'- ' + filtro if filtro else ''}"

    def consulta_rentabilidad_productos(self, filtro=None):
        df_filtrado = self.df[self.df['Tipo'] == 'VENTA']
        if filtro:
            df_filtrado = df_filtrado[df_filtrado['Categoria'] == filtro]
        
        rentabilidad = df_filtrado.groupby('Producto').agg({
            'Valor': 'sum',
            'Margen': 'mean',
            'Costo': 'sum'
        }).assign(
            Rentabilidad=lambda x: (x['Valor'] - x['Costo']).round(2)
        ).nlargest(10, 'Rentabilidad')[['Valor', 'Margen', 'Rentabilidad']]
        return rentabilidad.reset_index().values.tolist(), f"Rentabilidad Productos {'- ' + filtro if filtro else ''}"

    def consulta_gastos_region(self, filtro=None):
        df_filtrado = self.df[self.df['Tipo'] == 'GASTO']
        if filtro:
            df_filtrado = df_filtrado[df_filtrado['Region'] == filtro]
        
        gastos = df_filtrado.groupby('Region').agg({
            'Valor': ['sum', 'count']
        }).round(2)
        return gastos.reset_index().values.tolist(), f"Gastos por Región {'- ' + filtro if filtro else ''}"

    def consulta_eficiencia_vendedores(self, filtro=None):
        df_filtrado = self.df[self.df['Tipo'] == 'VENTA']
        if filtro:
            df_filtrado = df_filtrado[df_filtrado['Region'] == filtro]
        
        eficiencia = df_filtrado.groupby('Vendedor').agg({
            'Valor': 'sum',
            'Margen': 'mean',
            'Producto': 'count'
        }).rename(columns={'Producto': 'N° Ventas'}).round(2)
        return eficiencia.reset_index().values.tolist(), f"Eficiencia Vendedores {'- ' + filtro if filtro else ''}"

    def consulta_tendencia_diaria(self, filtro=None):
        df_filtrado = self.df[self.df['Tipo'] == 'VENTA']
        if filtro:
            df_filtrado = df_filtrado[df_filtrado['Categoria'] == filtro]
        
        tendencia = df_filtrado.groupby('Dia').agg({
            'Valor': 'sum',
            'Producto': 'count'
        }).rename(columns={'Valor': 'Total Ventas', 'Producto': 'N° Ventas'}).round(2)
        return tendencia.reset_index().values.tolist(), f"Tendencia Diaria {'- ' + filtro if filtro else ''}"

    def obtener_consultas(self):
        return {
            "1": self.consulta_ventas_region,
            "2": self.consulta_top_productos,
            "3": self.consulta_margen_categoria,
            "4": self.consulta_ventas_vendedor,
            "5": self.consulta_estadisticas,
            "6": self.consulta_ventas_mensuales,
            "7": self.consulta_rentabilidad_productos,
            "8": self.consulta_gastos_region,
            "9": self.consulta_eficiencia_vendedores,
            "10": self.consulta_tendencia_diaria
        }