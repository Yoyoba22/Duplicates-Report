import os
import re
import psycopg2 as pg
from datetime import datetime
import pandas as pd
from sqlalchemy import create_engine, exc
from tkinter import *
from tkinter import filedialog, messagebox
import matplotlib.pyplot as plt

class Jovana:
    def __init__(self, database, user, password, host, port):
        self.con = pg.connect(
            database=database,
            user=user,
            password=password,
            host=host,
            port=port
        )
        self.engine = create_engine(f'postgresql+psycopg2://{user}:{password}@{host}:{port}/{database}')
        
    def extract_date_from_filename(self, file_path):
        date_pattern = re.compile(r'\d{2}\.\d{2}')
        match = date_pattern.search(file_path)
        if match:
            return datetime.strptime(match.group(), '%d.%m').date().replace(year=datetime.now().year)
        else:
            raise ValueError("Date not found in the filename")

    def read_excel_file(self, file_path, date, osoba):
        print(f"Attempting to read file: {file_path}")
        try:
            df = pd.read_excel(file_path, usecols=[0, 12], header=None, skiprows=1)
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            raise

        print("Initial DataFrame head:\n", df.head())

        df.columns = ['identifier', 'status']
        df['datum'] = date
        df['osoba'] = osoba

        df['identifier'] = df['identifier'].astype(str)
        df['status'] = df['status'].astype(str)

        df = df[['identifier', 'status', 'datum', 'osoba']]
        print(f"Processed DataFrame head:\n{df.head()}")
        return df

    def insert_data_into_db(self, df, table_name):
        print(f"Attempting to insert data into table: {table_name}")
        try:
            df.to_sql(table_name, self.engine, if_exists='append', index=False)
            print("Data inserted successfully.")
        except exc.SQLAlchemyError as e:
            print(f"Error inserting data into database: {e}")
            raise

    def export_all_data_to_excel(self, table_name, output_file_path):
        print(f"Attempting to export data from table {table_name} to Excel file: {output_file_path}")
        try:
            query = f'SELECT * FROM "{table_name}"'
            df = pd.read_sql(query, self.con)
            df.to_excel(output_file_path, index=False)
            print("Data exported successfully.")
        except Exception as e:
            print(f"Error exporting data to Excel: {e}")
            raise

    def generate_pie_chart(self, table_name, osoba, month=None):
        print(f"Generating pie chart for {osoba} with month filter: {month}")
        try:
            query = f"""
                SELECT status, COUNT(*) as count 
                FROM "{table_name}" 
                WHERE status IN ('kein Duplikat', 'Done', 'In Arbeit') OR status IS NULL
            """
            params = []

            if osoba != "All":
                query += " AND osoba = %s"
                params.append(osoba)
            if month:
                query += " AND EXTRACT(MONTH FROM datum) = %s"
                params.append(month)

            query += " GROUP BY status"
            title = f'Status Distribution for {osoba}' + (f' in Month {month}' if month else '')

            df = pd.read_sql(query, self.con, params=params)

            if df.empty:
                messagebox.showinfo("Info", "No data available for the selected criteria.")
                return

            total_query = f"""
                SELECT COUNT(*) as total 
                FROM "{table_name}" 
                WHERE status IN ('kein Duplikat', 'Done', 'In Arbeit', 'nan')
            """
            if osoba != "All":
                total_query += " AND osoba = %s"
                params_total = [osoba]
            if month:
                total_query += " AND EXTRACT(MONTH FROM datum) = %s"
                params_total = params_total if 'params_total' in locals() else []
                params_total.append(month)

            total_count_df = pd.read_sql(total_query, self.con, params=params_total if 'params_total' in locals() else None)
            total_count = total_count_df['total'].iloc[0] if not total_count_df.empty else 0

            df_filtered = df[df['status'].notna()]

            plt.figure(figsize=(8, 8))
            plt.pie(df_filtered['count'], labels=df_filtered['status'], autopct='%1.1f%%', startangle=140)
            plt.title(title)

            plt.figtext(0.5, 0.01, f"Total: {total_count}", ha="center", fontsize=12)

            plt.show()

        except Exception as e:
            print(f"Error generating pie chart: {e}")
            messagebox.showerror("Error", str(e))

    def generate_line_chart(self):
        print("Generating line chart for status totals per month")
        try:
            query = f"""
                SELECT EXTRACT(MONTH FROM datum) AS month, status, COUNT(*) AS count
                FROM "Jovana"
                WHERE status IN ('kein Duplikat', 'Done', 'In Arbeit')
                GROUP BY EXTRACT(MONTH FROM datum), status
                ORDER BY month, status
            """
            df = pd.read_sql(query, self.con)

            if df.empty:
                messagebox.showinfo("Info", "No data available for the line chart.")
                return

            plt.figure(figsize=(10, 6))
            
            for status in df['status'].unique():
                status_df = df[df['status'] == status]
                plt.plot(status_df['month'], status_df['count'], marker='o', label=status)

            plt.xlabel('Month')
            plt.ylabel('Total Count')
            plt.title('Total Count per Month by Status')
            plt.legend(title='Status')
            plt.grid(True)
            plt.xticks(range(1, 13))
            plt.tight_layout()
            plt.show()

        except Exception as e:
            print(f"Error generating line chart: {e}")
            messagebox.showerror("Error", str(e))

def select_file():
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=(("Excel files", "*.xlsx"),))
    if file_path:
        selected_file_label.config(text=file_path)
        return file_path
    else:
        return None

def import_data():
    file_path = select_file()
    if not file_path:
        messagebox.showerror("Error", "No file selected.")
        return

    try:
        date = jovana.extract_date_from_filename(file_path)
        df = jovana.read_excel_file(file_path, date, osoba_var.get())
        jovana.insert_data_into_db(df, 'Jovana')
        messagebox.showinfo("Success", "Data imported successfully.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def export_data():
    output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=(("Excel files", "*.xlsx"),))
    if output_file_path:
        try:
            jovana.export_all_data_to_excel('Jovana', output_file_path)
            messagebox.showinfo("Success", "Data exported successfully.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

def open_pie_chart_window():
    chart_window = Toplevel(root)
    chart_window.title("Generate Pie Chart")

    chart_osoba_var = StringVar(chart_window)
    chart_osoba_var.set("All")  

    chart_month_var = StringVar(chart_window)
    chart_month_var.set("All") 

    Label(chart_window, text="Select osoba:").pack(pady=10)
    OptionMenu(chart_window, chart_osoba_var, "All", "Jovana", "Vukasin", "Jelena", "Ivana").pack(pady=10)

    Label(chart_window, text="Select month:").pack(pady=10)
    OptionMenu(chart_window, chart_month_var, "All", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12").pack(pady=10)

    Button(chart_window, text="Generate Pie Chart", command=lambda: generate_chart(chart_osoba_var, chart_month_var)).pack(pady=10)

def generate_chart(osoba_var, month_var):
    month = month_var.get()
    if month == "All":
        month = None
    jovana.generate_pie_chart('Jovana', osoba_var.get(), month)

def open_line_chart_window():
    line_chart_window = Toplevel(root)
    line_chart_window.title("Generate Line Chart")

    Label(line_chart_window, text="Generating line chart...").pack(pady=10)

    Button(line_chart_window, text="Generate Line Chart", command=jovana.generate_line_chart).pack(pady=10)


database = 'Jovana'
user = 'postgres'
password = 'Yoyoba22'
host = 'localhost'
port = '5432'


jovana = Jovana(database, user, password, host, port)







root = Tk()
root.title("Data Import and Export")

osoba_var = StringVar(root)
osoba_var.set("Jovana")

Label(root, text="Select osoba:").pack(pady=10)
OptionMenu(root, osoba_var, "Jovana", "Vukasin", "Jelena", "Ivana").pack(pady=10)

Button(root, text="Import Data", command=import_data).pack(pady=10)
Button(root, text="Export All Data to Excel", command=export_data).pack(pady=10)


menu = Menu(root)
root.config(menu=menu)

chart_menu = Menu(menu, tearoff=0)
menu.add_cascade(label="Chart Options", menu=chart_menu)
chart_menu.add_command(label="Generate Pie Chart", command=open_pie_chart_window)
chart_menu.add_command(label="Generate Line Chart", command=open_line_chart_window)

selected_file_label = Label(root, text="No file selected")
selected_file_label.pack(pady=10)

root.mainloop()
