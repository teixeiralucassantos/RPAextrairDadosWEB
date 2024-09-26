import tkinter as tk
from tkinter import messagebox, ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from openpyxl import Workbook
import time

# Configuração da janela principal
class DataExtractorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Web Data Exporter")
        self.geometry("800x600")
        self.configure(bg="#f0f0f0")

        # Estilo da Treeview
        self.style = ttk.Style()
        self.style.configure("Treeview", font=("Arial", 12), rowheight=30)
        self.style.configure("Treeview.Heading", font=("Arial", 14, "bold"), background="#4CAF50", foreground="white")
        self.style.map("Treeview.Heading", background=[('active', '#45a049')])

        # Criação da Treeview
        self.data_tree = ttk.Treeview(self, columns=("ID", "Due Date", "Invoice"), show="headings")
        self.data_tree.heading("ID", text="ID")
        self.data_tree.heading("Due Date", text="Due Date")
        self.data_tree.heading("Invoice", text="Invoice")
        self.data_tree.column("ID", anchor=tk.CENTER, width=100)
        self.data_tree.column("Due Date", anchor=tk.CENTER, width=300)
        self.data_tree.column("Invoice", anchor=tk.CENTER, width=400)
        self.data_tree.pack(pady=20, padx=20, fill=tk.BOTH, expand=True)

        # Botão de exportar
        self.export_button = tk.Button(self, text="Export to Excel", command=self.export_to_excel, font=("Arial", 14), bg="#4CAF50", fg="white")
        self.export_button.pack(pady=20)

        self.extract_data()

    def extract_data(self):
        # Configurações do Chrome
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        service = Service(executable_path="C:\\Users\\User\\Downloads\\chromedriver.exe")
        driver = webdriver.Chrome(service=service, options=chrome_options)

        driver.get("https://rpachallengeocr.azurewebsites.net/")
        time.sleep(3)

        for _ in range(3):  # Número de páginas a extrair
            table_element = driver.find_element(By.XPATH, '//*[@id="tableSandbox"]')
            rows = table_element.find_elements(By.TAG_NAME, "tr")

            for row in rows:
                if not row.text.startswith("#"):
                    data = row.text.split(" ")
                    self.data_tree.insert("", "end", values=(data[0], data[1], data[2]))

            # Clicando no botão "Próximo"
            next_button = driver.find_element(By.XPATH, '//*[@id="tableSandbox_next"]')
            next_button.click()
            time.sleep(2)

        driver.quit()

    def export_to_excel(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Extracted Data"

        # Adicionando cabeçalhos
        sheet.append(["ID", "Due Date", "Invoice"])

        for item in self.data_tree.get_children():
            row_data = self.data_tree.item(item)["values"]
            sheet.append(row_data)

        # Salvando o arquivo
        workbook.save("C:\\Users\\User\\Downloads\\Extracted_Data.xlsx")
        messagebox.showinfo("Success", "Data exported to Excel successfully!")

# Iniciar a aplicação
if __name__ == "__main__":
    app = DataExtractorApp()
    app.mainloop()
