import pandas as pd

class ExcelGenerator:
    def __init__(self, file_name):
        self.file_name = file_name
        self.sheets_added = False
        self.writer = pd.ExcelWriter(file_name, engine='xlsxwriter')


    def add_sheet(self, df, sheet_name):
        """Adiciona um DataFrame como nova aba no Excel"""
        df.to_excel(self.writer, sheet_name=sheet_name, index=False)
        self.sheets_added = True
        print(f"Planilha '{sheet_name}' adicionada com sucesso!")

    def save(self):
        """Salva o arquivo Excel."""
        if self.sheets_added:
            self.writer.close()
            return self.file_name
        else:
            print("Nenhuma planilha foi adicionada ao relatório.")
            return None  