import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font

class FinancialReportGenerator:
    def __init__(self, input_file, output_file):
        self.input_file = input_file
        self.output_file = output_file
        self.journal_entries = self._load_journal_entries()
        self.balance_sheet = {"Assets": 0, "Liabilities": 0, "Equity": 0}
        self.profit_and_loss = {"Income": 0, "Expenses": 0}
    
    def _load_journal_entries(self):
        """Load journal entries from an Excel file and validate input format."""
        try:
            df = pd.read_excel(self.input_file)
            required_columns = {'Date', 'Description', 'Type', 'Debit', 'Credit'}
            if not required_columns.issubset(df.columns):
                raise ValueError(f"Input Excel file must contain columns: {required_columns}")
            return df
        except FileNotFoundError:
            raise FileNotFoundError(f"File {self.input_file} not found.")
    
    def _categorize_entries(self):
        """Categorize journal entries into Balance Sheet and P&L accounts."""
        for _, row in self.journal_entries.iterrows():
            entry_type = row['Type'].strip().capitalize()
            debit = row['Debit']
            credit = row['Credit']
            
            if entry_type in ['Asset', 'Liability', 'Equity']:
                if entry_type == 'Asset':
                    self.balance_sheet['Assets'] += debit - credit
                elif entry_type == 'Liability':
                    self.balance_sheet['Liabilities'] += credit - debit
                elif entry_type == 'Equity':
                    self.balance_sheet['Equity'] += credit - debit
            elif entry_type in ['Income', 'Expense']:
                if entry_type == 'Income':
                    self.profit_and_loss['Income'] += credit - debit
                elif entry_type == 'Expense':
                    self.profit_and_loss['Expenses'] += debit - credit
            else:
                raise ValueError(f"Unknown entry type: {entry_type}")

    def _save_to_excel(self):
        """Generate the Balance Sheet and Profit & Loss statement and save them into an Excel file."""
        with pd.ExcelWriter(self.output_file, engine='openpyxl') as writer:
            self._save_balance_sheet(writer)
            self._save_profit_and_loss(writer)
        print(f"Financial reports saved to {self.output_file}")

    def _save_balance_sheet(self, writer):
        """Save Balance Sheet details to an Excel sheet."""
        total_equity_liabilities = self.balance_sheet['Liabilities'] + self.balance_sheet['Equity']
        balance_data = {
            'Category': ['Assets', 'Liabilities', 'Equity', 'Total Equity and Liabilities'],
            'Amount': [self.balance_sheet['Assets'], 
                       self.balance_sheet['Liabilities'], 
                       self.balance_sheet['Equity'], 
                       total_equity_liabilities]
        }
        df_balance = pd.DataFrame(balance_data)
        df_balance.to_excel(writer, sheet_name='Balance Sheet', index=False)

        # Apply some basic formatting
        worksheet = writer.sheets['Balance Sheet']
        for col in ['A', 'B']:
            worksheet[f'{col}1'].font = Font(bold=True)
            worksheet.column_dimensions[col].width = 25
    
    def _save_profit_and_loss(self, writer):
        """Save Profit & Loss details to an Excel sheet."""
        profit_or_loss = self.profit_and_loss['Income'] - self.profit_and_loss['Expenses']
        p_and_l_data = {
            'Category': ['Income', 'Expenses', 'Profit/Loss'],
            'Amount': [self.profit_and_loss['Income'], 
                       self.profit_and_loss['Expenses'], 
                       profit_or_loss]
        }
        df_p_and_l = pd.DataFrame(p_and_l_data)
        df_p_and_l.to_excel(writer, sheet_name='Profit & Loss', index=False)

        # Apply some basic formatting
        worksheet = writer.sheets['Profit & Loss']
        for col in ['A', 'B']:
            worksheet[f'{col}1'].font = Font(bold=True)
            worksheet.column_dimensions[col].width = 25
    
    def generate_reports(self):
        """Main method to categorize entries and generate Excel reports."""
        self._categorize_entries()
        self._save_to_excel()

if __name__ == "__main__":
    input_file = 'journal_entries.xlsx'  # Input Excel file containing journal entries
    output_file = 'financial_reports.xlsx'  # Output file to save financial reports
    report_generator = FinancialReportGenerator(input_file, output_file)
    report_generator.generate_reports()
