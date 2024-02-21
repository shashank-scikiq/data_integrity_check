import pandas as pd
import os
from openpyxl import Workbook

class DataProcessor:
    def __init__(self, fname: str):
        if os.path.exists(fname):
            self.df = pd.read_csv(fname)
        else:
            return
        
    def handle_missing_values(self):
        result_series_isnull = self.df.isnull().sum()
        result_series_isin = self.df.apply(lambda x: x.isin(['', None, 'NA', 'N/A', 'NaN', 'nan']).sum())
        result_series_combined = result_series_isnull + result_series_isin
        result_df = pd.DataFrame(result_series_combined, columns=['Null_Values'])
        result_df['Column_Name'] = result_series_combined.index  
        result_df = result_df[['Column_Name', 'Null_Values']]  
        return result_df
    
    def handle_delivery_pincode(self):
        self.df['delivery_pincode'] = self.df['delivery_pincode'].astype(str)
        a = self.df['delivery_pincode'].apply(lambda x: len(str(x)) != 6)
        pincode_counts_delivery = self.df[a]['delivery_pincode'].value_counts()
        result_df_delivery = pd.DataFrame({'Delivery Pincode Error': pincode_counts_delivery.index, 'Delivery_Pincode Error Count': pincode_counts_delivery.values})
        return result_df_delivery

    def handle_seller_pincode(self):
        self.df['seller_pincode'] = self.df['seller_pincode'].astype(str)
        b = self.df['seller_pincode'].apply(lambda x: len(str(x)) != 8)
        pincode_counts = self.df[b]['seller_pincode'].value_counts()
        result_df_seller = pd.DataFrame({'Seller Pincode Error': pincode_counts.index, 'Seller_Pincode Error Count': pincode_counts.values})
        return result_df_seller

    
   
    def save_to_excel(self, output_path: str, *dataframes):
        combined_df = pd.concat(dataframes, axis=1)  # Concatenate dataframes horizontally
        combined_excel_path = f"{output_path}.xlsx"

        # Create Excel writer using openpyxl
        with pd.ExcelWriter(combined_excel_path, engine='openpyxl') as writer:
            combined_df.fillna(0, inplace=True)  # Fill NaN values with 0

            combined_df.to_excel(writer, index=False, sheet_name='Details')
            
            # Save handle_missing_values data horizontally to a new sheet
            result_df_missing_values = self.handle_missing_values()
            result_df_missing_values.T.fillna(0).to_excel(writer, header=False, index=True, sheet_name='Summary')

            # Sum of 'Seller_Pincode Error Count' in the summary sheet
            summary_sheet = writer.sheets['Summary']
            summary_sheet['A3'] = 'Wrong Values'
            
            # Add 0 values to specific ranges (B3 to H3 and K3 to M3)
            for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'K', 'L', 'M']:
                summary_sheet[f'{col}3'] = 0
            
            summary_sheet['I3'] = combined_df['Seller_Pincode Error Count'].sum()

            # Sum of 'Delivery_Pincode Error Count' in the summary sheet
            summary_sheet['J3'] = combined_df['Delivery_Pincode Error Count'].sum()

# Your existing code
csv_file_path = "data.csv"
output_path = "Report"

data_processor = DataProcessor(csv_file_path)
result_df_delivery = data_processor.handle_delivery_pincode()
result_df_seller = data_processor.handle_seller_pincode()

# Use the updated method to save to a single Excel file with multiple sheets
data_processor.save_to_excel(output_path, result_df_delivery, result_df_seller)







