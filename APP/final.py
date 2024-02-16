import pandas as pd

class DataProcessor:
    def __init__(self, csv_file_path):
        self.df = pd.read_csv(csv_file_path)

    def handle_missing_values(self):
        result_series = self.df.isnull().sum()
        result_df = pd.DataFrame(result_series, columns=['Null_Count'])
        result_df['Column_Name'] = result_series.index  # Add column names to the DataFrame
        return result_df

    def handle_delivery_pincode(self):
        self.df['delivery_pincode'] = self.df['delivery_pincode'].astype(str)
        a = self.df['delivery_pincode'].apply(lambda x: len(x) > 6)
        pincode_counts_delivery = self.df[a]['delivery_pincode'].value_counts()
        result_df_delivery = pd.DataFrame({'Delivery_Pincode': pincode_counts_delivery.index, 'Count': pincode_counts_delivery.values})
        return result_df_delivery

    def handle_seller_pincode(self):
        self.df['seller_pincode'] = self.df['seller_pincode'].astype(str)
        b = self.df['seller_pincode'].apply(lambda x: len(x) > 6)
        pincode_counts = self.df[b]['seller_pincode'].value_counts()
        result_df_seller = pd.DataFrame({'Pincode': pincode_counts.index, 'Count': pincode_counts.values})
        return result_df_seller

    def save_to_excel(self, result_df, result_df_delivery, result_df_seller, output_file):
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            # Null_Count table
            result_df.to_excel(writer, sheet_name='Combined', index=False, startrow=1, startcol=0)
            worksheet = writer.sheets['Combined']
            worksheet.write(0, 0, 'Null Values Table', writer.book.add_format({'bold': True}))
            
            # Delivery_Pincode table
            result_df_delivery.to_excel(writer, sheet_name='Combined', startrow=result_df.shape[0] + 6, startcol=0, index=False, header=['Delivery_Pincode', 'Count'])
            worksheet.write(result_df.shape[0] + 5, 0, 'Delivery Pincode Table', writer.book.add_format({'bold': True}))

            # Seller_Pincode table
            result_df_seller.to_excel(writer, sheet_name='Combined', startrow=result_df.shape[0] + result_df_delivery.shape[0] + 9, startcol=0, index=False, header=['Pincode', 'Count'])
            worksheet.write(result_df.shape[0] + result_df_delivery.shape[0] + 8, 0, 'Seller Pincode Table', writer.book.add_format({'bold': True}))

# Example usage:
csv_file_path = "data.csv"
output_file = "output_combined.xlsx"

data_processor = DataProcessor(csv_file_path)
result_df = data_processor.handle_missing_values()
result_df_delivery = data_processor.handle_delivery_pincode()
result_df_seller = data_processor.handle_seller_pincode()
data_processor.save_to_excel(result_df, result_df_delivery, result_df_seller, output_file)
