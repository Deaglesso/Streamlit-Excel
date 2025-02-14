import pandas as pd
import subprocess
import sys
import streamlit as st
from io import BytesIO


# Function to calculate row price
def calculate_row_price(sku, price_dict):
    total_price = 0
    sku = str(sku)
    products = sku.split('/+')
    products = [product.strip('/ ').casefold() for product in products]  # Convert all to lowercase

    # Create a case-insensitive price dictionary
    price_dict_lower = {key.casefold(): value for key, value in price_dict.items()}

    shoe_count = len([p for p in products if p != 'çorap'])
    has_other_products = any(p != 'çorap' for p in products)

    for product in products:
        if product == 'çorap':
            total_price += 3 if has_other_products else price_dict_lower.get('çorap', 0)
        else:
            total_price += price_dict_lower.get(product, 0)

    discount = shoe_count - 1 if shoe_count > 1 else 0
    return total_price - discount



# Function to mark duplicate order numbers
def mark_duplicate_order_numbers(df, order_column='订单号'):
    order_counts = df[order_column].value_counts()
    df['Duplicate'] = df[order_column].apply(
        lambda x: '⛔' if pd.notna(x) and order_counts.get(x, 0) > 1 else ''
    )
    return df


# Function to check for mismatching total price
def check_mismatch_total_price(df):
    df['Mismatch'] = df.apply(
        lambda row: '⚠️' if pd.notna(row['Unnamed: 4']) and row['Unnamed: 4'] != row['Total Price'] else '',
        axis=1
    )
    return df


# Function to process the sales report
def process_sales_report(uploaded_file, price_dict):
    df = pd.read_excel(uploaded_file)

    # Process data
    df = mark_duplicate_order_numbers(df)
    df['Total Price'] = df['SKU'].apply(lambda sku: calculate_row_price(sku, price_dict))
    df = check_mismatch_total_price(df)

    total_sum = df['Total Price'].sum()
    totals = pd.DataFrame({'订单号': ['TOTAL'], 'SKU': [''], '买家姓名': [''], '运单号': [''], 'Unnamed: 4': [''], 'Duplicate': [''], 'Total Price': [total_sum], 'Mismatch': ['']}, index=[len(df)])
    df = pd.concat([df, totals])

    df.rename(columns={'Unnamed: 4': 'Given Price'}, inplace=True)
    df = df[['订单号', 'SKU', '买家姓名', '运单号', 'Given Price', 'Total Price', 'Mismatch', 'Duplicate']]

    return df


# Streamlit app interface
def main():
    st.set_page_config(layout="wide")
    st.title("Sales Report Processor")
    st.markdown("Upload your sales report in Excel format to process and view the results.")

    # File upload button
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    
    if uploaded_file is not None:
        st.info("Processing file...")
        price_dict = {
    'Glue': 36.5,
    'Street': 33,
    'Boat': 35,
    'Runner': 30,
    'Çorap': 9.5,
    'Takna': 34,
    'Boris': 30,
    'Airheels': 33,
    'Lost': 33,
    'FlexStride': 31,
    'Smith': 31,
    'Tabanlık 3cm': 7.5,
    'Tabanlık 4.5cm': 8,
    'Tabanlık 6cm': 10,
    'Pinor': 28,
    'Classic': 34,
    'Fashion': 28,
    'Lea': 33,
    'Stapper': 32,
    'Huddo': 33,
    'Clas': 37,
    'Stach': 26,
    'Confy': 26,
    'Vigo': 36,
    'Yamore': 36,
    'Long Boat': 35,
    'Casual': 30,
    'Breath': 22,
    'Warm Boat': 32,
}     

        # Process the sales report
        df_result = process_sales_report(uploaded_file, price_dict)
        
        # Show the processed DataFrame
        st.dataframe(df_result, use_container_width=True)

        # Save to an in-memory BytesIO object
        with BytesIO() as buffer:
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df_result.to_excel(writer, index=False, sheet_name="Report")
            
            # Get the file content and provide download
            st.download_button(
                label="Download Processed File",
                data=buffer.getvalue(),
                file_name="processed_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# Run the app
if __name__ == "__main__":
    main()
