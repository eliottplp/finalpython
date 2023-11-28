import pandas as pd
import os
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from fpdf import FPDF
from sklearn.linear_model import LinearRegression
from sklearn.model_selection import train_test_split
import tkinter as tk
from tkinter import filedialog, messagebox
DARK_BLUE = '#003366'
LIGHT_BLUE = '#66CCFF'
class Client:
    def __init__(self, name, date_birth, city_birth, email, phone, gender, job, country, postal_code):
        self.name = name
        self.date_birth = date_birth
        self.city_birth = city_birth
        self.email = email
        self.phone = phone
        self.gender = gender
        self.job = job
        self.country = country
        self.postal_code = postal_code

    def show_info(self):
        info = (f"Name: {self.name}\nDate of Birth: {self.date_birth}\nCity of Birth: {self.city_birth}\n"
                f"Email: {self.email}\nPhone: {self.phone}\nGender: {self.gender}\n"
                f"Job: {self.job}\nCountry: {self.country}\nPostal Code: {self.postal_code}")
        print(info)


def create_client_ui():
    client_window = tk.Toplevel()
    client_window.title("Create New Client")
    client_window.configure(bg=DARK_BLUE)

    # client creation
    def submit_client():
        name = entry_name.get()
        date_birth = entry_date_birth.get()
        city_birth = entry_city_birth.get()
        email = entry_email.get()
        phone = entry_phone.get()
        gender = entry_gender.get()
        job = entry_job.get()
        country = entry_country.get()
        postal_code = entry_postal_code.get()

        new_client = Client(name, date_birth, city_birth, email, phone, gender, job, country, postal_code)
        new_client.show_info()

        # Save information in Excel file
        new_client_data = pd.DataFrame([{
            'Name': name,
            'Date of Birth': date_birth,
            'City of Birth': city_birth,
            'Email': email,
            'Phone': phone,
            'Gender': gender,
            'Job': job,
            'Country': country,
            'Postal Code': postal_code
        }])

        excel_file = 'clients.xlsx'
        if os.path.exists(excel_file):
            existing_data = pd.read_excel(excel_file)
            updated_data = existing_data.append(new_client_data, ignore_index=True)
        else:
            updated_data = new_client_data

        updated_data.to_excel(excel_file, index=False)
        messagebox.showinfo("Info", f"Client {name} Created")
        client_window.destroy()


    tk.Label(client_window, text="Name").grid(row=0, column=0)
    entry_name = tk.Entry(client_window)
    entry_name.grid(row=0, column=1)

    tk.Label(client_window, text="Date of Birth").grid(row=1, column=0)
    entry_date_birth = tk.Entry(client_window)
    entry_date_birth.grid(row=1, column=1)

    tk.Label(client_window, text="City of Birth").grid(row=2, column=0)
    entry_city_birth = tk.Entry(client_window)
    entry_city_birth.grid(row=2, column=1)

    tk.Label(client_window, text="Email").grid(row=3, column=0)
    entry_email = tk.Entry(client_window)
    entry_email.grid(row=3, column=1)

    tk.Label(client_window, text="Phone").grid(row=4, column=0)
    entry_phone = tk.Entry(client_window)
    entry_phone.grid(row=4, column=1)

    tk.Label(client_window, text="Gender").grid(row=5, column=0)
    entry_gender = tk.Entry(client_window)
    entry_gender.grid(row=5, column=1)

    tk.Label(client_window, text="Job").grid(row=6, column=0)
    entry_job = tk.Entry(client_window)
    entry_job.grid(row=6, column=1)

    tk.Label(client_window, text="Country").grid(row=7, column=0)
    entry_country = tk.Entry(client_window)
    entry_country.grid(row=7, column=1)

    tk.Label(client_window, text="Postal Code").grid(row=8, column=0)
    entry_postal_code = tk.Entry(client_window)
    entry_postal_code.grid(row=8, column=1)

    submit_button = tk.Button(client_window, text="Submit", command=submit_client)
    submit_button.grid(row=9, column=0, columnspan=2, pady=10)

def get_key_numbers(salesdf, clientsdf):
    clientsdf['Date of Birth'] = pd.to_datetime(clientsdf['Date of Birth'], errors='coerce')
    last_30_days = datetime.now() - pd.Timedelta(days=30)
    sales_last_30_days = salesdf[salesdf['Purchase Time'] > last_30_days].shape[0]
    new_clients_last_30_days = clientsdf[clientsdf['Date of Birth'] > last_30_days].shape[0]
    revenue_this_year = salesdf[salesdf['Purchase Time'].dt.year == datetime.now().year].sum(numeric_only=True)['Total Price']
    
    return sales_last_30_days, new_clients_last_30_days, revenue_this_year



def create_pdf_report(salesdf, clientsdf):
    # Save plots as images
    def save_plot(fig, filename):
        fig.savefig(filename, bbox_inches='tight')
        plt.close(fig)

    # Plot 1: Sales Over Time
    fig, ax = plt.subplots(figsize=(10, 5))
    sales_over_time = salesdf.set_index('Purchase Time').resample('M').sum()['Total Price']
    ax.plot(sales_over_time, marker='o', linestyle='-')
    ax.set_title('Sales Over Time (Monthly Aggregated)')
    ax.set_xlabel('Time')
    ax.set_ylabel('Total Sales')
    save_plot(fig, 'sales_over_time.png')

    # Plot 2: Sales by Country
    fig, ax = plt.subplots(figsize=(10, 5))
    sales_by_country = salesdf.groupby('Client Country').sum(numeric_only=True)['Total Price']
    sales_by_country.plot(kind='bar', ax=ax)
    ax.set_title('Sales by Country')
    ax.set_xlabel('Country')
    ax.set_ylabel('Total Sales')
    save_plot(fig, 'sales_by_country.png')

    # Plot 3: Sales by Product
    fig, ax = plt.subplots(figsize=(10, 5))
    sales_by_product = salesdf.groupby('Product').sum(numeric_only=True)['Total Price']
    sales_by_product.plot(kind='bar', ax=ax)
    ax.set_title('Sales by Product')
    ax.set_xlabel('Product')
    ax.set_ylabel('Total Sales')
    save_plot(fig, 'sales_by_product.png')

    # Create PDF 
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, 'Sales Report', 0, 1, 'C') 

    # Key numbers
    pdf.set_font('Arial', '', 12)
    sales_last_30_days, new_clients_last_30_days, revenue_this_year = get_key_numbers(salesdf, clientsdf)
    pdf.cell(0, 10, f'Sales in Last 30 Days: {sales_last_30_days}', ln=True)
    pdf.cell(0, 10, f'New Clients in Last 30 Days: {new_clients_last_30_days}', ln=True)
    pdf.cell(0, 10, f'Revenue This Year: ${revenue_this_year:,.2f}', ln=True)

    # Adding plots to PDF
    for image in ['sales_over_time.png', 'sales_by_country.png', 'sales_by_product.png']:
        pdf.image(image, x=10, y=None, w=180)
        pdf.ln(0) 



    pdf.output('sales_report.pdf')

def analyze_sales_and_create_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        try:
            salesdf = pd.read_excel(file_path)
            clientsdf = pd.read_excel('clients.xlsx')  
            create_pdf_report(salesdf, clientsdf)
            messagebox.showinfo("Success", "PDF report has been created successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

#add new sale
products = {
    "Ergonomic Office Chair": 200,
    "Convertible Sofa Bed": 400,
    "Expandable Dining Table": 350,
    "Modular Shelving Unit": 150,
    "Smart Coffee Table": 300,
    "Vintage Vanity Dresser": 250,
    "Adjustable Standing Desk": 400,
    "Outdoor Patio Set": 500,
    "Children's Bunk Bed with Storage": 450,
    "Recliner with Massage and Heat Functions": 600
}

def add_sale_ui():
    sale_window = tk.Toplevel()
    sale_window.title("Add New Sale")
    sale_window.configure(bg=DARK_BLUE)

    # Sale submission
    def submit_sale():
        email = entry_email.get()
        product_name = product_var.get()
        quantity = int(entry_quantity.get())
        unit_price = products[product_name]
        total_price = quantity * unit_price
        purchase_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        new_sale = {
            'Client Email': email,
            'Product': product_name,
            'Unit Price': unit_price,
            'Quantity': quantity,
            'Total Price': total_price,
            'Purchase Time': purchase_time
        }
        update_sales_file(new_sale)

        messagebox.showinfo("Success", f"Sale of {quantity} x {product_name} recorded")
        sale_window.destroy()

    # update the sales Excel file
    def update_sales_file(new_sale):
        sales_file = 'sales.xlsx'
        try:
            if os.path.exists(sales_file):
                existing_data = pd.read_excel(sales_file)
                updated_data = existing_data.append(new_sale, ignore_index=True)
            else:
                updated_data = pd.DataFrame([new_sale])

            updated_data.to_excel(sales_file, index=False)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while saving the sale: {e}")

    # Email entry
    tk.Label(sale_window, text="Client Email").grid(row=0, column=0)
    entry_email = tk.Entry(sale_window)
    entry_email.grid(row=0, column=1)

    # Product dropdown
    tk.Label(sale_window, text="Product").grid(row=1, column=0)
    product_var = tk.StringVar(sale_window)
    product_var.set(list(products.keys())[0])  # default value
    dropdown_product = tk.OptionMenu(sale_window, product_var, *products.keys())
    dropdown_product.grid(row=1, column=1)

    # Quantity entry
    tk.Label(sale_window, text="Quantity").grid(row=2, column=0)
    entry_quantity = tk.Entry(sale_window)
    entry_quantity.grid(row=2, column=1)

    # Submit button
    submit_button = tk.Button(sale_window, text="Submit Sale", command=submit_sale)
    submit_button.grid(row=3, column=0, columnspan=2, pady=10)


#predict sales next month
def preprocess_data_for_quantity(salesdf):
    salesdf['Purchase Time'] = pd.to_datetime(salesdf['Purchase Time'])
    salesdf['Year'] = salesdf['Purchase Time'].dt.year
    salesdf['Month'] = salesdf['Purchase Time'].dt.month
    monthly_sales = salesdf.groupby(['Product', 'Year', 'Month']).sum(numeric_only=True).reset_index()
    return monthly_sales




def train_predict_quantity_model(product_data):
    product_data = product_data.copy()  
    product_data['TimeIndex'] = product_data['Year'] * 12 + product_data['Month']
    X = product_data[['TimeIndex']]
    y = product_data['Quantity']

    model = LinearRegression()
    model.fit(X, y)

    last_time_index = product_data['TimeIndex'].iloc[-1]
    next_month_quantity = model.predict([[last_time_index + 1]])
    return round(next_month_quantity[0]) 


def predict_next_month_quantity(salesdf):
    monthly_sales = preprocess_data_for_quantity(salesdf)
    quantity_predictions = {}

    for product in monthly_sales['Product'].unique():
        product_data = monthly_sales[monthly_sales['Product'] == product]
        quantity_predictions[product] = train_predict_quantity_model(product_data)

    return quantity_predictions


def predict_sales_ui():
    prediction_window = tk.Toplevel()
    prediction_window.title("Predict Sales for Next Month")
    prediction_window.configure(bg=DARK_BLUE)

    # Function to handle the prediction and display results
    def perform_prediction():
        try:
            salesdf = pd.read_excel('sales.xlsx')  
            next_month_predictions = predict_next_month_quantity(salesdf)

            # Displaying predictions
            result_text = "\nNext Month Quantity Predictions:\n"
            for product, prediction in next_month_predictions.items():
                result_text += f"{product}: {int(prediction)} units\n"

            messagebox.showinfo("Predictions", result_text)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    predict_button = tk.Button(prediction_window, text="Predict Sales", command=perform_prediction)
    predict_button.pack(pady=10)

def main():
    window = tk.Tk()
    window.title("Client and Sales Analysis")
    window.configure(bg=DARK_BLUE)

    tk.Button(window, text="Create New Client", command=create_client_ui).pack(pady=10)
    tk.Button(window, text="Analyze Excel File Clients and Sales", command=analyze_sales_and_create_pdf).pack(pady=10)
    tk.Button(window, text="Add New Sale", command=add_sale_ui).pack(pady=10)
    tk.Button(window, text="Predict Sales for Next Month", command=predict_sales_ui).pack(pady=10)
    tk.Button(window, text="Quit", command=window.quit).pack(pady=10)

    window.mainloop()

if __name__ == "__main__":
    main()