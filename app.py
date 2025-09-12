from flask import Flask, render_template, request, session, redirect, url_for
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import os
import uuid
from datetime import datetime
import socket # Import the socket library
import numpy as np # Import numpy for horizontal bar chart positioning

app = Flask(__name__)
app.secret_key = "supersecretkey" # Needed for sessions

# 🔹 Absolute paths
BASE_DIR = r"C:\Users\okuhl\financial_data_app"
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
STATIC_FOLDER = os.path.join(BASE_DIR, "static")
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure folders exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(STATIC_FOLDER, exist_ok=True)

@app.context_processor
def inject_year():
    """Inject current year for copyright in all templates"""
    return {'current_year': datetime.now().year}

def generate_graph(df, graph_type, filename, is_month_detail=False, selected_month_data=None):
    """
    Generates and saves a graph based on the specified type.
    is_month_detail: True if generating graph for a single month's detail.
    selected_month_data: DataFrame row for month detail if is_month_detail is True.
    """
    plt.figure(figsize=(10, 5)) # Default size for overall graphs

    if is_month_detail and selected_month_data is not None:
        # Month detail graphs are smaller
        plt.figure(figsize=(6, 4))
        income = selected_month_data['Income']
        expense = selected_month_data['Expense']
        month_name = selected_month_data['Month']

        if graph_type == 'column':
            plt.bar(['Income', 'Expense'], [income, expense], color=['green', 'orange'])
            plt.title(f"{month_name} - Income vs Expense", color='blue', fontsize=16)
            plt.ylabel("Amount", color='blue')
        elif graph_type == 'bar':
            plt.barh(['Income', 'Expense'], [income, expense], color=['green', 'orange'])
            plt.title(f"{month_name} - Income vs Expense", color='blue', fontsize=16)
            plt.xlabel("Amount", color='blue')
        elif graph_type == 'pie':
            # Only generate if there's data to show (e.g., income or expense > 0)
            if income > 0 or expense > 0:
                plt.pie([income, expense], labels=['Income', 'Expense'], colors=['green', 'orange'], autopct='%1.1f%%', startangle=90)
                plt.title(f"{month_name} - Income vs Expense", color='blue', fontsize=16)
            else:
                plt.text(0.5, 0.5, "No data for pie chart", horizontalalignment='center', verticalalignment='center', transform=plt.gca().transAxes)
                plt.title(f"{month_name} - Income vs Expense (No Data)", color='blue', fontsize=16)
    else: # Overall graphs
        if graph_type == 'line':
            plt.plot(df['Month'], df['Income'], label='Income', marker='o', color='green')
            plt.plot(df['Month'], df['Expense'], label='Expense', marker='o', color='orange')
            plt.title('Income & Expense Trends', color='blue', fontsize=16)
            plt.xlabel('Month', color='blue')
            plt.ylabel('Amount', color='blue')
            plt.xticks(rotation=45) # Rotate month labels
            plt.grid(True)
            plt.legend()
        elif graph_type == 'column':
            width = 0.35
            x = np.arange(len(df['Month']))
            plt.bar(x - width/2, df['Income'], width, label='Income', color='green')
            plt.bar(x + width/2, df['Expense'], width, label='Expense', color='orange')
            plt.xticks(x, df['Month'], rotation=45) # Rotate month labels
            plt.title('Income & Expense Trends', color='blue', fontsize=16)
            plt.xlabel('Month', color='blue')
            plt.ylabel('Amount', color='blue')
            plt.legend()
        elif graph_type == 'bar':
            height = 0.35
            y = np.arange(len(df['Month']))
            plt.barh(y - height/2, df['Income'], height, label='Income', color='green')
            plt.barh(y + height/2, df['Expense'], height, label='Expense', color='orange')
            plt.yticks(y, df['Month'])
            plt.title('Income & Expense Trends', color='blue', fontsize=16)
            plt.xlabel('Amount', color='blue')
            plt.ylabel('Month', color='blue')
            plt.legend()
        elif graph_type == 'pie':
            # For overall pie chart, sum up total income and total expense
            total_income = df['Income'].sum()
            total_expense = df['Expense'].sum()
            
            if total_income > 0 or total_expense > 0:
                plt.pie([total_income, total_expense], labels=['Total Income', 'Total Expense'], colors=['green', 'orange'], autopct='%1.1f%%', startangle=90)
                plt.title('Overall Income vs Expense Distribution', color='blue', fontsize=16)
            else:
                plt.text(0.5, 0.5, "No data for pie chart", horizontalalignment='center', verticalalignment='center', transform=plt.gca().transAxes)
                plt.title("Overall Distribution (No Data)", color='blue', fontsize=16)
                
    plt.tight_layout()
    graph_path = os.path.join(STATIC_FOLDER, filename)
    plt.savefig(graph_path)
    plt.close()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part"
    file = request.files['file']
    if file.filename == '':
        return "No selected file"
    if not file.filename.endswith('.xlsx'):
        return render_template('index.html', error="Only Excel files (.xlsx) are allowed!")

    # Save uploaded file with unique name
    unique_name = f"data_{uuid.uuid4().hex}.xlsx"
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_name)
    file.save(file_path)

    # Save file path in session so we can reload later
    session['last_file'] = file_path

    # Read Excel
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        return f"Error reading Excel file: {e}"

    # Check required columns
    required_columns = ['Month', 'Income', 'Expense']
    if not all(col in df.columns for col in required_columns):
        return f"Excel file must contain columns: {required_columns}"

    # Generate initial line graph (default)
    generate_graph(df, 'line', 'graph_line.png')
    
    # Summary stats
    highest_income_idx = df['Income'].idxmax()
    lowest_income_idx = df['Income'].idxmin()
    highest_expense_idx = df['Expense'].idxmax()
    lowest_expense_idx = df['Expense'].idxmin()

    summary = {
        'months': df['Month'].tolist(),
        'highest_income_month': df.at[highest_income_idx, 'Month'],
        'highest_income': df.at[highest_income_idx, 'Income'],
        'lowest_income_month': df.at[lowest_income_idx, 'Month'],
        'lowest_income': df.at[lowest_income_idx, 'Income'],
        'highest_expense_month': df.at[highest_expense_idx, 'Month'],
        'highest_expense': df.at[highest_expense_idx, 'Expense'],
        'lowest_expense_month': df.at[lowest_expense_idx, 'Month'],
        'lowest_expense': df.at[lowest_expense_idx, 'Expense'],
        'avg_income': round(df['Income'].mean(), 2),
        'avg_expense': round(df['Expense'].mean(), 2)
    }

    return render_template('result.html', **summary, graph_file='graph_line.png', graph_type='line')

@app.route('/result', methods=['GET', 'POST'])
def show_result_graph():
    if 'last_file' not in session:
        return redirect(url_for('index'))
    
    file_path = session['last_file']
    df = pd.read_excel(file_path)

    # Determine graph type based on GET or POST request
    if request.method == 'GET':
        graph_type = 'line' # Default to line chart for GET requests
    else: # POST request
        graph_type = request.form.get('graph_type', 'line')
    
    filename = f"graph_{graph_type}.png"
    generate_graph(df, graph_type, filename)

    # Summary stats (re-calculated to pass back to template)
    highest_income_idx = df['Income'].idxmax()
    lowest_income_idx = df['Income'].idxmin()
    highest_expense_idx = df['Expense'].idxmax()
    lowest_expense_idx = df['Expense'].idxmin()

    summary = {
        'months': df['Month'].tolist(),
        'highest_income_month': df.at[highest_income_idx, 'Month'],
        'highest_income': df.at[highest_income_idx, 'Income'],
        'lowest_income_month': df.at[lowest_income_idx, 'Month'],
        'lowest_income': df.at[lowest_income_idx, 'Income'],
        'highest_expense_month': df.at[highest_expense_idx, 'Month'],
        'highest_expense': df.at[highest_expense_idx, 'Expense'],
        'lowest_expense_month': df.at[lowest_expense_idx, 'Month'],
        'lowest_expense': df.at[lowest_expense_idx, 'Expense'],
        'avg_income': round(df['Income'].mean(), 2),
        'avg_expense': round(df['Expense'].mean(), 2)
    }
    
    return render_template('result.html', **summary, graph_file=filename, graph_type=graph_type)

@app.route('/month_detail', methods=['POST'])
def month_detail():
    if 'last_file' not in session:
        return "No file available. Please upload a file first."

    file_path = session['last_file']
    df = pd.read_excel(file_path)

    selected_month = request.form['month']
    graph_type = request.form.get('month_graph_type', 'column') # Default to column for month detail
    row = df[df['Month'] == selected_month].iloc[0].to_dict() # Convert to dict for easier passing

    # Generate graph for selected month
    month_graph_filename = f'month_graph_{selected_month}_{graph_type}.png'
    generate_graph(df, graph_type, month_graph_filename, is_month_detail=True, selected_month_data=row)

    return render_template(
        'month_detail.html',
        month=selected_month,
        income=row['Income'],
        expense=row['Expense'],
        graph_file=month_graph_filename,
        graph_type=graph_type
    )

def find_available_port(start_port=5000):
    """
    Finds an available port by checking for a listening socket.
    """
    port = start_port
    while True:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(('localhost', port))
                return port
        except OSError:
            port += 1
            if port > 6000: # Safety break to avoid infinite loop
                raise RuntimeError("Could not find an available port in the range 5000-6000.")

if __name__ == '__main__':
    port = find_available_port()
    print(f"Server is running on http://127.0.0.1:{port}")
    app.run(debug=True, port=port)