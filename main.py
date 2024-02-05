from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
from flask_bootstrap import Bootstrap
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
import base64

app = Flask(__name__)

excel_name = "IT-Assignment-Three.xlsx"
Bootstrap(app)
xl = pd.ExcelFile(excel_name)
sheets = xl.sheet_names

product_df = pd.read_excel('IT-Assignment-Two.xlsx', sheet_name='Sheet1', skiprows=1)
product_df['ID'] = range(1, len(product_df) + 1)

student_df = pd.read_excel('IT-Assignment-One.xlsx', sheet_name='Sheet1', skiprows=1)

student_df = student_df.iloc[:-2]

student_df['ID'] = range(1, len(student_df) + 1)
student_df['Full Name'] = student_df['First Name'] + ' ' + student_df['Last Name']
@app.route('/')
def home():
    return render_template('index.html')


@app.route('/employee_index')
def employee_index():
    return render_template('employee_index.html', sheets=sheets)


@app.route('/sheet/<sheet_name>')
def display_sheet(sheet_name):
    employee_df = pd.read_excel(excel_name, sheet_name=sheet_name)
    employee_df = employee_df.iloc[:-2]
    table_html = employee_df.to_html(index=False, classes='table table-striped')

    graphs = []
    for column in employee_df.columns:
        if employee_df[column].dtype in ['int64', 'float64']:
            plt.figure()
            employee_df[column].plot(kind='bar')
            plt.xlabel('Index')
            plt.ylabel(column)
            plt.title(f'Graph for {column}')
            plt.tight_layout()

            img_buffer = BytesIO()
            plt.savefig(img_buffer, format='png')
            img_buffer.seek(0)

            graph_html = f'<img src="data:image/png;base64,{base64.b64encode(img_buffer.read()).decode()}">'
            graphs.append(graph_html)

    return render_template('sheet.html', sheet_name=sheet_name, table=table_html, graphs=graphs)


@app.route('/product_index')
def product_index():
    return render_template('product_index.html')


@app.route('/product_search', methods=['POST'])
def product_search():
    search_term = request.form['search']
    filtered_df = product_df[product_df['Product Name'].str.contains(search_term, case=False)].reset_index(drop=True)
    products = filtered_df.to_dict(orient='records')
    return render_template('products_result.html', products=products)


@app.route('/product/<int:id>')
def product(id):
    product_info = product_df[product_df['ID'] == id].iloc[0].to_dict()  # Filter by the unique identifier column
    return render_template('product.html', product_info=product_info)


@app.route('/student_index')
def student_index():
        return render_template('student_index.html')

@app.route('/student_search', methods=['POST'])
def student_search():
    search_term = request.form['search']
    filtered_df = student_df[student_df['Full Name'].str.contains(search_term, case=False)].reset_index(drop=True)
    students = filtered_df.to_dict(orient='records')
    return render_template('students_result.html', students=students)


@app.route('/student/<int:id>')
def student(id):
    student_info = student_df[student_df['ID'] == id].iloc[0].to_dict()  # Filter by the unique identifier column
    return render_template('student.html', student_info=student_info)

if __name__ == '__main__':
    app.run(debug=True)
