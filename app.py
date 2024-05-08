#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from flask import Flask, render_template, request
from CodeCLH import run_process

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    if request.method == 'POST':
        # Get the form data
        from_date = request.form['from_date']
        to_date = request.form['to_date']
        email = request.form['email']

        # Splitting the date strings into day, month, and year
        a3, a2, a1 = map(int, from_date.split('-'))
        b3, b2, b1 = map(int, to_date.split('-'))

        # Call the function to run the process with form inputs
        output_files = run_process(a1, a2, a3, b1, b2, b3, email)

        # Check if output_files is not empty (i.e., report was generated successfully)
        report_generated = bool(output_files)

        # Pass the result to the template
        return render_template('index.html', report_generated=report_generated, output_files=output_files)

if __name__ == '__main__':
    app.run(port=8000, debug=True)

