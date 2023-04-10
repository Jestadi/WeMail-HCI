from flask import Flask, render_template
from myprogram import start_program

app = Flask(__name__)
app.template_folder = '/Users/rahuljestadi/Desktop/PROJECTS/WE-MAIL/templates/'
app.static_folder = '/Users/rahuljestadi/Desktop/PROJECTS/WE-MAIL/static/'

@app.route('/')
def home():
    return render_template('home.html')

@app.route('/start-program')
def start_program_endpoint():
    start_program()
    return render_template('home2.html')

if __name__== "__main__":
    app.run(debug=True)