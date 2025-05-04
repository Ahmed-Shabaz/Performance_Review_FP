from flask import Flask, request, render_template, redirect, url_for, flash, jsonify
from flask_sqlalchemy import SQLAlchemy
import pandas as pd
import os
from flaskwebgui import FlaskUI
import win32api
import win32net
from datetime import datetime
import pytz
# Constants
# INPUT_DIR = r"\\bfs\Projects-Main\VBA\BMS_Project Management\Exe"
INPUT_DIR = r"C:\\Users\\SHABHAZ AHMED\\OneDrive\\Desktop\\Performance_Review_Final_Plain\\instance"
DEFAULT_EXCEL_FILE_PATH = os.path.join(INPUT_DIR, "reviews.db")
EMPLOYEE_DATA_FILE_PATH = 'Employee_Data.xlsx'


IST = pytz.timezone("Asia/Kolkata")
# Initialize Flask app
app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{DEFAULT_EXCEL_FILE_PATH}'
app.secret_key = '#$%^&*'
db = SQLAlchemy(app)

def username():
    try:
        dc_name = win32net.NetGetAnyDCName()
        username = win32api.GetUserName()
        user_info = win32net.NetUserGetInfo(dc_name, username, 2)
        full_name = user_info["full_name"]
        return full_name
    except Exception as e:
        return "Unknown User"

class Review(db.Model):
    ID = db.Column(db.Integer, primary_key=True)
    Employee_Name = db.Column(db.String(100))
    Department = db.Column(db.String(100))
    Employee_ID = db.Column(db.String(100))
    Designation = db.Column(db.String(100))
    Reporting_Manager = db.Column(db.String(100))
    Date = db.Column(db.String(100))
    Role = db.Column(db.String(100))
    Sourcing = db.Column(db.String(100))
    Quality = db.Column(db.String(100))
    Quantity = db.Column(db.String(100))
    Domain_Knowledge = db.Column(db.String(100))
    Extra_Miler = db.Column(db.String(100))
    Attendance = db.Column(db.String(100))
    Achieved_Goals = db.Column(db.String(100))
    Next_Goals = db.Column(db.String(100))
    Comments = db.Column(db.String(100))
    Submitted_By = db.Column(db.String(200), nullable=False)
    Submitted_Date = db.Column(db.String(200), nullable=False)

    def __init__(self, Employee_Name, Department, Employee_ID, Designation, Reporting_Manager, Date, Role, Sourcing, Quality, Quantity, Domain_Knowledge, Extra_Miler, Attendance, Achieved_Goals, Next_Goals, Comments,Submitted_By, Submitted_Date):
        self.Employee_Name = Employee_Name
        self.Department = Department
        self.Employee_ID = Employee_ID
        self.Designation = Designation
        self.Reporting_Manager = Reporting_Manager
        self.Date = Date
        self.Role = Role
        self.Sourcing = Sourcing
        self.Quality = Quality
        self.Quantity = Quantity
        self.Domain_Knowledge = Domain_Knowledge
        self.Extra_Miler = Extra_Miler
        self.Attendance = Attendance
        self.Achieved_Goals = Achieved_Goals
        self.Next_Goals = Next_Goals
        self.Comments = Comments
        self.Submitted_By = Submitted_By
        self.Submitted_Date = Submitted_Date

@app.route("/")
def index():
    df = pd.read_excel(EMPLOYEE_DATA_FILE_PATH)
    Employee_Names = df['Employee Name'].tolist()
    return render_template("index.html", Employee_Names=Employee_Names)

@app.route("/get_employee_data", methods=["POST"])
def get_employee_data():
    Employee_Name = request.form['Employee_Name']
    df = pd.read_excel(EMPLOYEE_DATA_FILE_PATH)
    employee_data = df[df['Employee Name'] == Employee_Name].to_dict(orient='records')[0]
    return jsonify(employee_data)

@app.route("/submit", methods=["POST"])
def submit():
    Employee_Name = request.form['Employee_Name']
    Department = request.form['Department']
    Employee_ID = request.form['Employee_ID']
    Designation = request.form['Designation']
    Reporting_Manager = request.form['Reporting_Manager']
    Date = request.form['Date']
    Role = request.form['Role']
    Sourcing = request.form.get('Sourcing', '')
    Quality = request.form.get('Quality', '')
    Quantity = request.form.get('Quantity', '')
    Domain_Knowledge = request.form.get('Domain_Knowledge', '')
    Extra_Miler = request.form.get('Extra_Miler', '')
    Attendance = request.form.get('Attendance', '')
    Achieved_Goals = request.form['Achieved_Goals']
    Next_Goals = request.form['Next_Goals']
    Comments = request.form['Comments']
    Submitted_By = username()
    Submitted_Date = datetime.now(IST).strftime("%d-%m-%Y %H:%M:%S")
    new_review = Review(Employee_Name, Department, Employee_ID, Designation, Reporting_Manager, Date, Role, Sourcing, Quality, Quantity, Domain_Knowledge, Extra_Miler, Attendance, Achieved_Goals, Next_Goals, Comments, Submitted_By, Submitted_Date)
    db.session.add(new_review)
    db.session.commit()

    # Prepare data for writing to Excel
    data = {
        'Employee Name': [Employee_Name],
        'Department': [Department],
        'Employee ID': [Employee_ID],
        'Designation': [Designation],
        'Reporting Manager': [Reporting_Manager],
        'Date': [Date],
        'Role': [Role],
        'Sourcing': [Sourcing],
        'Quality': [Quality],
        'Quantity': [Quantity],
        'Domain Knowledge': [Domain_Knowledge],
        'Extra Miler': [Extra_Miler],
        'Attendance / Punctuality': [Attendance],
        'Achieved Goals Set in Previous Review': [Achieved_Goals],
        'Goals for Next Review Period': [Next_Goals],
        'Comments': [Comments]
    }

    # df = pd.DataFrame(data)
    # file_path = 'Performance_Reviews.xlsx'

    # file_exists = os.path.isfile(file_path)

    # if file_exists:
    #     with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
    #         start_row = writer.sheets['Sheet1'].max_row
    #         df.to_excel(writer, index=False, header=False, startrow=start_row)
    # else:
    #     with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    #         df.to_excel(writer, index=False)

    flash('Submitted successfully!', 'success')
    return redirect(url_for('success'))

@app.route("/success")
def success():
    return render_template("success.html")

if __name__ == "__main__":
    with app.app_context():
        db.create_all()


    ui = FlaskUI(app=app, server="flask")
    ui.run()