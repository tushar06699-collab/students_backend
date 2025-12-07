from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from pymongo import MongoClient
from bson.objectid import ObjectId

app = Flask(__name__)
CORS(app)

# ------------------ MONGODB ATLAS SETUP ------------------
MONGO_URI = "mongodb+srv://school_students:Tushar2007@cluster0.upoywck.mongodb.net/school_erp?retryWrites=true&w=majority"
client = MongoClient(MONGO_URI)
db = client["school_erp"]
students_col = db["students"]

# ------------------ GET ALL STUDENTS ------------------
@app.route("/students", methods=["GET"])
def get_students():
    students = list(students_col.find({}, {"_id": 1, "rollno": 1, "panno": 1,
                                           "student_name": 1, "father_name": 1, "mother_name": 1,
                                           "class_name": 1, "gender": 1, "aadharno": 1, "session": 1}))
    for s in students:
        s["id"] = str(s.pop("_id"))
    # Sort by class number and roll number
    def class_sort_key(s):
        import re
        cls = s.get("class_name", "")
        cls_num = int(re.search(r"\d+", cls).group()) if re.search(r"\d+", cls) else -1
        roll = int(s.get("rollno", 0)) if s.get("rollno") else 0
        return (cls_num, roll)
    students.sort(key=class_sort_key)
    return jsonify(students)

# ------------------ ADD STUDENT ------------------
@app.route("/students", methods=["POST"])
def add_student():
    data = request.json
    student = {
        "rollno": str(data.get("rollno", "")),
        "panno": str(data.get("panno", "")),
        "student_name": str(data.get("student_name", "")),
        "father_name": str(data.get("father_name", "")),
        "mother_name": str(data.get("mother_name", "")),
        "class_name": str(data.get("class_name", "")),
        "gender": str(data.get("gender", "")),
        "aadharno": str(data.get("aadharno", "")),
        "session": str(data.get("session", ""))
    }
    students_col.insert_one(student)
    return jsonify({"message": "Student added successfully"})

# ------------------ IMPORT EXCEL ------------------
@app.route("/import_excel", methods=["POST"])
def import_excel():
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file part"}), 400

        file = request.files["file"]
        df = pd.read_excel(file)

        required_cols = ["rollno", "panno", "student_name", "father_name",
                         "mother_name", "class_name", "gender", "aadharno", "session"]

        for col in required_cols:
            if col not in df.columns:
                return jsonify({"error": f"Missing column: {col}"}), 400

        students = []
        for _, row in df.iterrows():
            roll = row["rollno"]
            if pd.isnull(roll):
                roll = ""
            else:
                roll = str(int(roll)) if isinstance(roll, float) else str(roll)

            student = {
                "rollno": roll,
                "panno": str(row["panno"]),
                "student_name": str(row["student_name"]),
                "father_name": str(row["father_name"]),
                "mother_name": str(row["mother_name"]),
                "class_name": str(row["class_name"]),
                "gender": str(row["gender"]),
                "aadharno": str(row["aadharno"]),
                "session": str(row.get("session", ""))
            }
            students.append(student)

        if students:
            students_col.insert_many(students)

        return jsonify({"message": "Excel imported successfully"})

    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ------------------ DELETE ONE STUDENT ------------------
@app.route("/students/<id>", methods=["DELETE"])
def delete_student(id):
    students_col.delete_one({"_id": ObjectId(id)})
    return jsonify({"message": "Student deleted"})

# ------------------ DELETE SELECTED STUDENTS ------------------
@app.route("/students/delete_selected", methods=["POST"])
def delete_selected():
    data = request.json
    ids = data.get("ids", [])
    if not ids:
        return jsonify({"error": "No IDs provided"}), 400

    object_ids = [ObjectId(sid) for sid in ids]
    students_col.delete_many({"_id": {"$in": object_ids}})
    return jsonify({"message": "Selected students deleted"})

# ------------------ DELETE ALL STUDENTS ------------------
@app.route("/students/delete_all", methods=["DELETE"])
def delete_all_students():
    students_col.delete_many({})
    return jsonify({"message": "All students deleted"})

# ------------------ DOWNLOAD EXCEL FORMAT ------------------
@app.route("/download_format", methods=["GET"])
def download_format():
    wb = Workbook()
    ws = wb.active
    ws.title = "Student Format"

    headers = [
        "rollno", "panno", "student_name", "father_name", "mother_name",
        "class_name", "gender", "aadharno", "session"
    ]
    ws.append(headers)

    # CLASS DROPDOWN
    class_list = [
        "Nursery", "LKG", "UKG",
        "1st", "2nd", "3rd", "4th", "5th",
        "6th", "7th", "8th",
        "9th", "10th",
        "11th Arts", "11th Commerce", "11th Science",
        "12th Arts", "12th Commerce", "12th Science"
    ]
    class_options = ",".join(class_list)
    dv_class = DataValidation(type="list", formula1=f'"{class_options}"')
    ws.add_data_validation(dv_class)
    dv_class.add("F2:F500")

    # GENDER DROPDOWN
    dv_gender = DataValidation(type="list", formula1='"Male,Female,Other"')
    ws.add_data_validation(dv_gender)
    dv_gender.add("G2:G500")

    filepath = "student_format.xlsx"
    wb.save(filepath)

    return send_file(filepath, as_attachment=True)

@app.route("/", methods=["GET"])
def home():
    return "Backend Running", 200


# ------------------ RUN APP ------------------
if __name__ == "__main__":
    app.run(debug=True)
