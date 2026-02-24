from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from pymongo import MongoClient
from bson.objectid import ObjectId
import cloudinary
import cloudinary.uploader
import requests
import tempfile
import os
import zipfile
import shutil

app = Flask(__name__)
CORS(app)

# ================= CLOUDINARY CONFIG =================
cloudinary.config(
    cloud_name="djq1jjet6",
    api_key="635839659646439",
    api_secret="jx2ysIgjN6zGC71X23EvDS_9faI"
)

# ================= MONGODB CONFIG =================
MONGO_URI = "mongodb+srv://school_students:Tushar2007@cluster0.upoywck.mongodb.net/school_erp?retryWrites=true&w=majority"
client = MongoClient(MONGO_URI)
db = client["school_erp"]
students_col = db["students"]


def to_bool(value):
    if isinstance(value, bool):
        return value
    text = str(value or "").strip().lower()
    return text in {"1", "true", "yes", "y", "on"}

# ================= IMAGE FROM URL =================
def upload_to_cloudinary(image_url):
    if not image_url:
        return ""

    try:
        r = requests.get(image_url, timeout=10)
        if r.status_code != 200:
            return ""

        with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as f:
            f.write(r.content)
            temp_path = f.name

        result = cloudinary.uploader.upload(
            temp_path,
            folder="school_students"
        )

        os.remove(temp_path)
        return result.get("secure_url", "")

    except Exception as e:
        print("Image upload error:", e)
        return ""



def normalize_admission_no(value):
    """Normalize admission number so 1001 and 1001.0 map to same key."""
    if value is None:
        return ""
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return ""

    # Excel often sends numeric IDs as float text (e.g. "1001.0")
    if text.endswith(".0"):
        text = text[:-2]

    return text.strip()


def build_zip_image_map(extract_dir):
    """Map normalized admission_no -> image path (supports nested folders + any case extension)."""
    image_map = {}
    allowed = {".jpg", ".jpeg", ".png", ".webp"}

    for root, _, files in os.walk(extract_dir):
        for file_name in files:
            base, ext = os.path.splitext(file_name)
            if ext.lower() not in allowed:
                continue

            key = normalize_admission_no(base)
            if not key:
                continue

            full_path = os.path.join(root, file_name)
            # First match wins, avoids random overwrite
            if key not in image_map:
                image_map[key] = full_path

    return image_map

# ================= IMPORT EXCEL + ZIP IMAGES =================
@app.route("/import_excel_with_images", methods=["POST"])
def import_excel_with_images():
    if "excel" not in request.files:
        return jsonify({"error": "Excel file required"}), 400

    excel = request.files["excel"]
    zip_file = request.files.get("images")

    df = pd.read_excel(excel)
    extract_dir = tempfile.mkdtemp()

    matched_photos = 0
    image_map = {}

    try:
        if zip_file and zip_file.filename:
            with zipfile.ZipFile(zip_file, "r") as zip_ref:
                zip_ref.extractall(extract_dir)
            image_map = build_zip_image_map(extract_dir)

        students = []

        for _, row in df.iterrows():
            admission_no = normalize_admission_no(row.get("admission_no", ""))
            photo_url = ""

            img_path = image_map.get(admission_no)
            if img_path and os.path.exists(img_path):
                try:
                    res = cloudinary.uploader.upload(
                        img_path,
                        folder="school_students"
                    )
                    photo_url = res.get("secure_url", "")
                    if photo_url:
                        matched_photos += 1
                except Exception as e:
                    print(f"Photo upload error for admission_no={admission_no}:", e)

            students.append({
                "admission_no": admission_no,
                "rollno": normalize_admission_no(row.get("rollno", "")),
                "panno": str(row.get("panno", "")).strip(),
                "student_name": str(row.get("student_name", "")).strip(),
                "father_name": str(row.get("father_name", "")).strip(),
                "mother_name": str(row.get("mother_name", "")).strip(),
                "class_name": str(row.get("class_name", "")).strip(),
                "section": str(row.get("section", "")).strip(),
                "gender": str(row.get("gender", "")).strip(),
                "dob": str(row.get("dob", "")).strip(),
                "aadharno": normalize_admission_no(row.get("aadharno", "")),
                "parent_mobile": normalize_admission_no(row.get("parent_mobile", "")),
                "parent_email": str(row.get("parent_email", "")).strip(),
                "address": str(row.get("address", "")).strip(),
                "session": str(row.get("session", "")).strip(),
                "new_admission": to_bool(row.get("new_admission", False)),
                "photo_url": photo_url
            })

        if students:
            students_col.insert_many(students)

        return jsonify({
            "message": f"Imported {len(students)} students successfully",
            "students_imported": len(students),
            "photos_matched": matched_photos,
            "photos_missing": max(0, len(students) - matched_photos)
        })
    finally:
        shutil.rmtree(extract_dir, ignore_errors=True)

# ================= GET ALL STUDENTS =================

# ================= ADD STUDENT (FORM + IMAGE) =================
@app.route("/students", methods=["POST"])
def add_student():
    form = request.form
    photo = request.files.get("photo")

    photo_url = ""
    if photo:
        res = cloudinary.uploader.upload(
            photo,
            folder="school_students"
        )
        photo_url = res["secure_url"]

    student = {
        "admission_no": form.get("admission_no", ""),
        "rollno": form.get("rollno", ""),
        "panno": form.get("panno", ""),
        "student_name": form.get("student_name", ""),
        "father_name": form.get("father_name", ""),
        "mother_name": form.get("mother_name", ""),
        "class_name": form.get("class_name", ""),
        "section": form.get("section", ""),
        "gender": form.get("gender", ""),
        "dob": form.get("dob", ""),
        "session": form.get("session", ""),
        "parent_mobile": form.get("parent_mobile", ""),
        "parent_email": form.get("parent_email", ""),
        "address": form.get("address", ""),
        "new_admission": to_bool(form.get("new_admission", "false")),
        "photo_url": photo_url
    }

    students_col.insert_one(student)
    return jsonify({"message": "Student added successfully"})

# ================= IMPORT EXCEL (IMAGE URL COLUMN) =================
@app.route("/import_excel", methods=["POST"])
def import_excel():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    df = pd.read_excel(file)

    students = []

    for _, row in df.iterrows():
        cloud_img = upload_to_cloudinary(row.get("photo_url", ""))

        students.append({
            "admission_no": str(row.get("admission_no", "")).strip(),
            "rollno": str(row.get("rollno", "")).strip(),
            "panno": str(row.get("panno", "")).strip(),
            "student_name": str(row.get("student_name", "")).strip(),
            "father_name": str(row.get("father_name", "")).strip(),
            "mother_name": str(row.get("mother_name", "")).strip(),
            "class_name": str(row.get("class_name", "")).strip(),
            "section": str(row.get("section", "")).strip(),
            "dob": str(row.get("dob", "")).strip(),
            "gender": str(row.get("gender", "")).strip(),
            "aadharno": str(row.get("aadharno", "")).strip(),
            "parent_mobile": str(row.get("parent_mobile", "")).strip(),
            "parent_email": str(row.get("parent_email", "")).strip(),
            "address": str(row.get("address", "")).strip(),
            "new_admission": to_bool(row.get("new_admission", False)),
            "photo_url": cloud_img,
            "session": str(row.get("session", "")).strip()
        })

    if students:
        students_col.insert_many(students)

    return jsonify({"message": "Students imported successfully"})



@app.route("/students/by-admission/<admission_no>", methods=["GET"])
def get_student_by_admission(admission_no):
    admission_no = str(admission_no or "").strip()
    if not admission_no:
        return jsonify({"success": False, "message": "Missing admission number"}), 400

    student = students_col.find_one({"admission_no": admission_no})
    if not student:
        return jsonify({"success": False, "message": "Student not found"}), 404

    student["_id"] = str(student["_id"])
    return jsonify({"success": True, "student": student})

@app.route("/students", methods=["GET"])
def get_students():
    session = str(request.args.get("session", "")).strip()
    class_name = str(request.args.get("class_name", request.args.get("class", ""))).strip()

    q = {}
    if session:
        q["session"] = session
    if class_name:
        q["class_name"] = class_name

    students = list(students_col.find(q))
    for s in students:
        s["_id"] = str(s["_id"])
    return jsonify(students)


@app.route("/students/<id>", methods=["PUT"])
def update_student(id):
    try:
        update_data = {}

        # Support both JSON updates and multipart form updates with photo upload.
        if request.content_type and "multipart/form-data" in request.content_type:
            form = request.form
            photo = request.files.get("photo")

            fields = [
                "admission_no", "rollno", "panno", "student_name", "father_name", "mother_name",
                "class_name", "section", "gender", "dob", "session", "aadharno",
                "parent_mobile", "parent_email", "address", "photo_url", "new_admission"
            ]
            for f in fields:
                if f in form:
                    update_data[f] = form.get(f, "")

            if "new_admission" in update_data:
                update_data["new_admission"] = to_bool(update_data["new_admission"])

            if photo:
                res = cloudinary.uploader.upload(photo, folder="school_students")
                update_data["photo_url"] = res.get("secure_url", "")
        else:
            update_data = request.json or {}
            if "new_admission" in update_data:
                update_data["new_admission"] = to_bool(update_data["new_admission"])

        students_col.update_one(
            {"_id": ObjectId(id)},
            {"$set": update_data}
        )
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400

@app.route("/students/<id>", methods=["GET"])
def get_student(id):
    try:
        student = students_col.find_one({"_id": ObjectId(id)})
        if not student:
            return jsonify({"error": "Student not found"}), 404

        student["_id"] = str(student["_id"])
        return jsonify(student)
    except:
        return jsonify({"error": "Invalid ID"}), 400

# ================= DELETE ONE =================
@app.route("/students/<id>", methods=["DELETE"])
def delete_student(id):
    students_col.delete_one({"_id": ObjectId(id)})
    return jsonify({"message": "Student deleted"})

# ================= DELETE ALL =================
@app.route("/students/delete_all", methods=["DELETE"])
def delete_all_students():
    students_col.delete_many({})
    return jsonify({"message": "All students deleted"})

# ================= DOWNLOAD EXCEL FORMAT =================
@app.route("/download_format", methods=["GET"])
def download_format():
    wb = Workbook()
    ws = wb.active
    ws.title = "Student Import Format"

    headers = [
        "admission_no", "rollno", "panno", "student_name",
        "father_name", "mother_name", "class_name", "section",
        "dob", "gender", "aadharno",
        "parent_mobile", "parent_email", "address",
         "session"
    ]
    ws.append(headers)

    dv_gender = DataValidation(type="list", formula1='"Male,Female,Other"')
    ws.add_data_validation(dv_gender)
    dv_gender.add("J2:J1000")

    file_path = "student_import_format.xlsx"
    wb.save(file_path)
    return send_file(file_path, as_attachment=True)
@app.route("/portal/student/<student_id>", methods=["GET"])
def portal_get_student(student_id):
    try:
        student = students_col.find_one({"_id": ObjectId(student_id)})

        if not student:
            return jsonify({"success": False, "message": "Student not found"}), 404

        student["_id"] = str(student["_id"])

        return jsonify({
            "success": True,
            "student": {
                "id": student["_id"],
                "name": student.get("student_name", ""),
                "class_name": student.get("class_name", ""),
                "section": student.get("section", ""),
                "roll": student.get("rollno", ""),
                "photo_url": student.get("photo_url", ""),
                "session": student.get("session", ""),
                "eligible": True,
                "release_rollno": True,
                "release_result": True
            }
        })

    except Exception as e:
        return jsonify({"success": False, "message": "Invalid ID"}), 400

# ================= HOME =================
@app.route("/", methods=["GET"])
def home():
    return "Student Backend Running", 200

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
