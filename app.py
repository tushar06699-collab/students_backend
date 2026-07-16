from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from datetime import datetime, timedelta
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
import hashlib
import hmac
import math
import random

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
teachers_col = db["teachers"]
student_edit_requests_col = db["student_edit_requests"]
teacher_attendance_col = db["teacher_attendance"]

students_col.create_index("regno", unique=True, sparse=True)

IST_OFFSET = timedelta(hours=5, minutes=30)
QR_ROTATE_SECONDS = 30
QR_SECRET = os.environ.get("TEACHER_ATTENDANCE_QR_SECRET", "teacher-attendance-secret")
ATTENDANCE_MIN_OUT_MINUTES = int(os.environ.get("ATTENDANCE_MIN_OUT_MINUTES", "30"))
ATTENDANCE_SITE_LAT = os.environ.get("ATTENDANCE_SITE_LAT", "").strip()
ATTENDANCE_SITE_LON = os.environ.get("ATTENDANCE_SITE_LON", "").strip()
ATTENDANCE_SITE_RADIUS_M = float(os.environ.get("ATTENDANCE_SITE_RADIUS_M", "300") or 300)
ATTENDANCE_REQUIRE_GPS = os.environ.get("ATTENDANCE_REQUIRE_GPS", "0").strip().lower() in {"1", "true", "yes", "on", "y"}
ATTENDANCE_LATE_AFTER = os.environ.get("ATTENDANCE_LATE_AFTER", "09:15").strip()
TEXTBEE_ENABLED = os.environ.get("TEXTBEE_ENABLED", "0").strip().lower() in {"1", "true", "yes", "on", "y"}
TEXTBEE_API_BASE = os.environ.get("TEXTBEE_API_BASE", "https://api.textbee.dev").strip().rstrip("/")
TEXTBEE_DEVICE_ID = os.environ.get("TEXTBEE_DEVICE_ID", "").strip()
TEXTBEE_API_KEY = os.environ.get("TEXTBEE_API_KEY", "").strip()


def now_ist():
    return datetime.utcnow() + IST_OFFSET


def to_iso(dt):
    return dt.isoformat(timespec="seconds")


def parse_hhmm(value, fallback_h=9, fallback_m=15):
    text = str(value or "").strip()
    try:
        parts = text.split(":")
        if len(parts) >= 2:
            h = int(parts[0])
            m = int(parts[1])
            if 0 <= h <= 23 and 0 <= m <= 59:
                return h, m
    except Exception:
        pass
    return fallback_h, fallback_m


def qr_slot_for(dt):
    return int(dt.timestamp() // QR_ROTATE_SECONDS)


def qr_signature(slot):
    msg = str(slot).encode("utf-8")
    key = QR_SECRET.encode("utf-8")
    return hmac.new(key, msg, hashlib.sha256).hexdigest()[:12]


def build_qr_code(slot):
    return f"TCH-ATTN|{slot}|{qr_signature(slot)}"


def verify_qr_code(code):
    parts = str(code or "").strip().split("|")
    if len(parts) != 3 or parts[0] != "TCH-ATTN":
        return False
    try:
        slot = int(parts[1])
    except Exception:
        return False
    sig = parts[2].strip()
    now_slot = qr_slot_for(now_ist())
    for probe in (now_slot - 1, now_slot, now_slot + 1):
        if slot == probe and hmac.compare_digest(sig, qr_signature(probe)):
            return True
    return False


def distance_meters(lat1, lon1, lat2, lon2):
    # Haversine formula
    r = 6371000.0
    p1 = math.radians(lat1)
    p2 = math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def normalize_phone_for_sms(phone):
    p = str(phone or "").strip()
    if not p:
        return ""
    digits = "".join(ch for ch in p if ch.isdigit() or ch == "+")
    if digits.startswith("+"):
        return digits
    num = "".join(ch for ch in digits if ch.isdigit())
    if len(num) == 10:
        return "+91" + num
    if len(num) == 12 and num.startswith("91"):
        return "+" + num
    if len(num) > 0:
        return "+" + num
    return ""


def send_textbee_sms(recipient, message):
    if not TEXTBEE_ENABLED:
        return {"sent": False, "reason": "disabled"}
    if not TEXTBEE_DEVICE_ID or not TEXTBEE_API_KEY:
        return {"sent": False, "reason": "missing_config"}
    raw = str(recipient or "").strip()
    to = normalize_phone_for_sms(raw)
    if not to:
        return {"sent": False, "reason": "invalid_phone", "recipient_raw": raw}
    url = f"{TEXTBEE_API_BASE}/api/v1/gateway/devices/{TEXTBEE_DEVICE_ID}/send-sms"
    headers = {
        "x-api-key": TEXTBEE_API_KEY,
        "Content-Type": "application/json",
    }
    payload = {
        "recipients": [to],
        "message": str(message or "").strip(),
    }
    try:
        resp = requests.post(url, json=payload, headers=headers, timeout=15)
        ok = 200 <= resp.status_code < 300
        data = {}
        try:
            data = resp.json()
        except Exception:
            data = {"raw": resp.text[:500]}
        return {"sent": ok, "status": resp.status_code, "response": data, "recipient": to, "recipient_raw": raw}
    except Exception as ex:
        return {"sent": False, "reason": str(ex), "recipient": to, "recipient_raw": raw}


def to_bool(value):
    if isinstance(value, bool):
        return value
    text = str(value or "").strip().lower()
    return text in {"1", "true", "yes", "y", "on"}


def session_variants(session_value):
    """
    Build tolerant session variants so queries match common formats:
    2025_26, 2025-26, 2025/26, 2025 26, 2025_2026
    """
    s = str(session_value or "").strip()
    if not s:
        return []

    variants = {s}
    compact = s.replace(" ", "")
    variants.add(compact)
    variants.add(compact.replace("-", "_"))
    variants.add(compact.replace("/", "_"))
    variants.add(compact.replace("_", "-"))
    variants.add(compact.replace("_", "/"))
    variants.add(compact.replace("-", "/"))
    variants.add(compact.replace("/", "-"))

    # Try start/end year normalization if pattern contains a separator.
    for sep in ["_", "-", "/", " "]:
        if sep in s:
            parts = [p for p in s.split(sep) if p]
            if len(parts) >= 2 and parts[0].isdigit():
                start = parts[0]
                end = parts[1]
                if len(end) == 2:
                    full_end = start[:2] + end
                    variants.add(f"{start}_{end}")
                    variants.add(f"{start}-{end}")
                    variants.add(f"{start}/{end}")
                    variants.add(f"{start}_{full_end}")
                    variants.add(f"{start}-{full_end}")
                    variants.add(f"{start}/{full_end}")
                break

    return list(variants)

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


def normalize_photo_id(value):
    """Normalize photo_id similar to admission numbers."""
    return normalize_admission_no(value)


def generate_student_regno(reserved=None):
    """Create a unique 6 digit student registration number."""
    if reserved is None:
        reserved = set()
    for _ in range(100):
        regno = str(random.randint(100000, 999999))
        if regno not in reserved and not students_col.find_one({"regno": regno}, {"_id": 1}):
            reserved.add(regno)
            return regno
    raise RuntimeError("Unable to generate unique registration number")


def ensure_student_regno(student):
    if not student:
        return student
    regno = str(student.get("regno", "")).strip()
    if len(regno) == 6 and regno.isdigit():
        student["regno"] = regno
        return student
    regno = generate_student_regno()
    students_col.update_one({"_id": student["_id"]}, {"$set": {"regno": regno}})
    student["regno"] = regno
    return student


def normalize_teacher_code(value):
    """Normalize teacher code to 4 digits when numeric."""
    if value is None:
        return ""
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return ""
    if text.endswith(".0"):
        text = text[:-2]
    digits = "".join(ch for ch in text if ch.isdigit())
    if digits:
        return digits.zfill(4) if len(digits) <= 4 else digits
    return text


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


def build_zip_image_map_with_normalizer(extract_dir, normalizer):
    """Map normalized key -> image path using any custom normalizer."""
    image_map = {}
    allowed = {".jpg", ".jpeg", ".png", ".webp"}

    for root, _, files in os.walk(extract_dir):
        for file_name in files:
            base, ext = os.path.splitext(file_name)
            if ext.lower() not in allowed:
                continue
            key = normalizer(base)
            if not key:
                continue
            full_path = os.path.join(root, file_name)
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
        reserved_regnos = set()

        for _, row in df.iterrows():
            admission_no = normalize_admission_no(row.get("admission_no", ""))
            photo_id = normalize_photo_id(row.get("photo_id", ""))
            photo_url = ""

            img_path = image_map.get(photo_id) or image_map.get(admission_no)
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
                "regno": generate_student_regno(reserved_regnos),
                "photo_id": photo_id,
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
        "regno": generate_student_regno(),
        "admission_no": form.get("admission_no", ""),
        "photo_id": form.get("photo_id", ""),
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
    return jsonify({"message": "Student added successfully", "regno": student["regno"]})

# ================= IMPORT EXCEL (IMAGE URL COLUMN) =================
@app.route("/import_excel", methods=["POST"])
def import_excel():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    df = pd.read_excel(file)

    students = []
    reserved_regnos = set()

    for _, row in df.iterrows():
        cloud_img = upload_to_cloudinary(row.get("photo_url", ""))

        students.append({
            "regno": generate_student_regno(reserved_regnos),
            "photo_id": normalize_photo_id(row.get("photo_id", "")),
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

    ensure_student_regno(student)
    student["_id"] = str(student["_id"])
    return jsonify({"success": True, "student": student})

@app.route("/students", methods=["GET"])
def get_students():
    session = str(request.args.get("session", "")).strip()
    class_name = str(request.args.get("class_name", request.args.get("class", ""))).strip()

    q = {}
    if class_name:
        q["class_name"] = class_name

    # Primary: exact session filter when provided.
    # Compatibility fallback: many legacy rows may have missing/old session values.
    students = []
    reserved_regnos = set()
    if session:
        q_session = dict(q)
        variants = session_variants(session)
        q_session["session"] = {"$in": variants} if variants else session
        students = list(students_col.find(q_session))

        # Optional strict mode: never fallback to all sessions.
        strict = str(request.args.get("strict", "")).strip().lower() in {"1","true","yes","y","on"}
        if not strict:
            # If session-filtered dataset is too small, fallback to all records for the class.
            # This keeps older frontend pages working when session data is inconsistent.
            if len(students) <= 1:
                students = list(students_col.find(q))
    else:
        students = list(students_col.find(q))

    for s in students:
        ensure_student_regno(s)
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

        update_data.pop("regno", None)

        students_col.update_one(
            {"_id": ObjectId(id)},
            {"$set": update_data}
        )
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400


# ================= STUDENT EDIT REQUESTS =================
EDITABLE_FIELDS = {
    "photo_id", "admission_no", "rollno", "panno", "student_name", "father_name",
    "mother_name", "class_name", "section", "gender", "dob", "session",
    "aadharno", "parent_mobile", "parent_email", "address", "photo_url", "new_admission"
}
PHOTO_DATA_FIELD = "photo_data"


def filter_edit_changes(changes):
    if not isinstance(changes, dict):
        return {}
    cleaned = {}
    for k, v in changes.items():
        if k in EDITABLE_FIELDS:
            cleaned[k] = v
        if k == PHOTO_DATA_FIELD:
            cleaned[k] = v
    if "new_admission" in cleaned:
        cleaned["new_admission"] = to_bool(cleaned["new_admission"])
    return cleaned


@app.route("/student-edit-requests", methods=["POST"])
def create_student_edit_request():
    data = request.json or {}
    student_id = str(data.get("student_id", "")).strip()
    teacher_name = str(data.get("teacher_name", "")).strip()
    session = str(data.get("session", "")).strip()
    changes = filter_edit_changes(data.get("changes") or {})

    if not student_id:
        return jsonify({"success": False, "message": "student_id is required"}), 400
    if not changes:
        return jsonify({"success": False, "message": "No valid changes provided"}), 400

    try:
        student = students_col.find_one({"_id": ObjectId(student_id)})
        if not student:
            return jsonify({"success": False, "message": "Student not found"}), 404
    except Exception:
        return jsonify({"success": False, "message": "Invalid student_id"}), 400

    original = {k: student.get(k, "") for k in changes.keys()}
    snapshot = {
        "student_name": student.get("student_name", ""),
        "class_name": student.get("class_name", student.get("class", "")),
        "rollno": student.get("rollno", ""),
        "admission_no": student.get("admission_no", ""),
        "photo_url": student.get("photo_url", "")
    }
    req_doc = {
        "student_id": student_id,
        "teacher_name": teacher_name,
        "session": session,
        "changes": changes,
        "original": original,
        "student_snapshot": snapshot,
        "status": "pending",
        "created_at": datetime.utcnow(),
        "updated_at": datetime.utcnow()
    }
    ins = student_edit_requests_col.insert_one(req_doc)
    return jsonify({"success": True, "request_id": str(ins.inserted_id)})


@app.route("/student-edit-requests", methods=["GET"])
def list_student_edit_requests():
    status = str(request.args.get("status", "")).strip().lower()
    session = str(request.args.get("session", "")).strip()
    q = {}
    if status:
        q["status"] = status
    if session:
        q["session"] = session
    rows = list(student_edit_requests_col.find(q).sort("created_at", -1))
    for r in rows:
        r["_id"] = str(r["_id"])
    return jsonify({"success": True, "requests": rows})


@app.route("/student-edit-requests/<req_id>", methods=["GET"])
def get_student_edit_request(req_id):
    try:
        doc = student_edit_requests_col.find_one({"_id": ObjectId(req_id)})
    except Exception:
        doc = None
    if not doc:
        return jsonify({"success": False, "message": "Request not found"}), 404
    doc["_id"] = str(doc["_id"])
    return jsonify({"success": True, "request": doc})


@app.route("/student-edit-requests/<req_id>/approve", methods=["POST"])
def approve_student_edit_request(req_id):
    try:
        doc = student_edit_requests_col.find_one({"_id": ObjectId(req_id)})
    except Exception:
        doc = None
    if not doc:
        return jsonify({"success": False, "message": "Request not found"}), 404
    if doc.get("status") != "pending":
        return jsonify({"success": False, "message": "Request already processed"}), 400

    changes = filter_edit_changes(doc.get("changes") or {})
    if not changes:
        return jsonify({"success": False, "message": "No valid changes to apply"}), 400

    # Handle photo upload from data URL if provided
    if PHOTO_DATA_FIELD in changes and changes.get(PHOTO_DATA_FIELD):
        try:
            res = cloudinary.uploader.upload(changes[PHOTO_DATA_FIELD], folder="school_students")
            changes["photo_url"] = res.get("secure_url", "")
        except Exception as e:
            return jsonify({"success": False, "message": f"Photo upload failed: {e}"}), 400
        changes.pop(PHOTO_DATA_FIELD, None)

    students_col.update_one({"_id": ObjectId(doc["student_id"])}, {"$set": changes})
    student_edit_requests_col.update_one(
        {"_id": ObjectId(req_id)},
        {"$set": {"status": "approved", "reviewed_at": datetime.utcnow(), "updated_at": datetime.utcnow()}}
    )
    return jsonify({"success": True})


@app.route("/student-edit-requests/<req_id>/reject", methods=["POST"])
def reject_student_edit_request(req_id):
    try:
        doc = student_edit_requests_col.find_one({"_id": ObjectId(req_id)})
    except Exception:
        doc = None
    if not doc:
        return jsonify({"success": False, "message": "Request not found"}), 404
    if doc.get("status") != "pending":
        return jsonify({"success": False, "message": "Request already processed"}), 400

    student_edit_requests_col.update_one(
        {"_id": ObjectId(req_id)},
        {"$set": {"status": "rejected", "reviewed_at": datetime.utcnow(), "updated_at": datetime.utcnow()}}
    )
    return jsonify({"success": True})

@app.route("/students/<id>", methods=["GET"])
def get_student(id):
    try:
        student = students_col.find_one({"_id": ObjectId(id)})
        if not student:
            return jsonify({"error": "Student not found"}), 404

        ensure_student_regno(student)
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

# ================= ADD TEACHER (FORM + IMAGE) =================
@app.route("/teachers", methods=["POST"])
def add_teacher():
    form = request.form
    photo = request.files.get("photo")
    teacher_code = normalize_teacher_code(form.get("teacher_code", ""))
    if not teacher_code or not teacher_code.isdigit() or len(teacher_code) != 4:
        return jsonify({"error": "teacher_code must be exactly 4 digits"}), 400

    photo_url = ""
    if photo:
        res = cloudinary.uploader.upload(photo, folder="school_teachers")
        photo_url = res.get("secure_url", "")
    else:
        # allow direct URL copy when no file is uploaded
        photo_url = form.get("photo_url", "").strip()

    teacher = {
        "teacher_code": teacher_code,
        "employee_id": form.get("employee_id", "").strip(),
        "teacher_name": form.get("teacher_name", "").strip(),
        "father_name": form.get("father_name", "").strip(),
        "mother_name": form.get("mother_name", "").strip(),
        "gender": form.get("gender", "").strip(),
        "dob": form.get("dob", "").strip(),
        "joining_date": form.get("joining_date", "").strip(),
        "qualification": form.get("qualification", "").strip(),
        "designation": form.get("designation", "").strip(),
        "subject": form.get("subject", "").strip(),
        "mobile": form.get("mobile", "").strip(),
        "email": form.get("email", "").strip(),
        "address": form.get("address", "").strip(),
        "session": form.get("session", "").strip(),
        "photo_url": photo_url
    }

    teachers_col.insert_one(teacher)
    return jsonify({"message": "Teacher added successfully"})


@app.route("/teachers/import_excel_with_images", methods=["POST"])
def import_teachers_excel_with_images():
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
            image_map = build_zip_image_map_with_normalizer(extract_dir, normalize_teacher_code)

        teachers = []
        for _, row in df.iterrows():
            teacher_code = normalize_teacher_code(row.get("teacher_code", ""))
            photo_url = ""

            img_path = image_map.get(teacher_code)
            if img_path and os.path.exists(img_path):
                try:
                    res = cloudinary.uploader.upload(img_path, folder="school_teachers")
                    photo_url = res.get("secure_url", "")
                    if photo_url:
                        matched_photos += 1
                except Exception as e:
                    print(f"Teacher photo upload error for teacher_code={teacher_code}:", e)

            teachers.append({
                "teacher_code": teacher_code,
                "employee_id": str(row.get("employee_id", "")).strip(),
                "teacher_name": str(row.get("teacher_name", "")).strip(),
                "father_name": str(row.get("father_name", "")).strip(),
                "mother_name": str(row.get("mother_name", "")).strip(),
                "gender": str(row.get("gender", "")).strip(),
                "dob": str(row.get("dob", "")).strip(),
                "joining_date": str(row.get("joining_date", "")).strip(),
                "qualification": str(row.get("qualification", "")).strip(),
                "designation": str(row.get("designation", "")).strip(),
                "subject": str(row.get("subject", "")).strip(),
                "mobile": normalize_admission_no(row.get("mobile", "")),
                "email": str(row.get("email", "")).strip(),
                "address": str(row.get("address", "")).strip(),
                "session": str(row.get("session", "")).strip(),
                "photo_url": photo_url
            })

        if teachers:
            teachers_col.insert_many(teachers)

        return jsonify({
            "message": f"Imported {len(teachers)} teachers successfully",
            "teachers_imported": len(teachers),
            "photos_matched": matched_photos,
            "photos_missing": max(0, len(teachers) - matched_photos)
        })
    finally:
        shutil.rmtree(extract_dir, ignore_errors=True)


@app.route("/teachers/download_format", methods=["GET"])
def download_teacher_format():
    wb = Workbook()
    ws = wb.active
    ws.title = "Teacher Import Format"

    headers = [
        "teacher_code", "employee_id", "teacher_name", "father_name", "mother_name",
        "gender", "dob", "joining_date", "qualification", "designation", "subject",
        "mobile", "email", "address", "session"
    ]
    ws.append(headers)

    dv_gender = DataValidation(type="list", formula1='"Male,Female,Other"')
    ws.add_data_validation(dv_gender)
    dv_gender.add("F2:F1000")

    file_path = "teacher_import_format.xlsx"
    wb.save(file_path)
    return send_file(file_path, as_attachment=True)


# ================= GET ALL TEACHERS =================
@app.route("/teachers", methods=["GET"])
def get_teachers():
    session = str(request.args.get("session", "")).strip()
    designation = str(request.args.get("designation", "")).strip()
    subject = str(request.args.get("subject", "")).strip()

    q = {}
    if designation:
        q["designation"] = designation
    if subject:
        q["subject"] = subject

    teachers = []
    if session:
        q_session = dict(q)
        variants = session_variants(session)
        q_session["session"] = {"$in": variants} if variants else session
        teachers = list(teachers_col.find(q_session))
        if len(teachers) <= 1:
            teachers = list(teachers_col.find(q))
    else:
        teachers = list(teachers_col.find(q))

    for t in teachers:
        t["_id"] = str(t["_id"])
    return jsonify(teachers)


# ================= GET SINGLE TEACHER =================
@app.route("/teachers/<id>", methods=["GET"])
def get_teacher(id):
    try:
        teacher = teachers_col.find_one({"_id": ObjectId(id)})
        if not teacher:
            return jsonify({"error": "Teacher not found"}), 404
        teacher["_id"] = str(teacher["_id"])
        return jsonify(teacher)
    except Exception:
        return jsonify({"error": "Invalid ID"}), 400


# ================= UPDATE TEACHER =================
@app.route("/teachers/<id>", methods=["PUT"])
def update_teacher(id):
    try:
        update_data = {}

        if request.content_type and "multipart/form-data" in request.content_type:
            form = request.form
            photo = request.files.get("photo")

            fields = [
                "teacher_code", "employee_id", "teacher_name", "father_name", "mother_name",
                "gender", "dob", "joining_date", "qualification", "designation",
                "subject", "mobile", "email", "address", "session", "photo_url"
            ]
            for f in fields:
                if f in form:
                    update_data[f] = form.get(f, "").strip()
            if "teacher_code" in update_data:
                update_data["teacher_code"] = normalize_teacher_code(update_data["teacher_code"])
                if update_data["teacher_code"] and (not update_data["teacher_code"].isdigit() or len(update_data["teacher_code"]) != 4):
                    return jsonify({"success": False, "error": "teacher_code must be exactly 4 digits"}), 400

            if photo:
                res = cloudinary.uploader.upload(photo, folder="school_teachers")
                update_data["photo_url"] = res.get("secure_url", "")
        else:
            update_data = request.json or {}
            if "teacher_code" in update_data:
                update_data["teacher_code"] = normalize_teacher_code(update_data.get("teacher_code", ""))
                if update_data["teacher_code"] and (not update_data["teacher_code"].isdigit() or len(update_data["teacher_code"]) != 4):
                    return jsonify({"success": False, "error": "teacher_code must be exactly 4 digits"}), 400

        teachers_col.update_one({"_id": ObjectId(id)}, {"$set": update_data})
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400


# ================= DELETE TEACHER =================
@app.route("/teachers/<id>", methods=["DELETE"])
def delete_teacher(id):
    teachers_col.delete_one({"_id": ObjectId(id)})
    return jsonify({"message": "Teacher deleted"})


# ================= DELETE ALL TEACHERS =================
@app.route("/teachers/delete_all", methods=["DELETE"])
def delete_all_teachers():
    teachers_col.delete_many({})
    return jsonify({"message": "All teachers deleted"})

# ================= DOWNLOAD EXCEL FORMAT =================
@app.route("/download_format", methods=["GET"])
def download_format():
    wb = Workbook()
    ws = wb.active
    ws.title = "Student Import Format"

    headers = [
        "photo_id", "admission_no", "rollno", "panno", "student_name",
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

        ensure_student_regno(student)
        student["_id"] = str(student["_id"])

        return jsonify({
            "success": True,
            "student": {
                "id": student["_id"],
                "regno": student.get("regno", ""),
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


@app.route("/teacher-attendance/qr/current", methods=["GET"])
def teacher_attendance_qr_current():
    now = now_ist()
    slot = qr_slot_for(now)
    expires_in = QR_ROTATE_SECONDS - (int(now.timestamp()) % QR_ROTATE_SECONDS)
    return jsonify({
        "success": True,
        "slot": slot,
        "qr_code": build_qr_code(slot),
        "expires_in": expires_in,
        "generated_at": to_iso(now)
    })


@app.route("/teacher-attendance/mark", methods=["POST"])
def teacher_attendance_mark():
    data = request.json or {}
    teacher_id_raw = str(data.get("teacher_id", "")).strip()
    teacher_id = normalize_teacher_code(teacher_id_raw)
    teacher_name = str(data.get("teacher_name", "")).strip()
    qr_code = str(data.get("qr_code", "")).strip()
    session = str(data.get("session", "")).strip()
    device = str(data.get("device", "")).strip()

    lat = data.get("lat", None)
    lon = data.get("lon", None)

    if not teacher_id_raw:
        return jsonify({"success": False, "message": "teacher_id is required"}), 400
    if not qr_code:
        return jsonify({"success": False, "message": "qr_code is required"}), 400
    if not verify_qr_code(qr_code):
        return jsonify({"success": False, "message": "Invalid or expired QR code"}), 400

    teacher = None
    q_or = []
    if teacher_id:
        q_or.append({"teacher_code": teacher_id})
    if teacher_id_raw:
        q_or.append({"teacher_code": teacher_id_raw})
        q_or.append({"employee_id": teacher_id_raw})
    if q_or:
        teacher = teachers_col.find_one({"$or": q_or})
    if not teacher and ObjectId.is_valid(teacher_id_raw):
        teacher = teachers_col.find_one({"_id": ObjectId(teacher_id_raw)})
    if not teacher and teacher_name:
        teacher = teachers_col.find_one({"teacher_name": {"$regex": f"^{teacher_name}$", "$options": "i"}})
    if not teacher:
        return jsonify({"success": False, "message": "Teacher not found"}), 404

    resolved_teacher_code = normalize_teacher_code(teacher.get("teacher_code", "")) or teacher_id or teacher_id_raw
    if not teacher_name:
        teacher_name = str(teacher.get("teacher_name", "")).strip()
    teacher_employee_id = str(teacher.get("employee_id", "")).strip()
    teacher_mobile = str(teacher.get("mobile", "") or teacher.get("phone", "")).strip()
    teacher_mongo_id = str(teacher.get("_id", "")).strip()
    aliases = []
    for a in [resolved_teacher_code, teacher_id_raw, teacher_employee_id, teacher_mongo_id]:
        a = str(a or "").strip()
        if a and a not in aliases:
            aliases.append(a)

    if lat is not None and lon is not None:
        try:
            lat = float(lat)
            lon = float(lon)
        except Exception:
            return jsonify({"success": False, "message": "Invalid location coordinates"}), 400
    else:
        lat = None
        lon = None

    site_distance = None
    if ATTENDANCE_SITE_LAT and ATTENDANCE_SITE_LON and lat is not None and lon is not None:
        try:
            site_lat = float(ATTENDANCE_SITE_LAT)
            site_lon = float(ATTENDANCE_SITE_LON)
            site_distance = distance_meters(site_lat, site_lon, lat, lon)
        except Exception:
            site_distance = None

    if ATTENDANCE_REQUIRE_GPS:
        if lat is None or lon is None:
            return jsonify({"success": False, "message": "Location is required"}), 400
        if site_distance is not None and site_distance > ATTENDANCE_SITE_RADIUS_M:
            return jsonify({
                "success": False,
                "message": f"Outside allowed attendance area ({int(site_distance)}m)"
            }), 400

    now = now_ist()
    today = now.date().isoformat()
    rec = teacher_attendance_col.find_one({"teacher_id": teacher_id, "date": today})

    late_h, late_m = parse_hhmm(ATTENDANCE_LATE_AFTER, 9, 15)
    late_cutoff = now.replace(hour=late_h, minute=late_m, second=0, microsecond=0)

    event = "IN"
    if rec and rec.get("in_time"):
        event = "OUT"
    if rec and rec.get("in_time") and rec.get("out_time"):
        return jsonify({
            "success": False,
            "message": "Attendance already marked for both IN and OUT today",
            "code": "ALREADY_MARKED"
        }), 409
    if event == "OUT":
        in_time = rec.get("in_time") if rec else None
        if isinstance(in_time, datetime):
            elapsed_minutes = max(0, int((now - in_time).total_seconds() // 60))
            if elapsed_minutes < ATTENDANCE_MIN_OUT_MINUTES:
                wait_minutes = ATTENDANCE_MIN_OUT_MINUTES - elapsed_minutes
                return jsonify({
                    "success": False,
                    "message": f"OUT attendance allowed after {ATTENDANCE_MIN_OUT_MINUTES} minutes from IN. Please wait {wait_minutes} minute(s).",
                    "code": "OUT_TOO_EARLY",
                    "in_time": to_iso(in_time),
                    "wait_minutes": wait_minutes
                }), 409

    payload = {
        "teacher_id": resolved_teacher_code,
        "teacher_name": teacher_name,
        "teacher_employee_id": teacher_employee_id,
        "teacher_mongo_id": teacher_mongo_id,
        "teacher_aliases": aliases,
        "session": session,
        "date": today,
        "updated_at": now,
    }
    if device:
        payload["device"] = device
    if lat is not None and lon is not None:
        payload["last_lat"] = lat
        payload["last_lon"] = lon
    if site_distance is not None:
        payload["site_distance_m"] = round(site_distance, 2)

    if event == "IN":
        payload["in_time"] = now
        payload["late_minutes"] = max(0, int((now - late_cutoff).total_seconds() // 60))
    else:
        payload["out_time"] = now
        in_time = rec.get("in_time")
        if isinstance(in_time, datetime):
            minutes = max(0, int((now - in_time).total_seconds() // 60))
            payload["working_minutes"] = minutes

    teacher_attendance_col.update_one(
        {"teacher_id": resolved_teacher_code, "date": today},
        {"$set": payload, "$setOnInsert": {"created_at": now}},
        upsert=True
    )

    saved = teacher_attendance_col.find_one({"teacher_id": resolved_teacher_code, "date": today}) or {}
    result = {
        "teacher_id": resolved_teacher_code,
        "teacher_name": saved.get("teacher_name", teacher_name),
        "date": today,
        "event": event,
        "in_time": to_iso(saved["in_time"]) if isinstance(saved.get("in_time"), datetime) else "",
        "out_time": to_iso(saved["out_time"]) if isinstance(saved.get("out_time"), datetime) else "",
        "late_minutes": int(saved.get("late_minutes", 0) or 0),
        "working_minutes": int(saved.get("working_minutes", 0) or 0),
        "site_distance_m": saved.get("site_distance_m"),
    }

    sms_text = (
        f"Dear {result['teacher_name']}, your attendance ({result['event']}) "
        f"has been marked on {result['date']}. "
        f"IN: {result['in_time'] or '-'} OUT: {result['out_time'] or '-'}."
    )
    sms_result = send_textbee_sms(teacher_mobile, sms_text)

    return jsonify({"success": True, "attendance": result, "sms": sms_result})


@app.route("/teacher-attendance/history", methods=["GET"])
def teacher_attendance_history():
    teacher_id_raw = str(request.args.get("teacher_id", "")).strip()
    teacher_id = normalize_teacher_code(teacher_id_raw)
    teacher_name = str(request.args.get("teacher_name", "")).strip()
    limit = request.args.get("limit", "20")
    try:
        limit_n = max(1, min(100, int(limit)))
    except Exception:
        limit_n = 20

    if not teacher_id_raw and not teacher_name:
        return jsonify({"success": False, "message": "teacher_id is required"}), 400

    q_or = []
    for v in [teacher_id_raw, teacher_id]:
        v = str(v or "").strip()
        if not v:
            continue
        q_or.append({"teacher_id": v})
        q_or.append({"teacher_aliases": v})
        q_or.append({"teacher_employee_id": v})
        q_or.append({"teacher_mongo_id": v})
    if teacher_name:
        q_or.append({"teacher_name": {"$regex": f"^{teacher_name}$", "$options": "i"}})

    query = {"$or": q_or} if q_or else {}
    rows = list(
        teacher_attendance_col.find(query)
        .sort("date", -1)
        .limit(limit_n)
    )
    out = []
    for r in rows:
        out.append({
            "id": str(r.get("_id", "")),
            "teacher_id": r.get("teacher_id", ""),
            "teacher_name": r.get("teacher_name", ""),
            "date": r.get("date", ""),
            "session": r.get("session", ""),
            "in_time": to_iso(r["in_time"]) if isinstance(r.get("in_time"), datetime) else "",
            "out_time": to_iso(r["out_time"]) if isinstance(r.get("out_time"), datetime) else "",
            "late_minutes": int(r.get("late_minutes", 0) or 0),
            "working_minutes": int(r.get("working_minutes", 0) or 0),
            "site_distance_m": r.get("site_distance_m"),
            "device": r.get("device", ""),
        })
    return jsonify({"success": True, "history": out})


@app.route("/teacher-attendance/admin/list", methods=["GET"])
def teacher_attendance_admin_list():
    date_from = str(request.args.get("date_from", "")).strip()
    date_to = str(request.args.get("date_to", "")).strip()
    teacher_id_raw = str(request.args.get("teacher_id", "")).strip()
    teacher_id = normalize_teacher_code(teacher_id_raw)
    teacher_name = str(request.args.get("teacher_name", "")).strip()
    limit = request.args.get("limit", "500")
    try:
        limit_n = max(1, min(5000, int(limit)))
    except Exception:
        limit_n = 500

    query = {}
    if date_from or date_to:
        query["date"] = {}
        if date_from:
            query["date"]["$gte"] = date_from
        if date_to:
            query["date"]["$lte"] = date_to
        if not query["date"]:
            query.pop("date", None)

    q_or = []
    for v in [teacher_id_raw, teacher_id]:
        v = str(v or "").strip()
        if not v:
            continue
        q_or.append({"teacher_id": v})
        q_or.append({"teacher_aliases": v})
        q_or.append({"teacher_employee_id": v})
        q_or.append({"teacher_mongo_id": v})
    if teacher_name:
        q_or.append({"teacher_name": {"$regex": teacher_name, "$options": "i"}})
    if q_or:
        query["$or"] = q_or

    rows = list(
        teacher_attendance_col.find(query)
        .sort([("date", -1), ("updated_at", -1)])
        .limit(limit_n)
    )
    out = []
    for r in rows:
        out.append({
            "id": str(r.get("_id", "")),
            "teacher_id": r.get("teacher_id", ""),
            "teacher_name": r.get("teacher_name", ""),
            "date": r.get("date", ""),
            "session": r.get("session", ""),
            "in_time": to_iso(r["in_time"]) if isinstance(r.get("in_time"), datetime) else "",
            "out_time": to_iso(r["out_time"]) if isinstance(r.get("out_time"), datetime) else "",
            "late_minutes": int(r.get("late_minutes", 0) or 0),
            "working_minutes": int(r.get("working_minutes", 0) or 0),
            "site_distance_m": r.get("site_distance_m"),
            "device": r.get("device", ""),
        })
    return jsonify({"success": True, "rows": out, "count": len(out)})


@app.route("/teacher-attendance/admin/edit", methods=["POST"])
def teacher_attendance_admin_edit():
    data = request.json or {}
    teacher_id_raw = str(data.get("teacher_id", "")).strip()
    teacher_id = normalize_teacher_code(teacher_id_raw)
    date = str(data.get("date", "")).strip()
    in_time = str(data.get("in_time", "")).strip()
    out_time = str(data.get("out_time", "")).strip()
    if not teacher_id_raw or not date:
        return jsonify({"success": False, "message": "teacher_id and date are required"}), 400

    q_or = []
    for v in [teacher_id_raw, teacher_id]:
        v = str(v or "").strip()
        if not v:
            continue
        q_or.append({"teacher_id": v})
        q_or.append({"teacher_aliases": v})
        q_or.append({"teacher_employee_id": v})
        q_or.append({"teacher_mongo_id": v})
    query = {"date": date, "$or": q_or}

    rec = teacher_attendance_col.find_one(query)
    if not rec:
        return jsonify({"success": False, "message": "Attendance record not found"}), 404

    tz = IST
    updates = {"updated_at": now_ist()}
    in_dt = None
    out_dt = None
    if in_time:
        try:
            in_dt = datetime.strptime(f"{date} {in_time}", "%Y-%m-%d %H:%M").replace(tzinfo=tz)
            updates["in_time"] = in_dt
        except Exception:
            return jsonify({"success": False, "message": "Invalid in_time format. Use HH:MM"}), 400
    else:
        updates["in_time"] = None

    if out_time:
        try:
            out_dt = datetime.strptime(f"{date} {out_time}", "%Y-%m-%d %H:%M").replace(tzinfo=tz)
            updates["out_time"] = out_dt
        except Exception:
            return jsonify({"success": False, "message": "Invalid out_time format. Use HH:MM"}), 400
    else:
        updates["out_time"] = None

    if in_dt and out_dt and out_dt < in_dt:
        return jsonify({"success": False, "message": "OUT time cannot be earlier than IN time"}), 400

    late_h, late_m = parse_hhmm(ATTENDANCE_LATE_AFTER, 9, 15)
    if in_dt:
        late_cutoff = in_dt.replace(hour=late_h, minute=late_m, second=0, microsecond=0)
        updates["late_minutes"] = max(0, int((in_dt - late_cutoff).total_seconds() // 60))
    else:
        updates["late_minutes"] = 0

    if in_dt and out_dt:
        updates["working_minutes"] = max(0, int((out_dt - in_dt).total_seconds() // 60))
    else:
        updates["working_minutes"] = 0

    teacher_attendance_col.update_one({"_id": rec["_id"]}, {"$set": updates})
    saved = teacher_attendance_col.find_one({"_id": rec["_id"]}) or {}
    return jsonify({
        "success": True,
        "record": {
            "teacher_id": saved.get("teacher_id", ""),
            "teacher_name": saved.get("teacher_name", ""),
            "date": saved.get("date", ""),
            "in_time": to_iso(saved["in_time"]) if isinstance(saved.get("in_time"), datetime) else "",
            "out_time": to_iso(saved["out_time"]) if isinstance(saved.get("out_time"), datetime) else "",
            "late_minutes": int(saved.get("late_minutes", 0) or 0),
            "working_minutes": int(saved.get("working_minutes", 0) or 0),
        }
    })


@app.route("/teacher-attendance/admin/delete", methods=["POST"])
def teacher_attendance_admin_delete():
    data = request.json or {}
    teacher_id_raw = str(data.get("teacher_id", "")).strip()
    teacher_id = normalize_teacher_code(teacher_id_raw)
    date = str(data.get("date", "")).strip()
    if not teacher_id_raw or not date:
        return jsonify({"success": False, "message": "teacher_id and date are required"}), 400

    q_or = []
    for v in [teacher_id_raw, teacher_id]:
        v = str(v or "").strip()
        if not v:
            continue
        q_or.append({"teacher_id": v})
        q_or.append({"teacher_aliases": v})
        q_or.append({"teacher_employee_id": v})
        q_or.append({"teacher_mongo_id": v})

    query = {"date": date, "$or": q_or}
    rec = teacher_attendance_col.find_one(query)
    if not rec:
        return jsonify({"success": False, "message": "Attendance record not found"}), 404

    teacher_attendance_col.delete_one({"_id": rec["_id"]})
    return jsonify({"success": True, "message": "Attendance deleted"})


@app.route("/teacher-attendance/send-sms", methods=["POST"])
def teacher_attendance_send_sms():
    data = request.json or {}
    teacher_name = str(data.get("teacher_name", "")).strip() or "Teacher"
    teacher_mobile = str(data.get("mobile", "")).strip()
    event = str(data.get("event", "")).strip().upper() or "IN"
    date = str(data.get("date", "")).strip() or now_ist().date().isoformat()
    in_time = str(data.get("in_time", "")).strip() or "-"
    out_time = str(data.get("out_time", "")).strip() or "-"

    if not teacher_mobile:
        return jsonify({"success": False, "message": "mobile is required"}), 400

    msg = (
        f"Dear {teacher_name}, your attendance ({event}) has been marked on {date}. "
        f"IN: {in_time} OUT: {out_time}."
    )
    sms_result = send_textbee_sms(teacher_mobile, msg)
    return jsonify({"success": bool(sms_result.get("sent")), "sms": sms_result})

# ================= HOME =================
@app.route("/", methods=["GET"])
def home():
    return "Student Backend Running", 200

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
