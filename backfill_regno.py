import ast
import random

from pymongo import MongoClient


def load_mongo_uri():
    with open("app.py", "r", encoding="utf-8") as f:
        tree = ast.parse(f.read(), filename="app.py")
    for node in tree.body:
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if getattr(target, "id", "") == "MONGO_URI":
                    return ast.literal_eval(node.value)
    raise RuntimeError("MONGO_URI not found in app.py")


client = MongoClient(load_mongo_uri(), serverSelectionTimeoutMS=20000, connectTimeoutMS=20000)
db = client["school_erp"]
students_col = db["students"]


def generate_student_regno(reserved):
    for _ in range(1000):
        regno = str(random.randint(100000, 999999))
        if regno not in reserved and not students_col.find_one({"regno": regno}, {"_id": 1}):
            reserved.add(regno)
            return regno
    raise RuntimeError("Unable to generate unique registration number")


def needs_regno(student):
    regno = str(student.get("regno", "")).strip()
    return not (len(regno) == 6 and regno.isdigit())


def main():
    students = list(students_col.find({}, {"regno": 1}))
    existing = {
        str(s.get("regno", "")).strip()
        for s in students
        if len(str(s.get("regno", "")).strip()) == 6 and str(s.get("regno", "")).strip().isdigit()
    }
    updated = 0

    for student in students:
        if not needs_regno(student):
            continue
        regno = generate_student_regno(existing)
        result = students_col.update_one(
            {"_id": student["_id"]},
            {"$set": {"regno": regno}},
        )
        updated += result.modified_count

    total = students_col.count_documents({})
    missing = students_col.count_documents({
        "$or": [
            {"regno": {"$exists": False}},
            {"regno": ""},
            {"regno": None},
        ]
    })
    all_regnos = [str(s.get("regno", "")).strip() for s in students_col.find({}, {"regno": 1})]
    valid_regnos = [r for r in all_regnos if len(r) == 6 and r.isdigit()]
    duplicate_count = len(valid_regnos) - len(set(valid_regnos))
    invalid_count = total - len(valid_regnos)
    print(f"students_total={total}")
    print(f"regno_backfilled={updated}")
    print(f"regno_missing_blank={missing}")
    print(f"regno_invalid={invalid_count}")
    print(f"regno_duplicates={duplicate_count}")


if __name__ == "__main__":
    main()
