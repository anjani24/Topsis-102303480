from flask import Flask, render_template, request, flash, send_file, jsonify
import pandas as pd
import numpy as np
import os
import smtplib
import re
from email.message import EmailMessage

app = Flask(__name__)
app.secret_key = "secretkey"

UPLOAD_FOLDER = "/tmp/uploads"
RESULT_FOLDER = "/tmp/results"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

# -------- TOPSIS FUNCTIONS --------

def normalize_matrix(matrix, weights):
    norm_matrix = matrix / np.sqrt((matrix**2).sum(axis=0))
    weighted_matrix = norm_matrix * weights
    return weighted_matrix

def calculate_ideal_solutions(weighted_matrix, impacts):
    ideal_best = []
    ideal_worst = []
    for i, impact in enumerate(impacts):
        if impact == '+':
            ideal_best.append(weighted_matrix[:, i].max())
            ideal_worst.append(weighted_matrix[:, i].min())
        elif impact == '-':
            ideal_best.append(weighted_matrix[:, i].min())
            ideal_worst.append(weighted_matrix[:, i].max())
    return np.array(ideal_best), np.array(ideal_worst)

def calculate_topsis(matrix, weights, impacts):
    weighted_matrix = normalize_matrix(matrix, weights)
    ideal_best, ideal_worst = calculate_ideal_solutions(weighted_matrix, impacts)
    distances_best = np.sqrt(((weighted_matrix - ideal_best)**2).sum(axis=1))
    distances_worst = np.sqrt(((weighted_matrix - ideal_worst)**2).sum(axis=1))
    scores = distances_worst / (distances_best + distances_worst)
    return scores

# -------- EMAIL FUNCTION --------

def send_email_with_attachment(to_email, file_path):
    sender_email = os.getenv("SENDER_EMAIL")
    sender_password = os.getenv("SENDER_PASSWORD")

    msg = EmailMessage()
    msg["Subject"] = "Your TOPSIS Result File"
    msg["From"] = sender_email
    msg["To"] = to_email
    msg.set_content("Your TOPSIS result file is attached.")

    with open(file_path, "rb") as f:
        file_data = f.read()

    msg.add_attachment(
        file_data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=os.path.basename(file_path),
    )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender_email, sender_password)
        server.send_message(msg)

# -------- EMAIL VALIDATION --------

def is_valid_email(email):
    pattern = r"^[\w\.-]+@[\w\.-]+\.\w+$"
    return re.match(pattern, email)

# -------- ROUTES --------

@app.route("/", methods=["GET", "POST"])
def home():
    return render_template("index.html")

@app.route("/preview", methods=["POST"])
def preview():
    file = request.files["file"]
    ext = file.filename.split(".")[-1].lower()

    path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(path)

    if ext == "csv":
        df = pd.read_csv(path)
    else:
        df = pd.read_excel(path)

    return df.head(5).to_json(orient="records")

@app.route("/run", methods=["POST"])
def run_topsis():
    file = request.files["file"]
    weights = request.form["weights"]
    impacts = request.form["impacts"]
    email = request.form["email"]

    if not is_valid_email(email):
        return jsonify({"error": "Invalid email format!"})

    if "," not in weights or "," not in impacts:
        return jsonify({"error": "Weights and impacts must be comma-separated!"})

    weights_list = list(map(float, weights.split(",")))
    impacts_list = impacts.split(",")

    if len(weights_list) != len(impacts_list):
        return jsonify({"error": "Number of weights must equal number of impacts!"})

    if not all(i in ["+", "-"] for i in impacts_list):
        return jsonify({"error": "Impacts must be + or - only!"})

    ext = file.filename.split(".")[-1].lower()
    path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(path)

    if ext == "csv":
        data = pd.read_csv(path)
    else:
        data = pd.read_excel(path)

    if data.shape[1] < 3:
        return jsonify({"error": "Input file must have at least three columns (ID, Criteria1, Criteria2, ...)."})

    criteria_matrix = data.iloc[:, 1:].values

    if not np.issubdtype(criteria_matrix.dtype, np.number):
        return jsonify({"error": "Criteria columns must contain only numeric values."})

    if len(weights_list) != criteria_matrix.shape[1]:
        return jsonify({"error": "Weights must match number of criteria columns!"})

    scores = calculate_topsis(criteria_matrix, weights_list, impacts_list)
    data["TOPSIS Score"] = scores
    data["Rank"] = data["TOPSIS Score"].rank(ascending=False).astype(int)

    result_file = os.path.join(RESULT_FOLDER, "topsis_result.xlsx")
    data.to_excel(result_file, index=False)

    send_email_with_attachment(email, result_file)

    return jsonify({"success": True, "download": "/download"})

@app.route("/download")
def download():
    return send_file(
        os.path.join(RESULT_FOLDER, "topsis_result.xlsx"),
        as_attachment=True,
        download_name="topsis_result.xlsx",
    )

if __name__ == "__main__":
    app.run(debug=True)