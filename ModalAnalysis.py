import os
import shutil
import hashlib
from datetime import datetime

def check_and_move_file():

    # ==========================================================
    # USER INPUT
    # ==========================================================

    folder = input("Enter main folder path: ")

    pre_path = os.path.join(folder, "pre-process")
    post_path = os.path.join(folder, "post-process")

    # ==========================================================
    # CREATE POST-PROCESS FOLDER IF NOT PRESENT
    # ==========================================================

    os.makedirs(post_path, exist_ok=True)

    try:

        # ==========================================================
        # CHECK PRE-PROCESS FOLDER
        # ==========================================================

        if not os.path.exists(pre_path):

            print(f"\nERROR: pre-process folder not found:\n{pre_path}")
            return

        # ==========================================================
        # FIND .OUT FILES
        # ==========================================================

        files = [
            f for f in os.listdir(pre_path)
            if f.endswith(".out")
        ]

        if not files:

            print("\nNo .out file found in pre-process folder.")
            return

        # ==========================================================
        # SELECT FIRST OUT FILE
        # ==========================================================

        filename = files[0]

        file_path = os.path.join(pre_path, filename)

        # ==========================================================
        # READ ORIGINAL OUT FILE
        # ==========================================================

        with open(file_path, 'r') as file:

            lines = file.readlines()

        # ==========================================================
        # FLEXIBLE HEADER DETECTION
        # ==========================================================

        keywords = ["Subcase", "Mode", "Frequency"]

        start_index = -1

        for i, line in enumerate(lines):

            if all(word in line for word in keywords):

                start_index = i + 1
                break

        if start_index == -1:

            print("\nFile does not contain required modal records.")
            return

        # ==========================================================
        # INITIALIZE VARIABLES
        # ==========================================================

        previous_line = None
        transition_line = None

        prev_mode = None
        prev_freq = None

        trans_mode = None
        trans_freq = None

        # ==========================================================
        # FIND TRANSITION
        # ==========================================================

        for line in lines[start_index:]:

            parts = line.split()

            if len(parts) < 3:
                continue

            try:

                # ==================================================
                # ORIGINAL USER LOGIC (UNCHANGED)
                # ==================================================

                mode = int(parts[1])
                frequency = float(parts[2])

                # ==================================================
                # RIGID BODY MODE
                # ==================================================

                if abs(frequency) < 1e-2:

                    previous_line = line.strip()

                    prev_mode = mode
                    prev_freq = frequency

                    continue

                # ==================================================
                # FIRST STRUCTURAL MODE
                # ==================================================

                else:

                    transition_line = line.strip()

                    trans_mode = mode
                    trans_freq = frequency

                    break

            except Exception:

                continue

        # ==========================================================
        # RESULT FILE GENERATION
        # ==========================================================

        result_filename = filename.replace(".out", "_result.html")

        result_path = os.path.join(post_path, result_filename)

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # ==========================================================
        # HTML + CSS CONTENT
        # ==========================================================

        html_content = f"""

<!DOCTYPE html>

<html>

<head>

<meta charset="UTF-8">

<title>Eka Mobility Modal Analysis Report</title>

<style>

@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700&display=swap');

*{{
    margin:0;
    padding:0;
    box-sizing:border-box;
}}

body{{
    font-family:'Montserrat',sans-serif;
    background:#edf3f8;
    padding:40px;
    color:#1a1a1a;
}}

.container{{
    max-width:1400px;
    margin:auto;
    background:white;
    border-radius:18px;
    overflow:hidden;
    box-shadow:0 10px 30px rgba(0,0,0,0.12);
}}

/* ===================================================== */
/* HEADER */
/* ===================================================== */

.header{{
    background:linear-gradient(135deg,#005fa3,#0074bc);
    color:white;
    padding:30px 40px;

    display:flex;
    justify-content:space-between;
    align-items:center;
    flex-wrap:wrap;
}}

.header-left{{
    display:flex;
    align-items:center;
    gap:25px;
}}

.logo1{{
    width:90px;
    background:white;
    padding:10px;
    border-radius:12px;
}}

.logo2{{
    width:180px;
}}

.title-group h1{{
    font-size:42px;
    font-weight:700;
}}

.title-group h2{{
    font-size:18px;
    font-weight:400;
    margin-top:6px;
}}

/* ===================================================== */
/* BLOCKS */
/* ===================================================== */

.block{{
    padding:35px 40px;
    border-bottom:1px solid #dde7f0;
}}

.section-title{{
    font-size:28px;
    color:#005fa3;
    margin-bottom:30px;
    font-weight:700;
    border-left:6px solid #0074bc;
    padding-left:15px;
}}

/* ===================================================== */
/* METRICS */
/* ===================================================== */

.metrics{{
    display:grid;
    grid-template-columns:repeat(auto-fit,minmax(260px,1fr));
    gap:22px;
}}

.card{{
    background:#f8fbff;
    border:1px solid #dbe8f3;
    border-radius:16px;
    padding:25px;
}}

.card-title{{
    color:#6b7280;
    font-size:14px;
    margin-bottom:12px;
    text-transform:uppercase;
    letter-spacing:1px;
}}

.card-value{{
    color:#0074bc;
    font-size:22px;
    font-weight:700;
    word-break:break-word;
}}

/* ===================================================== */
/* SUMMARY */
/* ===================================================== */

.summary-grid{{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:25px;
}}

.summary-box{{
    background:#f8fbff;
    border-left:8px solid #0074bc;
    border-radius:14px;
    padding:25px;
}}

.summary-box h3{{
    color:#005fa3;
    margin-bottom:18px;
    font-size:22px;
}}

.summary-item{{
    margin-bottom:14px;
    font-size:17px;
    line-height:1.6;
}}

.summary-item span{{
    font-weight:600;
    color:#005fa3;
}}

/* ===================================================== */
/* INTERPRETATION BLOCK */
/* ===================================================== */

.interpretation-block{{
    margin-top:35px;
}}

.interpretation-title{{
    font-size:28px;
    color:#005fa3;
    margin-bottom:25px;
    font-weight:700;
    border-left:6px solid #0074bc;
    padding-left:15px;
}}

.ok{{
    margin-top:25px;
    background:#e8f7ec;
    color:#1f7a38;
    padding:16px;
    border-radius:10px;
    border-left:6px solid #28a745;
    font-weight:600;
}}

.warning{{
    margin-top:25px;
    background:#fff4df;
    color:#b7791f;
    padding:16px;
    border-radius:10px;
    border-left:6px solid #ff9800;
    font-weight:600;
}}

.error{{
    margin-top:25px;
    background:#ffe5e5;
    color:#c62828;
    padding:16px;
    border-radius:10px;
    border-left:6px solid #e53935;
    font-weight:600;
}}

/* ===================================================== */
/* HASH */
/* ===================================================== */

.hash-box{{
    margin-top:25px;
    padding:20px;
    background:#0b456b;
    border-radius:12px;
    color:white;
    font-family:monospace;
    word-break:break-all;
    line-height:1.8;
}}

.footer{{
    margin-top:30px;
    text-align:center;
    color:#d0d0d0;
    font-size:13px;
}}

@media(max-width:900px){{

    .summary-grid{{
        grid-template-columns:1fr;
    }}

    .header{{
        flex-direction:column;
        align-items:flex-start;
        gap:20px;
    }}

    .title-group h1{{
        font-size:32px;
    }}
}}

</style>

</head>

<body>

<div class="container">

    <!-- HEADER -->

    <div class="header">

        <div class="header-left">

            <img src="ekalogo1.jpeg" class="logo1">

            <div class="title-group">

                <h1>Modal Analysis Report</h1>

                <h2>R&D - CAE Department </h2>

            </div>

        </div>

        <img src="eka_logo_final.jpg" class="logo2">

    </div>

    <!-- KEY METRICS -->

    <div class="block">

        <div class="section-title">
            Key Metrics
        </div>

        <div class="metrics">

            <div class="card">
                <div class="card-title">File Name</div>
                <div class="card-value">{filename}</div>
            </div>

            <div class="card">
                <div class="card-title">Processed On</div>
                <div class="card-value">{now}</div>
            </div>

            <div class="card">
                <div class="card-title">Last Rigid Mode</div>
                <div class="card-value">{prev_mode}</div>
            </div>

            <div class="card">
                <div class="card-title">First Structural Mode</div>
                <div class="card-value">{trans_mode}</div>
            </div>

        </div>

    </div>

    <!-- TRANSITION SUMMARY -->

    <div class="block">

        <div class="section-title">
            Transition Summary
        </div>

        <div class="summary-grid">

            <div class="summary-box">

                <h3>Rigid Body Information</h3>

                <div class="summary-item">
                    <span>Last Rigid Mode :</span> {prev_mode}
                </div>

                <div class="summary-item">
                    <span>Rigid Frequency :</span> {prev_freq:.6f} Hz
                </div>

            </div>

            <div class="summary-box">

                <h3>Structural Information</h3>

                <div class="summary-item">
                    <span>First Structural Mode :</span> {trans_mode}
                </div>

                <div class="summary-item">
                    <span>Structural Frequency :</span> {trans_freq:.6f} Hz
                </div>

            </div>

        </div>

        <!-- INTERPRETATION BLOCK -->

        <div class="interpretation-block">

            <div class="interpretation-title">
                Interpretation
            </div>
"""

        # ==========================================================
        # INTERPRETATION
        # ==========================================================

        if prev_mode == 6:

            html_content += """

            <div class="ok">

                OK : System is properly constrained
                with 6 rigid body modes.

            </div>
"""

        elif prev_mode and prev_mode < 6:

            html_content += """

            <div class="warning">

                WARNING : Under-constrained system detected.

            </div>
"""

        elif prev_mode and prev_mode > 6:

            html_content += """

            <div class="warning">

                WARNING : Possible over-constraint or
                modeling issue detected.

            </div>
"""

        else:

            html_content += """

            <div class="error">

                ERROR : Unable to determine system interpretation.

            </div>
"""

        # ==========================================================
        # PRINT TRANSITION MODE
        # ==========================================================

        if trans_mode is not None:

            html_content += f"""

            <div class="summary-item"
                 style="
                    margin-top:22px;
                    font-size:18px;
                    font-weight:600;
                    color:#005fa3;
                 ">

                Transition occurs at Mode : {trans_mode}

            </div>
"""
            
        # ==========================================================
        # SHA256 HASH
        # ==========================================================

        report_hash = hashlib.sha256(
            html_content.encode()
        ).hexdigest()

        html_content += f"""

        </div>

    </div>

    <!-- HASH SECTION -->

    <div class="block">

        <div class="section-title">
            Integrity Verification
        </div>

        <div class="hash-box">

            <b>SHA256 Hash:</b><br><br>

            {report_hash}

        </div>

        <div class="footer">

            © Copyright Confidential eka Mobility India Pvt.Ltd

        </div>

    </div>

</div>

</body>

</html>
"""

        # ==========================================================
        # WRITE HTML FILE
        # ==========================================================

        with open(result_path, "w", encoding="utf-8") as result_file:

            result_file.write(html_content)

        # ==========================================================
        # MAKE HTML FILE READ ONLY
        # ==========================================================

        os.chmod(result_path, 0o444)

        # ==========================================================
        # COPY LOGOS
        # ==========================================================

        current_dir = os.path.dirname(os.path.abspath(__file__))

        shutil.copy(
            os.path.join(current_dir, "ekalogo1.jpeg"),
            post_path
        )

        shutil.copy(
            os.path.join(current_dir, "eka_logo_final.jpg"),
            post_path
        )

        # ==========================================================
        # SAVE HASH FILE
        # ==========================================================

        hash_file_path = result_path + ".sha256"

        with open(hash_file_path, "w") as hash_file:

            hash_file.write(report_hash)

        # ==========================================================
        # MAKE HASH FILE READ ONLY
        # ==========================================================

        os.chmod(hash_file_path, 0o444)

        # ==========================================================
        # FILE MOVE SECTION (ORIGINAL LOGIC)
        # ==========================================================

        destination = os.path.join(post_path, filename)

        shutil.move(file_path, destination)

        print(f"File moved to post-process folder:\n{destination}")

        # ==========================================================
        # PRINT STATUS
        # ==========================================================

        print(f"\nResult saved at:\n{result_path}")

        print(f"\nSHA256 file saved at:\n{hash_file_path}")

    except Exception as e:

        print("\n===================================")
        print("ERROR OCCURRED")
        print("===================================")
        print(e)
        print("===================================\n")

# ==========================================================
# RUN
# ==========================================================

check_and_move_file()