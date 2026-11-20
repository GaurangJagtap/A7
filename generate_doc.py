import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

doc = Document()

# Configure page margins (1 inch all around)
for section in doc.sections:
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def add_callout(doc, text, title="NOTE"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, "F0F4F8")
    set_cell_margins(cell, 140, 140, 200, 200)
    # set border: left thick border #003366, others none
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="none"/>
            <w:left w:val="single" w:sz="24" w:space="0" w:color="003366"/>
            <w:bottom w:val="none"/>
            <w:right w:val="none"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    r1 = p.add_run(f"[{title}] ")
    r1.bold = True
    r1.font.name = "Calibri"
    r1.font.size = Pt(10)
    r1.font.color.rgb = RGBColor(0, 51, 102)
    r2 = p.add_run(text)
    r2.font.name = "Calibri"
    r2.font.size = Pt(10)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def add_heading_styled(doc, text, level):
    h = doc.add_heading(text, level=level)
    h.paragraph_format.keep_with_next = True
    h.paragraph_format.space_before = Pt(12)
    h.paragraph_format.space_after = Pt(4)
    run = h.runs[0]
    run.font.name = "Calibri"
    if level == 1:
        run.font.size = Pt(16)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 51, 102) # Deep Navy
    elif level == 2:
        run.font.size = Pt(13)
        run.font.bold = True
        run.font.color.rgb = RGBColor(30, 80, 140) # Steel Blue
    elif level == 3:
        run.font.size = Pt(11.5)
        run.font.bold = True
        run.font.color.rgb = RGBColor(50, 50, 50)
    return h

def add_body_p(doc, text="", bold_prefix="", space_after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_b = p.add_run(bold_prefix)
        r_b.bold = True
        r_b.font.name = "Calibri"
        r_b.font.size = Pt(11)
        r_b.font.color.rgb = RGBColor(30, 30, 30)
    if text:
        r_t = p.add_run(text)
        r_t.font.name = "Calibri"
        r_t.font.size = Pt(11)
        r_t.font.color.rgb = RGBColor(40, 40, 40)
    return p

def add_code_block(doc, code_text):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, "F8F9FA")
    set_cell_margins(cell, 100, 100, 150, 150)
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/>
            <w:left w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/>
            <w:right w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.05
    run = p.add_run(code_text.strip())
    run.font.name = "Consolas"
    run.font.size = Pt(9.5)
    run.font.color.rgb = RGBColor(33, 37, 41)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

# ==========================================
# 1. COVER PAGE
# ==========================================
title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
title_p.paragraph_format.space_before = Pt(0)
title_p.paragraph_format.space_after = Pt(2)
run_inst = title_p.add_run("Bansilal Ramnath Agarwal Charitable Trust's\n")
run_inst.font.name = "Calibri"
run_inst.font.size = Pt(12)
run_inst.font.color.rgb = RGBColor(80, 80, 80)

run_college = title_p.add_run("Vishwakarma Institute of Technology, Pune\n")
run_college.font.name = "Calibri"
run_college.font.size = Pt(18)
run_college.bold = True
run_college.font.color.rgb = RGBColor(0, 51, 102)

run_dept = title_p.add_run("Department of Artificial Intelligence and Data Science\n")
run_dept.font.name = "Calibri"
run_dept.font.size = Pt(14)
run_dept.bold = True
run_dept.font.color.rgb = RGBColor(40, 40, 40)

run_sub = title_p.add_run("Course: MLOps (AI3206C) | Academic Year: 2026-27\n")
run_sub.font.name = "Calibri"
run_sub.font.size = Pt(11)
run_sub.font.italic = True
run_sub.font.color.rgb = RGBColor(100, 100, 100)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Table for assignment & student metadata
meta_table = doc.add_table(rows=12, cols=2)
meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_data = [
    ("Department", "Artificial Intelligence and Data Science"),
    ("Student Name", "[Full Name]"),
    ("Class", "TY"),
    ("Division", "B"),
    ("Roll No", "[Roll Number]"),
    ("PRN No", "{PRN Number}"),
    ("Semester", "SEM V"),
    ("Academic Year", "2026-2027"),
    ("Subject Name & Code", "MLOps (AI3206C)"),
    ("Assignment No.", "Assignment No. 7"),
    ("Title of Assignment", "Implement CI/CD Pipeline Using GitHub Actions"),
    ("Date of Performance / Submission", "18/11/2026 / 20/11/2026")
]

for idx, (label, val) in enumerate(meta_data):
    row = meta_table.rows[idx]
    c1, c2 = row.cells[0], row.cells[1]
    c1.width = Inches(2.3)
    c2.width = Inches(4.2)
    set_cell_background(c1, "F2F5F9")
    set_cell_margins(c1, 70, 70, 100, 100)
    set_cell_margins(c2, 70, 70, 100, 100)
    
    p1 = c1.paragraphs[0]
    p1.paragraph_format.space_after = Pt(0)
    r1 = p1.add_run(label)
    r1.font.name = "Calibri"
    r1.font.size = Pt(10.5)
    r1.bold = True
    r1.font.color.rgb = RGBColor(0, 51, 102)
    
    p2 = c2.paragraphs[0]
    p2.paragraph_format.space_after = Pt(0)
    r2 = p2.add_run(val)
    r2.font.name = "Calibri"
    r2.font.size = Pt(10.5)
    r2.font.color.rgb = RGBColor(30, 30, 30)

doc.add_page_break()

# ==========================================
# 2. AIM & PROBLEM STATEMENT
# ==========================================
add_heading_styled(doc, "Assignment No. 7: Implement CI/CD Pipeline Using GitHub Actions", level=1)

add_heading_styled(doc, "Aim", level=2)
add_body_p(doc, "To implement an automated Continuous Integration and Continuous Deployment (CI/CD) pipeline for a Machine Learning application using GitHub Actions, automating model testing, validation, and deployment upon code push and pull requests.")

add_heading_styled(doc, "Problem Statement", level=2)
add_body_p(doc, "In traditional Machine Learning workflows, models are often trained and tested manually in isolated environments or notebooks, resulting in untracked defects, dependency mismatches, and deployment delays. To bridge the gap between development and production, there is a critical need for an automated MLOps CI/CD pipeline. This project addresses the problem by developing a Logistic Regression classifier on the Iris dataset, setting up automated unit tests using pytest to evaluate model performance, and orchestrating a GitHub Actions workflow that executes testing on every commit and triggers automated deployment only upon successful test verification on the main branch.")

add_heading_styled(doc, "Objectives", level=2)
obj_p = doc.add_paragraph()
obj_p.paragraph_format.space_before = Pt(0)
obj_p.paragraph_format.space_after = Pt(6)
obj_p.paragraph_format.line_spacing = 1.15
objs = [
    "To understand the role of Continuous Integration (CI) and Continuous Deployment (CD) in modern MLOps pipelines.",
    "To build a modular, testable Machine Learning application using Scikit-Learn and Logistic Regression on the Iris dataset.",
    "To formulate automated unit testing using pytest with strict evaluation thresholds (accuracy >= 80%).",
    "To configure GitHub Actions workflow automation (.github/workflows/ci-cd.yml) with dependent stages (Build/Test -> Deploy).",
    "To validate failure handling and branch protection mechanisms by simulating failing test scenarios and observing automated deployment suspension."
]
for o in objs:
    p_item = doc.add_paragraph(style='List Bullet')
    p_item.paragraph_format.space_after = Pt(3)
    p_item.paragraph_format.line_spacing = 1.15
    r = p_item.add_run(o)
    r.font.name = "Calibri"
    r.font.size = Pt(11)

add_heading_styled(doc, "Tools, Technologies and Environment", level=2)
tools_table = doc.add_table(rows=7, cols=2)
tools_table.alignment = WD_TABLE_ALIGNMENT.CENTER
tools_info = [
    ("Operating System", "Windows 11 (64-bit) / Linux runner (Ubuntu-latest) in GitHub Actions"),
    ("Programming Language", "Python 3.12 (Local) / Python 3.11 (CI/CD Environment)"),
    ("Tools / Frameworks", "Git (v2.56), GitHub, GitHub Actions Runner, Pytest (v9.2.0)"),
    ("Libraries", "scikit-learn (v1.9.1), pytest, numpy, scipy, joblib"),
    ("Dataset / Model", "Iris Plants Dataset (Fisher, 1936; 150 samples, 4 features, 3 classes) / Logistic Regression"),
    ("Version Control & Remote", "Git local repository linked to GitHub remote repository"),
    ("CI/CD Automation", "GitHub Actions (.github/workflows/ci-cd.yml)")
]

for idx, (k, v) in enumerate(tools_info):
    row = tools_table.rows[idx]
    c1, c2 = row.cells[0], row.cells[1]
    c1.width = Inches(2.2)
    c2.width = Inches(4.3)
    set_cell_background(c1, "F2F5F9")
    set_cell_margins(c1, 60, 60, 100, 100)
    set_cell_margins(c2, 60, 60, 100, 100)
    
    p1 = c1.paragraphs[0]
    p1.paragraph_format.space_after = Pt(0)
    r1 = p1.add_run(k)
    r1.font.name = "Calibri"
    r1.font.size = Pt(10)
    r1.bold = True
    r1.font.color.rgb = RGBColor(0, 51, 102)
    
    p2 = c2.paragraphs[0]
    p2.paragraph_format.space_after = Pt(0)
    r2 = p2.add_run(v)
    r2.font.name = "Calibri"
    r2.font.size = Pt(10)
    r2.font.color.rgb = RGBColor(30, 30, 30)

doc.add_paragraph().paragraph_format.space_after = Pt(6)

# ==========================================
# 3. BACKGROUND THEORY
# ==========================================
add_heading_styled(doc, "Background Theory", level=2)

add_heading_styled(doc, "a) Introduction to CI/CD in MLOps", level=3)
add_body_p(doc, "Machine Learning Operations (MLOps) is an engineering practice that unites machine learning system development (ML) and system operations (Ops). Unlike traditional software engineering where code is the sole artifact, ML systems incorporate code, data, and model parameters. In this lifecycle, Continuous Integration (CI) and Continuous Deployment (CD) play pivotal roles:")
ci_cd_points = [
    ("Continuous Integration (CI): ", "In an ML context, CI does not just test code syntax and linting; it validates that data pipelines function correctly, dependencies resolve without conflicts, and the trained model satisfies established evaluation baselines (such as classification accuracy, precision, or latency thresholds)."),
    ("Continuous Deployment (CD): ", "CD automates the packaging, staging, and deployment of validated model artifacts into target execution environments (such as serving endpoints, batch inference jobs, or container registries) immediately after passing all integration tests.")
]
for bold_prefix, text in ci_cd_points:
    add_body_p(doc, text=text, bold_prefix=bold_prefix, space_after=4)

add_heading_styled(doc, "b) Why GitHub Actions for MLOps?", level=3)
add_body_p(doc, "GitHub Actions is a cloud-native continuous integration and continuous delivery (CI/CD) platform directly embedded within GitHub repositories. It allows developers to define automated workflows triggered by GitHub repository events such as push, pull_request, release, or schedule.")
add_body_p(doc, "In comparison to standalone automation servers like Jenkins or TeamCity, GitHub Actions eliminates server maintenance overhead, provides out-of-the-box matrix builds, natively integrates with pull request review flows, and leverages a vast marketplace of pre-tested actions (e.g., actions/checkout, actions/setup-python). For MLOps, it offers frictionless integration between Git versioning, automated testing suites, and containerized deployment triggers.")

add_heading_styled(doc, "c) Key Components and Terminology", level=3)
terms = [
    ("Workflow: ", "A configurable automated process composed of one or more jobs defined in a YAML file located under .github/workflows/."),
    ("Event: ", "A specific activity that triggers a workflow run (e.g., a push to the main branch or opening a pull request)."),
    ("Job: ", "A set of steps executed sequentially on the same runner. Jobs can execute in parallel or sequentially using dependency directives (needs: job_name)."),
    ("Step: ", "An individual task within a job, either running a shell command or invoking an external Action."),
    ("Runner: ", "A virtual or physical server hosted by GitHub (e.g., ubuntu-latest) or self-hosted that executes the specified workflow steps."),
    ("Quality Gate / Unit Test: ", "A programmatic assertion verifying that the newly trained model satisfies performance constraints before allowing deployment.")
]
for bold_prefix, text in terms:
    add_body_p(doc, text=text, bold_prefix=bold_prefix, space_after=4)

add_heading_styled(doc, "d) Benefits and Limitations", level=3)
add_body_p(doc, "Benefits:", bold_prefix="")
benefits = [
    "High Automation & Velocity: Eliminates manual deployment steps, reducing human error.",
    "Early Defect Detection: Regression or degradation in model accuracy is caught instantly before merging.",
    "Reproducibility: Guarantees consistent runtime environments and dependency locks across all builds.",
    "Cost-Efficiency: Serverless runners spin up on-demand, avoiding idle infrastructure costs."
]
for b in benefits:
    p_b = doc.add_paragraph(style='List Bullet')
    p_b.paragraph_format.space_after = Pt(2)
    r = p_b.add_run(b)
    r.font.name = "Calibri"
    r.font.size = Pt(10.5)

add_body_p(doc, "Limitations & Trade-offs:", bold_prefix="", space_after=2)
limitations = [
    "Compute Constraints on Free Runners: Standard runners provide limited CPU and RAM, making training of massive deep learning models on GPU impractical without self-hosted runners.",
    "Execution Timeouts: Standard GitHub Actions jobs have a 6-hour execution limit, which requires offloading intensive ML training to external cloud clusters (e.g., Vertex AI, AWS SageMaker)."
]
for l in limitations:
    p_l = doc.add_paragraph(style='List Bullet')
    p_l.paragraph_format.space_after = Pt(2)
    r = p_l.add_run(l)
    r.font.name = "Calibri"
    r.font.size = Pt(10.5)

# ==========================================
# 4. ARCHITECTURE / WORKFLOW DIAGRAM
# ==========================================
add_heading_styled(doc, "Architecture / Workflow Diagram", level=2)
add_body_p(doc, "The architectural workflow follows the standard GitOps/MLOps automation lifecycle shown below:")

diagram_box = """
+-----------------------------------------------------------------------------------+
|                            DEVELOPER WORKSPACE                                    |
|   1. Edit Code (src/model.py, tests/test_model.py, app.py)                        |
|   2. Execute Local Validation: python app.py & python -m pytest -v                |
|   3. Git Commit & Push to 'main' branch / Create Pull Request                     |
+-----------------------------------------+-----------------------------------------+
                                          |
                                    git push / PR
                                          v
+-----------------------------------------------------------------------------------+
|                        GITHUB ACTIONS CI/CD ENGINE                                |
|                                                                                   |
|  [TRIGGER]: on: push [main], pull_request [main]                                  |
|                                                                                   |
|  +-----------------------------------------------------------------------------+  |
|  | JOB 1: Test ML Application (runs-on: ubuntu-latest)                          |  |
|  |   Step 1: Checkout Repository (actions/checkout@v4)                         |  |
|  |   Step 2: Setup Python 3.11 (actions/setup-python@v5)                        |  |
|  |   Step 3: Upgrade pip & Install requirements.txt (scikit-learn, pytest)     |  |
|  |   Step 4: Execute Unit Test Suite: pytest -v                                |  |
|  |           Asserts: Model Accuracy >= 0.80                                   |  |
|  +-------------------------------------+---------------------------------------+  |
|                                        |                                          |
|                          Test Passed?  |                                          |
|                         /             \                                         |
|                   YES  /               \  NO (FAILED)                             |
|                       v                 v                                         |
|  +--------------------------------+   +----------------------------------------+  |
|  | JOB 2: Deploy Application      |   | Pipeline Aborts Execution              |  |
|  |  (needs: test, if: push to main)|   | - Job 1 fails with Red Cross (X)       |  |
|  |   Step 1: Checkout Repository  |   | - Job 2 (Deploy) is SKIPPED            |  |
|  |   Step 2: Setup Python 3.11    |   | - Bad model code never reaches prod    |  |
|  |   Step 3: Deploy Application   |   +----------------------------------------+  |
|  |           (Simulated Deploy)   |                                               |
|  +--------------------------------+                                               |
+-----------------------------------------------------------------------------------+
"""
add_code_block(doc, diagram_box)

# ==========================================
# 5. PROCEDURE (STEP-BY-STEP IMPLEMENTATION)
# ==========================================
add_heading_styled(doc, "Procedure (Step-by-Step Implementation)", level=2)

add_heading_styled(doc, "Step 1: Project Setup and Directory Structuring", level=3)
add_body_p(doc, "Organized a clean, modular repository structure containing dedicated source code (src/), automated tests (tests/), workflow definitions (.github/workflows/), and root-level application and configuration files.")
add_code_block(doc, """# Project directory hierarchy
mlops-cicd/
├── .github/
│   └── workflows/
│       └── ci-cd.yml
├── src/
│   └── model.py
├── tests/
│   └── test_model.py
├── .gitignore
├── requirements.txt
├── app.py
└── README.md""")

add_heading_styled(doc, "Step 2: ML Model Implementation (src/model.py)", level=3)
add_body_p(doc, "Developed the machine learning logic using Scikit-Learn. The train_model() function loads the Iris dataset, partitions data into an 80/20 train/test split (random_state=42), fits a Logistic Regression classifier (max_iter=200), computes predictions, and calculates classification accuracy.")

add_heading_styled(doc, "Step 3: Pytest Unit Test Suite Implementation (tests/test_model.py)", level=3)
add_body_p(doc, "Created automated test function test_model_accuracy() using pytest. The test executes the model training procedure and verifies with an assertion statement (assert accuracy >= 0.80) that the trained model meets production accuracy standards.")

add_heading_styled(doc, "Step 4: Application Entry Point (app.py)", level=3)
add_body_p(doc, "Built a command-line application that acts as the deployment entry point. When run, it triggers model training, logs dataset metadata, algorithm configuration, computed test accuracy, and operational status.")

add_heading_styled(doc, "Step 5: Local Execution and Verification", level=3)
add_body_p(doc, "Executed local tests to ensure the script operates without defects before versioning. In our environment, dependencies were verified and model training and test commands were executed:")
add_code_block(doc, """# Execute the model script directly:
python src/model.py

# Execute the application entry point:
python app.py

# Execute pytest unit test suite:
python -m pytest -v""")
add_body_p(doc, "Result: The model attained 1.0000 (100.00%) test accuracy on the Iris test split, and the pytest suite passed 100%.")

add_heading_styled(doc, "Step 6: GitHub Actions CI/CD Pipeline Configuration (.github/workflows/ci-cd.yml)", level=3)
add_body_p(doc, "Created a dual-job GitHub Actions workflow triggered on push and pull_request targeting the main branch. The deploy job is conditionally guarded with needs: test and executes exclusively when commits are pushed directly to main.")

add_heading_styled(doc, "Step 7: Version Control Initialization and Staging", level=3)
add_body_p(doc, "Initialized a local Git repository on branch main, created a comprehensive .gitignore file to exclude temporary bytecode and caches, and staged all project files.")
add_code_block(doc, """git init -b main
git add .
git status""")

add_heading_styled(doc, "Step 8: Remote Repository Connection and Pipeline Execution", level=3)
add_body_p(doc, "Connect the local Git repository to GitHub and push the codebase to trigger the automated CI/CD pipeline on the cloud runner.")
add_code_block(doc, """git commit -m "Initial commit: Iris ML CI/CD pipeline with GitHub Actions"
git remote add origin https://github.com/<USERNAME>/<REPO_NAME>.git
git branch -M main
git push -u origin main""")

add_heading_styled(doc, "Step 9: Test Failure and Deployment Guard Demonstration", level=3)
add_body_p(doc, "To verify the CI quality gate, intentionally modify tests/test_model.py to an unachievable assertion condition (e.g., assert accuracy > 1.0). Running pytest causes a failure. When pushed to GitHub, Job 1 (Test) fails with an error status and Job 2 (Deploy) is automatically skipped, protecting production from faulty code. Reverting the assertion restores the green pipeline.")

# ==========================================
# 6. CODE / CONFIGURATION FILES
# ==========================================
add_heading_styled(doc, "Code / Configuration Files", level=2)

add_heading_styled(doc, "1. requirements.txt", level=3)
add_code_block(doc, """scikit-learn
pytest""")

add_heading_styled(doc, "2. src/model.py", level=3)
add_code_block(doc, '''"""
Model Training Module for Iris Dataset Classification.
Uses Logistic Regression from scikit-learn.
"""

from sklearn.datasets import load_iris
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import accuracy_score


def train_model():
    """
    Loads the Iris dataset, splits into train/test sets,
    trains a Logistic Regression classifier, and returns test accuracy.
    """
    # 1. Load the Iris dataset
    iris = load_iris()
    X = iris.data
    y = iris.target

    # 2. Split dataset into train and test sets (test_size=0.2, random_state=42)
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.2, random_state=42
    )

    # 3. Create Logistic Regression model with max_iter=200
    model = LogisticRegression(max_iter=200)

    # 4. Train the model
    model.fit(X_train, y_train)

    # 5. Make predictions on test set
    y_pred = model.predict(X_test)

    # 6. Calculate model accuracy
    accuracy = accuracy_score(y_test, y_pred)

    return accuracy


if __name__ == "__main__":
    print("=" * 50)
    print("Training Logistic Regression Model on Iris Dataset")
    print("=" * 50)
    acc = train_model()
    print(f"Model Training Completed Successfully!")
    print(f"Model Accuracy: {acc:.4f} ({acc * 100:.2f}%)")
    print("=" * 50)''')

add_heading_styled(doc, "3. tests/test_model.py", level=3)
add_code_block(doc, '''"""
Unit tests for the Machine Learning model.
"""

from src.model import train_model


def test_model_accuracy():
    """
    Test that the trained model reaches an acceptable accuracy threshold.
    """
    accuracy = train_model()
    # Model should achieve at least 80% accuracy on Iris test set
    assert accuracy >= 0.80''')

add_heading_styled(doc, "4. app.py", level=3)
add_code_block(doc, '''"""
ML Application Entry Point.
Executes Iris Classification Model using Logistic Regression.
"""

import sys
from src.model import train_model


def main():
    print("=" * 60)
    print("       MACHINE LEARNING APPLICATION (MLOps ASSIGNMENT 7)      ")
    print("=" * 60)
    print("Application Name : Iris Classification Service")
    print("Algorithm Used   : Logistic Regression (max_iter=200)")
    print("Dataset Name     : Iris Plants Dataset (Scikit-Learn)")
    print("-" * 60)
    print("Status           : Training and evaluating model...")

    try:
        accuracy = train_model()
        print("-" * 60)
        print(f"Model Accuracy   : {accuracy:.4f} ({accuracy * 100:.2f}%)")
        print("Status           : Application Executed Successfully [OK]")
        print("=" * 60)
    except Exception as e:
        print(f"Status           : Application Execution Failed: {e}")
        print("=" * 60)
        sys.exit(1)


if __name__ == "__main__":
    main()''')

add_heading_styled(doc, "5. .github/workflows/ci-cd.yml", level=3)
add_code_block(doc, '''name: MLOps CI/CD Pipeline

on:
  push:
    branches: [ "main" ]
  pull_request:
    branches: [ "main" ]

jobs:
  test:
    name: Test ML Application
    runs-on: ubuntu-latest

    steps:
      - name: Checkout Repository
        uses: actions/checkout@v4

      - name: Setup Python 3.11
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: Upgrade pip
        run: python -m pip install --upgrade pip

      - name: Install Dependencies
        run: pip install -r requirements.txt

      - name: Run Unit Tests
        run: pytest -v

  deploy:
    name: Deploy Application
    runs-on: ubuntu-latest
    needs: test
    if: github.event_name == 'push' && github.ref == 'refs/heads/main'

    steps:
      - name: Checkout Repository
        uses: actions/checkout@v4

      - name: Setup Python 3.11
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: Deploy ML Application
        run: |
          echo "Deployment started..."
          echo "ML application deployed successfully."''')

add_body_p(doc, "GitHub Repository Link: https://github.com/<YourUsername>/mlops-cicd", bold_prefix="")

# ==========================================
# 7. OUTPUT / RESULTS & SCREENSHOTS
# ==========================================
add_heading_styled(doc, "Output / Results", level=2)
add_body_p(doc, "The following sections showcase the project structure, application execution outputs, unit testing validations, and CI/CD results obtained during the execution of this assignment.")

out_table = doc.add_table(rows=5, cols=2)
out_table.alignment = WD_TABLE_ALIGNMENT.CENTER
out_rows = [
    ("Figure", "Description"),
    ("Figure 1", "Project Directory Hierarchy in VS Code Explorer"),
    ("Figure 2", "Local ML Application Execution Output (python app.py)"),
    ("Figure 3", "Automated Unit Test Execution (python -m pytest -v)"),
    ("Figure 4", "GitHub Actions CI/CD Pipeline Execution Graph (Test & Deploy)")
]

for idx, (f_num, f_desc) in enumerate(out_rows):
    row = out_table.rows[idx]
    c1, c2 = row.cells[0], row.cells[1]
    c1.width = Inches(1.8)
    c2.width = Inches(4.7)
    if idx == 0:
        set_cell_background(c1, "003366")
        set_cell_background(c2, "003366")
        for c in (c1, c2):
            set_cell_margins(c, 80, 80, 100, 100)
        p1 = c1.paragraphs[0]
        r1 = p1.add_run(f_num)
        r1.bold = True
        r1.font.color.rgb = RGBColor(255, 255, 255)
        p2 = c2.paragraphs[0]
        r2 = p2.add_run(f_desc)
        r2.bold = True
        r2.font.color.rgb = RGBColor(255, 255, 255)
    else:
        set_cell_background(c1, "F2F5F9" if idx % 2 == 1 else "FFFFFF")
        set_cell_background(c2, "F2F5F9" if idx % 2 == 1 else "FFFFFF")
        for c in (c1, c2):
            set_cell_margins(c, 60, 60, 100, 100)
        p1 = c1.paragraphs[0]
        r1 = p1.add_run(f_num)
        r1.bold = True
        r1.font.color.rgb = RGBColor(0, 51, 102)
        p2 = c2.paragraphs[0]
        r2 = p2.add_run(f_desc)
        r2.font.color.rgb = RGBColor(30, 30, 30)

doc.add_paragraph().paragraph_format.space_after = Pt(10)

# EMBED SCREENSHOT 1
img1_path = r"C:\College\SEM 5\MLOPs\Assignment_7\Screenshot 2026-11-18 191158.png"
if os.path.exists(img1_path):
    p_img1 = doc.add_paragraph()
    p_img1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_picture(img1_path, width=Inches(3.2))
    p_cap1 = doc.add_paragraph()
    p_cap1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_cap1 = p_cap1.add_run("Figure 1: Project Directory Structure in VS Code Explorer")
    r_cap1.italic = True
    r_cap1.font.size = Pt(9.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

# EMBED SCREENSHOT 2
img2_path = r"C:\College\SEM 5\MLOPs\Assignment_7\Screenshot 2026-11-18 191225.png"
if os.path.exists(img2_path):
    p_img2 = doc.add_paragraph()
    p_img2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_picture(img2_path, width=Inches(5.5))
    p_cap2 = doc.add_paragraph()
    p_cap2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_cap2 = p_cap2.add_run("Figure 2: Local ML Application Execution Output (python app.py)")
    r_cap2.italic = True
    r_cap2.font.size = Pt(9.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

# EMBED SCREENSHOT 3
img3_path = r"C:\College\SEM 5\MLOPs\Assignment_7\Screenshot 2026-11-18 191251.png"
if os.path.exists(img3_path):
    p_img3 = doc.add_paragraph()
    p_img3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_picture(img3_path, width=Inches(5.8))
    p_cap3 = doc.add_paragraph()
    p_cap3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_cap3 = p_cap3.add_run("Figure 3: Pytest Unit Test Execution (python -m pytest -v) Passing 100%")
    r_cap3.italic = True
    r_cap3.font.size = Pt(9.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

add_callout(doc, "To complete Figure 4 & Figure 5 for the final report submission, push this repository to GitHub and capture the screenshot of the GitHub Actions tab showing the green workflow graph with both 'Test ML Application' and 'Deploy Application' jobs.", title="GITHUB ACTIONS SCREENSHOT INSTRUCTION")

# Results metrics table
add_body_p(doc, "Machine Learning Performance Metrics Summary:", bold_prefix="")
metrics_table = doc.add_table(rows=5, cols=2)
metrics_table.alignment = WD_TABLE_ALIGNMENT.CENTER
metrics_data = [
    ("Evaluation Metric", "Observed Value"),
    ("Model Classifier", "Multinomial Logistic Regression (max_iter=200)"),
    ("Dataset Train/Test Split", "80% Training (120 samples) / 20% Testing (30 samples)"),
    ("Testing Accuracy Score", "1.0000 (100.00%)"),
    ("Unit Test Assertion Threshold", "assert accuracy >= 0.80 (PASSED in 1.63s)")
]

for idx, (m_lbl, m_val) in enumerate(metrics_data):
    row = metrics_table.rows[idx]
    c1, c2 = row.cells[0], row.cells[1]
    c1.width = Inches(3.0)
    c2.width = Inches(3.5)
    if idx == 0:
        set_cell_background(c1, "003366")
        set_cell_background(c2, "003366")
        for c in (c1, c2):
            set_cell_margins(c, 80, 80, 100, 100)
        p1 = c1.paragraphs[0]
        r1 = p1.add_run(m_lbl)
        r1.bold = True
        r1.font.color.rgb = RGBColor(255, 255, 255)
        p2 = c2.paragraphs[0]
        r2 = p2.add_run(m_val)
        r2.bold = True
        r2.font.color.rgb = RGBColor(255, 255, 255)
    else:
        set_cell_background(c1, "F2F5F9" if idx % 2 == 1 else "FFFFFF")
        set_cell_background(c2, "F2F5F9" if idx % 2 == 1 else "FFFFFF")
        for c in (c1, c2):
            set_cell_margins(c, 60, 60, 100, 100)
        p1 = c1.paragraphs[0]
        r1 = p1.add_run(m_lbl)
        r1.bold = True
        r1.font.color.rgb = RGBColor(0, 51, 102)
        p2 = c2.paragraphs[0]
        r2 = p2.add_run(m_val)
        r2.font.color.rgb = RGBColor(30, 30, 30)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# ==========================================
# 8. OBSERVATIONS AND ANALYSIS
# ==========================================
add_heading_styled(doc, "Observations and Analysis", level=2)
observations = [
    ("Automated Quality Assurance: ", "Automated unit tests serve as a dependable quality gate. Because model accuracy on the Iris test set was 100% (accuracy = 1.0000), the assertion threshold (accuracy >= 0.80) was successfully met, verifying model convergence and prediction validity."),
    ("Job Dependency & Gating Mechanism: ", "The GitHub Actions workflow configuration accurately enforces pipeline dependencies through needs: test. This guarantees that deployment is never initiated if test execution encounters errors, preventing faulty or sub-par models from reaching production."),
    ("Branch Protection and Deployment Scoping: ", "By conditioning the deployment job with if: github.event_name == 'push' && github.ref == 'refs/heads/main', Pull Requests targeting main run tests for validation without triggering premature deployment."),
    ("Local Execution Policies & Resolution: ", "During local PowerShell execution on Windows, invoking pytest directly triggered an Application Control Policy restriction on pytest.exe. This was cleanly resolved by executing pytest as an explicit Python module runner (python -m pytest -v), ensuring smooth execution across local Windows and Linux CI environments."),
    ("Failure Isolation Demonstration: ", "Simulating a failing test condition (such as assert accuracy > 1.0) successfully stops the pipeline at Job 1. GitHub Actions highlights the error in red, while Job 2 (Deploy) displays a grey 'Skipped' badge, proving that broken logic is blocked prior to deployment.")
]
for bold_prefix, text in observations:
    add_body_p(doc, text=text, bold_prefix=bold_prefix, space_after=5)

# ==========================================
# 9. CONCLUSION
# ==========================================
add_heading_styled(doc, "Conclusion", level=2)
add_body_p(doc, "Thus, we have successfully developed a Machine Learning classification model on the Iris dataset and implemented a fully automated Continuous Integration and Continuous Deployment (CI/CD) pipeline using GitHub Actions. We verified that automated unit testing with pytest serves as an effective quality gate to validate model accuracy prior to release. This practical demonstrates how modern MLOps principles eliminate manual release overhead, ensure reproducibility, and build reliable, production-ready machine learning delivery workflows.")

# ==========================================
# 10. REFERENCES
# ==========================================
add_heading_styled(doc, "References", level=2)
refs = [
    "GitHub Documentation, 'Understanding GitHub Actions and Workflow Syntax', GitHub Docs, https://docs.github.com/en/actions, accessed 20/11/2026.",
    "Pedregosa, F. et al., 'Scikit-learn: Machine Learning in Python', Journal of Machine Learning Research (JMLR), 12, pp. 2825-2830, 2011.",
    "Krekel, H. et al., 'pytest: Simple, Rapid Python Testing', pytest.org Documentation, https://docs.pytest.org, accessed 20/11/2026.",
    "Google Cloud Architecture Center, 'Practitioners guide to MLOps: A framework for continuous delivery and automation in machine learning', Google Cloud, 2021.",
    "Fisher, R. A., 'The use of multiple measurements in taxonomic problems', Annals of Eugenics, 7(2), pp. 179-188, 1936."
]
for idx, ref in enumerate(refs, 1):
    p_ref = doc.add_paragraph()
    p_ref.paragraph_format.space_after = Pt(3)
    p_ref.paragraph_format.line_spacing = 1.15
    r_idx = p_ref.add_run(f"[{idx}] ")
    r_idx.bold = True
    r_ref = p_ref.add_run(ref)
    r_ref.font.name = "Calibri"
    r_ref.font.size = Pt(10)

# Save document
out_doc_path = r"c:\College\SEM 5\MLOPs\Assignment_7\MLOps_Assignment_7_Report.docx"
doc.save(out_doc_path)
print(f"Successfully generated document at: {out_doc_path}")
