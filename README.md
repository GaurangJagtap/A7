# Assignment No. 7: Implement CI/CD Pipeline Using GitHub Actions

## 1. Project Title
**MLOps CI/CD Pipeline for Iris Classification using GitHub Actions**

---

## 2. Project Description
This project demonstrates an end-to-end basic Continuous Integration and Continuous Deployment (CI/CD) pipeline for a Machine Learning application. The application trains and evaluates a Logistic Regression classifier on the Iris dataset. Automation is managed using GitHub Actions, ensuring that every code change is validated through automated unit testing before proceeding to deployment.

---

## 3. Aim
To develop a simple Machine Learning application, set up unit tests using `pytest`, and build an automated CI/CD pipeline using GitHub Actions that triggers on code push and pull requests targeting the `main` branch.

---

## 4. Technologies Used
- **Programming Language:** Python 3.11 / 3.12
- **Machine Learning Library:** `scikit-learn`
- **Testing Framework:** `pytest`
- **Version Control:** Git & GitHub
- **CI/CD Automation:** GitHub Actions

---

## 5. Project Structure
```text
mlops-cicd/
│
├── .github/
│   └── workflows/
│       └── ci-cd.yml          # GitHub Actions CI/CD pipeline definition
│
├── src/
│   └── model.py               # Model training and evaluation logic
│
├── tests/
│   └── test_model.py          # Pytest unit tests for model accuracy
│
├── requirements.txt           # Project dependencies
├── app.py                     # Main CLI application entry point
├── .gitignore                 # Git ignore file for Python artifacts
└── README.md                  # Project documentation
```

---

## 6. How the ML Model Works
The model is implemented in `src/model.py`:
1. **Dataset:** Loads the standard Iris flower dataset (150 samples across 3 classes: Setosa, Versicolour, Virginica).
2. **Train/Test Split:** Splits features ($X$) and targets ($y$) with an 80/20 train-test ratio (`test_size=0.2`) and fixed random seed (`random_state=42`).
3. **Algorithm:** Uses `LogisticRegression(max_iter=200)` to ensure convergence across multi-class predictions.
4. **Evaluation:** Measures accuracy using `accuracy_score(y_test, y_pred)`. Typically achieves 100% accuracy on this test split.

---

## 7. How to Install Dependencies
Create and activate a virtual environment (recommended), then install requirements:

```bash
# On Windows (PowerShell):
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# Install requirements:
pip install -r requirements.txt
```

---

## 8. How to Execute the Application
Run the main application script:

```bash
python app.py
```

Expected output:
```text
============================================================
       MACHINE LEARNING APPLICATION (MLOps ASSIGNMENT 7)      
============================================================
Application Name : Iris Classification Service
Algorithm Used   : Logistic Regression (max_iter=200)
Dataset Name     : Iris Plants Dataset (Scikit-Learn)
------------------------------------------------------------
Status           : Training and evaluating model...
------------------------------------------------------------
Model Accuracy   : 1.0000 (100.00%)
Status           : Application Executed Successfully [OK]
============================================================
```

---

## 9. How to Execute Unit Tests
Run `pytest` in verbose mode:

```bash
pytest -v
```

Expected output:
```text
tests/test_model.py::test_model_accuracy PASSED                      [100%]
============================== 1 passed in ...s ===============================
```

---

## 10. Explanation of the CI/CD Workflow
The file `.github/workflows/ci-cd.yml` defines the automated pipeline:
- **Triggers:**
  - `push` to branch `main`
  - `pull_request` targeting branch `main`
- **Job 1: `test` (Test ML Application):**
  - Runs on `ubuntu-latest`.
  - Sets up Python 3.11.
  - Upgrades `pip` and installs `requirements.txt`.
  - Executes `pytest -v`.
  - If tests fail, the job fails and stops subsequent execution.
- **Job 2: `deploy` (Deploy Application):**
  - Depends on `test` (`needs: test`).
  - Condition: Runs **only** for `push` events to `main` (`github.event_name == 'push' && github.ref == 'refs/heads/main'`).
  - Pull requests trigger tests only, preventing unintended deployment.
  - Simulates deployment with echo status messages.

---

## 11. Explanation of Continuous Integration (CI)
**Continuous Integration** is the practice where developers regularly merge their code changes into a central repository. Automated builds and tests run automatically to catch integration issues and bugs early before code is merged into production.

---

## 12. Explanation of Continuous Deployment (CD)
**Continuous Deployment** is the practice of automatically releasing validated code changes to the production or staging environment without manual intervention, provided all stages of the CI pipeline succeed.

---

## 13. How to Create and Connect a GitHub Repository
1. Log in to [GitHub](https://github.com).
2. Click **New Repository** (or **+** > **New repository**).
3. Name it (e.g., `mlops-assignment-7` or `mlops-cicd`).
4. Keep it **Public** (or Private) and leave initialize with README/gitignore unchecked.
5. In your local terminal, run:
   ```bash
   git remote add origin https://github.com/<YOUR_USERNAME>/<YOUR_REPOSITORY_NAME>.git
   git branch -M main
   git push -u origin main
   ```

---

## 14. How to View GitHub Actions Results
1. Navigate to your repository on GitHub.
2. Click on the **Actions** tab at the top.
3. Select the latest workflow run named **MLOps CI/CD Pipeline**.
4. You will see two stages: **Test ML Application** and **Deploy Application**.
5. Click on either job to view detailed execution logs.

---

## 15. How to Demonstrate the Test Failure Scenario
To demonstrate pipeline failure and protection against bad deployments:
1. In `tests/test_model.py`, change:
   ```python
   assert accuracy >= 0.80
   ```
   to an impossible threshold:
   ```python
   assert accuracy >= 0.99
   ```
   *(or if accuracy is 1.0, set `assert accuracy > 1.0` or `assert accuracy >= 1.01`)*
2. Run `pytest -v` locally to observe the failure (`FAILED tests/test_model.py::test_model_accuracy`).
3. Commit and push the failure:
   ```bash
   git add tests/test_model.py
   git commit -m "Demonstrate CI test failure"
   git push
   ```
4. Check GitHub Actions: The **Test ML Application** job fails, and **Deploy Application** is skipped.
5. Revert the test assertion back to `assert accuracy >= 0.80`.
6. Commit and push again to see the pipeline pass with both jobs succeeding!
