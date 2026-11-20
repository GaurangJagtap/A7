"""
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
    main()
