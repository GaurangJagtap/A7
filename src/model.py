"""
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
    print("=" * 50)
