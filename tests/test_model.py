"""
Unit tests for the Machine Learning model.
"""

from src.model import train_model


def test_model_accuracy():
    """
    Test that the trained model reaches an acceptable accuracy threshold.
    """
    accuracy = train_model()
    # Model should achieve at least 80% accuracy on Iris test set
    assert accuracy >= 0.80
