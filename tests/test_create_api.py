import os
import sys

sys.path.insert(0, os.path.abspath(".."))

import numpy as np
import pandas as pd
import pytest

import pycaret.classification
import pycaret.datasets
import pycaret.regression


def test_classification_create_api():

    # loading dataset
    data = pycaret.datasets.get_data("blood")

    # initialize setup
    clf1 = pycaret.classification.setup(
        data,
        target="Class",
        html=False,
        n_jobs=1,
    )

    # train model
    lr = pycaret.classification.create_model("lr")

    # create api
    pycaret.classification.create_api(lr, "blood_api")
    assert 1 == 1


def test_regression_create_api():

    # loading dataset
    data = pycaret.datasets.get_data("boston")

    # initialize setup
    reg1 = pycaret.regression.setup(
        data,
        target="medv",
        html=False,
        n_jobs=1,
    )

    # train model
    lr = pycaret.regression.create_model("lr")

    # create api
    pycaret.regression.create_api(lr, "boston_api")
    assert 1 == 1


if __name__ == "__main__":
    test_classification_create_api()
    test_regression_create_api()
