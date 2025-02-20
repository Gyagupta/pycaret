import os
import sys

sys.path.insert(0, os.path.abspath(".."))

import daal4py
import pandas as pd
import sklearn

import pycaret.classification
import pycaret.datasets


def test_engines_setup_global_args():
    """Tests the setting of engines using global arguments in setup."""

    juice_dataframe = pycaret.datasets.get_data("juice")
    exp = pycaret.classification.ClassificationExperiment()

    # init setup
    exp.setup(
        juice_dataframe,
        target="Purchase",
        remove_multicollinearity=True,
        multicollinearity_threshold=0.95,
        log_experiment=True,
        html=False,
        session_id=123,
        n_jobs=1,
        engines={"lr": "sklearnex"},
    )

    #### Default Model Engine ----
    assert exp.get_engine("lr") == "sklearnex"
    model = exp.create_model("lr")
    assert isinstance(
        model, daal4py.sklearn.linear_model.logistic_path.LogisticRegression
    )


def test_engines_global_methods():
    """Tests the setting of engines using methods like set_engine (global changes)."""

    juice_dataframe = pycaret.datasets.get_data("juice")
    exp = pycaret.classification.ClassificationExperiment()

    # init setup
    exp.setup(
        juice_dataframe,
        target="Purchase",
        remove_multicollinearity=True,
        multicollinearity_threshold=0.95,
        log_experiment=True,
        html=False,
        session_id=123,
        n_jobs=1,
        engines={"lr": "sklearnex"},
    )

    assert exp.get_engine("lr") == "sklearnex"

    #### Globally reset engine ----
    exp._set_engine("lr", "sklearn")
    assert exp.get_engine("lr") == "sklearn"
    model = exp.create_model("lr")
    assert isinstance(model, sklearn.linear_model._logistic.LogisticRegression)


def test_create_model_engines_local_args():
    """Tests the setting of engines for create_model using local args."""

    juice_dataframe = pycaret.datasets.get_data("juice")
    exp = pycaret.classification.ClassificationExperiment()

    # init setup
    exp.setup(
        juice_dataframe,
        target="Purchase",
        remove_multicollinearity=True,
        multicollinearity_threshold=0.95,
        log_experiment=True,
        html=False,
        session_id=123,
        n_jobs=1,
    )

    #### Default Model Engine ----
    assert exp.get_engine("lr") == "sklearn"
    model = exp.create_model("lr")
    assert isinstance(model, sklearn.linear_model._logistic.LogisticRegression)

    #### Override model engine locally ----
    model = exp.create_model("lr", engine="sklearnex")
    assert isinstance(
        model, daal4py.sklearn.linear_model.logistic_path.LogisticRegression
    )
    # Original engine should remain the same
    assert exp.get_engine("lr") == "sklearn"


def test_compare_models_engines_local_args():
    """Tests the setting of engines for compare_models using local args."""

    juice_dataframe = pycaret.datasets.get_data("juice")
    exp = pycaret.classification.ClassificationExperiment()

    # init setup
    exp.setup(
        juice_dataframe,
        target="Purchase",
        remove_multicollinearity=True,
        multicollinearity_threshold=0.95,
        log_experiment=True,
        html=False,
        session_id=123,
        n_jobs=1,
    )

    #### Default Model Engine ----
    assert exp.get_engine("lr") == "sklearn"
    model = exp.compare_models(include=["lr"])

    assert isinstance(model, sklearn.linear_model._logistic.LogisticRegression)
    # Original engine should remain the same
    assert exp.get_engine("lr") == "sklearn"

    #### Override model engine locally ----
    model = exp.compare_models(include=["lr"], engines={"lr": "sklearnex"})
    assert isinstance(
        model, daal4py.sklearn.linear_model.logistic_path.LogisticRegression
    )
    # Original engine should remain the same
    assert exp.get_engine("lr") == "sklearn"
    model = exp.compare_models(include=["lr"])
    assert isinstance(model, sklearn.linear_model._logistic.LogisticRegression)


def test_all_sklearnex_models():

    ALGORITHMS_LIST = ["lr", "knn", "rbfsvm"]

    juice_dataframe = pycaret.datasets.get_data("juice")
    exp = pycaret.classification.ClassificationExperiment()

    # init setup
    exp.setup(
        juice_dataframe,
        target="Purchase",
        remove_multicollinearity=True,
        multicollinearity_threshold=0.95,
        log_experiment=True,
        html=False,
        session_id=123,
        n_jobs=1,
    )

    for algo in ALGORITHMS_LIST:
        model = exp.create_model(algo)
        parent_library = model.__module__
        assert parent_library.startswith("sklearn") == True

    for algo in ALGORITHMS_LIST:
        model = exp.create_model(algo, engine="sklearnex")
        parent_library = model.__module__
        if algo == "rbfsvm" or algo == "knn":
            assert parent_library.startswith("sklearnex") == True
        else:
            assert parent_library.startswith("daal4py") == True
