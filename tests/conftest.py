import pytest
from synthetic import build_clean_deck, build_problem_deck


@pytest.fixture(scope="session")
def clean_deck(tmp_path_factory):
    return build_clean_deck(tmp_path_factory.mktemp("decks") / "clean.pptx")


@pytest.fixture(scope="session")
def problem_deck(tmp_path_factory):
    return build_problem_deck(tmp_path_factory.mktemp("decks") / "problem.pptx")
