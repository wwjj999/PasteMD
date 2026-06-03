from pastemd.utils.fs import sanitize_filename


def test_sanitize_filename_keeps_regular_names():
    assert sanitize_filename("Chapter 1: Intro") == "Chapter 1_ Intro"


def test_sanitize_filename_avoids_windows_reserved_names():
    assert sanitize_filename("CON") == "CON_"
    assert sanitize_filename("con") == "con_"
    assert sanitize_filename("AUX.txt") == "AUX_.txt"
    assert sanitize_filename("LPT9") == "LPT9_"


def test_sanitize_filename_strips_trailing_dots_and_spaces():
    assert sanitize_filename("report. ") == "report"
    assert sanitize_filename("...") == "document"
