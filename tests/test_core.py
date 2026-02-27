from pathlib import Path

from sprut_importer.core import execution_code, parse_file_name


def test_parse_file_name() -> None:
    code, name = parse_file_name(Path("КУМП.422410.031-01 ПИ Пульта управления.xls"))
    assert code == "КУМП.422410.031-01"
    assert name == "ПИ Пульта управления"


def test_execution_code_with_suffix() -> None:
    assert execution_code("КУМП.422410.031-01", 0) == "КУМП.422410.031-01"
    assert execution_code("КУМП.422410.031-01", 1) == "КУМП.422410.031-02"


def test_execution_code_without_suffix() -> None:
    assert execution_code("КУМП.325811.013", 0) == "КУМП.325811.013"
    assert execution_code("КУМП.325811.013", 2) == "КУМП.325811.013-02"
