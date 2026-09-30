from make_synthetic import build_rows, write_xlsx
from voucher.verify import main, verify


def test_verify_synthetic_counts(synthetic_path):
    rep = verify(synthetic_path)
    assert rep["원본_비어있지않은행"] == rep["파서_행수"] == rep["다운로드_원본분개행"] == len(build_rows())
    assert rep["다운로드_전표요약건"] == rep["전표수(회계단위+전표관리단위+기표번호)"] == 60
    assert rep["차이"] == []
    assert rep["파일첨부여부구분"] == {"문자열 FALSE": 160}


def test_verify_cli(synthetic_path, capsys):
    assert main(["x", str(synthetic_path)]) == 0
    out = capsys.readouterr().out
    assert "차이 없음" in out
    assert "원재료 매입" not in out  # 전표 내용(적요)은 출력하지 않음
