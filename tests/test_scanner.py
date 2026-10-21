import cv2
import numpy as np
from openpyxl import load_workbook

from parts import ExcelStore, Part
from scanner import Scanner, make_sample


def test_real_qr_to_excel(tmp_path):
    image = tmp_path / "label.png"
    make_sample(image)
    detections = Scanner().detect(cv2.imread(str(image)))
    assert len(detections) == 1
    assert detections[0].part == Part("Steel Bracket", "P-001", "B-12")
    path = tmp_path / "scans.xlsx"
    assert ExcelStore(path).add(detections[0].part).added
    assert not ExcelStore(path).add(detections[0].part).added
    wb = load_workbook(path)
    assert wb["manufacturing"].max_row == 2
    wb.close()


def test_blank_frame():
    assert Scanner().detect(np.full((400, 400, 3), 255, dtype=np.uint8)) == []
