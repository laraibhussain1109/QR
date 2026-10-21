"""OpenCV QR detection shared by the camera UI and image smoke checks."""
from dataclasses import dataclass

import cv2
import numpy as np

from parts import Part, parse_payload


@dataclass
class Detection:
    payload: str
    corners: np.ndarray
    part: Part | None
    error: str = ""


class Scanner:
    def __init__(self):
        self.detector = cv2.QRCodeDetector()

    def detect(self, frame):
        found, texts, points, _ = self.detector.detectAndDecodeMulti(frame)
        if not found or not any(texts):
            text, corners, _ = self.detector.detectAndDecode(frame)
            texts, points = ([text], [corners[0]]) if text and corners is not None else ([], [])
        results = []
        for text, corners in zip(texts, points if points is not None else []):
            if not text:
                continue
            try:
                results.append(Detection(text, corners, parse_payload(text)))
            except ValueError as exc:
                results.append(Detection(text, corners, None, str(exc)))
        return results


def make_sample(path):
    payload = '{"part_name":"Steel Bracket","part_id":"P-001","bin_number":"B-12"}'
    code = cv2.QRCodeEncoder_create().encode(payload)
    code = cv2.copyMakeBorder(code, 4, 4, 4, 4, cv2.BORDER_CONSTANT, value=255)
    code = cv2.resize(code, None, fx=10, fy=10, interpolation=cv2.INTER_NEAREST)
    if not cv2.imwrite(str(path), code):
        raise OSError(f"Could not save {path}")
