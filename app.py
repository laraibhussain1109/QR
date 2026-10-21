"""Run `python app.py --help` for camera and image-scanning options."""
from __future__ import annotations

import argparse
from pathlib import Path
import sys
import threading
import time

import cv2
from PyQt6.QtCore import QPoint, QPointF, Qt, QThread, pyqtSignal
from PyQt6.QtGui import QColor, QFont, QImage, QPainter, QPen, QPixmap, QPolygonF
from PyQt6.QtWidgets import (
    QApplication, QComboBox, QFileDialog, QHBoxLayout, QLabel, QLineEdit,
    QMainWindow, QPushButton, QSpinBox, QVBoxLayout, QWidget,
)

from parts import ExcelStore
from scanner import Scanner, make_sample


class CameraWorker(QThread):
    frame_ready = pyqtSignal(object, object)
    notice = pyqtSignal(str)
    recorded = pyqtSignal(object, str)

    def __init__(self, camera, output, backend):
        super().__init__()
        self.camera, self.output, self.backend = camera, output, backend
        self.frame_available = threading.Event()
        self.frame_available.set()

    def run(self):
        cap = None
        try:
            store = ExcelStore(self.output)
            store.initialize()
            cap = cv2.VideoCapture(self.camera, self.backend)
            if not cap.isOpened():
                raise RuntimeError("Camera unavailable. Check the camera index, permissions and other camera apps.")
            scanner = Scanner()
            recent = {}
            last_scan = 0
            detections = []
            failures = 0
            while not self.isInterruptionRequested():
                ok, frame = cap.read()
                if not ok:
                    failures += 1
                    if failures >= 10:
                        raise RuntimeError("Camera stopped delivering frames. Reconnect it and press Start.")
                    self.msleep(50)
                    continue
                failures = 0
                now = time.monotonic()
                if now - last_scan >= 0.2:
                    detections = scanner.detect(frame)
                    last_scan = now
                    recent = {key: stamp for key, stamp in recent.items() if now - stamp < 2}
                    for detection in detections:
                        if detection.payload in recent:
                            continue
                        recent[detection.payload] = now
                        if detection.part is None:
                            self.notice.emit("Unrecognized label: " + detection.error)
                            continue
                        try:
                            result = store.add(detection.part)
                            if result.added:
                                message = "Saved to Excel"
                            elif result.part != detection.part:
                                message = "Duplicate Part ID with different details — existing Excel row retained"
                            else:
                                message = "Already recorded — duplicate skipped"
                            self.recorded.emit(detection.part, message)
                        except Exception as exc:
                            self.notice.emit(f"Excel save failed; will retry: {exc}")
                # Apply backpressure: one outstanding frame keeps the UI responsive.
                if self.frame_available.is_set():
                    self.frame_available.clear()
                    self.frame_ready.emit(frame.copy(), list(detections))
                self.msleep(30)
        except Exception as exc:
            self.notice.emit(str(exc))
        finally:
            if cap is not None:
                cap.release()


class MainWindow(QMainWindow):
    def __init__(self, camera=0, output="manufacturing.xlsx"):
        super().__init__()
        self.worker = None
        self.closing = False
        self.setWindowTitle("Manufacturing • QR Part Scanner")
        self.resize(1100, 800)
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        controls = QHBoxLayout()
        self.camera = QSpinBox()
        self.camera.setRange(0, 99)
        self.camera.setValue(camera)
        self.backend = QComboBox()
        self.backend.addItem("Automatic camera backend", cv2.CAP_ANY)
        if sys.platform == "win32":
            self.backend.addItem("Windows DirectShow", cv2.CAP_DSHOW)
        self.output = QLineEdit(str(Path(output).expanduser().resolve()))
        self.browse = QPushButton("Choose Excel file…")
        self.browse.clicked.connect(self.choose_file)
        self.start = QPushButton("Start camera")
        self.start.clicked.connect(self.start_camera)
        self.stop = QPushButton("Stop")
        self.stop.setEnabled(False)
        self.stop.clicked.connect(self.stop_camera)
        for widget in (QLabel("Camera"), self.camera, self.backend, self.output,
                       self.browse, self.start, self.stop):
            controls.addWidget(widget)
        layout.addLayout(controls)
        self.view = QLabel("Press Start camera and present a QR label")
        self.view.setMinimumSize(640, 400)
        self.view.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.view.setStyleSheet("background:#101c29;color:#dbe9f7;border-radius:8px")
        layout.addWidget(self.view, 1)
        self.details = QLabel("Part Name: —    |    Part ID: —    |    BIN Number: —")
        self.details.setTextFormat(Qt.TextFormat.PlainText)
        self.details.setWordWrap(True)
        self.details.setStyleSheet("font-size:20px;padding:12px;background:#e5f0f7;color:#12354b")
        layout.addWidget(self.details)
        self.status = QLabel("Ready. Keep the workbook closed in Excel while scanning.")
        self.status.setTextFormat(Qt.TextFormat.PlainText)
        self.status.setWordWrap(True)
        layout.addWidget(self.status)

    def choose_file(self):
        path, _ = QFileDialog.getSaveFileName(self, "Select workbook", self.output.text(), "Excel (*.xlsx)")
        if path:
            self.output.setText(path if path.lower().endswith(".xlsx") else path + ".xlsx")

    def set_running(self, running):
        for widget in (self.start, self.camera, self.backend, self.output, self.browse):
            widget.setEnabled(not running)
        self.stop.setEnabled(running)

    def start_camera(self):
        if self.worker is not None:
            return
        self.set_running(True)
        self.status.setText("Opening camera…")
        self.worker = CameraWorker(self.camera.value(), self.output.text(), self.backend.currentData())
        self.worker.frame_ready.connect(self.show_frame)
        self.worker.notice.connect(self.status.setText)
        self.worker.recorded.connect(self.show_record)
        self.worker.finished.connect(self.worker_finished)
        self.worker.start()

    def stop_camera(self):
        if self.worker:
            self.worker.requestInterruption()
            self.stop.setEnabled(False)

    def worker_finished(self):
        self.worker.deleteLater()
        self.worker = None
        self.set_running(False)
        if self.closing:
            self.close()

    def show_record(self, part, message):
        self.status.setText(f"{message}: {part.part_id}")

    def show_frame(self, frame, detections):
        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        h, w, _ = rgb.shape
        image = QImage(rgb.data, w, h, rgb.strides[0], QImage.Format.Format_RGB888).copy()
        # Qt text rendering supports non-ASCII part names directly on the camera view.
        painter = QPainter(image)
        painter.setPen(QPen(QColor("#44ff99"), 3))
        painter.setFont(QFont("Sans Serif", max(12, w // 60)))
        valid = []
        for detection in detections:
            polygon = QPolygonF([QPointF(float(x), float(y)) for x, y in detection.corners])
            painter.drawPolygon(polygon)
            if detection.part:
                part = detection.part
                valid.append(part)
                x = max(5, int(min(detection.corners[:, 0])))
                y = max(25, int(min(detection.corners[:, 1])) - 10)
                label = f"{part.name}  |  BIN {part.bin_number}"
                rect = painter.fontMetrics().boundingRect(label).adjusted(-4, -4, 4, 4)
                rect.moveBottomLeft(QPoint(x, y + 4))
                painter.fillRect(rect, QColor(0, 0, 0, 210))
                painter.drawText(x, y, label)
        painter.end()
        self.details.setText("\n".join(f"Part Name: {p.name}    |    Part ID: {p.part_id}    |    BIN Number: {p.bin_number}" for p in valid)
                             or "No valid part label currently visible")
        self.view.setPixmap(QPixmap.fromImage(image).scaled(
            self.view.size(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        if self.worker:
            self.worker.frame_available.set()

    def closeEvent(self, event):
        if self.worker is not None:
            self.closing = True
            self.stop_camera()
            self.status.setText("Stopping camera before closing…")
            event.ignore()
        else:
            event.accept()


def main():
    parser = argparse.ArgumentParser(description="Scan QR part labels into duplicate-safe Excel rows.")
    parser.add_argument("--camera", type=int, default=0)
    parser.add_argument("--output", default="manufacturing.xlsx")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--scan-image", type=Path, help="Scan a still image without opening the GUI")
    mode.add_argument("--make-sample", type=Path, help="Write a sample QR label PNG and exit")
    args = parser.parse_args()
    if args.make_sample:
        make_sample(args.make_sample)
        print(f"Sample QR saved: {args.make_sample}")
        return 0
    if args.scan_image:
        frame = cv2.imread(str(args.scan_image))
        if frame is None:
            parser.error("Cannot read the input image.")
        detections = Scanner().detect(frame)
        if not detections:
            print("No QR code decoded.", file=sys.stderr)
            return 1
        store = ExcelStore(args.output)
        invalid = False
        for detection in detections:
            if detection.part is None:
                print(detection.error, file=sys.stderr)
                invalid = True
            else:
                result = store.add(detection.part)
                print(f"{'Saved' if result.added else 'Duplicate skipped'}: {detection.part}")
        return int(invalid)
    application = QApplication(sys.argv[:1])
    window = MainWindow(args.camera, args.output)
    window.show()
    return application.exec()


if __name__ == "__main__":
    sys.exit(main())
