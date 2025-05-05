import sys
import os
import pymupdf  # PyMuPDF
import time  # Import time for timestamp generation
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QVBoxLayout, QPushButton, QLabel,
    QComboBox, QTextEdit, QWidget, QProgressBar, QMessageBox
)
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import QIcon

# Define colors
COLORS = {
    "Red": (1, 0, 0),
    "Green": (0, 1, 0),
    "Blue": (0, 0, 1),
    "Purple": (0.5, 0, 0.5),
    "Orange": (1, 0.647, 0),
    "Pink": (1, 0.75, 0.8),
    "Cyan": (0, 1, 1),
    "Light_Green": (0.5, 1, 0),
    "Light_Blue": (0.678, 0.847, 0.902),  # Sky Blue
    "Gray": (0.5, 0.5, 0.5),
    "Dark_Blue": (0, 0, 0.5),
    "Dark_Green": (0, 0.5, 0),
    "Yellow": (1, 1, 0)
}
DefaultColor = COLORS["Yellow"]
DefaultColorRGB = [int(DefaultColor[0] * 255), int(DefaultColor[1] * 255), int(DefaultColor[2] * 255)]
OUTPUT_DIR = os.path.abspath('Result')

class PDFHighlighterWorker(QThread):
    progress_signal = pyqtSignal(int)  # Signal to update progress bar
    status_signal = pyqtSignal(str)  # Signal to send status messages to the GUI
    finished_signal = pyqtSignal(str)  # Signal to notify when processing is complete

    def __init__(self, input_pdfs, output_pdfs, search_terms, highlight_color):
        super().__init__()
        self.input_pdfs = input_pdfs  # List of input PDF paths
        self.output_pdfs = output_pdfs  # List of output PDF paths
        self.search_terms = search_terms
        self.highlight_color = highlight_color

    def run(self):
        try:
            total_tasks = len(self.search_terms) * sum(pymupdf.open(pdf).page_count for pdf in self.input_pdfs)
            completed_tasks = 0
            color = COLORS.get(self.highlight_color, DefaultColor)

            for input_pdf, output_pdf in zip(self.input_pdfs, self.output_pdfs):
                doc = pymupdf.open(input_pdf)
                for idx, search_text in enumerate(self.search_terms, start=1):
                    self.status_signal.emit(f"Searching for term {idx}/{len(self.search_terms)}: '{search_text}' in {os.path.basename(input_pdf)}")
                    for page_num in range(doc.page_count):
                        page = doc.load_page(page_num)
                        self.status_signal.emit(f"Processing file: {os.path.basename(input_pdf)}, page {page_num + 1}/{doc.page_count}...")
                        text_instances = page.search_for(search_text)
                        for inst in text_instances:
                            highlight = page.add_highlight_annot(inst)
                            highlight.set_colors(stroke=color)
                            highlight.update()

                        # Update progress
                        completed_tasks += 1
                        progress = int((completed_tasks / total_tasks) * 100)
                        self.progress_signal.emit(progress)

                if not os.path.exists(OUTPUT_DIR):
                    os.makedirs(OUTPUT_DIR, exist_ok=True)

                output_path = os.path.join(OUTPUT_DIR, output_pdf)
                doc.save(output_path)
                self.status_signal.emit(f"Saved highlighted PDF to: {output_path}")

            self.finished_signal.emit("All PDFs processed successfully.")
        except Exception as e:
            self.status_signal.emit(f"Error: {e}")
            self.finished_signal.emit("")

class PDFHighlighterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Highlighter")

        # Adjust the path for the icon
        if hasattr(sys, '_MEIPASS'):
            icon_path = os.path.join(sys._MEIPASS, "icon.ico")
        else:
            icon_path = "icon.ico"

        self.setWindowIcon(QIcon(icon_path))  # Set the application icon
        self.setGeometry(300, 300, 600, 600)

        # Initialize variables
        self.pdf1_path = None
        self.pdf2_path = None
        self.txt_path = None
        self.highlight_color = None

        # Initialize UI components
        self.init_ui()

    def update_color_preview(self, color):
        """Update the color preview box based on the selected color."""
        rgb = COLORS[color]
        self.color_preview.setStyleSheet(f"background-color: rgb({int(rgb[0] * 255)}, {int(rgb[1] * 255)}, {int(rgb[2] * 255)});")

    def init_ui(self):
        """Initialize the user interface."""
        layout = QVBoxLayout()

        # Labels for file selection
        self.label_pdf1 = QLabel("First PDF: Not Selected")
        self.label_pdf2 = QLabel("Second PDF: Not Selected")
        self.label_txt = QLabel("Text File: Not Selected")
        layout.addWidget(self.label_pdf1)
        layout.addWidget(self.label_pdf2)
        layout.addWidget(self.label_txt)

        # Buttons for file selection
        btn_select_pdf1 = QPushButton("Select First PDF")
        btn_select_pdf1.clicked.connect(self.select_pdf1)
        layout.addWidget(btn_select_pdf1)

        btn_select_pdf2 = QPushButton("Select Second PDF (optional)")
        btn_select_pdf2.clicked.connect(self.select_pdf2)
        layout.addWidget(btn_select_pdf2)

        btn_select_txt = QPushButton("Select Text File")
        btn_select_txt.clicked.connect(self.select_txt)
        layout.addWidget(btn_select_txt)

        # ComboBox for selecting highlight color
        self.color_selector = QComboBox()
        self.color_selector.addItems(COLORS.keys())
        self.color_selector.setCurrentText("Yellow")
        self.color_selector.currentTextChanged.connect(self.select_color)

        # Color preview box
        self.color_preview = QLabel("Select Highlight Color:")
        # self.color_preview.setFixedSize(110, 15)
        self.color_preview.setStyleSheet(f"background-color: rgb({DefaultColorRGB[0]}, {DefaultColorRGB[1]}, {DefaultColorRGB[2]});")
        layout.addWidget(self.color_preview)

        # Update color preview when color is changed
        self.color_selector.currentTextChanged.connect(self.update_color_preview)
        self.color_selector.setCurrentText(self.highlight_color)
        layout.addWidget(self.color_selector)

        # Button to start processing
        self.btn_process = QPushButton("Highlight PDFs")
        self.btn_process.clicked.connect(self.start_processing)
        layout.addWidget(self.btn_process)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)  # Hide percentage text
        layout.addWidget(self.progress_bar)

        # TextEdit for log messages
        self.log_widget = QTextEdit()
        self.log_widget.setReadOnly(True)
        layout.addWidget(self.log_widget)

        # Set the central widget
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def set_interface_enabled(self, enabled):
        """Enable or disable the interface."""
        self.centralWidget().setEnabled(enabled)

    def select_pdf1(self):
        """Select the first PDF file."""
        file_path, _ = QFileDialog.getOpenFileName(self, "Select First PDF", "", "PDF Files (*.pdf)")
        if file_path:
            self.pdf1_path = file_path
            self.label_pdf1.setText(f"First PDF: {os.path.basename(file_path)}")

    def select_pdf2(self):
        """Select the second PDF file (optional)."""
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Second PDF", "", "PDF Files (*.pdf)")
        if file_path:
            self.pdf2_path = file_path
            self.label_pdf2.setText(f"Second PDF: {os.path.basename(file_path)}")

    def select_txt(self):
        """Select the text file containing search terms."""
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Text File", "", "Text Files (*.txt)")
        if file_path:
            self.txt_path = file_path
            self.label_txt.setText(f"Text File: {os.path.basename(file_path)}")

    def select_color(self, color):
        """Select the highlight color."""
        self.highlight_color = color

    def handle_existing_file(self, output_file):
        """Handle cases where the output file already exists."""
        if os.path.exists(output_file):
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("File Already Exists")
            msg_box.setText(f"The file '{os.path.basename(output_file)}' already exists.")
            msg_box.setInformativeText("What would you like to do?")
            overwrite_button = msg_box.addButton("Overwrite", QMessageBox.AcceptRole)
            rename_button = msg_box.addButton("Rename", QMessageBox.ActionRole)
            cancel_button = msg_box.addButton("Cancel", QMessageBox.RejectRole)
            msg_box.exec_()

            clicked_button = msg_box.clickedButton()
            if clicked_button == overwrite_button:
                return output_file
            elif clicked_button == rename_button:
                timestamp = time.strftime("%d.%m.%y_%H%M")
                return os.path.join(OUTPUT_DIR, f'{os.path.basename(output_file).split(".")[0]}_{timestamp}.pdf')
            elif clicked_button == cancel_button:
                QMessageBox.information(self, "Operation Cancelled", "The operation has been cancelled.")
                return None
        return output_file

    def start_processing(self):
        """Start the PDF highlighting process."""
        if not self.pdf1_path or not self.txt_path:
            QMessageBox.warning(self, "Missing Input", "Please select at least one PDF file and a text file.")
            return

        with open(self.txt_path, "r") as file:
            search_terms = file.read().splitlines()

        # Store input and output PDFs in lists
        input_pdfs = [self.pdf1_path]
        output_pdfs = [os.path.basename(self.pdf1_path).replace(".pdf", "_highlighted.pdf")]

        # Add the second PDF if it is provided
        if self.pdf2_path:
            input_pdfs.append(self.pdf2_path)
            output_pdfs.append(os.path.basename(self.pdf2_path).replace(".pdf", "_highlighted.pdf"))

        # Handle existing files
        for i, output_file in enumerate(output_pdfs):
            output_path = os.path.join(OUTPUT_DIR, output_file)
            new_output_path = self.handle_existing_file(output_path)
            if new_output_path is None:  # User canceled the operation
                self.set_interface_enabled(True)
                return
            output_pdfs[i] = new_output_path

        # Disable the interface
        self.set_interface_enabled(False)

        # Create and start the worker
        self.worker = PDFHighlighterWorker(input_pdfs, output_pdfs, search_terms, self.highlight_color)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.status_signal.connect(self.log_message)
        self.worker.finished_signal.connect(self.processing_finished)
        self.worker.start()

    def update_progress(self, value):
        """Update the progress bar."""
        self.progress_bar.setValue(value)

    def log_message(self, message):
        """Log messages to the QTextEdit."""
        self.log_widget.append(message)

    def processing_finished(self, message):
        """Handle the completion of the PDF processing."""
        self.set_interface_enabled(True)  # Re-enable the interface
        self.log_message(message)
        QMessageBox.information(self, "Processing Complete", message)
        


def main():
    app = QApplication(sys.argv)
    window = PDFHighlighterApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()