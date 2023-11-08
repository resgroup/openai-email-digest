from dotenv import load_dotenv
from digest_emails.digest import get_summary

import sys
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
    QSpacerItem,
    QSizePolicy,
)
from PyQt6.QtGui import QFont, QIcon


# This would be your email summarization function
def generate_email_summary():
    # For now, we will just return a placeholder string
    # Replace this with your actual email summary generation code
    return "This is where your email summary will appear."


class EmailSummaryApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowIcon(QIcon("icon.png"))

        # Set the title and initial size of the window
        self.setWindowTitle("Email Summary Generator")
        self.setGeometry(100, 100, 600, 400)

        # Create the central widget
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        # Create the button and text edit
        self.button = QPushButton("Generate Summary", self)
        self.text_edit = QTextEdit(self)

        font = QFont()
        font.setPointSize(14)  # Set the size you want here
        self.text_edit.setFont(font)

        # Connect the button's click signal to the function that will handle it
        self.button.clicked.connect(self.on_generate_summary_clicked)

        # Set up the layout
        layout = QVBoxLayout()
        layout.addWidget(self.button)

        # Create a spacer item with the height approximately of a button and add it to the layout
        spacer = QSpacerItem(
            20,
            self.button.sizeHint().height(),
            QSizePolicy.Policy.Minimum,
            QSizePolicy.Policy.Fixed,
        )
        layout.addSpacerItem(spacer)

        layout.addWidget(self.text_edit)

        self.central_widget.setLayout(layout)

    def on_generate_summary_clicked(self):
        # Get the email summary
        summary = get_summary()
        # Display the summary in the text box
        self.text_edit.setPlainText(summary)


def main():
    load_dotenv()

    app = QApplication(sys.argv)
    main_window = EmailSummaryApp()
    main_window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
