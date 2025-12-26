# proposal_gui.py  ── first GUI pass
# ----------------------------------------------------------
# pip install PyQt5 xlwings
# usage:  py proposal_gui.py
# ----------------------------------------------------------

from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QTextEdit,
    QPushButton, QGridLayout, QMessageBox
)
from PyQt5.QtCore import Qt
import sys
from extract_proposal import extract_proposal_data   # same folder


class ProposalEditor(QWidget):
    def __init__(self, data: dict):
        super().__init__()
        self.setWindowTitle("Proposal Review")

        # store original / edited data
        self.data = data.copy()

        # ── build the form ───────────────────────────────────
        grid = QGridLayout(self)

        # single-line fields
        self.job_edit     = QLineEdit(data["JOB"])
        self.contact_edit = QLineEdit(data["CONTACT"])
        self.phone_edit   = QLineEdit(data["PHONE"])
        self.email_edit   = QLineEdit(data["EMAIL"])
        self.price_edit   = QLineEdit(data["TOTAL_PRICE"])

        # make them stretch horizontally
        for w in (self.job_edit, self.contact_edit,
                  self.phone_edit, self.email_edit, self.price_edit):
            w.setMinimumWidth(300)

        # multi-line
        self.scope_edit = QTextEdit()
        self.scope_edit.setPlainText(data["SCOPE_OF_WORK"])   # keep hard returns
        self.scope_edit.setMinimumSize(600, 300)

        # layout
        grid.addWidget(QLabel("Job"),        0, 0); grid.addWidget(self.job_edit,     0, 1)
        grid.addWidget(QLabel("Contact"),    1, 0); grid.addWidget(self.contact_edit, 1, 1)
        grid.addWidget(QLabel("Phone"),      2, 0); grid.addWidget(self.phone_edit,   2, 1)
        grid.addWidget(QLabel("Email"),      3, 0); grid.addWidget(self.email_edit,   3, 1)
        grid.addWidget(QLabel("Total Price"),4, 0); grid.addWidget(self.price_edit,   4, 1)

        grid.addWidget(QLabel("Scope of Work"), 5, 0, Qt.AlignTop)
        grid.addWidget(self.scope_edit,         5, 1)

        # send button bottom-right
        send_btn = QPushButton("Send")
        send_btn.clicked.connect(self.on_send)
        grid.addWidget(send_btn, 6, 1, Qt.AlignRight)

    # slot: collect edits, store, and (for now) just print
    def on_send(self):
        self.data["JOB"]           = self.job_edit.text()
        self.data["CONTACT"]       = self.contact_edit.text()
        self.data["PHONE"]         = self.phone_edit.text()
        self.data["EMAIL"]         = self.email_edit.text()
        self.data["TOTAL_PRICE"]   = self.price_edit.text()
        self.data["SCOPE_OF_WORK"] = self.scope_edit.toPlainText()

        QMessageBox.information(self, "Saved & sending", "Values captured – starting upload…")
        self.close()        # run_gui_editor will now return editor.data


def main():
    # pull from Excel; if anything is wrong show a dialog
    try:
        proposal = extract_proposal_data()
    except Exception as exc:
        app = QApplication(sys.argv)
        QMessageBox.critical(None, "Extraction error", str(exc))
        sys.exit()

    app = QApplication(sys.argv)
    editor = ProposalEditor(proposal)
    editor.show()
    sys.exit(app.exec_())


def run_gui_editor(data):
    app = QApplication(sys.argv)
    editor = ProposalEditor(data)
    editor.show()
    app.exec_()
    return editor.data        # returns the (possibly edited) dict



if __name__ == "__main__":
    main()
