import pandas as pd
from docx import Document
import tkinter as tk
from tkinter import filedialog, messagebox
import difflib
import os
import time

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet

    REPORTLAB_SUPPORT = True
except ImportError:
    REPORTLAB_SUPPORT = False


def parse_text_to_dataframe(filepath):
    parsed_data = []
    known_topics = {}
    next_item_id = 0
    current_topic = None
    current_item_id = None

    with open(filepath, 'r', encoding='utf-8') as file:
        for line in file:
            line = line.strip()
            if not line:
                continue

            if line.startswith('-'):
                if current_topic is not None:
                    note_text = line.lstrip('-').strip()
                    parsed_data.append({
                        'Item': current_item_id,
                        'TopicName': current_topic,
                        'Notes': note_text
                    })
            else:
                raw_topic = line
                matched_id = None

                for known_id, known_name in known_topics.items():
                    # Threshold increased to 0.9 to prevent Topic1/Topic2 collisions
                    similarity = difflib.SequenceMatcher(None, known_name.lower(), raw_topic.lower()).ratio()
                    if similarity > 0.9:
                        matched_id = known_id
                        break

                if matched_id is not None:
                    current_item_id = matched_id
                    current_topic = known_topics[matched_id]
                else:
                    current_item_id = next_item_id
                    current_topic = raw_topic
                    known_topics[current_item_id] = current_topic
                    next_item_id += 1

    df = pd.DataFrame(parsed_data)
    for i in range(1, 10):
        df[f'Num{i}'] = ''

    return df[['Num1', 'Num2', 'Num3', 'Num4', 'Item', 'Notes', 'Num5', 'Num6', 'Num7', 'Num8', 'Num9', 'TopicName']]


class DocumentConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Topic Converter & Autocorrect")
        self.root.geometry("650x380")

        self.label = tk.Label(root, text="Select your .txt notes file:", pady=10)
        self.label.pack()

        self.btn_browse = tk.Button(root, text="Browse & Process File", command=self.process_file, width=25, pady=5)
        self.btn_browse.pack()

        self.status = tk.Text(root, height=14, width=75, state=tk.DISABLED)
        self.status.pack(pady=10)

    def log_status(self, message):
        self.status.config(state=tk.NORMAL)
        self.status.insert(tk.END, message + "\n")
        self.status.see(tk.END)
        self.status.config(state=tk.DISABLED)
        self.root.update()

    def process_file(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
        )
        if not filepath:
            return

        self.log_status(f"Loading: {os.path.basename(filepath)}\n")
        epoch = int(time.time())

        try:
            df = parse_text_to_dataframe(filepath)

            # Export to CSV
            csv_path = os.path.abspath(f'backend_output_{epoch}.csv')
            df.to_csv(csv_path, index=False)
            self.log_status(f"Saved CSV:\n  {csv_path}\n")

            # Export to XLSX
            xlsx_path = os.path.abspath(f'backend_output_{epoch}.xlsx')
            df.to_excel(xlsx_path, index=False)
            self.log_status(f"Saved XLSX:\n  {xlsx_path}\n")

            grouped_data = df.groupby('TopicName')['Notes'].apply(list).to_dict()

            # Export to DOCX
            docx_path = os.path.abspath(f'grouped_topics_output_{epoch}.docx')
            doc = Document()
            doc.add_heading('Grouped Topical Output', 0)

            for topic_title, notes_list in grouped_data.items():
                doc.add_heading(topic_title, level=1)
                for note in notes_list:
                    doc.add_paragraph(f"- {str(note).strip()}")
                doc.add_paragraph()

            doc.save(docx_path)
            self.log_status(f"Saved DOCX:\n  {docx_path}\n")

            # Export to PDF without Word Automation (No Dialogs)
            if REPORTLAB_SUPPORT:
                pdf_path = os.path.abspath(f'grouped_topics_output_{epoch}.pdf')
                pdf_doc = SimpleDocTemplate(pdf_path, pagesize=letter)
                styles = getSampleStyleSheet()
                elements = []

                elements.append(Paragraph('Grouped Topical Output', styles['Title']))
                for topic_title, notes_list in grouped_data.items():
                    elements.append(Paragraph(topic_title, styles['Heading1']))
                    for note in notes_list:
                        elements.append(Paragraph(f"• {str(note).strip()}", styles['Normal']))
                    elements.append(Spacer(1, 12))

                pdf_doc.build(elements)
                self.log_status(f"Saved PDF (Silently):\n  {pdf_path}\n")
            else:
                self.log_status("ReportLab missing. PDF skipped.\n")

            messagebox.showinfo("Success", "Processing complete! Check the console log for absolute file paths.")

        except Exception as e:
            self.log_status(f"Error: {str(e)}")
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = DocumentConverterGUI(root)
    root.mainloop()
