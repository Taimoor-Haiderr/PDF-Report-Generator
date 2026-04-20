import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pdfencrypt import StandardEncryption
import csv
import json
import os
from datetime import datetime
import matplotlib.pyplot as plt

class PDFReportGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Professional PDF Report Generator")
        self.root.geometry("1180x820")
        self.root.configure(bg="#f8fafc")

        self.data_list = []
        self.history = []
        self.logo_path = None

        self.create_empty_test_files()
        self.create_professional_gui()

    def create_empty_test_files(self):
        if not os.path.exists("test_students.csv"):
            with open("test_students.csv", "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["name", "id", "email", "department", "details"])
        if not os.path.exists("test_companies.json"):
            with open("test_companies.json", "w", encoding="utf-8") as f:
                json.dump([], f, indent=2)

    def create_professional_gui(self):
        header = tk.Frame(self.root, bg="#1e3a8a", height=100)
        header.pack(fill=tk.X)
        tk.Label(header, text="Professional PDF Report Generator",
                 font=("Helvetica", 26, "bold"), bg="#1e3a8a", fg="white").pack(pady=25)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)

        self.create_input_tab()
        self.create_load_tab()
        self.create_preview_tab()
        self.create_generate_tab()
        self.create_history_tab()

    def create_input_tab(self):
        tab = ttk.Frame(self.notebook, padding=25)
        self.notebook.add(tab, text=" Add Data ")

        frame = ttk.LabelFrame(tab, text=" Enter New Record ", padding=20)
        frame.pack(fill=tk.X, pady=10)

        self.report_type = tk.StringVar(value="Student")
        (ttk.Radiobutton(frame, text="Student Report", variable=self.report_type, value="Student")
         .grid(row=0, column=0, padx=10))
        (ttk.Radiobutton(frame, text="Company Report", variable=self.report_type, value="Company")
         .grid(row=0, column=1, padx=10))

        ttk.Label(frame, text="Name:").grid(row=1, column=0, sticky="w", pady=8)
        self.name_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.name_var, width=55).grid(row=1, column=1, pady=8, padx=15)

        ttk.Label(frame, text="ID:").grid(row=2, column=0, sticky="w", pady=8)
        self.id_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.id_var, width=55).grid(row=2, column=1, pady=8, padx=15)

        ttk.Label(frame, text="Email:").grid(row=3, column=0, sticky="w", pady=8)
        self.email_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.email_var, width=55).grid(row=3, column=1, pady=8, padx=15)

        ttk.Label(frame, text="Department/Role:").grid(row=4, column=0, sticky="w", pady=8)
        self.dept_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.dept_var, width=55).grid(row=4, column=1, pady=8, padx=15)

        ttk.Label(frame, text="Details:").grid(row=5, column=0, sticky="nw", pady=8)
        self.details_text = tk.Text(frame, height=7, width=55)
        self.details_text.grid(row=5, column=1, pady=8, padx=15)

        ttk.Button(tab, text="Add Record", command=self.add_record).pack(pady=25)

    def add_record(self):
        if not self.name_var.get() or not self.id_var.get() or not self.email_var.get():
            messagebox.showwarning("Missing", "Name, ID and Email are required!")
            return

        record = {
            "type": self.report_type.get(),
            "name": self.name_var.get().strip(),
            "id": self.id_var.get().strip(),
            "email": self.email_var.get().strip(),
            "department": self.dept_var.get().strip(),
            "details": self.details_text.get("1.0", tk.END).strip(),
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

        self.data_list.append(record)
        messagebox.showinfo("Success", f"Record for '{record['name']}' added!")

        self.name_var.set("")
        self.id_var.set("")
        self.email_var.set("")
        self.dept_var.set("")
        self.details_text.delete("1.0", tk.END)

    def create_load_tab(self):
        tab = ttk.Frame(self.notebook, padding=30)
        self.notebook.add(tab, text=" Load Data ")

        ttk.Button(tab, text="Load test_students.csv", command=self.load_csv).pack(pady=12)
        ttk.Button(tab, text=" Load test_companies.json", command=self.load_json).pack(pady=12)

    def load_csv(self):
        self.load_file("test_students.csv", "Student")

    def load_json(self):
        self.load_file("test_companies.json", "Company")

    def load_file(self, filename, default_type):
        if not os.path.exists(filename):
            messagebox.showerror("Missing", f"{filename} not found!")
            return
        try:
            if filename.endswith(".csv"):
                with open(filename, newline="", encoding="utf-8") as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        row["type"] = default_type
                        self.data_list.append(row)
            else:
                with open(filename, encoding="utf-8") as f:
                    data = json.load(f)
                    for item in data:
                        item["type"] = default_type
                        self.data_list.append(item)
            messagebox.showinfo("Success", f"Loaded {len(self.data_list)} records!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{str(e)}")

    def create_preview_tab(self):
        tab = ttk.Frame(self.notebook, padding=25)
        self.notebook.add(tab, text=" Preview Records ")

        ttk.Button(tab, text="🔍 Show All Records", command=self.show_preview).pack(pady=15)

        self.preview_tree = ttk.Treeview(tab, columns=("name", "id", "email", "department", "type"),
                                         show="headings", height=20)
        for col in ("name", "id", "email", "department", "type"):
            self.preview_tree.heading(col, text=col.capitalize())
        self.preview_tree.column("email", width=280)
        self.preview_tree.pack(fill=tk.BOTH, expand=True, pady=15)

    def show_preview(self):
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        if not self.data_list:
            messagebox.showinfo("Empty", "No records yet!")
            return
        for r in self.data_list:
            self.preview_tree.insert("", "end", values=(
                r.get("name", ""), r.get("id", ""), r.get("email", ""),
                r.get("department", ""), r.get("type", "")
            ))

    def create_generate_tab(self):
        tab = ttk.Frame(self.notebook, padding=30)
        self.notebook.add(tab, text=" Generate PDF ")

        ttk.Label(tab, text="Choose Company Logo (Optional)", font=("Helvetica", 11)).pack(anchor="w", pady=5)
        ttk.Button(tab, text="📸 Select Logo", command=self.choose_logo).pack(pady=8)

        (ttk.Button(tab, text="Generate Student Report", command=lambda: self.generate_pdf("Student"))
         .pack(pady=15, ipadx=30))
        (ttk.Button(tab, text="Generate Company Report", command=lambda: self.generate_pdf("Company"))
         .pack(pady=15, ipadx=30))

    def choose_logo(self):
        file = filedialog.askopenfilename(filetypes=[("Image Files", "*.png *.jpg *.jpeg")])
        if file:
            self.logo_path = file
            messagebox.showinfo("Logo Selected", f"Logo loaded: {os.path.basename(file)}")

    def create_chart(self):
        if not self.data_list:
            return None
        names = [r.get("name", "Unknown")[:10] for r in self.data_list[:8]]
        values = [len(r.get("details", "")) for r in self.data_list[:8]]

        plt.figure(figsize=(10, 5))
        bars = plt.bar(names, values, color="#1e40af")
        plt.title("Performance Summary (Detail Length)", fontsize=14, pad=20)
        plt.xlabel("Name", fontsize=12)
        plt.ylabel("Length of Details", fontsize=12)
        plt.xticks(rotation=45, ha='right')
        plt.grid(axis='y', alpha=0.3)

        for bar in bars:
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height + 2,
                    f'{int(height)}', ha='center', va='bottom', fontsize=10)

        chart_path = "temp_chart.png"
        plt.savefig(chart_path, bbox_inches='tight', dpi=200)
        plt.close()
        return chart_path


    def generate_pdf(self, report_type: str):
        if not self.data_list:
            messagebox.showwarning("No Data", "Please add or load data first!")
            return

        try:
            os.makedirs("Reports", exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"Reports/{report_type}_Report_{timestamp}.pdf"

            from reportlab.lib.pdfencrypt import StandardEncryption

            encrypt = StandardEncryption(
                userPassword="1234",  # open password
                ownerPassword="admin123",  # admin password
                canPrint=0,  # disable printing
                canModify=0,  #  disable editing
                canCopy=0,  #  disable copying
                canAnnotate=0  #  disable comments
            )

            doc = SimpleDocTemplate(
                filename,
                pagesize=letter,
                encrypt=encrypt,
                rightMargin=40,
                leftMargin=40,
                topMargin=50,
                bottomMargin=40
            )
            styles = getSampleStyleSheet()

            # ---------- PROFESSIONAL STYLES ----------
            title_style = ParagraphStyle(
                'TitleStyle',
                parent=styles['Title'],
                fontSize=20,
                textColor=colors.darkblue,
                alignment=TA_CENTER,
                spaceAfter=25
            )

            normal_style = ParagraphStyle(
                'NormalStyle',
                parent=styles['Normal'],
                fontSize=10,
                leading=14
            )

            elements = []

            # ---------- LOGO ----------
            if self.logo_path and os.path.exists(self.logo_path):
                try:
                    img = Image(self.logo_path, width=120, height=60)
                    img.hAlign = 'CENTER'
                    elements.append(img)
                    elements.append(Spacer(1, 15))
                except:
                    pass

            # ---------- TITLE ----------
            elements.append(Paragraph(f"<b>{report_type.upper()} PERFORMANCE REPORT</b>", title_style))

            elements.append(Paragraph(
                f"Generated on: {datetime.now().strftime('%d %B %Y at %H:%M')}",
                styles['Normal']
            ))
            elements.append(Spacer(1, 25))

            # ---------- TABLE ----------
            table_data = [[
                "Name", "ID", "Email", "Department", "Details"
            ]]

            for r in self.data_list:
                if r.get("type") == report_type:
                    table_data.append([
                        Paragraph(r.get("name", "-"), normal_style),
                        Paragraph(r.get("id", "-"), normal_style),
                        Paragraph(r.get("email", "-"), normal_style),
                        Paragraph(r.get("department", "-"), normal_style),
                        Paragraph(r.get("details", "-"), normal_style),
                    ])

            #  PERFECT WIDTH (NO OVERFLOW)
            table = Table(
                table_data,
                colWidths=[85, 60, 140, 100, 135]
            )

            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),

                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 11),

                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),

                ('GRID', (0, 0), (-1, -1), 0.7, colors.grey),

                ('ROWBACKGROUNDS', (0, 1), (-1, -1),
                 [colors.whitesmoke, colors.white]),

                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),

                ('LEFTPADDING', (0, 0), (-1, -1), 6),
                ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                ('TOPPADDING', (0, 0), (-1, -1), 6),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ]))

            elements.append(table)

            # ---------- CHART ----------
            chart_path = self.create_chart()
            if chart_path and os.path.exists(chart_path):
                elements.append(Spacer(1, 30))
                elements.append(Paragraph("<b>Performance Summary</b>", styles['Heading2']))
                elements.append(Spacer(1, 10))

                chart_img = Image(chart_path, width=360, height=200)
                chart_img.hAlign = 'CENTER'
                elements.append(chart_img)

            # ---------- BUILD ----------
            doc.build(elements)

            # ---------- HISTORY ----------
            self.history.append({
                "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "filename": os.path.basename(filename),
                "type": report_type
            })
            self.refresh_history()

            messagebox.showinfo("Success", f"Report Generated:\n{filename}")

        except Exception as e:
            messagebox.showerror("Error", str(e))

            messagebox.showinfo("Success", f" Professional PDF Generated!\n\nSaved as:"
                                           f"\n{filename}\nPassword: 1234")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate PDF:\n{str(e)}")

    def create_history_tab(self):
        tab = ttk.Frame(self.notebook, padding=30)
        self.notebook.add(tab, text=" History ")

        ttk.Label(tab, text="Generated Reports History", font=("Helvetica", 14, "bold")).pack(anchor="w",
                                                                                              pady=10)

        self.history_tree = ttk.Treeview(tab, columns=("date", "filename", "type"), show="headings",
                                         height=20)
        self.history_tree.heading("date", text="Date & Time")
        self.history_tree.heading("filename", text="Report File")
        self.history_tree.heading("type", text="Type")
        self.history_tree.column("date", width=180)
        self.history_tree.column("filename", width=450)
        self.history_tree.pack(fill=tk.BOTH, expand=True, pady=15)

    def refresh_history(self):
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        for entry in reversed(self.history[-30:]):
            self.history_tree.insert("", "end", values=(
                entry["date"],
                entry["filename"],
                entry["type"]
            ))

    def generate_all_reports(self):
        if not self.data_list:
            messagebox.showwarning("No Data", "No records available!")
            return
        self.generate_pdf("Student")
        self.generate_pdf("Company")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFReportGenerator(root)
    root.mainloop()