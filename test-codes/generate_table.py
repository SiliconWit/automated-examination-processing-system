import tabula
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

def generate_pdf(data):
    # Create a new PDF document
    pdf_filename = "student_scores.pdf"
    doc = SimpleDocTemplate(pdf_filename, pagesize=letter)

    # Create a list of data rows for the table
    table_data = [list(data.columns)] + data.values.tolist()

    # Create a Table object with the data
    table = Table(table_data)

    # Apply table styles to highlight passed and failed students
    table_style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (2, 1), (2, -1), colors.green),  # Highlight 'Status' column for Passed
        ('BACKGROUND', (2, 2), (2, -1), colors.red)     # Highlight 'Status' column for Failed
    ])

    table.setStyle(table_style)

    # Build the PDF document
    doc.build([table])

if __name__ == "__main__":
    # Sample data: Student scores and pass/fail status
    data = pd.DataFrame({
        "Name": ["Alice", "Bob", "Charlie", "David"],
        "Score": [55, 30, 70, 38],
        "Status": ["Passed", "Failed", "Passed", "Failed"]
    })

    # Filter data for passed and failed students
    passed_students = data[data["Status"] == "Passed"]
    failed_students = data[data["Status"] == "Failed"]

    # Generate the PDF
    generate_pdf(data)

    print("PDF document 'student_scores.pdf' generated.")

