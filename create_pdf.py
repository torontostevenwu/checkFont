from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch

def create_sample_pdf(file_path="sample.pdf"):
    c = canvas.Canvas(file_path, pagesize=letter)
    width, height = letter  # Keep for potential positioning calculations

    # Times New Roman
    c.setFont("Times-Roman", 12)
    c.drawString(1 * inch, height - 1 * inch, "This is Times New Roman text.")

    # Arial
    c.setFont("Helvetica", 12) # Helvetica is a common substitute for Arial in PDF
    c.drawString(1 * inch, height - 1.5 * inch, "This is Arial text.")

    c.save()
    print(f"{file_path} created successfully.")

if __name__ == "__main__":
    create_sample_pdf()
