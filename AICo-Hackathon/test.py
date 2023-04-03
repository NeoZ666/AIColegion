from fpdf import FPDF

pdf = FPDF()
pdf.add_page()
pdf.set_font("times", style='b', size=24)
# pdf.set_font("helvetica", size=24)
pdf.cell(w=40, h=10, txt="CERTIFICATE", border=0, align="C", link="https://github.com/PyFPDF/fpdf2")
pdf.output("hyperlink.pdf")