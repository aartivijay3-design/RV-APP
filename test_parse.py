
import sys, os
sys.path.insert(0, r"C:\Users\Vijay.Aarti\Documents\Claude\Projects\bawa-rv-app")
os.chdir(r"C:\Users\Vijay.Aarti\Documents\Claude\Projects\bawa-rv-app")
from app import parse_document, read_word

# Test Erdt Reiseverlauf Word doc (NO truncation)
data = open(r"C:\Users\Vijay.Aarti\Downloads\Bestätiger RV Erdt Japan.docx", "rb").read()
text = read_word(data, max_chars=None)
print("Full doc chars:", len(text))
result = parse_document(text)
print("Clients:", result["client_names"][:4])
print("Destination:", result["destination_de"])
print("Dates:", result["travel_start"], "->", result["travel_end"])
print("Hotels found:", len(result["hotels"]))
for h in result["hotels"]:
    print(" ", h["check_in"], "-", h["check_out"], str(h["nights"])+"n:", h["hotel_name"])

print()
# Test Groh Rechnung PDF
import pypdf, io as io2
pdf_data = open(r"C:\Users\Vijay.Aarti\Downloads\Groh Indien Rechnung 29.07.25.pdf", "rb").read()
reader = pypdf.PdfReader(io2.BytesIO(pdf_data))
pdf_text = "\n".join(p.extract_text() or "" for p in reader.pages)
result2 = parse_document(pdf_text)
print("Rechnung - Hotels:", len(result2["hotels"]))
for h in result2["hotels"]:
    print(" ", h["hotel_name"], h["check_in"], "->", h["check_out"])
