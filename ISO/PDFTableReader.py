import tabula
tables = tabula.read_pdf("CertC20590U.pdf", pages = 1)
print(tables)