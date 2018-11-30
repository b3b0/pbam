import sys
import os
import pdfkit
f = open('nul', 'w')
sys.stdout = f
if __name__ == "__main__":
    file = str(sys.argv[1])
    pdf = str(sys.argv[2])
    exe = str(sys.argv[3])
    config = pdfkit.configuration(wkhtmltopdf= exe)
    pdfkit.from_file(file, pdf, configuration = config)
