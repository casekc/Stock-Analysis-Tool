nia_dir="C:\Users\cummi\Desktop\webscrap\bin\net_income_analyses"
hpa_dir="C:\Users\cummi\Desktop\webscrap\bin\historical_price_analyses"
nia_output_dir="C:\Users\cummi\Desktop\webscrap\bin\converted_files\net_income_analyses"
hpa_output_dir="C:\Users\cummi\Desktop\webscrap\bin\converted_files\historical_price_analyses"

# Convert files using LibreOffice
"C:\Program Files\LibreOffice\program\soffice.exe" --headless --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$nia_output_dir" "$nia_dir"/*.xls*

"C:\Program Files\LibreOffice\program\soffice.exe" --headless --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$hpa_output_dir" "$hpa_dir"/*.xls*

