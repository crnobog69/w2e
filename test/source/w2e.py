try:
    import docx
    import pandas as pd
    from pathlib import Path
    from openpyxl.styles import Alignment
except ImportError:
    print("Required packages are not installed. Please run:")
    print("pip install -r requirements.txt")
    exit(1)

def get_column_width(col_series):
    # Handle different data types for column width calculation
    max_length = 0
    for value in col_series:
        try:
            # Convert any value to string and get its length
            str_value = str(value) if value is not None else ''
            max_length = max(max_length, len(str_value))
        except:
            continue
    return max_length

def word_to_excel(word_file, excel_file):
    # Load the Word document
    doc = docx.Document(word_file)
    
    # Initialize data storage with single column
    rows = []
    
    # Process document sequentially
    for element in doc.element.body:
        if element.tag.endswith('tbl'):  # Table
            table = doc.tables[len([e for e in doc.element.body[:doc.element.body.index(element)] 
                                  if e.tag.endswith('tbl')])]
            
            # Add blank line before table
            rows.append([''])
            
            for row in table.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                if any(row_data):  # Only add non-empty rows
                    rows.append(row_data)
            
            # Add blank line after table
            rows.append([''])
                
        elif element.tag.endswith('p'):  # Paragraph
            text = element.text.strip()
            if text:
                rows.append([text])
    
    # Create DataFrame without headers
    final_df = pd.DataFrame(rows)
    
    # Save to Excel with formatting
    try:
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            # Write DataFrame without headers
            final_df.to_excel(writer, 
                            sheet_name='Document Content',
                            index=False,
                            header=False)
            worksheet = writer.sheets['Document Content']
            
            # Adjust column widths without considering headers
            for idx in range(final_df.shape[1]):
                max_length = max(
                    final_df[idx].astype(str).apply(len).max(),
                    0  # No header to consider
                )
                col_letter = chr(65 + idx) if idx < 26 else chr(64 + (idx//26)) + chr(65 + (idx%26))
                worksheet.column_dimensions[col_letter].width = min(max_length + 2, 100)
            
            # Adjust row heights and wrap text using proper alignment
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical='top')
                worksheet.row_dimensions[row[0].row].height = None  # Auto-height
    
    except Exception as e:
        print(f"Warning: Error while saving file: {str(e)}")
        return False
    
    return True

if __name__ == "__main__":
    # Get current directory
    current_dir = Path('.')
    # Create output directory if it doesn't exist
    output_dir = current_dir / 'ex'
    output_dir.mkdir(exist_ok=True)
    
    # Find all .docx files
    docx_files = list(current_dir.glob('*.docx'))
    
    if not docx_files:
        print("No .docx files found in the current directory!")
        exit(1)
        
    print(f"Found {len(docx_files)} Word files to process...")
    
    for word_path in docx_files:
        try:
            # Create output path with same name but .xlsx extension in 'ex' folder
            excel_path = output_dir / f"{word_path.stem}.xlsx"
            
            success = word_to_excel(word_path, excel_path)
            if success:
                print(f"Converted: {word_path.name} → {excel_path.name}")
        except Exception as e:
            print(f"Error converting {word_path.name}: {str(e)}")
    
    print("\nProcessing complete! Check the 'ex' folder for output files.")
