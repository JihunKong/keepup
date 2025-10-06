
import os
import docx

def convert_docx_to_txt_in_current_directory():
    """
    Finds all .docx files in the script's directory, extracts their text,
    and saves it to corresponding .txt files.
    """
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    for filename in os.listdir(current_dir):
        if filename.endswith('.docx'):
            docx_path = os.path.join(current_dir, filename)
            txt_path = os.path.join(current_dir, filename.replace('.docx', '.txt'))
            
            try:
                document = docx.Document(docx_path)
                full_text = [para.text for para in document.paragraphs]
                
                with open(txt_path, 'w', encoding='utf-8') as f:
                    f.write('\n'.join(full_text))
                
                print(f"Successfully converted '{filename}' to .txt")
            
            except Exception as e:
                print(f"Error processing {filename}: {e}")

if __name__ == "__main__":
    convert_docx_to_txt_in_current_directory()
