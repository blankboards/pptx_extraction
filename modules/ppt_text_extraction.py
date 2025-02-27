from pptx import Presentation 
from modules.config import OUTPUT_DIR_2
import os

def extract_text_from_ppt(file_path):
    os.makedirs(OUTPUT_DIR_2, exist_ok=True)
    text_output = []
    presentation = Presentation(file_path) 
    
    for slide_number, slide in enumerate(presentation.slides, start=1):
        slide_folder = os.path.join(OUTPUT_DIR_2, f"slide_{slide_number}", "image")
        os.makedirs(slide_folder, exist_ok=True)
        slide_text = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        slide_text.append(run.text)
        text_output.append(f"\n\n@@@Slide_{slide_number}@@@\n" + "\n".join(slide_text))
        with open(os.path.join(OUTPUT_DIR_2, f"slide_{slide_number}", f"slide_{slide_number}_texts.txt"), "w", encoding="utf-8") as f:
            text = "".join(slide_text)
            f.write("@@@content@@@\n"+text+"\n\n")
    return text_output

def extract_metadata(file_path):
    try:
        presentation = Presentation(file_path)
        props = presentation.core_properties

        metadata = {
            "Title": props.title or "N/A",
            "Author": props.author or "N/A",
            "Subject": props.subject or "N/A",
            "Keywords": props.keywords or "N/A",
            "Comments": props.comments or "N/A",
            "Last Modified By": props.last_modified_by or "N/A",
            "Created": props.created.strftime('%Y-%m-%d %H:%M:%S') if props.created else "N/A",
            "Modified": props.modified.strftime('%Y-%m-%d %H:%M:%S') if props.modified else "N/A",
            "Category": props.category or "N/A",
            "Content Status": props.content_status or "N/A",
            "Identifier": props.identifier or "N/A",
            "Language": props.language or "N/A",
            "Revision": props.revision or "N/A"
        }

        return metadata

    except Exception as e:
        return {"Error": str(e)}