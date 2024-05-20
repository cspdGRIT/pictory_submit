import os
import zipfile
import re
import shutil
import boto3
import logging
from pptx import Presentation
from pptx.util import Inches, Pt

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def extract_media_from_pptx(pptx_file, output_dir, bucket_name, aws_access_key_id, aws_secret_access_key):
    try:
        # Create a temporary ZIP file name
        zip_file = os.path.join(output_dir, os.path.splitext(os.path.basename(pptx_file))[0] + ".zip")

        # Copy the PPTX file to the output directory
        output_pptx_file = os.path.join(output_dir, os.path.basename(pptx_file))
        shutil.copy(pptx_file, output_pptx_file)

        # Open the ZIP file
        with zipfile.ZipFile(zip_file, 'w') as zip_ref:
            # Add the PPTX file to the ZIP file
            zip_ref.write(output_pptx_file, os.path.basename(output_pptx_file))

        # Extract media files from the "ppt/media" folder
        media_paths = {}
        media_positions = {}
        s3 = boto3.client('s3', aws_access_key_id=aws_access_key_id, aws_secret_access_key=aws_secret_access_key)
        with zipfile.ZipFile(zip_file, 'r') as zip_ref:
            for zip_info in zip_ref.infolist():
                if zip_info.filename.startswith("ppt/media/"):
                    media_path = os.path.join(output_dir, os.path.basename(zip_info.filename))
                    with open(media_path, "wb") as f:
                        f.write(zip_ref.read(zip_info.filename))
                    slide_num = int(re.search(r'\d+', zip_info.filename).group())
                    if slide_num not in media_paths:
                        media_paths[slide_num] = []
                        media_positions[slide_num] = []
                    media_paths[slide_num].append(media_path)

                    # Load the PowerPoint presentation and get the media position
                    presentation = Presentation(output_pptx_file)
                    slide = presentation.slides[slide_num - 1]
                    for shape in slide.shapes:
                        if hasattr(shape, "image"):
                            if shape.image.blob == zip_ref.read(zip_info.filename):
                                media_positions[slide_num].append({
                                    "left": shape.left,
                                    "top": shape.top,
                                    "width": shape.width,
                                    "height": shape.height
                                })
                                # Upload media file to S3
                                filename = os.path.basename(media_path)
                                s3_key = f"slide_{slide_num}/{filename}"
                                try:
                                    s3.upload_file(media_path, bucket_name, s3_key)
                                except Exception as e:
                                    logger.error(f"Failed to upload media {media_path} to S3: {e}")
                                break

        # Remove the temporary PPTX file
        os.remove(output_pptx_file)

        return media_paths, media_positions

    except FileNotFoundError:
        print(f"Error: File '{pptx_file}' not found.")
        return {}, {}
    except Exception as e:
        print(f"An error occurred during PPTX to media conversion: {e}")
        return {}, {}

def extract_slide_text(pptx_file, output_dir):
    try:
        # Copy the PPTX file to the output directory
        output_pptx_file = os.path.join(output_dir, os.path.basename(pptx_file))
        shutil.copy(pptx_file, output_pptx_file)

        # Load PowerPoint presentation
        presentation = Presentation(output_pptx_file)
        text_content = {}

        # Extract text content from each slide
        for i, slide in enumerate(presentation.slides):
            slide_text = []
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    slide_text.append(shape.text.strip())
            text_content[i+1] = "\n".join(slide_text)

        # Remove the temporary PPTX file
        os.remove(output_pptx_file)

        return text_content

    except FileNotFoundError:
        print(f"Error: File '{pptx_file}' not found.")
        return {}
    except Exception as e:
        print(f"An error occurred during text extraction: {e}")
        return {}

def extract_slide_text_and_media(pptx_file, output_dir, bucket_name, aws_access_key_id, aws_secret_access_key):
    media_paths, media_positions = extract_media_from_pptx(pptx_file, output_dir, bucket_name, aws_access_key_id, aws_secret_access_key)
    text_content = extract_slide_text(pptx_file, output_dir)

    # Pair slide text with media paths and positions
    paired_data = {}
    for slide_num, text in text_content.items():
        if slide_num in media_paths:
            paired_data[slide_num] = {
                "text": text,
                "media_info": [{"path": path, "position": position} for path, position in zip(media_paths[slide_num], media_positions[slide_num])]
            }
        else:
            paired_data[slide_num] = {"text": text, "media_info": []}

    return paired_data

# Example usage
pptx_file = "jpictory.pptx"
bucket_name = "ppt2video"
aws_access_key_id = ""
aws_secret_access_key = ""
aws_session_token = ""

output_dir = "output_media"
paired_data = extract_slide_text_and_media(pptx_file, output_dir, bucket_name, aws_access_key_id, aws_secret_access_key)

# Print paired data (slide number, slide text, media paths)
for slide_num, data in paired_data.items():
    print(f"Slide {slide_num}:")
    print(f"Text:\n{data['text']}")
    print("Media Info:")
    for media_info in data["media_info"]:
        print(f"- Path: {media_info['path']}")
        print(f"  Position: Left={media_info['position']['left']}, Top={media_info['position']['top']}, Width={media_info['position']['width']}, Height={media_info['position']['height']}")
    print()
