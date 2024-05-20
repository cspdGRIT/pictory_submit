# PowerPoint to Video Conversion Application

This application extracts text and media from PowerPoint files, uploads the media to an S3 bucket, and generates paired data for a chat application.

## System Design

The system consists of the following main components:

1. **PowerPoint Text and Media Extraction**
2. **S3 Media Upload**
3. **Paired Data Generation for Chat Application**

### PowerPoint Text and Media Extraction

The PowerPoint text and media extraction component is responsible for extracting text and media from PowerPoint files. It uses the following functions:

1. `extract_media_from_pptx`
2. `extract_slide_text`
3. `extract_slide_text_and_media`

#### `extract_media_from_pptx`

This function extracts media files from a PowerPoint file, uploads them to an S3 bucket, and returns the media paths and positions. It performs the following steps:

1. Creates a temporary ZIP file from the PowerPoint file
2. Copies the PowerPoint file to the output directory
3. Opens the ZIP file and extracts media files from the "ppt/media" folder
4. Loads the PowerPoint presentation and gets the media position for each extracted file
5. Uploads the media files to the S3 bucket with the slide number as the folder name
6. Removes the temporary PowerPoint file
7. Returns the media paths and positions

#### `extract_slide_text`

This function extracts text content from each slide in a PowerPoint file. It performs the following steps:

1. Copies the PowerPoint file to the output directory
2. Loads the PowerPoint presentation
3. Extracts text content from each slide
4. Removes the temporary PowerPoint file
5. Returns the text content

#### `extract_slide_text_and_media`

This function combines the functionality of `extract_media_from_pptx` and `extract_slide_text` to extract both text and media from a PowerPoint file. It performs the following steps:

1. Calls `extract_media_from_pptx` to extract media files and upload them to S3
2. Calls `extract_slide_text` to extract text content from the PowerPoint file
3. Pairs the slide text with the corresponding media paths and positions
4. Returns the paired data

### S3 Media Upload

The S3 media upload component is integrated into the `extract_media_from_pptx` function, which uploads each media file to the S3 bucket with the slide number as the folder name.

### Paired Data Generation for Chat Application

The paired data generation component creates a dictionary that pairs slide text with media paths and positions. This data is used by the chat application to display the appropriate content for each slide.

## Configuration

The application requires the following configuration:

- Python 3.x
- Required Python packages (listed in requirements.txt)
- AWS credentials with access to an S3 bucket

## Installation

1. Clone the repository or download the source code.
2. Create a virtual environment (optional but recommended).
3. Install the required packages using pip:

   ```bash
   pip install -r requirements.txt
