# ppt2png

A Python tool that converts a PowerPoint file (.ppt or .pptx) into PNG images, generating one image per slide.

## Description

The `ppt_to_png.py` script uses LibreOffice in headless mode to convert a PowerPoint file to PNG images. This is useful for:
- Extracting each slide as an image.
- Reusing visual content.
- Archiving presentations as images.

## Requirements

- **Python 3:** The script uses built-in modules (os, subprocess, sys), so no additional Python libraries are needed.
- **LibreOffice:** Must be installed on your system and accessible from the command line.

### Installing LibreOffice

- **macOS:**  
  Using Homebrew:  
  `brew install --cask libreoffice`
- **Linux (Debian/Ubuntu):**  
  `sudo apt-get install libreoffice`
- **Windows:**  
  Download the installer from the official LibreOffice website: [LibreOffice Download](https://www.libreoffice.org/download/download/).

## Installation and Setup

1. **Clone the repository:**
   - `git clone https://github.com/your_username/ppt2png.git`
   - `cd ppt2png`

2. **Create and activate a virtual environment (optional but recommended):**
   - Create the virtual environment:  
     `python3 -m venv venv`
   - Activate it:  
     - On macOS/Linux: `source venv/bin/activate`  
     - On Windows: `venv\Scripts\activate`

3. **Install dependencies:**
   Since this project does not use external Python libraries, the `requirements.txt` file is empty. If you add dependencies later, install them with:  
   `pip install -r requirements.txt`

## How to Use the Script

To run the script and convert a PowerPoint file to PNG images, use the following command:

   `python ppt_to_png.py path_to_file.ppt`

- Replace `path_to_file.ppt` with the path to your PowerPoint file.
- The script will create a folder `/ppt` (or the folder specified in the `output_dir` variable) and save one PNG image for each slide.

## Example

If you have a file named `presentation.ppt` in the same directory, run:

   `python ppt_to_png.py presentation.ppt`

The resulting images will be saved in the `/ppt` folder.

## Additional Notes

- Ensure you have write permissions to the output folder.
- You can modify the `output_dir` variable in the script to change the location where the images are saved.

## Contributions

Contributions are welcome. If you'd like to improve the script or add new features, fork the repository and submit a pull request.

## License

This project is licensed under the MIT License. See the LICENSE file for details.
