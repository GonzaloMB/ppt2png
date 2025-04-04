import os
import subprocess
import sys


def convert_ppt_to_png(ppt_file, output_dir):
    """
    Converts a PowerPoint file to PNG images for each slide.

    Requires LibreOffice to be installed.
    """
    # Check if the PPT file exists
    if not os.path.exists(ppt_file):
        print(f"The file '{ppt_file}' does not exist.")
        sys.exit(1)

    # Create the output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Command to convert PPT to PNG
    command = [
        "libreoffice",
        "--headless",
        "--convert-to",
        "png",
        "--outdir",
        output_dir,
        ppt_file,
    ]

    try:
        subprocess.run(command, check=True)
        print("Conversion completed successfully.")
    except subprocess.CalledProcessError as e:
        print("Error during conversion:", e)
        sys.exit(1)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script.py file.ppt")
        sys.exit(1)

    ppt_file = sys.argv[1]
    output_dir = "/ppt-png"  # Output folder as requested
    convert_ppt_to_png(ppt_file, output_dir)
