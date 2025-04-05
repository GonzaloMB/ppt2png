#!/usr/bin/env python3
import os
import subprocess
import sys
from pathlib import Path


def convert_ppt_to_png(ppt_file, parent_output_dir="ppt-png"):
    if not os.path.exists(ppt_file):
        print(f"The file '{ppt_file}' does not exist.")
        sys.exit(1)

    base_name = Path(ppt_file).stem

    # Create parent_output_dir/ppt_file_name/ e.g. ppt-png/presentation/
    if not os.path.exists(parent_output_dir):
        os.makedirs(parent_output_dir, mode=0o755)

    output_dir = os.path.join(parent_output_dir, base_name)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir, mode=0o755)

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

    rename_slides(output_dir, base_name)


def rename_slides(folder, base_name):
    all_files = os.listdir(folder)
    png_files = [f for f in all_files if f.lower().endswith(".png")]

    # Sort them to ensure a natural order
    png_files.sort()

    for index, filename in enumerate(png_files, start=1):
        old_path = os.path.join(folder, filename)
        new_filename = f"slide{index}.png"
        new_path = os.path.join(folder, new_filename)
        os.rename(old_path, new_path)

    print(f"Renamed {len(png_files)} slides to slideX.png format.")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python convert-ppt-to-png.py file.pptx")
        sys.exit(1)

    ppt_file = sys.argv[1]
    convert_ppt_to_png(ppt_file)
