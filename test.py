import subprocess
import os

def convert_docx_to_md(input_docx, output_md):
    try:
        # Ensure the input file exists
        if not os.path.exists(input_docx):
            raise FileNotFoundError(f"Input file '{input_docx}' does not exist.")
        
        # Construct the pandoc command
        command = [
            "pandoc",
            input_docx,
            "-o", output_md,
            "--extract-media=./media"
        ]

        # Execute the command
        subprocess.run(command, check=True)

        print(f"Conversion successful! Markdown file saved at: {output_md}")
        print(f"Extracted media saved in './media' folder.")

    except subprocess.CalledProcessError as e:
        print("An error occurred during conversion:", e)
    except Exception as e:
        print("An error occurred:", e)


# Example usage
input_file = "Hướng dẫn sử dụng Omni Facebook 2024.docx"  # Replace with your .docx file
output_file = "test.md"  # Desired output .md file
convert_docx_to_md(input_file, output_file)
