import subprocess
import os
import re


def convert_docx_to_markdown(input_docx, output_md):
    """
    Converts a .docx file to Markdown using Pandoc and cleans up the output.

    :param input_docx: Path to the input .docx file.
    :param output_md: Path to the output .md file.
    """
    try:
        # Ensure the input file exists
        if not os.path.exists(input_docx):
            raise FileNotFoundError(f"Input file '{input_docx}' does not exist.")
        
        # Create a media folder for extracted images
        media_folder = os.path.join(os.path.dirname(output_md), "media")
        if not os.path.exists(media_folder):
            os.makedirs(media_folder)

        # Construct the pandoc command
        command = [
            "pandoc",
            input_docx,
            "-o", output_md,
            "--extract-media=./media",  # Extract images to media folder
            "--columns=100",            # Prevents word wrapping for better table handling
            "--wrap=none",              # Avoids wrapping in the table content
            "--from", "docx",
            "--to", "markdown"
        ]

        # Execute the command
        subprocess.run(command, check=True)

        print(f"Conversion successful! Markdown file saved at: {output_md}")
        print(f"Extracted media saved in './media' folder.")

        # Clean up the Markdown file
        clean_markdown_file(output_md)

    except subprocess.CalledProcessError as e:
        print("An error occurred during conversion:", e)
    except Exception as e:
        print("An error occurred:", e)


def clean_markdown_file(md_file):
    """
    Cleans up the converted Markdown file by removing unnecessary attributes 
    and fixing formatting issues.

    :param md_file: Path to the Markdown file to be cleaned.
    """
    try:
        with open(md_file, "r", encoding="utf-8") as file:
            content = file.read()

        # Remove image size attributes like {width="..." height="..."}
        content = re.sub(r'\{width=".*?"\s*height=".*?"\}', '', content)

        # Fix table formatting by ensuring consistent table row separators
        content = re.sub(r'(?<=\+)=+', '-', content)  # Replace '=' with '-' for Markdown tables

        # Remove unwanted attributes like {.underline}
        content = re.sub(r'\{\.underline\}', '', content)

        # Write the cleaned content back to the file
        with open(md_file, "w", encoding="utf-8") as file:
            file.write(content)

        print(f"Cleaned up Markdown file: {md_file}")
        print("Fixed table formatting and removed unwanted attributes.")

    except Exception as e:
        print("An error occurred during cleanup:", e)


# Example usage
if __name__ == "__main__":
    input_file = "example.docx"  # Replace with your .docx file
    output_file = "example.md"  # Desired output .md file
    convert_docx_to_markdown(input_file, output_file)
