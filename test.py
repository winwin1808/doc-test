import subprocess
import os
import re

def convert_docx_to_md_with_cleanup(input_docx, output_md):
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

        # Clean up the Markdown file to remove image size attributes
        cleanup_markdown_file(output_md)

    except subprocess.CalledProcessError as e:
        print("An error occurred during conversion:", e)
    except Exception as e:
        print("An error occurred:", e)


def cleanup_markdown_file(md_file):
    try:
        with open(md_file, "r", encoding="utf-8") as file:
            content = file.read()

        # Regex to remove {width="..." height="..."}
        cleaned_content = re.sub(r'\{width=".*?"\s*height=".*?"\}', '', content)

        # Write the cleaned content back to the file
        with open(md_file, "w", encoding="utf-8") as file:
            file.write(cleaned_content)

        print(f"Cleaned up Markdown file: {md_file}")
        print("Removed size attributes for images.")

    except Exception as e:
        print("An error occurred during cleanup:", e)


# Example usage
if __name__ == "__main__":
    input_file = "example.docx"  # Replace with your .docx file
    output_file = "example.md"  # Desired output .md file
    convert_docx_to_md_with_cleanup(input_file, output_file)
