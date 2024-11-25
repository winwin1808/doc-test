import subprocess
import os
import re


def convert_docx_to_markdown(input_docx, output_md):
    """
    Converts a .docx file to Markdown using Pandoc and fixes issues with tables.
    """
    try:
        if not os.path.exists(input_docx):
            raise FileNotFoundError(f"Input file '{input_docx}' does not exist.")
        
        media_folder = os.path.join(os.path.dirname(output_md), "media")
        if not os.path.exists(media_folder):
            os.makedirs(media_folder)

        command = [
            "pandoc",
            input_docx,
            "-o", output_md,
            "--extract-media=./media",
            "--columns=100",
            "--wrap=none",
            "--from", "docx",
            "--to", "markdown"
        ]

        subprocess.run(command, check=True)

        print(f"Conversion successful! Markdown file saved at: {output_md}")
        print(f"Extracted media saved in './media' folder.")

        fix_table_formatting(output_md)

    except subprocess.CalledProcessError as e:
        print("An error occurred during conversion:", e)
    except Exception as e:
        print("An error occurred:", e)


def fix_table_formatting(md_file):
    """
    Fixes complex Markdown tables by simplifying them to standard table syntax.
    """
    try:
        with open(md_file, "r", encoding="utf-8") as file:
            content = file.read()

        # Remove extra separators like '=' in tables
        content = re.sub(r'\+\=+\+', r'+---+', content)

        # Standardize table headers and separators
        content = re.sub(r'\| +\| +', '|', content)  # Fix spacing in table rows
        content = re.sub(r'^\| [^\|]* \|$', lambda m: fix_table_row(m.group()), content, flags=re.MULTILINE)

        # Remove excessive empty table rows
        content = re.sub(r'\| +\| +', '|', content)

        with open(md_file, "w", encoding="utf-8") as file:
            file.write(content)

        print(f"Table formatting fixed in: {md_file}")

    except Exception as e:
        print("An error occurred while fixing table formatting:", e)


def fix_table_row(row):
    """
    Ensures consistent column alignment for a Markdown table row.
    """
    columns = row.strip('|').split('|')
    fixed_columns = [col.strip() for col in columns]
    return '| ' + ' | '.join(fixed_columns) + ' |'


# Example usage
if __name__ == "__main__":
    input_file = "example.docx"  # Replace with your .docx file
    output_file = "example.md"  # Desired output .md file
    convert_docx_to_markdown(input_file, output_file)
