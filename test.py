import os
import re
import base64

def convert_base64_images(md_file, output_md):
    """
    Converts Base64-encoded images in a Markdown file to local image files.
    
    :param md_file: Path to the original Markdown file.
    :param output_md: Path to the output Markdown file with updated image paths.
    """
    try:
        # Read the Markdown file
        with open(md_file, "r", encoding="utf-8") as file:
            content = file.read()

        # Create a folder to save images
        media_folder = os.path.join(os.path.dirname(output_md), "media")
        if not os.path.exists(media_folder):
            os.makedirs(media_folder)

        # Find all Base64-encoded images
        base64_images = re.findall(r'\[.*?\]:\s*<data:image/.*?;base64,(.*?)>', content)

        # Replace Base64-encoded images with local file paths
        for idx, base64_data in enumerate(base64_images):
            try:
                # Decode the Base64 image
                image_data = base64.b64decode(base64_data)
                image_name = f"image_{idx + 1}.png"
                image_path = os.path.join(media_folder, image_name)

                # Save the image
                with open(image_path, "wb") as img_file:
                    img_file.write(image_data)

                # Replace the Base64 reference in the Markdown file
                content = content.replace(
                    f"<data:image/png;base64,{base64_data}>",
                    f"./media/{image_name}"
                )

                print(f"Extracted image saved as: {image_path}")

            except Exception as e:
                print(f"Failed to decode and save image {idx + 1}: {e}")

        # Save the updated Markdown file
        with open(output_md, "w", encoding="utf-8") as file:
            file.write(content)

        print(f"Converted Markdown file saved at: {output_md}")
        print(f"Images saved in: {media_folder}")

    except Exception as e:
        print(f"An error occurred: {e}")


# Example usage
if __name__ == "__main__":
    input_md = "test1.md"  # Path to the original Markdown file
    output_md = "output.md"      # Path to the output Markdown file
    convert_base64_images(input_md, output_md)
