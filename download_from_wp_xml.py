import os
import re
import urllib.request
import xml.etree.ElementTree as ET

# Supported image file extensions
IMAGE_EXTENSIONS = ('.jpg', '.jpeg', '.png', '.gif', '.svg', '.webp')

def extract_and_download_images(xml_file_path, output_folder="institute_blog_images"):
    """
    Parses a WordPress WXR export file and downloads all associated blog images.
    """
    if not os.path.exists(xml_file_path):
        print(f"Error: The XML file '{xml_file_path}' was not found.")
        return

    # Create destination directory
    os.makedirs(output_folder, exist_ok=True)

    print(f"Parsing XML file: {xml_file_path}...")
    
    # Define XML namespaces used by WordPress WXR export
    namespaces = {
        'wp': 'http://wordpress.org/export/1.2/',
        'content': 'http://purl.org/rss/1.0/modules/content/'
    }

    try:
        tree = ET.parse(xml_file_path)
        root = tree.getroot()
    except ET.ParseError as e:
        print(f"Failed to parse XML file: {e}")
        return

    channel = root.find('channel')
    if channel is None:
        print("Invalid WordPress XML export file format.")
        return

    image_urls = set()

    # Iterate through all items in the XML export
    for item in channel.findall('item'):
        post_type = item.findtext('wp:post_type', namespaces=namespaces)

        # 1. Extract direct attachment URLs
        if post_type == 'attachment':
            attachment_url = item.findtext('wp:attachment_url', namespaces=namespaces)
            if attachment_url:
                clean_url = attachment_url.strip()
                if clean_url.lower().endswith(IMAGE_EXTENSIONS):
                    image_urls.add(clean_url)

        # 2. Extract images embedded in post content (<img src="...">)
        content = item.findtext('content:encoded', namespaces=namespaces)
        if content:
            img_srcs = re.findall(r'<img[^>]+src=["\'](.*?)["\']', content, re.IGNORECASE)
            for src in img_srcs:
                clean_src = src.strip().split('?')[0]  # Remove URL parameters if any
                if clean_src.lower().endswith(IMAGE_EXTENSIONS):
                    image_urls.add(clean_src)

    print(f"Found {len(image_urls)} unique image URLs.")

    # Request headers to prevent 403 Forbidden errors from server security filters
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36'
    }

    downloaded_count = 0
    skipped_count = 0

    # Download each image locally
    for url in sorted(image_urls):
        filename = os.path.basename(url.split('?')[0])
        if not filename:
            continue

        filepath = os.path.join(output_folder, filename)

        # Skip downloading if the image already exists
        if os.path.exists(filepath):
            print(f"[SKIPPED] Already exists: {filename}")
            skipped_count += 1
            continue

        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req) as response, open(filepath, 'wb') as out_file:
                out_file.write(response.read())
            print(f"[DOWNLOADED] {filename}")
            downloaded_count += 1
        except Exception as e:
            print(f"[ERROR] Failed to download {url}: {e}")

    print(f"\nDone! Downloaded: {downloaded_count} | Skipped: {skipped_count} | Total: {len(image_urls)}")

if __name__ == "__main__":
    # Specify your XML file path here
    XML_FILENAME = "export.xml"  
    
    # Specify your desired local target folder name here
    FOLDER_NAME = "institute_blog_images"
    
    extract_and_download_images(XML_FILENAME, FOLDER_NAME)