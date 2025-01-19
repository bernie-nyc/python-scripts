import re

# Define the affiliate ID to append
affiliate_id = "cr2tm-20"

# Prompt the user for the input filename
input_filename = input("Enter the input filename (with extension): ")

# Prompt the user for the output filename
output_filename = input("Enter the output filename (with extension): ")

# Define a regex pattern to find Amazon links
amazon_link_pattern = r"https://www\.amazon\.com/dp/[\w\d]+"

try:
    # Read the input file
    with open(input_filename, 'r') as file:
        content = file.read()
    
    # Find all Amazon links in the file
    amazon_links = re.findall(amazon_link_pattern, content)
    
    # Replace each Amazon link by appending the affiliate tag
    for link in amazon_links:
        if "?tag=" not in link:  # Avoid duplicating the affiliate tag if it already exists
            updated_link = f"{link}?tag={affiliate_id}"
            content = content.replace(link, updated_link)
    
    # Remove Roman numeral section headings
    content = re.sub(r"^\s*[IVXLCDM]+\.\s*", "", content, flags=re.MULTILINE)
    
    # Write the updated content to the output file
    with open(output_filename, 'w') as file:
        file.write(content)
    
    print(f"File successfully processed. Updated content saved to: {output_filename}")

except FileNotFoundError:
    print(f"Error: The file '{input_filename}' was not found. Please check the filename and try again.")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
