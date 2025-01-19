import re

# File paths
input_file = "essay.txt"  # Replace with your input file name
output_file = "updated_essay.txt"  # Replace with the desired output file name

# Affiliate ID
affiliate_id = "cr2tm-20"

# Regex pattern to match Amazon product links
amazon_link_pattern = r"(https://www\.amazon\.com/dp/\w+)"

# Read the input file
with open(input_file, "r") as file:
    content = file.read()

# Find all Amazon links and append the affiliate tag
def add_affiliate_tag(match):
    base_url = match.group(1)
    return f"{base_url}?tag={affiliate_id}"

# Replace Amazon links in the content
content = re.sub(amazon_link_pattern, add_affiliate_tag, content)

# Remove Roman numeral section headings (e.g., I. Introduction -> Introduction)
content = re.sub(r"^\s*[IVXLCDM]+\.\s*", "", content, flags=re.MULTILINE)

# Write the updated content to the output file
with open(output_file, "w") as file:
    file.write(content)

print(f"Processing complete. Updated file saved as: {output_file}")
