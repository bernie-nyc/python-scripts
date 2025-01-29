import fitz  # PyMuPDF
import difflib
import os

def compare_documents():
    print("\n--- Document Comparison Tool ---")

    # Prompt user for file paths
    source_path = input("Enter the path to the source (control) document: ").strip()
    comparison_path = input("Enter the path to the comparison document: ").strip()

    # Verify files exist
    if not os.path.exists(source_path):
        print(f"Error: Source file not found at {source_path}")
        return
    if not os.path.exists(comparison_path):
        print(f"Error: Comparison file not found at {comparison_path}")
        return

    # Open the documents
    source_doc = fitz.open(source_path)
    comparison_doc = fitz.open(comparison_path)

    # Create a new document for results
    output_doc = fitz.open()

    # Iterate through pages
    for page_num in range(min(len(source_doc), len(comparison_doc))):
        source_page = source_doc[page_num]
        comparison_page = comparison_doc[page_num]

        # Extract words from pages
        source_words = source_page.get_text("text").split()
        comparison_words = comparison_page.get_text("text").split()

        # Compute differences using difflib
        diff = list(difflib.ndiff(source_words, comparison_words))

        added_words = {word[2:] for word in diff if word.startswith('+ ')}
        missing_words = {word[2:] for word in diff if word.startswith('- ')}

        # Create a new page for results
        new_page = output_doc.new_page(width=comparison_page.rect.width, height=comparison_page.rect.height)
        new_page.show_pdf_page(new_page.rect, comparison_doc, page_num)

        # Highlight added words
        for word in added_words:
            text_instances = new_page.search_for(word)
            for inst in text_instances:
                highlight = new_page.add_highlight_annot(inst)
                highlight.set_colors(stroke=(1, 0.75, 0.8))  # Pink highlight

        # Annotate missing words
        if missing_words:
            annotation_text = f"Omitted words: {', '.join(list(missing_words)[:10])}..."
            new_page.insert_text((50, 50), annotation_text, fontsize=10, color=(1, 0, 0))

    # Save the output document
    output_path = os.path.join(os.getcwd(), "Comparison_Result.pdf")
    output_doc.save(output_path)
    output_doc.close()
    source_doc.close()
    comparison_doc.close()

    print(f"\nComparison completed. Output saved to {output_path}")

if __name__ == "__main__":
    compare_documents()
