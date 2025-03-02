import os


def get_all_pdf_files(directory):
    """
    Recursively finds all PDF files in the given directory and subdirectories.
    :param directory: Root directory to search.
    :return: List of file paths.
    """
    pdf_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith(".pdf"):
                pdf_files.append(os.path.join(root, file))
    return pdf_files
