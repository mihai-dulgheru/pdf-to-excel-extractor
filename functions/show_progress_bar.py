import sys


def show_progress_bar(processed, total):
    """
    Displays a simple text progress bar in the console.
    """
    progress = int((processed / total) * 100)
    bar_length = 50  # length of the progress bar in characters
    filled_length = int(bar_length * progress / 100)
    bar = '#' * filled_length + '-' * (bar_length - filled_length)

    sys.stdout.write(f"\rProcessing PDF files: [{bar}] {progress}%")
    sys.stdout.flush()

    if processed == total:
        print()  # line break after progress reaches 100%
