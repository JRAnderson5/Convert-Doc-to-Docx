import os
import win32com.client as win32

def convert_doc_to_docx(input_folder, output_folder):
    """
    Converts all .doc files in a specified input folder to .docx format
    and saves them to an output folder.

    Args:
        input_folder (str): The path to the folder containing the .doc files.
        output_folder (str): The path to the folder where the converted .docx
                             files will be saved.
    """
    # Ensure the output folder exists, create it if not
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"Created output folder: {output_folder}")

    # Connect to Microsoft Word application
    # This line attempts to open a Word application in the background
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False  # Keep Word hidden
    except Exception as e:
        print(f"Error: Could not open Microsoft Word application. "
              f"Please ensure Word is installed and accessible. Details: {e}")
        return

    # Define the Word file format for DOCX (wdFormatDocumentDefault)
    # This constant tells Word to save the document in its default format, which is .docx
    wdFormatDocumentDefault = 16

    # Iterate through all files in the input folder
    for filename in os.listdir(input_folder):
        if filename.endswith(".doc"):
            input_filepath = os.path.join(input_folder, filename)
            # Create the output filename by replacing .doc with .docx
            output_filename = filename.replace(".doc", ".docx")
            output_filepath = os.path.join(output_folder, output_filename)

            print(f"Attempting to convert: {filename}")
            try:
                # Open the .doc document
                doc = word.Documents.Open(input_filepath)
                # Save the document as .docx
                doc.SaveAs2(output_filepath, FileFormat=wdFormatDocumentDefault)
                # Close the document without saving changes (as we just saved it as .docx)
                doc.Close(False)
                print(f"Successfully converted {filename} to {output_filename}")
            except Exception as e:
                print(f"Failed to convert {filename}. Error: {e}")
        else:
            print(f"Skipping non-.doc file: {filename}")

    # Quit the Word application
    word.Quit()
    print("\nConversion process completed.")

if __name__ == "__main__":
    # --- Configuration ---
    # IMPORTANT: Replace these paths with your actual input and output folders.
    # Example:
    # input_directory = "C:\\Users\\YourUser\\Documents\\OldDocs"
    # output_directory = "C:\\Users\\YourUser\\Documents\\ConvertedDocs"

    # Get the current working directory of the script
    current_dir = os.getcwd()

    # Define input and output directories relative to the script's location
    # You can change these to absolute paths if needed
    input_directory = os.path.join(current_dir, "input_docs")
    output_directory = os.path.join(current_dir, "output_docs")

    # Create dummy input files for demonstration if they don't exist
    # In a real scenario, you would place your actual .doc files here.
    if not os.path.exists(input_directory):
        os.makedirs(input_directory)
        print(f"Created dummy input folder: {input_directory}")
        # Create a dummy .doc file (this will be an empty file, but serves for testing)
        # For a real .doc file, you'd need to manually create one in Word.
        try:
            # This creates an empty file, which Word can open but won't have content.
            # For a true test, place an actual .doc file in 'input_docs'.
            with open(os.path.join(input_directory, "sample.doc"), "w") as f:
                f.write("This is a dummy .doc file content.")
            print("Created dummy 'sample.doc' in input_docs. Please replace with actual .doc files for full functionality.")
        except Exception as e:
            print(f"Could not create dummy .doc file: {e}")


    print(f"\nInput folder set to: {input_directory}")
    print(f"Output folder set to: {output_directory}\n")

    # Run the conversion
    convert_doc_to_docx(input_directory, output_directory)
