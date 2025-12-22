import os
import sys
import comtypes.client

def log(message):
    print(message)
    sys.stdout.flush()

def main():
    if len(sys.argv) < 3:
        sys.exit(1)

    # Use absolute paths - critical for Windows COM
    input_path = os.path.abspath(sys.argv[1])
    output_path = os.path.abspath(sys.argv[2])

    log("Process: Launching PowerPoint Virtual Engine...")

    # Initialize PowerPoint COM object
    powerpoint = None
    try:
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        # Keep it hidden
        powerpoint.Visible = 1 

        log("Process: Indexing slides and master layouts...")
        
        # Open the presentation
        # 1: ReadOnly, 0: Untitled, 0: WithWindow
        deck = powerpoint.Presentations.Open(input_path, 1, 0, 0)
        
        # FormatType 32 = PDF
        deck.SaveAs(output_path, 32)
        deck.Close()
        
        log("Success: PPT converted to PDF.")
        sys.exit(0)

    except Exception as e:
        log(f"Critical Error: {str(e)}")
        sys.exit(1)
    finally:
        if powerpoint:
            powerpoint.Quit()

if __name__ == "__main__":
    main()