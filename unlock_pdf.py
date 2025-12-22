import os
import sys
import pikepdf

def log(message):
    """Sends real-time updates to Node.js backend"""
    print(message)
    sys.stdout.flush()

def main():
    # Args: script.py <input_path> <output_path> <password>
    if len(sys.argv) < 4:
        log("Error: Missing parameters for decryption.")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]
    password = sys.argv[3]

    log("Process: Initializing Decryption Engine...")

    try:
        # 1. Attempt to open with password
        log("Process: Authenticating security credentials...")
        with pikepdf.open(input_path, password=password) as pdf:
            
            # 2. Check if the PDF was actually encrypted
            log("Process: Stripping encryption headers and restrictions...")
            
            # 3. Save as a new file (unencrypted by default)
            pdf.save(output_path)
            
        log("Success: PDF unlocked and restrictions removed.")
        sys.exit(0)

    except pikepdf.PasswordError:
        log("Error: Incorrect password. Decryption failed.")
        sys.exit(1)
    except Exception as e:
        log(f"Critical Error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()