import os
import sys
import pikepdf

def log(message):
    print(message)
    sys.stdout.flush()

def main():
    # Args: script.py <input_path> <output_path> <user_password>
    if len(sys.argv) < 4:
        log("Error: Missing parameters for encryption.")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]
    user_pass = sys.argv[3]

    log("Process: Initializing Security Vault...")

    try:
        with pikepdf.open(input_path) as pdf:
            log("Process: Applying AES-256 bit encryption...")
            
            # Setting permissions: we lock printing and extracting by default
            # but allow the user to open it with the password.
            no_permissions = pikepdf.Permissions(extract=False, print_lowres=False)
            
            encryption_settings = pikepdf.Encryption(
                user=user_pass, 
                owner="omnigrid_admin", # Internal owner key
                allow=no_permissions
            )

            pdf.save(output_path, encryption=encryption_settings)
            
        log("Success: PDF encrypted and locked.")
        sys.exit(0)

    except Exception as e:
        log(f"Critical Error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()