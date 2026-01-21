import pandas as pd
import os


def create_bat_file(email):
    bat_content = f"""@echo off
echo Creating registry entries...

REM Create MSIPC node
reg add "HKEY_CURRENT_USER\\Software\\Classes\\Local Settings\\Software\\Microsoft\\MSIPC" /f

REM Create aip-addin node
reg add "HKEY_CURRENT_USER\\Software\\Classes\\Local Settings\\Software\\Microsoft\\MSIPC\\aip-addin" /f

REM Create RMSUser value under aip-addin node
reg add "HKEY_CURRENT_USER\\Software\\Classes\\Local Settings\\Software\\Microsoft\\MSIPC\\aip-addin" /v "RMSUser" /t REG_SZ /d "{email}" /f

REM Create UPN value under aip-addin node
reg add "HKEY_CURRENT_USER\\Software\\Classes\\Local Settings\\Software\\Microsoft\\MSIPC\\aip-addin" /v "UPN" /t REG_SZ /d "{email}" /f

echo Registry entries created successfully.
pause
"""
    return bat_content


def main():
    # Read the Excel file
    excel_file = "emails.xlsx"  # Replace with your Excel file name
    df = pd.read_excel(excel_file)

    # Get emails from the first column
    emails = df.iloc[:, 0].tolist()

    # Create a directory to store bat files
    output_dir = "bat_files"
    os.makedirs(output_dir, exist_ok=True)

    # Generate bat files for each email
    for i, email in enumerate(emails, 1):
        bat_content = create_bat_file(email)
        file_name = f"create_registry_entries_{i}.bat"
        file_path = os.path.join(output_dir, file_name)

        with open(file_path, "w") as f:
            f.write(bat_content)

        print(f"Created {file_name}")


if __name__ == "__main__":
    main()
