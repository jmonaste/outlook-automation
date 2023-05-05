# Outlook Mail Automation

This Python script automates the processing of Outlook emails, allowing you to efficiently manage and extract important information from your inbox. By leveraging the `win32com` library, the script provides the following features:

- **Unread Email Retrieval**: The script retrieves unread emails from the Outlook inbox based on a specific subject, enabling you to focus on crucial messages that require attention.

- **Data Extraction**: It extracts relevant data, such as ID and date, from the email body. This feature helps you streamline information retrieval and analysis, saving valuable time and effort.

- **CSV Export**: The extracted data is stored in a CSV file, facilitating further data manipulation and integration with other systems or tools.

## Requirements

To run the script, ensure that you have the following prerequisites installed:

- Python 3.x
- `win32com` library
- `csv` library

## Usage

1. Clone this repository or download the script file `outlook_mail_automation.py` to your local machine.

2. Install the required libraries by running the following command:
    ```
    pip install pywin32
    ```

3. Update the script's configuration:
    - Modify the subject in the `unread_messages` line to match the desired email subject you want to process.
    - Customize the `ouput_file_path` variable to specify the output CSV file path and name.

4. Run the script:
    ```
    python outlook_mail_automation.py
    ```

5. Check the specified CSV file for the extracted data.

Feel free to adjust and extend the script according to your specific requirements.

## Contribution

Contributions to this project are welcome! If you encounter any issues or have suggestions for improvements, please open an issue or submit a pull request.

## License

This project is licensed under the [MIT License](LICENSE).

## Disclaimer

This script is provided as-is without any warranty. Use it at your own risk.

Please ensure compliance with your organization's policies and guidelines when automating email processing.