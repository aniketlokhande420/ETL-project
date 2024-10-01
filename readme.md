# ETL Project

## Introduction
This project contains scripts to convert a given XML file into an XLSX file with specific data modifications. There are two Python scripts included:
- **`converter.py`**: Converts an XML file located in the current directory to an XLSX file in the same directory.
- **`converter-api.py`**: Provides the same functionality but is implemented as an API. It accepts a Google Drive URL for an XML file and outputs an XLSX file available for download.

## How to Run the Scripts

### Prerequisites
- Clone the project to your local machine.
- Should have python installed.
- You can either use a virtual environment or install the dependencies in your root Python environment.

### Installing Dependencies
1. Open a command prompt terminal.
2. Run the following command to install the necessary packages:
    ```bash
    pip install -r requirements.txt
    ```

### Running `converter.py`
1. Place the input XML file in the project directory. You can use the provided `Input.xml` file or rename your own file to `Input.xml`.
2. Run the following command in the terminal:
    ```bash
    python converter.py
    ```
3. After execution, you should see `output.xlsx` generated in the same directory. You can open this file in Microsoft Excel.

### Running `converter-api.py`
1. Start the local server to accept incoming requests. Run the following command:
    ```bash
    python converter-api.py
    ```
2. You should see a message like:
    ```
    Running on http://127.0.0.1:5000
    ```
   This confirms that the server is up and running.

3. In **Postman**, create a new POST request and use the following cURL command:
    ```bash
    curl --location 'http://127.0.0.1:5000/convert' \
    --header 'Content-Type: application/json' \
    --data '{"xml_url": "https://drive.google.com/file/d/1fbwTF0bWoseNJGgpCjl6fJdpaIb260nQ/view"}'
    ```

4. You can replace the `"xml_url"` in the body with your own Google Drive URL. Ensure the file is publicly accessible.

### Handling the API Response
- Upon receiving a **status 200** response, this means the request was successful.
- To download the file, click on the three dots in the Postman response tab and select **Save response to file**.
- Save the file to your desired location.
- Now you can open this file with ms-excel.

## Notes
- Ensure that the XML file is shared publicly on Google Drive for the API to access it.
- If you encounter any issues, ensure that dependencies are properly installed and the correct file paths are used.
