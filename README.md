# Automated Redmi 6 Review Analysis 

This project automates the analysis of product reviews for the Redmi 6 smartphone using Cohere's AI APIs for sentiment classification and summarization. It integrates with Google Sheets to read raw reviews, perform analysis, and write results (sentiment and summary) back into the sheet.

## Features
- Automatically fetches reviews from a Google Sheet**
- Uses Cohere to analyze sentiment (Positive, Neutral, Negative)**
- Saves sentiment and summary data back into the same Google Sheet**
- Generates a pie chart of sentiment distribution (optional for reporting)**


## Setup Instructions

1.  **Clone the repository** (if you have one):
    ```bash
    git clone your_repository_url
    cd your_project_directory
    ```
    (If you're not using Git, just navigate to your project folder in your terminal.)

2.  **Install the required Python libraries:**
    ```bash
    pip install -r requirements.txt
    ```

3. **Set up Google Sheets access**: 
     - Create a service account on Google Cloud Console.
     - Download the credentials.json and place it in your project folder.
     - Share your Google Sheet with the service account's email with Editor access.

4. **Set up Cohere API**:
     - Sign up at cohere.com and get your API key.
     - Add it to your environment or directly in the script.

4.  **Create a `.env` file** in the root directory of your project.

5.  **Add your Cohere API key to the `.env` file:**
    ```
    COHERE_API_KEY=your_actual_cohere_cloud_api_key_here
    ```
    **Important:** Replace `your_actual_cohere_cloud_api_key_here` with your actual API key from cohere. **Do not commit this file to version control for security reasons.**


## Tools and libraries
1. **Python**
2. **GSpread**
3. **Cohere API**
4. **Google Sheets API**


## How to Run the Assistant

1.  **Navigate to the project directory** in your terminal:
    ```bash
    cd your_project_directory
    ```

2.  **Run the main Python script** (e.g., if your script is named `app.py`):
    ```bash
    python app.py
    ```

    This will execute the script, which in the example provided, sends a query to the cohere model and prints the response to your terminal.

## Example Usage

The current example in `app.py` sends the following prompt to the cohere model: