import gspread
import cohere
from dotenv import load_dotenv
import matplotlib.pyplot as plt
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from gspread_formatting import batch_updater
from collections import Counter
import os

# Load environment variables from .env file
load_dotenv()

# --- Google Sheets Setup ---
credentials_file = "C:\\Users\\DELL\\Downloads\\automate-review-analysis-4f50e640dbf2.json"
spreadsheet_name = "Copy of Team data"
worksheet_name = "Redmi6"

# --- Cohere Setup ---
co = cohere.Client(os.getenv("COHERE_API_KEY"))

try:
    # Authenticate with Google Sheets
    gc = gspread.service_account(filename=credentials_file)
    sh = gc.open(spreadsheet_name)
    worksheet = sh.worksheet(worksheet_name)
    data = worksheet.get_all_values()

    if data:
        header = data[0]
        try:
            review_column_index = header.index("Review")
            group_column_index = header.index("Group")

            # Step 1: Ensure 'AI Summary' and 'AI Sentiment' columns exist or create them
            if "AI Summary" not in header:
                worksheet.update_cell(1, len(header) + 1, "AI Summary")
                header.append("AI Summary")
            if "AI Sentiment" not in header:
                worksheet.update_cell(1, len(header) + 1, "AI Sentiment")
                header.append("AI Sentiment")

            summary_col_index = header.index("AI Summary") + 1
            sentiment_col_index = header.index("AI Sentiment") + 1

            print("Processing reviews...")

            # Step 2: Loop through each row and process Group 3 reviews
            for i, row in enumerate(data[1:], start=2):  # start=2 because row 1 is header
                if len(row) > max(review_column_index, group_column_index) and row[group_column_index] == '3':
                    review = row[review_column_index]
                    try:
                        print("\n---")
                        print(f"Row {i} Review:", review)

                        # Summarize if review is long
                        if len(review.split()) >= 30:
                            try:
                                prompt = f"Summarize this customer review:\n{review}"
                                response = co.chat(
                                    model="command-r-plus",
                                    message=prompt
                                )
                                summary = response.text
                            except Exception as e:
                                print(f"Summary error for row {i}: {e}")
                                summary = review  # fallback
                        else:
                            summary = review
                            print("Review too short for summary") 

                        print("Summary:", summary)

                        # Classify Sentiment using your fine-tuned model
                        response = co.classify(
                            model="7f41b709-3888-4ba2-860c-5c67e7961d1c-ft",  # Your fine-tuned model ID
                            inputs=[review]
                        )
                        sentiment = response.classifications[0].prediction
                        print("Sentiment:", sentiment)

                        # Write results back to the sheet
                        worksheet.update_cell(i, summary_col_index, summary)
                        worksheet.update_cell(i, sentiment_col_index, sentiment)

                    except Exception as e:
                        print(f"Error processing row {i}: {e}")
        except ValueError as e:
            print(f"Error: One or more required columns not found - {e}")
            exit()
            
    #Generate Pie Chart for Group 3 Sentiment

    def generate_group3_sentiment_pie_chart(worksheet, spreadsheet_id):
        all_rows = worksheet.get_all_values()
        header = all_rows[0]

        review_column_index = header.index("Review")
        group_column_index = header.index("Group")
        sentiment_col_index = header.index("AI Sentiment")

        group3_sentiments = []
        for row in all_rows[1:]:
            if len(row) > group_column_index and row[group_column_index] == '3':
                if len(row) > sentiment_col_index:
                    group3_sentiments.append(row[sentiment_col_index])

        if not group3_sentiments:
            print("No Group 3 sentiments found.")
            return

        sentiment_counts = Counter(group3_sentiments)
        chart_data_start_row = len(all_rows) + 3
        chart_data_start_col = 1  # Column A

        # Write headers
        worksheet.update_cell(chart_data_start_row, chart_data_start_col, "Sentiment")
        worksheet.update_cell(chart_data_start_row, chart_data_start_col + 1, "Count")

        # Write data rows
        for idx, (label, count) in enumerate(sentiment_counts.items()):
            worksheet.update_cell(chart_data_start_row + idx + 1, chart_data_start_col, label)
            worksheet.update_cell(chart_data_start_row + idx + 1, chart_data_start_col + 1, count)

        # Use Google Sheets API to insert a pie chart
        service = build('sheets', 'v4', credentials=worksheet.client.auth)

        requests = [{
            "addChart": {
                "chart": {
                    "spec": {
                        "title": "Sentiment Distribution for Group 3",
                        "pieChart": {  # Using pieChart instead of basicChart
                            "legendPosition": "RIGHT_LEGEND",
                            "domain": {
                                "sourceRange": {
                                    "sources": [{
                                        "sheetId": worksheet._properties['sheetId'],
                                        "startRowIndex": chart_data_start_row - 1,
                                        "endRowIndex": chart_data_start_row + len(sentiment_counts),
                                        "startColumnIndex": 0,
                                        "endColumnIndex": 1
                                    }]
                                }
                            },
                            "series": {
                                "sourceRange": {
                                    "sources": [{
                                        "sheetId": worksheet._properties['sheetId'],
                                        "startRowIndex": chart_data_start_row - 1,
                                        "endRowIndex": chart_data_start_row + len(sentiment_counts),
                                        "startColumnIndex": 1,
                                        "endColumnIndex": 2
                                    }]
                                }
                            },
                            "threeDimensional": False  # Optional: Set to True for 3D pie chart
                        }
                    },
                    "position": {
                        "overlayPosition": {
                            "anchorCell": {
                                "sheetId": worksheet._properties['sheetId'],
                                "rowIndex": chart_data_start_row - 1,
                                "columnIndex": 3
                            },
                            "offsetXPixels": 20,
                            "offsetYPixels": 20
                        }
                    }
                }
            }
        }]

        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests}
        ).execute()

        print("Pie chart for Group 3 sentiment inserted successfully.")

    
    
    generate_group3_sentiment_pie_chart(worksheet, sh.id)


except gspread.exceptions.APIError as e:
    print(f"Google Sheets API Error: {e}")
except Exception as e:
    print(f"Unexpected error: {e}")
