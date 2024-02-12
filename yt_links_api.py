import openpyxl as xl
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Authentication function
def get_authenticated_service():
    flow = InstalledAppFlow.from_client_secrets_file(
        'client_secrets.json',  # Replace with your downloaded JSON file
        scopes=['https://www.googleapis.com/auth/youtube.force-ssl']
    )
    credentials = flow.run_local_server(port=0)
    return build('youtube', 'v3', credentials=credentials)

# Load the input Excel file
input_workbook = xl.load_workbook('details.xlsx')  # Replace with your file path
input_sheet = input_workbook['Sheet5'] # respective sheet number to be modified

# Create a new Excel file for output links
output_workbook = xl.Workbook()
output_sheet = output_workbook.active
output_sheet['A1'] = 'Title'
output_sheet['B1'] = 'Link'

youtube = get_authenticated_service()

# Iterate through rows in the input sheet
row_index = 2
while True:
    title = input_sheet.cell(row=row_index, column=2).value
    date_time = input_sheet.cell(row=row_index, column=3).value
    stream_key = input_sheet.cell(row=row_index, column=4).value
    if input_sheet.cell(row=row_index, column=5).value == "off":
        dvr = False
    else:
        dvr = True

    if input_sheet.cell(row=row_index, column=6).value == "normal":
        latency = "normal"
    else:
        latency = "ultraLow"
    if input_sheet.cell(row=row_index, column=7).value == "on":
        auto_start = True
    else:
        auto_start = False

    if not title:
        break

    # Create the live broadcast using the YouTube API
    
    broadcast_snippet = {
        "scheduledStartTime": date_time,
        "title": title,
        }
    details = {
                "enableDvr": dvr,
                "latencyPreference": latency,
                "enableAutoStart": auto_start,
                "boundStreamId": stream_key
            }
    insert_response = youtube.liveBroadcasts().insert(
    part='snippet,status,contentDetails',
    body=dict(snippet=broadcast_snippet, status=dict(privacyStatus='unlisted'), contentDetails=details)

    ).execute()

    # request = youtube.liveBroadcasts().bind(
    #     id = insert_response['id'],
    #     streamId = stream_key
    # )
    # response = request.execute()

    # print(response)

    # request = youtube.liveBroadcasts().bind(
    #     id = insert_response['id'],
    #     part = "snippet",
    #     streamId = stream_key
    # )
    # bind_response = request.execute()

    # print(bind_response)

    live_stream_link = f"https://www.youtube.com/watch?v={insert_response['id']}"

    print(f"{title} with youtube link {live_stream_link} is created successfully.")

    # Write the link to the output Excel sheet
    output_sheet.cell(row=row_index, column=1).value = title
    output_sheet.cell(row=row_index, column=2).value = live_stream_link

    row_index += 1

# Save the output Excel file
output_workbook.save('output_links.xlsx')
