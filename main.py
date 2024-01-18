from fastapi import FastAPI
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()
# CORS middleware
origins = [
    "*",
]
 
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
 

@app.get("/")
async def root():
    return {"message": "Hello World"}

@app.post("/questions")
async def questions(params:dict):
    # Load the Excel file
    file_path = 'Questions.xlsx'
    df = pd.read_excel(file_path)

    grouped = df.groupby([f'Topic Name {params["language_key"]}', f'Topic Statement {params["language_key"]}'])

    # Initializing the final JSON structure
    final_json = {"data": []}

    # Iterating through each group
    for (topic, statement), group in grouped:
        topic_dict = {
            "id": int(group['Pair ID'].iloc[0]),  # Convert numpy.int64 to int
            "topic": topic,
            "statement": statement,
            "questions": []
        }

        # Grouping by questions within each topic
        question_groups = group.groupby(['Question ID', f'Question Text {params["language_key"]}'])
        for (question_id, question_text), q_group in question_groups:
            question_id = int(question_id)  # Convert numpy.int64 to int
            question_dict = {
                "id": question_id,
                "question": question_text,
                "choices": []
            }

            # Adding choices
            for i, row in q_group.iterrows():
                for choice_id in range(1, 6):
                    choice_key = f'Choice {choice_id} {params["language_key"]}'
                    if pd.notna(row[choice_key]):
                        choice_dict = {"id": int(row['Question ID']) * 10 + choice_id,  # Convert numpy.int64 to int
                                       "name": row[choice_key]}
                        question_dict["choices"].append(choice_dict)

            topic_dict["questions"].append(question_dict)

        final_json["data"].append(topic_dict)

    return final_json


@app.post("/report")
async def generate_report(user_data: dict):
    # Assuming your Excel file is named 'report.xlsx'

     # ... (your existing code)

    # Send email with the generated PDF attachment
    sender_email = "developer@agilecyber.com"  # Replace with your email
    receiver_email = user_data['mail']  # Replace with the recipient's email
    subject = "Report PDF"
    body = "Please find the attached report."

    # Create a MIME object
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))


    file_path = 'report.xlsx'

    # Read the Excel file into a DataFrame
    df = pd.read_excel(file_path)

    # Convert the data dictionary to a DataFrame
    data_df = pd.DataFrame(user_data["data"])

    # Merge the Excel DataFrame with the data DataFrame based on pair_id and responses
    merged_df = pd.merge(df, data_df, how='inner', left_on=['Question Pair', 'Response 1', 'Response 2'],
                        right_on=['pair_id', 'response_1', 'response_2'])

    # Extract English Text from the merged DataFrame
    english_texts = merged_df['English Text'].tolist()

    # Create a PDF file
    pdf_file_path = 'output_report.pdf'
    pdf = canvas.Canvas(pdf_file_path, pagesize=letter)
    pdf.setFont("Helvetica", 12)
    
    for i, text in enumerate(english_texts, start=1):
        pdf.drawString(50, 800 - i * 12, text)

    pdf.save()
    with open(pdf_file_path, "rb") as attachment:
        part = MIMEApplication(attachment.read(), Name="output_report.pdf")
        part['Content-Disposition'] = 'attachment; filename="output_report.pdf"'
        msg.attach(part)

    # Set up the SMTP server
    smtp_server = "smtp.gmail.com"  # Replace with your SMTP server
    smtp_port = 587  # Replace with the SMTP server port

    # Log in to the SMTP server
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login("developer@agilecyber.com", "sleiemskccrmtmdc")  # Replace with your email password

        # Send the email
        server.sendmail(sender_email, receiver_email, msg.as_string())
    return {"message": "Report Sent to Mail Successfully"}