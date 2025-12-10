from PIL import Image, ImageDraw, ImageFont
import os
from dotenv import load_dotenv
import pandas as pd
import smtplib
from email.message import EmailMessage
import mimetypes
from io import BytesIO
from matplotlib import font_manager


load_dotenv()


def SendEmail(sender_email,password,receiver_email,subject,body,attachment_path=None):
    msg = EmailMessage()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.set_content(body)

    if attachment_path:
        mime_type, _ = mimetypes.guess_type(attachment_path)
        mime_type, mime_subtype = mime_type.split('/')

        with open(attachment_path, 'rb') as file:
            msg.add_attachment(file.read(),
                               maintype=mime_type,
                               subtype=mime_subtype,
                               filename=os.path.basename(attachment_path))

    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(sender_email, password)
        server.send_message(msg)
        print(f"Email sent to {receiver_email}")




def create_image_with_text(input_image_path, text, position=(50, 50), font_size=40, text_color=(255, 255, 255), font_path=None):
    img = Image.open(input_image_path)
    draw = ImageDraw.Draw(img)
    if font_path:
        font = ImageFont.truetype(font_path, font_size)
    else:
        font_path_found = font_manager.findfont(font_manager.FontProperties(family='DejaVu Sans'))
        font = ImageFont.truetype(font_path_found, font_size)

    draw.text(position, text, fill=text_color, font=font)

    if not os.path.exists("output"):
        os.makedirs("output")

    output_path = f"output/{text}.png"
    img.save(output_path)
    print("Saved:", output_path)

    return output_path



def main():
    sender_email = os.getenv('EMAIL_USERNAME')
    password = os.getenv('EMAIL_PASSWORD')
    excel_file_path = os.getenv('EXCEL_FILE_PATH')

    df = pd.read_excel(excel_file_path)

    for index, row in df.iterrows():
        receiver_email = row['Email']
        name = row['Name']
        subject = "Invitation to Fresher Party"
        body = f"Hello {name},\n\nPlease find your Ticket  attached."

        input_image_path = os.getenv('Image_path')
        text_to_add = name
        image_path = create_image_with_text(
            input_image_path,
            text_to_add,
            position=(1650, 100),
            font_size=30,
            text_color=(0,0,0),
            font_path=None
        )


        SendEmail(sender_email, password, receiver_email, subject, body, image_path)


if __name__ == "__main__":
    main()
