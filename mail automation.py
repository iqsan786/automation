import pandas as pd
import yagmail

file_path=r'C:\Users\iqsha\OneDrive\Documents\intern details.xlsx'
df= pd.read_excel(file_path)
# Initialize yagmail client
sender_email = 'iqshanm24@gmail.com'
sender_password = 'wvbt zifi qlfk jcdp'
yag = yagmail.SMTP(sender_email, sender_password)

for index, row in df.iterrows():
    if row['STATUS'] == 'completed'and row['FEEDBACK STATUS']!='sent':
        recipient= row['EMAIL']
        name= row['NAME']
        subject = 'Feedback Request'
        body = f"""
        hi {name},
        I hope this email finds you well. We would like to request your feedback on the recent internship experience. Your insights are valuable to us and will help us improve our program.
        Please take a moment to fill out the feedback form at the following link: [Feedback Form Link]
        Thank you for your time and support.
        Best regards,
        VDart
        """
        try:
            yag.send(to=recipient,subject=subject,contents=[body])
            print(f'email sent to{name} at {recipient}')
            df.at[index,'FEEDBACK STATUS']='sent'
        except Exception as e:
            print("failed to send email to {name} at {recipient} error:{e}")
            
df.to_excel(file_path, index=False)
print("Emails sent and feedback status updated successfully.")