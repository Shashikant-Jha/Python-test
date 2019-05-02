import os, json, schedule, time, requests, base64, zipfile, ntpath, smtplib, email, ssl, shutil
from datetime import datetime, timedelta
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import numpy as np

class scheduledExtractor():
    
    serviceAccount = 'agent-export@analytics-bot-84386.iam.gserviceaccount.com'
    privateKey = '-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQDGJEPRamRhjz/n\nRb6J3pkpqte8wYT02SoaRnm9H1zkrHLftmedifuhd8ugpFZKkvv5HfuohyAbMsm4\ngxN+sK0FZiPnsYJs/3NVp4s7OTptV7Sqpa1hoMeXgbFQXOmL86Gpjgr1xtedLF/8\nF7efjpScRvaH9oLXAvv9fgsTmwSNyvScObI9/NDrOd5Bz7ePGJVM7Tz2YKmmZe5l\n0aldCTRHlwGjr7tLFeLQmPZwZosKBb2zCUSgKWWv7NdyaqdEL4cEir0nKyMrZhxv\nyDZ0uyy+oN6zngPSy9PmI8as+rlN1xHooqRfK3GMBD42e79G28n+AYENIS7RS27P\nWUmEGOGBAgMBAAECggEAJVOLmxXJ1z8fMLhIzrwqalkOzzY9j8qhTdXc9S3FWCbM\ndSxtlJX7200wBREwWFgcM6/cSsE54SXOLn4y2/j4fz3gobMk+NeeqJIpfnynbxXI\nqSUQ5oVdVENQXe1C6eR5nfFuSyKsV6WWZ20oYFyBznbn1nEv8MVVJ0npWEYps3Ql\ndlYY8kV7pP27U516ONBXM2ktmZkzi3ZD5NyjDkM9QesvBaMAxnXl5wID++mc3f6v\nmXgj5lEgApd4UYHQRpoPAJKR4fgz8Lf2UDDjSA5SkLVtRs1udTlvcp9SHzSFXQjs\nXRIXDLuBDQWFB7SCF0bnqmeyqxFLjQvvHeca96JyCQKBgQDq14CexjScg6XwgEwD\n34zCbKgHsZ94rdEOmGQiSvSa//NK3t/VjCsh8pT/ilJRPp4jlgnLVMBermCzJGmD\n7yMqJ1NlKvPBMFb5v97Z4D8qGvnW9l/2Z8qyWq0fx3CUIOCyWszgFbZagJxIufM0\nsUjTWFuHIsgB4D/IRlCzKN3EewKBgQDX/ksxZvo4at3ynx5PEjw9oJrIeWf2sMyf\n31UPvXlCNVd4i5+Ip5isxssg86X62Y44R7fod/hitEGkI6jvEfud4KAsOOLzJrri\n/BGZcJQsyzdEMvayQFVC+Bny8YLq4HFOOCWN5sldCVy4ICVOP6AsUyJogKPE160N\nTbm7qjQnMwKBgQCovzmY3Wuom6B9dHMqxVPr0Q/cl3Gz0ZJLHo21Zo7lYc18mzHT\nzOiheCJAjTAhWRFhGMro5Hlmj37EuwFm4EswWxm9tGN7CjU1OP31KQG6S7ADGv5R\nnLs19Zo5H6Jxjj5nan+U9YUW+xtR1uw/jLR7yR3buM5nKrAqRPxwAIl6SQKBgHXU\nW+DPdKFia0H4W+h24jYcb1P+JyEEwhxVEWpMyuG7h8RTJuy9wSRALoADawC1vkgl\nl8ZR7EasX0LT0UzaWpF+AOMfBW/wKPO63z79k1f8ZGHoC3yia+DkyAMojWrkles9\n5f7Lb/45JUOtBazyOMb1c0ffJXg00Er5o+EksN7RAoGBAOB0RJuDDoqCVLdH3ut6\nYhXfE/zImA+oNwFRPWyPUXiSago8PAJc6fb1XnsQqy5Ur8meYHQVmQ/ZGOCdHaa7\n8xCJLRAlFTZ0OxGU6lTx/uh5nMzaX8SPIN5jl8g7kR1v8Dpx44NR/Gu7uHgEBVgk\n6bqAj7vSHwfNWFxrs2w/9phL\n-----END PRIVATE KEY-----\n'
    scope = 'https://www.googleapis.com/auth/cloud-platform'
    token = ''
    currDateTime = ''
    reportMismatch = ''

    def __init__(self):
        self.currDateTime = datetime.today().strftime('%Y-%m-%d')
        self.exportAgent()
    
    def exportAgent(self):
        print("Check")
        self.token = self.getOauthToken()
        self.getAgentExport()
        self.reportMismatch = self.compareReports()
        print("value ", str(self.reportMismatch))
        self.sendEmail(self.reportMismatch)
        return
    
    def getOauthToken(self):
        payload = {"email": self.serviceAccount, "key": self.privateKey, "scopes": [self.scope]}
        resp = requests.post('https://guarded-caverns-38909.herokuapp.com/api/generateToken', json=payload)
        print(resp.json()['token'])
        return resp.json()['token']

    def getAgentExport(self):
        headers = {'Authorization': 'Bearer ' + self.token}
        resp = requests.post('https://dialogflow.googleapis.com/v2/projects/analytics-bot-84386/agent:export', headers=headers)
        # print(resp.json())
        agentZip = resp.json()['response']['agentContent']
        fileData = base64.b64decode(agentZip)
        newFileByteArray = bytearray(fileData)
        f = open('branding-line ' + str(self.currDateTime) + '.zip', 'w+b')
        f.write(newFileByteArray)
        f.close()
        self.unzipFile()
        self.createReport()
        return

    def unzipFile(self):
        zip_ref = zipfile.ZipFile('branding-line ' + str(self.currDateTime) + '.zip', 'r')
        zip_ref.extractall('branding-line ' + str(self.currDateTime))
        zip_ref.close()
        os.remove('branding-line ' + str(self.currDateTime) + '.zip')
        return

    def createReport(self):
        # entites
        path_to_json = 'branding-line ' + str(self.currDateTime) + '/entities'
        json_files = [pos_json for pos_json in os.listdir(path_to_json) if pos_json.endswith('_en.json')]
        jsons_data = pd.DataFrame(columns=['entities', 'synonyms'])
        for index, js in enumerate(json_files):
            with open(os.path.join(path_to_json, js)) as json_file:
                json_text = json.load(json_file)
                entities = json_text[0]['value']
                synonyms = np.array(json_text[0]['synonyms'])
                jsons_data.loc[index] = [entities, synonyms]

        # intents
        path_to_json_intents = 'branding-line ' + str(self.currDateTime) + '/intents'
        json_files_intents = [pos_json for pos_json in os.listdir(path_to_json_intents) if pos_json.endswith('_en.json')]
        jsons_data_intents = pd.DataFrame(columns=['intents', 'utterances', 'parameters'])
        for index, js in enumerate(json_files_intents):
            with open(os.path.join(path_to_json_intents, js)) as json_file:
                json_text = json.load(json_file)
                intent = str(ntpath.basename(str(json_file.name)))[:-17]
                utterance = ''
                params = ''
                for i in json_text:
                    for d in i['data']:
                        utterance = utterance + d['text']
                        if 'meta' in d:
                            params = params + d['meta'] + ' '
                    utterance = utterance + '\n'
                jsons_data_intents.loc[index] = [intent, utterance, params]
        writer = pd.ExcelWriter('branding-report ' + str(self.currDateTime) + '.xlsx', engine='xlsxwriter')
        jsons_data.to_excel(writer, sheet_name='Entities')
        jsons_data_intents.to_excel(writer, sheet_name='Intents')
        writer.save()
        shutil.rmtree('branding-line ' + str(self.currDateTime))
        return
    
    def compareReports(self):
        prev = pd.ExcelFile('branding-report ' + str(datetime.strftime(datetime.now() - timedelta(1), '%Y-%m-%d')) + '.xlsx')
        prevEntities = pd.read_excel(prev,'Entities')
        prevIntents = pd.read_excel(prev,'Intents')
        current = pd.ExcelFile('branding-report ' + str(datetime.today().strftime('%Y-%m-%d')) + '.xlsx')
        currentEntities = pd.read_excel(current,'Entities')
        currentIntents = pd.read_excel(current,'Intents')
        if currentEntities.equals(prevEntities):
            return True
        else:
            return False

    def sendEmail(self, param):
        print("yo", str(param))
        print("email function")
        subject = "Mail alert - attachement"
        body = "Emailing branding line bot report - contains entities and intents"
        sender_email = "jhashashi669@gmail.com"
        password = "zxnabpuypsfafdcm"
        blind_receivers = ["jhashashi669@gmail.com"]
        text = ''
        for receiver_email in blind_receivers:
            message = MIMEMultipart()
            message["From"] = sender_email
            message["To"] = receiver_email
            message["Subject"] = subject
            if param == False:
                message.attach(MIMEText(body, "plain"))
                filename = 'branding-report ' + str(self.currDateTime) + '.xlsx'
                with open(filename, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename= {filename}",)
                message.attach(part)
                text = message.as_string()
            else:
                body = body + "\n No Change detected"
                message.attach(MIMEText(body, "plain"))
                text = message.as_string()
            context = ssl.create_default_context()
            context.check_hostname = False
            context.verify_mode = ssl.CERT_NONE
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                server.login(sender_email, password)
                server.sendmail(sender_email, receiver_email, text)
        print("mail sent")

    def __del__(self):
        print("Object destroyed")

def job():
    obj = scheduledExtractor()
    del obj
    return

if __name__ == "__main__":
    schedule.every().day.at("14:36").do(job)
    while True:
        schedule.run_pending()
        time.sleep(0)
