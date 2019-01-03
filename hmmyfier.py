import urllib.request
import urllib.parse
import constants
import datetime
import mimetypes
import os
import re
import schedule
import smtplib
import time
import yaml

from email.message import EmailMessage


def load_config():
    return yaml.safe_load(open(constants.CONFIG_FILENAME))


def create_directory(directory):
    if (not os.path.exists(directory)):
        os.mkdir(directory)


def download_images(output_folder):
    url = constants.URL_BASE + config["subreddit"] + constants.URL_END
    values = {'sort': 'top',
              't': config["frequency"]}
    data = urllib.parse.urlencode(values)
    data = data.encode('utf-8')

    req = urllib.request.Request(url, data)

    try:
        resp = urllib.request.urlopen(req)
    except urllib.HTTPError as e:
        if e.code == 429:
            time.sleep(30)
            resp = urllib.request.urlopen(req)

    images = re.findall(r'(data-url=".+?")', str(resp.read()))

    image_count = 0
    for image in images:
        image_count = image_count + 1
        image_name = os.path.join(output_folder, str(image_count) + '.jpg')
        image_urls = re.findall(r'data-url="(.+?)"', image)

        for image_url in image_urls:
            urllib.request.urlretrieve(str(image_url), image_name)


def send_images_via_email(image_folder):
    now = datetime.datetime.now()
    msg = EmailMessage()
    msg["Subject"] = config["subreddit"] + " %s" % now.strftime("%Y-%m-%d")
    msg["From"] = constants.FROM_EMAIL
    msg["To"] = config["email_to"]

    for filename in os.listdir(image_folder):
        path = os.path.join(image_folder, filename)
        if not os.path.isfile(path):
            continue
        ctype, encoding = mimetypes.guess_type(path)
        maintype, subtype = ctype.split('/', 1)
        with open(path, 'rb') as fp:
            msg.add_attachment(fp.read(), maintype=maintype, subtype=subtype, filename=filename)

    server = smtplib.SMTP(constants.EMAIL_SMTP_ADDRESS)
    server.starttls()
    server.login(constants.EMAIL_USERNAME, constants.EMAIL_PASSWORD)
    server.send_message(msg)
    server.quit()


def hmmify_job():
    now = datetime.datetime.now()
    output_folder = os.path.join(config["outputFolder"], now.strftime("%Y-%m-%d"))
    create_directory(output_folder)
    download_images(output_folder)
    send_images_via_email(output_folder)


config = load_config()

if config["frequency"] == "week":
    schedule.every().friday.at(config["time"]).do(hmmify_job)
elif config["frequency"] == "day":
    schedule.every().day.at(config["time"]).do(hmmify_job)
else:
    raise ValueError('The "frequency" value in the config .yaml file is unknown. Suitable values are "week" or "day".')

while(1):
    schedule.run_pending()
    time.sleep(1)
