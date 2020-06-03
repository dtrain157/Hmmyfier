import urllib.request
import urllib.parse
import constants
import datetime
import logging
import mimetypes
import os
import re
import schedule
import smtplib
import time
import yaml
import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import pathlib

from email.message import EmailMessage


class AppServerSvc (win32serviceutil.ServiceFramework):
    _svc_name_ = "Hmmmyfier"
    _svc_display_name_ = "Hmmmyfier"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main()

    def load_config(self):
        os.chdir(pathlib.Path(__file__).parent.absolute())
        return yaml.safe_load(open(constants.CONFIG_FILENAME))

    def create_directory(self, directory):
        if (not os.path.exists(directory)):
            os.mkdir(directory)

    def download_images(self, output_folder):
        url = constants.URL_BASE + self._config["subreddit"] + constants.URL_END
        values = {'sort': 'top', 't': self._config["frequency"]}
        data = urllib.parse.urlencode(values)
        data = data.encode('utf-8')

        req = urllib.request.Request(url, data)

        trycounter = 3
        shouldTryAgain = True
        while (shouldTryAgain) & (trycounter > 0):
            try:
                resp = urllib.request.urlopen(req)
                shouldTryAgain = False
            except urllib.error.HTTPError as e:
                if e.code == 429:
                    time.sleep(30)
                    trycounter = trycounter - 1
                    shouldTryAgain = True
                else:
                    shouldTryAgain = False
                    self._logger.error(e)
            except Exception as e:
                self._logger.error(e)

        images = re.findall(r'(data-url=".+?")', str(resp.read()))

        image_count = 0
        for image in images:
            image_count = image_count + 1
            image_name = os.path.join(output_folder, str(image_count) + '.jpg')
            image_urls = re.findall(r'data-url="(.+?)"', image)

            for image_url in image_urls:
                urllib.request.urlretrieve(str(image_url), image_name)

    def send_images_via_email(self, image_folder):
        now = datetime.datetime.now()
        msg = EmailMessage()
        msg["Subject"] = self._config["subreddit"] + " %s" % now.strftime("%Y-%m-%d")
        msg["From"] = constants.FROM_EMAIL
        msg["To"] = self._config["email_to"]

        for filename in os.listdir(image_folder):
            path = os.path.join(image_folder, filename)
            if not os.path.isfile(path):
                continue
            ctype, _ = mimetypes.guess_type(path)
            maintype, subtype = ctype.split('/', 1)
            with open(path, 'rb') as fp:
                msg.add_attachment(fp.read(), maintype=maintype, subtype=subtype, filename=filename)

        server = smtplib.SMTP(constants.EMAIL_SMTP_ADDRESS)
        server.starttls()
        server.login(constants.EMAIL_USERNAME, constants.EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()

    def hmmify_job(self):
        now = datetime.datetime.now()
        output_folder = os.path.join(
            self._config["outputFolder"], now.strftime("%Y-%m-%d"))
        self.create_directory(output_folder)
        self.download_images(output_folder)
        self.send_images_via_email(output_folder)

    def main(self):
        logging.basicConfig(filename='/temp/hmmyfier.log', level=logging.ERROR, format='%(asctime)s %(levelname)s %(name)s %(message)s')
        self._logger = logging.getLogger(__name__)

        self._config = self.load_config()

        if self._config["frequency"] == "week":
            schedule.every().friday.at(self._config["time"]).do(self.hmmify_job)
        elif self._config["frequency"] == "day":
            schedule.every().day.at(self._config["time"]).do(self.hmmify_job)
        else:
            raise ValueError('The "frequency" value in the config .yaml file is unknown. Suitable values are "week" or "day".')

        rc = None
        while(rc != win32event.WAIT_OBJECT_0:
            schedule.run_pending()
            time.sleep(1)
            rc = win32event.WaitForSingleObject(self.hWaitStop, 24 * 60 * 60 * 1000)


if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(AppServerSvc)
