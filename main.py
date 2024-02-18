import imaplib
import ssl
import email
from email.header import decode_header
from bs4 import BeautifulSoup
from colorama import init, Fore, Style
import ctypes
import time

init(autoreset=True)


def set_console_size(width, height):
    try:
        ctypes.windll.kernel32.SetConsoleScreenBufferSize(ctypes.windll.kernel32.GetStdHandle(-11), width * 100 + height)
        ctypes.windll.kernel32.SetConsoleWindowInfo(ctypes.windll.kernel32.GetStdHandle(-11), True, ctypes.byref(ctypes.wintypes.SMALL_RECT(0, 0, width - 1, height - 1)))
    except Exception as e:
        print(f"Ошибка при установке размера консоли: {e}")

set_console_size(400, 120)

init(autoreset=True)

class EmailClient:
    def __init__(self, host, port, username=None, password=None):
        self.host = host
        self.port = port
        self.username = username
        self.password = password
        self.mail = None
        self.connected = False  # Добавляем флаг для проверки состояния подключения

    def connect(self):
        try:
            self.mail = imaplib.IMAP4_SSL(self.host, self.port)
            self.mail.login(self.username, self.password)
            if not self.connected:
                print(f"{Fore.GREEN}Успешное подключение к почтовому серверу.{Style.RESET_ALL}")
                self.connected = True
        except Exception as e:
            print(f"Ошибка подключения: {e}")
            self.connected = False  # Сбрасываем состояние подключения

    def change_credentials(self, new_username, new_password):
        self.username = new_username
        self.password = new_password

    def find_email_by_subject(self, subject):
        try:
            self.mail.select('inbox')
            _, data = self.mail.search(None, f'(SUBJECT "{subject}")')
            return data[0].split()
        except Exception as e:
            print(f"Ошибка при поиске почты: {e}")
            return []

    def fetch_email_content(self, email_id):
        try:
            _, msg_data = self.mail.fetch(email_id, '(RFC822)')
            msg = email.message_from_bytes(msg_data[0][1])
            return msg
        except Exception as e:
            print(f"Ошибка при получении содержимого письма: {e}")
            return None

    def extract_verification_codes(self, msg):
        try:
            verification_codes = []

            if msg.is_multipart():
                for part in msg.walk():
                    # Игнорируем вложения
                    if part.get_content_maintype() == 'multipart':
                        continue

                    payload = part.get_payload(decode=True)
                    soup = BeautifulSoup(payload, 'html.parser')

                    # Находим все элементы <td> с указанными стилями
                    td_elements = soup.find_all('td', {'style': 'background:#f1f1f1;margin-top:20px;font-family: arial,helvetica,sans-serif; mso-line-height-rule: exactly; font-size:30px; color:#202020; line-height:19px; line-height: 134%; letter-spacing: 10px;text-align: center;padding: 20px 0px !important;letter-spacing: 10px !important;border-radius: 4px;'})

                    # Извлекаем текст из всех подходящих элементов
                    verification_codes.extend([td.get_text(strip=True).strip() for td in td_elements])

            return verification_codes
        except Exception as e:
            print(f"Ошибка при извлечении кодов верификации: {e}")
            return []

    def disconnect(self):
        try:
            if self.mail:
                self.mail.logout()
                print("Отключение от почтового сервера.")
        except Exception as e:
            print(f"Ошибка при отключении от почтового сервера: {e}")

def read_settings(file_path):
    settings = {}
    try:
        with open(file_path, 'r') as file:
            for line in file:
                key, value = map(str.strip, line.strip().split(':'))
                settings[key] = value
    except Exception as e:
        print(f"Ошибка при чтении настроек из файла: {e}")

    return settings

if __name__ == "__main__":
    while True:
        settings = read_settings('settings.txt')
        host = settings.get('host', 'имя_почтового_сервера')
        port = int(settings.get('port', 993))

        credentials = input("Введите логин и пароль от почты в формате 'логин:пароль': ")
        username, password = credentials.split(':')

        client = EmailClient(host, port, username, password)
        client.connect()

        target_subject = "Epic Games - Email Verification"

        email_ids = client.find_email_by_subject(target_subject)
        if email_ids:
            email_id = email_ids[-1]
            msg = client.fetch_email_content(email_id)
            if msg:
                print(f"{Fore.GREEN}Успешное подключение к почтовому серверу.{Style.RESET_ALL}")
                print(f"Message Subject: {msg['Subject']}")
                print(f"Message Body:\n{msg.get_payload()}")
                verification_codes = client.extract_verification_codes(msg)
                if verification_codes:
                    verification_codes_string = ', '.join(verification_codes)
                    print(f"{Fore.YELLOW}Verification Codes: {verification_codes_string}{Style.RESET_ALL}")
                else:
                    print("Коды верификации не найдены в письме.")
            else:
                print("Ошибка при получении содержимого письма.")
        else:
            print(f"Письмо с темой {target_subject} не найдено.")

        client.disconnect()

        # Задержка перед запросом новых данных
        time.sleep(60)  # 60 секунд задержки