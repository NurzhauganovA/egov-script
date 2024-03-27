import datetime

from cryptography.hazmat._oid import NameOID
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
import qrcode
from xml.etree import ElementTree as ET
from io import BytesIO
from cryptography.hazmat.primitives.serialization import pkcs12
import base64


class SignContractWithEDSService:
    """
    Подпись договора с помощью ЭЦП
    Генерировать QR-код с данными подписи
    Использовать логику из документации ncanode.kz
    """

    @staticmethod
    def get_certificate() -> bytes:
        """ Получить сертификат в формате .p12 """

        try:
            certificate = open("eds/AUTH_RSA256_f76176815e5e30559813d90ab8aff7b0219b1795.p12", "rb")
        except KeyError:
            raise ValueError("Необходимо прикрепить сертификат для подписи договора")

        if certificate.name.startswith('AUTH'):
            raise ValueError("Вы должны прикрепить RSA сертификат для подписи договора")

        elif not certificate.name.endswith('.p12'):
            raise ValueError("Сертификат должен быть в формате .p12")

        return certificate.read()

    @staticmethod
    def check_verify_of_certificate(before_time, after_time):
        """ Проверка сертификата на валидность """

        date_today = datetime.datetime.now().date()
        if before_time.date() > date_today or after_time.date() < date_today:
            raise ValueError("Срок действия сертификата истек")
        else:
            return True

    @staticmethod
    def notification_user_about_certificate_expiration(before_time):
        """
            Уведомление пользователя о скором истечении срока действия сертификата.
            Уведомим пользователя за 2 недели до истечения срока действия сертификата
        """

        date_today = datetime.datetime.now().date()
        expiration_date = (before_time.date() - date_today).days

        if expiration_date <= 14:
            notification_text = f"Хотим уведомить вас о том, что срок действия вашего сертификата истечет через {expiration_date} дней"
            return notification_text

    def get_private_key(self, certificate_data, password: str):
        """ Получение закрытого ключа из сертификата в формате .p12 """

        try:
            private_key = pkcs12.load_key_and_certificates(certificate_data, password.encode())
            # print(private_key[1].signature)
            self.check_verify_of_certificate(private_key[1].not_valid_before_utc, private_key[1].not_valid_after_utc)
            return private_key
        except Exception as e:
            raise ValueError(f'{e}')

    @staticmethod
    def get_signature(private_key, hash_key):
        """ Получение подписи договора """
        signature = private_key.sign(
            hash_key,
            padding.PSS(
                mgf=padding.MGF1(hashes.SHA256()),
                salt_length=padding.PSS.MAX_LENGTH
            ),
            hashes.SHA256()
        )
        return signature

    @staticmethod
    def get_signature_base64(signature):
        """ Получение подписи договора в формате base64 """
        signature_base64 = base64.b64encode(signature)
        return signature_base64

    @staticmethod
    def data_to_xml(data):
        root = ET.Element("data")

        for key, value in data.items():
            node = ET.SubElement(root, key)
            node.text = value

        xml_string = ET.tostring(root, encoding='utf-8')
        return xml_string

    def get_data(self, certificate, password):
        """ Получение данных для подписи договора """

        private_key = self.get_private_key(certificate, password)
        public_key = private_key[1]
        subject = public_key.subject
        issuer = public_key.issuer

        data = dict(VALID="TRUE",
                    SERIAL_NUMBER=str(public_key.serial_number),
                    SIGNATURE_ALGORITHM=str(public_key.signature_algorithm_oid._name),
                    SIGNATURE=str(self.get_signature_base64(self.get_signature(private_key[0], public_key.fingerprint(hashes.SHA256()))).decode()),
                    COMMON_NAME=str(subject.get_attributes_for_oid(NameOID.COMMON_NAME)[0].value),
                    INN=str(subject.get_attributes_for_oid(NameOID.SERIAL_NUMBER)[0].value),
                    ISSURE_COMMON_NAME=str(issuer.get_attributes_for_oid(NameOID.COMMON_NAME)[0].value),
                    NOTIFICATION_TEXT=self.notification_user_about_certificate_expiration(public_key.not_valid_after_utc))
        return data
