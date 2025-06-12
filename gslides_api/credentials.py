import logging
import os
from typing import Optional

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import Resource, build

# import gslides
logger = logging.getLogger(__name__)


class Creds:
    # Initial version from the gslides package
    """The credentials object to build the connections to the APIs"""

    def __init__(self) -> None:
        """Constructor method"""
        self.crdtls: Optional[Credentials] = None
        self.sht_srvc: Optional[Resource] = None
        self.sld_srvc: Optional[Resource] = None
        self.drive_srvc: Optional[Resource] = None

    def set_credentials(self, credentials: Optional[Credentials]) -> None:
        """Sets the credentials

        :param credentials: :class:`google.oauth2.credentials.Credentials`
        :type credentials: :class:`google.oauth2.credentials.Credentials`

        """
        self.crdtls = credentials
        logger.info("Building sheets connection")
        self.sht_srvc = build("sheets", "v4", credentials=credentials)
        logger.info("Built sheets connection")
        logger.info("Building slides connection")
        self.sld_srvc = build("slides", "v1", credentials=credentials)
        logger.info("Built slides connection")
        logger.info("Building drive connection")
        self.drive_srvc = build("drive", "v3", credentials=credentials)
        logger.info("Built drive connection")

    @property
    def sheet_service(self) -> Resource:
        """Returns the connects to the sheets API

        :raises RuntimeError: Must run set_credentials before executing method
        :return: API connection
        :rtype: :class:`googleapiclient.discovery.Resource`
        """
        if self.sht_srvc:
            return self.sht_srvc
        else:
            raise RuntimeError("Must run set_credentials before executing method")

    @property
    def slide_service(self) -> Resource:
        """Returns the connects to the slides API

        :raises RuntimeError: Must run set_credentials before executing method
        :return: API connection
        :rtype: :class:`googleapiclient.discovery.Resource`
        """
        if self.sht_srvc:
            return self.sld_srvc
        else:
            raise RuntimeError("Must run set_credentials before executing method")

    @property
    def drive_service(self) -> Resource:
        """Returns the connects to the drive API

        :raises RuntimeError: Must run set_credentials before executing method
        :return: API connection
        :rtype: :class:`googleapiclient.discovery.Resource`
        """
        if self.drive_srvc:
            return self.drive_srvc
        else:
            raise RuntimeError("Must run set_credentials before executing method")


creds = Creds()


def initialize_credentials(credential_location: str):
    """

    :param credential_location:
    :return:
    """

    SCOPES = [
        "https://www.googleapis.com/auth/presentations",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    _creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(credential_location + "token.json"):
        _creds = Credentials.from_authorized_user_file(credential_location + "token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not _creds or not _creds.valid:
        if _creds and _creds.expired and _creds.refresh_token:
            _creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credential_location + "credentials.json", SCOPES
            )
            _creds = flow.run_local_server()
        # Save the credentials for the next run
        with open(credential_location + "token.json", "w") as token:
            token.write(_creds.to_json())
    creds.set_credentials(_creds)
