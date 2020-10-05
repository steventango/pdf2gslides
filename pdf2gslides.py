import os
import pathlib
import pickle
import random
import shutil
import subprocess
from threading import Timer
from typing import Tuple
import numpy as np
from googleapiclient.discovery import build, Resource
from googleapiclient.errors import HttpError
from googleapiclient.http import HttpRequest, MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ['https://www.googleapis.com/auth/drive.file']


def gdriveauth() -> Tuple[Resource, Resource]:
    """Authenticate with Google drive, maintain credentials and returns Drive
    and Slides services.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return build('drive', 'v3', credentials=creds), \
        build('slides', 'v1', credentials=creds)


def exponentialBackoff(
        request: HttpRequest,
        time: int = 0,
        tries: int = 0,
        max_tries: int = 8):
    """Periodically retry a failed request over an increasing amount of time to
    handle errors related to rate limits, network volume, or response time.
    https://developers.google.com/drive/api/v3/handle-errors
    """
    try:
        response = None
        while response is None:
            status, response = request.next_chunk()
            if status:
                print(status.resumable_progress / status.total_size * 100)
        file = request.execute()
        return file
    except HttpError as e:
        print(e)
        if tries < max_tries:
            t = Timer(
                time,
                exponentialBackoff,
                args=(
                    request,
                    (tries + 1) * 2 + random.random(),
                    tries + 1,
                    max_tries))
            t.start()
        else:
            raise Exception('Max tries reached.')


def odp2gslides(drive_service: Resource, slides_service: Resource):
    """Upload to Google drive and convert odp files to Google slides.
    """
    tempdir = os.fsencode('temp')
    for file in os.listdir(tempdir):
        filename = os.fsdecode(file)
        if filename.endswith('.odp'):
            file_metadata = {
                'name': filename.rstrip('.odp'),
                'mimeType': 'application/vnd.google-apps.presentation'
            }
            media = MediaFileUpload(
                f'temp/{filename}',
                mimetype='application/vnd.oasis.opendocument.presentation',
                chunksize=256 * 1024,
                resumable=True)
            request = drive_service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id')
            gslide = exponentialBackoff(request)

            # Close file after upload is done
            media.stream().close()

            # Fix text box formatting
            gslide_id = gslide['id']
            print(f'Uploaded {filename}, id: {gslide_id}')
            fixformatgslide(slides_service, gslide_id)


def fixformatgslide(service: Resource, pid: str):
    """Use Google Slides API to fix the slightly too small widths of the text
    boxes formatting caused by the Google Drive conversion from .odf to
    .gslides.
    """
    requests = []
    presentation = service.presentations().get(
        presentationId=pid).execute()

    for slide in presentation['slides']:
        for pageElement in slide['pageElements']:
            if 'shape' in pageElement:
                if pageElement['shape']['shapeType'] == 'TEXT_BOX':
                    # Precompute transformation with element reference frame
                    # https://developers.google.com/slides/how-tos/transform#element_reference_frames
                    A = pageElement['transform']
                    T2 = np.array([[1, 0, A['translateX']],
                                   [0, 1, A['translateY']],
                                   [0, 0, 1]])
                    SCALE_X = 1.1
                    B = np.array([[SCALE_X, 0, 0], [0, 1, 0], [0, 0, 1]])
                    T1 = np.array([[1, 0, -A['translateX']],
                                   [0, 1, -A['translateY']],
                                   [0, 0, 1]])

                    A_prime = np.matmul(np.matmul(T2, B), T1)
                    requests.append({
                        'updatePageElementTransform': {
                            'objectId': pageElement['objectId'],
                            'transform': {
                                'scaleX': A_prime[0][0],
                                'scaleY': A_prime[1][1],
                                'translateX': A_prime[0][2],
                                'translateY': A_prime[1][2],
                                'unit': 'EMU'
                            },
                            'applyMode': 'RELATIVE'
                        }
                    })

    body = {
        'requests': requests
    }
    service.presentations().batchUpdate(
        presentationId=pid,
        body=body).execute()
    print(f'Fixed GSlides formatting, id: {pid}')


def pdf2odp():
    """Convert PDF file to ODP intermediate using LibreOffice command:
    soffice --convert-to
    """
    if not os.path.exists('in'):
        os.makedirs('in')
    if not os.path.exists('temp'):
        os.makedirs('temp')

    indir = os.fsencode('in')
    for file in os.listdir(indir):
        filename = os.fsdecode(file)
        if filename.endswith('.pdf'):
            # https://ask.libreoffice.org/en/question/185693/pdf-to-ppt-and-excel/
            subprocess.call([
                'soffice',
                '--infilter=impress_pdf_import',
                '--convert-to',
                'odp:impress8',
                '--outdir',
                str(pathlib.Path('./temp/').absolute()),
                f'in/{filename}'],
                executable='C:\\Program Files\\LibreOffice\\program\\soffice.bin',
            )
            print(f'Converted {filename} to odp')


def main():
    drive_service, slides_service = gdriveauth()
    pdf2odp()
    odp2gslides(drive_service, slides_service)
    shutil.rmtree('temp')


if __name__ == '__main__':
    main()
