import os.path
from time import sleep

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/presentations"]

presentation_id = '12TkfhLj53aBA1ed0Yn_o3EfayLTXFOdy5GLZORVWdtI'
pages = []
images = {
    "Pagina_1": "https://media.licdn.com/dms/image/C4D03AQFPb4-AKqTeFw/profile-displayphoto-shrink_800_800/0/1589309927264?e=1710979200&v=beta&t=2AYu-5JZ_zhdHZl58z1VJDPRbxcDgQQtGYifZFig5t0",
    "Pagina_2": "https://freegoogleslidestemplates.com/wp-content/uploads/2016/12/Screenshot-2016-12-01-15.27.57.png",
    "Pagina_3": "https://shops-united.nl/wp-content/uploads/2021/06/hurby-pakket-versturen.png",
    "Pagina_4": "https://nl.devoteam.com/wp-content/uploads/sites/13/2021/05/API-management-lifecycle.png",
    "Pagina_5": "https://nl.devoteam.com/wp-content/uploads/sites/13/2023/03/pexels-ann-h-3482441-1296x864.jpg",
    "Pagina_6": "https://redis.com/wp-content/uploads/2023/10/Devoteam_cmyk.png?&auto=webp&quality=85,75&width=500",
}


def create_slide(presentation_id, page_id, insertion_index):
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    try:
        service = build("slides", "v1", credentials=creds)
        # Add a slide at index 1 using the predefined
        # 'BLANK' layout and the ID page_id.
        requests = [
            {
                "createSlide": {
                    "objectId": page_id,
                    "insertionIndex": insertion_index,
                    "slideLayoutReference": {
                        "predefinedLayout": "BLANK"
                    },
                },
            },
        ]

        # If you wish to populate the slide with elements,
        # add element create requests here, using the page_id.

        # Execute the request.
        body = {"requests": requests}
        response = (
            service.presentations()
            .batchUpdate(presentationId=presentation_id, body=body)
            .execute()
        )
        create_slide_response = response.get("replies")[0].get("createSlide")
        print(f"Created slide with ID:{(create_slide_response.get('objectId'))}")
    except HttpError as error:
        print(f"An error occurred: {error}")
        print("Slides not created")
        return error

    return response


def update_slide(presentation_id, page_id):
    element_id = 'IntroPageJur'
    image_id = "Pagina_1"
    IMAGE_URL = images.get('Pagina_1')
    pt350 = {"magnitude": 350, "unit": "PT"}
    emu4M = {"magnitude": 4000000, "unit": "EMU"}

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    try:
        service = build("slides", "v1", credentials=creds)
        requests = [
            {
                "createShape": {
                    "objectId": element_id,
                    "shapeType": "TEXT_BOX",
                    "elementProperties": {
                        "pageObjectId": page_id,
                        "size": {"height": pt350, "width": pt350},
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": 100,
                            "translateY": 20,
                            "unit": "PT",
                        },
                    },
                }
            },
            # Insert text into the box, using the supplied element ID.
            {
                "insertText": {
                    "objectId": element_id,
                    "insertionIndex": 0,
                    "text": "Presentatie DEVO TEAM\nJurriaan den Uyl",
                },
            },
            {
                "updateTextStyle": {
                    "objectId": element_id,
                    "textRange": {
                        "type": "FIXED_RANGE",
                        "startIndex": 0,
                        "endIndex": 21,
                    },
                    "style": {
                        "bold": True, "italic": True,
                        "foregroundColor": {
                            "opaqueColor": {
                                "rgbColor": {
                                    "blue": 94 / 255,
                                    "green": 72 / 255,
                                    "red": 248 / 255,
                                }
                            }
                        },
                    },
                    "fields": "bold,italic,foregroundColor",
                }
            },
            {
                "createImage": {
                    "objectId": image_id,
                    "url": IMAGE_URL,
                    "elementProperties": {
                        "pageObjectId": page_id,
                        "size": {"height": emu4M, "width": emu4M},
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": 1400000,
                            "translateY": 1000000,
                            "unit": "EMU",
                        },
                    },
                }
            },
        ]

        # Execute the request.
        body = {"requests": requests}
        response = (
            service.presentations()
            .batchUpdate(presentationId=presentation_id, body=body)
            .execute()
        )
        print("All systems go")
    except HttpError as error:
        print(f"An error occurred: {error}")
        print("Slides not created")
        return error

    return response


def update_slide2(presentation_id, page_id, element_image_id, main_text):
    element_id = element_image_id
    image_id = f"Image_%s" % element_image_id
    IMAGE_URL = images.get(element_image_id)
    pt350 = {"magnitude": 350, "unit": "PT"}
    pt550 = {"magnitude": 550, "unit": "PT"}
    emu4M = {"magnitude": 4000000, "unit": "EMU"}

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    try:
        service = build("slides", "v1", credentials=creds)
        requests = [
            {
                "createShape": {
                    "objectId": element_id,
                    "shapeType": "TEXT_BOX",
                    "elementProperties": {
                        "pageObjectId": page_id,
                        "size": {"height": pt350, "width": pt550},
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": 100,
                            "translateY": 20,
                            "unit": "PT",
                        },
                    },
                }
            },
            # Insert text into the box, using the supplied element ID.
            {
                "insertText": {
                    "objectId": element_id,
                    "insertionIndex": "0",
                    "text": f"Presentatie DEVO TEAM\n%s" % main_text,
                },
            },
            {
                "updateTextStyle": {
                    "objectId": element_id,
                    "textRange": {
                        "type": "FIXED_RANGE",
                        "startIndex": 0,
                        "endIndex": 21,
                    },
                    "style": {
                        "bold": True, "italic": True,
                        "foregroundColor": {
                            "opaqueColor": {
                                "rgbColor": {
                                    "blue": 94/255,
                                    "green": 72/255,
                                    "red": 248/255,
                                }
                            }
                        },
                    },
                    "fields": "bold,italic,foregroundColor",
                }
            },
            {
                "createImage": {
                    "objectId": image_id,
                    "url": IMAGE_URL,
                    "elementProperties": {
                        "pageObjectId": page_id,
                        "size": {"height": emu4M, "width": emu4M},
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": 1400000,
                            "translateY": 1000000,
                            "unit": "EMU",
                        },
                    },
                }
            },
        ]

        # If you wish to populate the slide with elements,
        # add element create requests here, using the page_id.

        # Execute the request.
        body = {"requests": requests}
        response = (
            service.presentations()
            .batchUpdate(presentationId=presentation_id, body=body)
            .execute()
        )
        print("All systems go")
    except HttpError as error:
        print(f"An error occurred: {error}")
        print("Slides not created")
        return error

    return response


def delete_slide(presentation_id, page_id):
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    try:
        service = build("slides", "v1", credentials=creds)
        requests = [
            {
                "deleteObject": {
                    "objectId": page_id,
                },
            },
        ]

        # Execute the request.
        body = {"requests": requests}
        response = (
            service.presentations()
            .batchUpdate(presentationId=presentation_id, body=body)
            .execute()
        )
        print("Deleted Slide")
    except HttpError as error:
        print(f"An error occurred: {error}")
        print("Could not delete Slide")
        return error

    return response


def main():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

if __name__ == "__main__":
    main()
    create_slide(presentation_id, "My_first_slide", 0)
    update_slide(presentation_id, "My_first_slide")
    input("Press Enter to continue...")
    delete_slide(presentation_id, "My_first_slide")
    create_slide(presentation_id, "My_second_slide", 0)
    update_slide2(presentation_id, "My_second_slide", "Pagina_2", "Gemaakt in Python, "
                                                                  "via de Google Slides API")
    input("Press Enter to continue...")
    delete_slide(presentation_id, "My_second_slide")
    create_slide(presentation_id, "My_third_slide", 0)
    update_slide2(presentation_id, "My_third_slide", "Pagina_3",
                  "Co-Founder Hurby, een tech gedreven logistieke startup\n-Product Owner en project manager"
                  " voor dev team Sri Lanka\n-Project leider voor meerdere API integraties met partners")
    input("Press Enter to continue...")
    delete_slide(presentation_id, "My_third_slide")
    create_slide(presentation_id, "My_fourth_slide", 0)
    update_slide2(presentation_id, "My_fourth_slide", "Pagina_4", "Veel met IT gedaan"
                  " maar geen formele opleiding\nIk hoop een beter idee van IT te krijgen"
                                                                  ", en hierdoor efficienter te werken")
    input("Press Enter to continue...")
    delete_slide(presentation_id, "My_fourth_slide")
    create_slide(presentation_id, "My_fifth_slide", 0)
    update_slide2(presentation_id, "My_fifth_slide", "Pagina_5", "Ambitie: Werken in de IT"
                                                                 " met voorkeur voor overheid/duurzame bedrijven")
    input("Press Enter to continue...")
    delete_slide(presentation_id, "My_fifth_slide")
    create_slide(presentation_id, "My_sixth_slide", 0)
    update_slide2(presentation_id, "My_sixth_slide", "Pagina_6", "Nieuwsgierig werken")
    input("Press Enter to continue...")
    delete_slide(presentation_id, "My_sixth_slide")
