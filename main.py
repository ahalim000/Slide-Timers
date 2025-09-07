#!/usr/bin/env python3

from __future__ import annotations
from typing import Tuple
import re
import uuid
import os
import click

from google.oauth2 import service_account
from googleapiclient.discovery import build

"""
Insert a video into Google Slides using a Service Account.

Prereqs
- Share the target Slides deck with the service account's email (or use domain-wide delegation).
- Enable the "Google Slides API" (and "Drive API" only if inserting Drive-hosted video) on your GCP project.
- pip install google-api-python-client google-auth
"""


# --- CONFIG ---
SERVICE_ACCOUNT_FILE = os.environ.get(
    "GOOGLE_APPLICATION_CREDENTIALS", "service-account.json"
)
# --- /CONFIG ---

TIMER_VIDEOS_MINUTES = {
    "1": "https://www.youtube.com/watch?v=OUQXwVIQw54",
    "2": "https://www.youtube.com/watch?v=R36tED8kLsU",
    "3": "https://www.youtube.com/watch?v=07T16GrtySQ",
    "4": "https://www.youtube.com/watch?v=zVHWhLme2NQ",
    "5": "https://www.youtube.com/watch?v=h6Hr8IJLpLo",
    "6": "https://www.youtube.com/watch?v=2Zdyu2lE-dM",
    "7": "https://www.youtube.com/watch?v=Em14gkKfczM",
    "8": "https://www.youtube.com/watch?v=xwgVlLn2Btc",
    "9": "https://www.youtube.com/watch?v=RA9KsHy6Ejo",
    "10": "https://www.youtube.com/watch?v=v77I_bByCDQ",
    "11": "https://www.youtube.com/watch?v=HyA4j66HLyk",
    "12": "https://www.youtube.com/watch?v=PVl4WRAThc8",
    "13": "https://www.youtube.com/watch?v=7Vfls3dJLZI",
    "14": "https://www.youtube.com/watch?v=gnk3l9uQUwc",
    "15": "https://www.youtube.com/watch?v=u_BcMXgws6Y",
    "16": "https://www.youtube.com/watch?v=OIs28WWNcjA",
    "17": "https://www.youtube.com/watch?v=TFQ4R7776zs",
    "18": "https://www.youtube.com/watch?v=tJan3T4h5ok",
    "19": "https://www.youtube.com/watch?v=38GV-w5v09g",
    "20": "https://www.youtube.com/watch?v=kxGWsHYITAw",
    "25": "https://www.youtube.com/watch?v=WctdLnMAOPg",
    "30": "https://www.youtube.com/watch?v=Xk24DMOInnQ",
}
TIMER_VIDEOS_SECONDS = {
    "10": "https://www.youtube.com/watch?v=tCDvOQI3pco",
    "15": "https://www.youtube.com/watch?v=5dlubcRwYnI",
    "20": "https://www.youtube.com/watch?v=7NWN3wivxhA",
    "25": "https://www.youtube.com/watch?v=FPElVeiASEc",
    "30": "https://www.youtube.com/watch?v=0yZcDeVsj_Y",
    "35": "https://www.youtube.com/watch?v=JbKrwBhFxLc",
    "40": "https://www.youtube.com/watch?v=jU0dZwoU4J0",
    "45": "https://www.youtube.com/watch?v=NFAUVAPWxiE",
    "50": "https://www.youtube.com/watch?v=ffdaK-kzgeg",
    "55": "https://www.youtube.com/watch?v=fE8jIMbFOVA&pp=0gcJCfwAo7VqN5tD",
    "60": "https://www.youtube.com/watch?v=OUQXwVIQw54",
}

SPEAKER_NOTES_PATTERN_MINUTES = r"((tiempo sugerido|tiempo recomendado|recomendación de tiempo|recommended time|suggested(?:\s+time)?)\s*[:\-–—]?\s*(?: (?:(?:para\s+)?diapositiva(?:s)?|for\s+slides)\s*\d+\s*[–—-]\s*\d+\s*[:\-–—]?\s* )?(\d+(?:\.\d+)?)(\s*[–—-]\s*\d+(?:\.\d+)?)?\s*(minutos|minutes|min|mins\.?)(?:\s*[-–—]?\s*(?:para\s+diapositivas|for\s+slides)\s*\d+\s*[–—-]\s*\d+)?)"
SPEAKER_NOTES_PATTERN_SECONDS = r"((tiempo sugerido|suggested time|recomendación de tiempo|recommended time)\s*[:\-]?\s*(\d+(?:\.\d+)?)(\s*[–—-]\s*\d+(?:\.\d+)?)?\s*(segundos|seconds|seg|sec|segs|secs))"
GRADE_PATTERN = r"(?:(?:Grade|Grado)\s+2|2(?:nd\s+grade|do\s+grado))"
EXIT_TICKET_PATTERN = r"(?:Demostración de Aprendizaje|Demostración del Aprendizaje|Demonstration of Learning)"
ATTRIBUTIONS_PATTERN = r"(?:Attributions|Atribuciones)"


def slides_service():
    # Gets Google Slides API
    scopes = ["https://www.googleapis.com/auth/presentations"]

    # Creates a Credentials instance from a service account json file
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=scopes
    )

    # Returns SlidesResource
    return build("slides", "v1", credentials=creds)


# Returns ID of Google Slides presentation
def get_presentation_id(presentation_url: str) -> str:
    return presentation_url.partition("/d/")[2].split("/edit")[0]


# Returns ID of YouTube video
def get_video_id(video_url: str) -> str:
    return video_url.partition("v=")[2]


# Prints a dotted line that extends to the full width of the terminal
def print_full_width_dotted_line():
    try:
        terminal_width = os.get_terminal_size().columns
        dotted_line = "." * terminal_width
        print(dotted_line)
    except OSError:
        # Handle cases where terminal size cannot be determined (e.g., not a TTY)
        print("Unable to determine terminal size. Printing a default dotted line.")
        print("." * 80)  # Print a default length line


# Returns time and unit suggested for slide in speaker notes or None if doesn't exist
def get_suggested_time_for_slide(slide) -> Tuple[int, str] | None:
    page_elements = (
        slide.get("slideProperties", {}).get("notesPage", {}).get("pageElements", [])
    )
    for page_element in page_elements:

        if "shape" not in page_element:
            continue

        shape = page_element["shape"]
        if "text" not in shape:
            continue

        text = shape["text"]
        full_text = ""
        for text_element in text.get("textElements", []):
            if "textRun" not in text_element:
                continue
            run_data = text_element["textRun"]
            if "content" not in run_data:
                continue
            content = run_data["content"]
            full_text += content
        # print(full_text)
        minutes_match = re.search(
            SPEAKER_NOTES_PATTERN_MINUTES, full_text, re.IGNORECASE
        )
        seconds_match = re.search(
            SPEAKER_NOTES_PATTERN_SECONDS, full_text, re.IGNORECASE
        )

        if minutes_match is not None:
            return int(minutes_match.group(3)), "minutes"

        if seconds_match is not None:
            return int(seconds_match.group(3)), "seconds"

    return None


"""
Returns true when slides intended for teacher at beginning of slide deck
are over and slides intended for students have begun.
"""


def is_presentation_started(slide) -> bool:
    for page_element in slide.get("pageElements", []):
        text_elements = (
            page_element.get("shape", {}).get("text", {}).get("textElements", [])
        )
        full_text = ""
        for text_element in text_elements:
            if "textRun" not in text_element:
                continue
            run_data = text_element["textRun"]
            if "content" not in run_data:
                continue
            content = run_data["content"]
            full_text += content
        match = re.search(GRADE_PATTERN, full_text, re.IGNORECASE)
        if match:
            return True

    return False


# Returns total number of slides in presentation
def get_total_slides_in_pres(slides):
    total = 0
    for _ in slides:
        total += 1

    return total


# Determines whether a slide is an exit ticket
def is_exit_ticket(slides, slide, slide_number) -> bool:
    total_slides = get_total_slides_in_pres(slides)

    for page_element in slide.get("pageElements", []):
        text_elements = (
            page_element.get("shape", {}).get("text", {}).get("textElements", [])
        )
        full_text = ""
        for text_element in text_elements:
            if "textRun" not in text_element:
                continue
            run_data = text_element["textRun"]
            if "content" not in run_data:
                continue
            content = run_data["content"]
            full_text += content
        match = re.search(EXIT_TICKET_PATTERN, full_text, re.IGNORECASE)
        if match and slide_number == total_slides - 1:
            return True

    return False


# Determines whether we're on the last slide (attributions slide)
def is_presentation_ended(slide) -> bool:
    for page_element in slide.get("pageElements", []):
        text_elements = (
            page_element.get("shape", {}).get("text", {}).get("textElements", [])
        )
        full_text = ""
        for text_element in text_elements:
            if "textRun" not in text_element:
                continue
            run_data = text_element["textRun"]
            if "content" not in run_data:
                continue
            content = run_data["content"]
            full_text += content
        match = re.search(ATTRIBUTIONS_PATTERN, full_text, re.IGNORECASE)
        if match:
            return True

    return False


# Returns slide number
def get_slide_number(slide):
    return int(slide["objectId"].lstrip("p"))


# Adds timer videos to slides
def add_videos(svc, pres, presentation_id):
    slides = pres.get("slides", [])
    presentation_started = False
    presentation_ended = False

    inserted_count = 0

    slide_suggested_time = None

    slide_duration = 2  # default time is 2 minutes
    time_unit = "minutes"

    for slide in slides:
        if "objectId" not in slide:
            continue

        slide_number = get_slide_number(slide)

        print(f"Slide {slide_number}")

        if not presentation_started:
            print("\nPresentation not started\n")
            print_full_width_dotted_line()

            presentation_started = is_presentation_started(slide)
            continue

        if presentation_ended or is_presentation_ended(slide):
            print("\nPresentation ended\n")
            continue

        # Check if slide is an exit ticket
        exit_ticket = is_exit_ticket(slides, slide, slide_number)

        # Get suggested time
        resp = get_suggested_time_for_slide(slide)

        # Exit ticket slides always last 10 minutes
        if exit_ticket:
            slide_duration = 10
            print(f"\nExit ticket: suggested time is 10 minutes")

        # If slide has suggested time, set suggested time and unit
        elif resp is not None:
            slide_suggested_time = int(resp[0])
            time_unit = resp[1]

        # If slide has suggested time, set it as the slide duration
        if slide_suggested_time is not None:
            slide_duration = slide_suggested_time

        page_id = slide["objectId"]
        video_object_id = f"video_{uuid.uuid4().hex[:8]}"

        if time_unit == "seconds":
            video_url = TIMER_VIDEOS_SECONDS[str(slide_duration)]
        else:
            video_url = TIMER_VIDEOS_MINUTES[str(slide_duration)]

        if video_url:
            video_id = get_video_id(video_url)

            # 16:9 box 80x45 pt, positioned near the bottom-right
            create_video = {
                "createVideo": {
                    "objectId": video_object_id,
                    "elementProperties": {
                        "pageObjectId": page_id,
                        "size": {
                            "width": {"magnitude": 80, "unit": "PT"},
                            "height": {"magnitude": 45, "unit": "PT"},
                        },
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": 630,  # Higher number is farther to right on X axis
                            "translateY": 350,  # Higher number is lower on Y axis
                            "unit": "PT",
                        },
                    },
                    "source": "YOUTUBE",
                    "id": video_id,
                }
            }

            # Tweak playback properties (autoplay, start/end, mute)
            update_props = {
                "updateVideoProperties": {
                    "objectId": video_object_id,
                    "videoProperties": {
                        "autoPlay": True,  # Make video autoplay in presentation mode
                    },
                    "fields": "autoPlay",
                }
            }

            resp = (
                svc.presentations()
                .batchUpdate(
                    presentationId=presentation_id,
                    body={"requests": [create_video, update_props]},
                )
                .execute()
            )

            if slide_suggested_time is None and not exit_ticket:
                print("\nSuggested time is None")
            elif slide_suggested_time is not None and not exit_ticket:
                print(f"\nSuggested time is {slide_suggested_time} {time_unit}")

            if time_unit == "seconds":
                print(
                    f"Inserted {slide_duration} second timer on slide {slide_number}\n"
                )
            else:
                print(
                    f"Inserted {slide_duration} minute timer on slide {slide_number}\n"
                )

            print_full_width_dotted_line()

            inserted_count += 1
            slide_suggested_time = None
            slide_duration = 2

            if exit_ticket:
                presentation_ended = True

    print(f"Added {inserted_count} video(s) to presentation {presentation_id}")


# Delete all videos from Google Slides presentation
def delete_videos(svc, presentation_id):
    presentation = (
        svc.presentations()
        .get(
            presentationId=presentation_id, fields="slides/pageElements(objectId,video)"
        )
        .execute()
    )

    requests = []
    for slide in presentation.get("slides", []):
        for page_element in slide.get("pageElements", []):
            if "video" in page_element:  # page element is a video
                requests.append(
                    {"deleteObject": {"objectId": page_element["objectId"]}}
                )

    if requests != []:
        svc.presentations().batchUpdate(
            presentationId=presentation_id, body={"requests": requests}
        ).execute()

    print(f"Deleted {len(requests)} video(s) from presentation {presentation_id}")


# ------------------------------------------Command Line Interface------------------------------------------
@click.group()
def cli():
    pass


@cli.command()
@click.argument("presentation-url")
def add(presentation_url: str):
    # SlidesResource object for interacting with Google Slides API
    svc = slides_service()

    # Extract ID from presentation_url
    presentation_id = get_presentation_id(presentation_url)

    pres = svc.presentations().get(presentationId=presentation_id).execute()

    add_videos(svc, pres, presentation_id)


@cli.command()
@click.argument("presentation-url")
def delete(presentation_url: str):
    # SlidesResource object for interacting with Google Slides API
    svc = slides_service()

    # Extract ID from presentation_url
    presentation_id = get_presentation_id(presentation_url)

    delete_videos(svc, presentation_id)


if __name__ == "__main__":
    cli()
