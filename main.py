#!/usr/bin/env python3

from __future__ import annotations
from typing import Tuple
import math
import re
import uuid
import json
import os
import sys
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
# PRESENTATION_ID = "1D8_iOt26vkyAZhs_Ec29BbRbUreNX41lwggwN6idrc"  # from https://docs.google.com/presentation/d/<ID>/edit
# --- /CONFIG ---

TIMER_VIDEOS = {
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
    "15": "https://www.youtube.com/watch?v=u_BcMXgws6Y",
    "20": "https://www.youtube.com/watch?v=kxGWsHYITAw",
    "25": "https://www.youtube.com/watch?v=WctdLnMAOPg",
}


RANGE_PATTERN = r"Diapositivas\s+(?P<start>\d+)\s*[–—-]\s*(?P<end>\d+)\s+Para esta lección,\s*planifique aproximadamente:\s*(?P<minutes>\d+)\s+minutos"
SPEAKER_NOTES_PATTERN = r"((tiempo sugerido|suggested time)\s*[:\-]?\s*(\d+(?:\.\d+)?)(\s*-\s*\d+(?:\.\d+)?)?\s*(minutos|minutes|min)?)"
UNWANTED_PATTERN = r".*\bpara\s+diapositivas\b[\s\S]*"
# TIME_PATTERN = r"(?:[0-5]?\d):[0-5]\d"
# SPEAKER_NOTES_PATTERN = r"^(?:Suggested\s+[Tt]ime|Tiempo\s+sugerido):\s*(?P<start>\d+)(?:\s*[–—-]\s*(?P<end>\d+))?\s*(?:min(?:s)?\.?|minutes?)\s*$"


def slides_service():
    # Gets Google Slides API
    scopes = ["https://www.googleapis.com/auth/presentations"]

    # Creates a Credentials instance from a service account json file
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=scopes
    )
    # If your domain uses Domain-Wide Delegation, uncomment and set a user to impersonate:
    # creds = creds.with_subject("user@yourdomain.com")

    # Returns SlidesResource
    return build("slides", "v1", credentials=creds)


# Prints a dotted line that extends to the full width of the terminal.
def print_full_width_dotted_line():
    try:
        # Get terminal size (columns, rows)
        terminal_width = os.get_terminal_size().columns

        # Create the dotted line string
        dotted_line = "." * terminal_width

        # Print the dotted line
        print(dotted_line)
    except OSError:
        # Handle cases where terminal size cannot be determined (e.g., not a TTY)
        print("Unable to determine terminal size. Printing a default dotted line.")
        print("." * 80)  # Print a default length line


# Returns ID of Google Slides presentation
def get_presentation_id(presentation_url: str) -> str:
    return presentation_url.partition("/d/")[2].split("/edit")[0]


# Returns ID of YouTube video
def get_video_id(video_url: str) -> str:
    return video_url.partition("v=")[2]


# Returns start, end, and total number of minutes for range of slides
def get_range(slide) -> Tuple[int, int, int] | None:
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
        match = re.match(RANGE_PATTERN, full_text, re.IGNORECASE)
        if match:
            start = match.group("start")
            end = match.group("end")
            minutes = match.group("minutes")

            return (int(start), int(end), int(minutes))

    return None


# Returns total number of slides in range
def get_num_slides_in_range(start: int, end: int):
    return end - start + 1


# Returns average number of minutes per slide in given range
def get_avg_time_per_slide_in_range(num_slides: int, minutes: int) -> int:
    avg = minutes / num_slides

    return int(math.floor(avg * 2) / 2)


# Returns time suggested for slide in speaker notes or None if doesn't exist
def get_suggested_time_for_slide(slide):
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

        match = re.match(SPEAKER_NOTES_PATTERN, full_text, re.IGNORECASE)
        unwanted_match = re.match(UNWANTED_PATTERN, full_text, re.IGNORECASE)
        if match is not None and unwanted_match is None:
            return int(match.group(3))

    return None


# Returns total number of slides in entire presentation
def get_total_slides_in_pres(slides):
    total = 0
    for _ in slides:
        total += 1

    return total


# Returns sum of all times suggested in speaker notes for slides in range
def add_all_suggested_times_in_range(slides, start: int, end: int):
    total = 0
    for slide in slides:
        if "objectId" not in slide:
            continue

        slide_number = int(slide["objectId"].lstrip("p"))
        slide_in_range = slide_number >= start and slide_number <= end

        if slide_in_range:
            suggested_time = get_suggested_time_for_slide(slide)

            if suggested_time is not None:
                total += int(suggested_time)

    return total


# Returns slide number
def get_slide_number(slide):
    return int(slide["objectId"].lstrip("p"))


# Returns true if slide in range, else false
def is_slide_in_range(slide, start, end) -> bool:
    slide_number = get_slide_number(slide)
    return slide_number >= start and slide_number <= end


# Adds timer videos to slides
def add_videos(svc, pres, presentation_id):
    slides = pres.get("slides", [])
    inserted_count = 0
    mins_assigned_so_far = 0

    range_start = 0
    range_end = 0
    range_total_mins = 0
    range_avg_mins_per_slide = 0
    range_mins_assigned_so_far = 0
    range_info_slide_number = 0

    slide_suggested_mins = 0

    slide_duration = 0

    for slide in slides:
        if "objectId" not in slide:
            continue

        slide_number = get_slide_number(slide)

        print(f"Slide {slide_number}")

        # Check if start of new range
        resp = get_range(slide)
        if resp is not None:
            range_info_slide_number = slide_number
            range_start = resp[0]
            range_end = resp[1]
            range_total_mins = resp[2]
            range_avg_mins_per_slide = get_avg_time_per_slide_in_range(
                get_num_slides_in_range(range_start, range_end), range_total_mins
            )

            # Warn user that total allotted time for slides in range may not match sum of suggested times for individual slides in range
            suggested_times_sum = add_all_suggested_times_in_range(
                slides, range_start, range_end
            )

            # print(
            #     f"\nWarning: The sum of the suggested times for slides {range_start}-{range_end}, {suggested_times_sum} minutes, might not match {range_total_mins} minutes, the time allotted for this range on slide {slide_number}. To stay within the time limit, slides in range {range_start}-{range_end} should average {range_avg_mins_per_slide} minute(s).\n"
            # )

        # Check if slide included in range
        slide_in_range = is_slide_in_range(slide, range_start, range_end)

        # Check if slide has suggested time
        resp = get_suggested_time_for_slide(slide)
        if resp is not None:
            slide_suggested_mins = int(resp)

        # If slide neither included in range NOR has own suggested time, skip
        if not slide_in_range and slide_suggested_mins == 0:
            print_full_width_dotted_line()
            continue

        # If slide included in range AND has own suggested time, issue warning and elicit user's desired time for slide
        if slide_in_range and slide_suggested_mins != 0:
            slide_duration = int(
                input(
                    f"\nWarning: The suggested time for slide {slide_number}, {slide_suggested_mins} minute(s), might put the presentation over or under {range_total_mins} minutes, the time allotted for slides in range {range_start}-{range_end} on slide {range_info_slide_number}. To stay within the time limit, slides {range_start}-{range_end} should average {range_avg_mins_per_slide} minute(s). Enter the number of minutes that you would like to spend on this slide:\n"
                )
            )
            mins_assigned_so_far += slide_duration
            range_mins_assigned_so_far += slide_duration

        # If slide included in range but DOESN'T have own suggested time, set duration to range's average minutes per slide
        elif slide_in_range:
            slide_duration = range_avg_mins_per_slide
            mins_assigned_so_far += slide_duration
            range_mins_assigned_so_far += slide_duration

        # If slide has own suggested time but NOT included in range, set duration to suggested time
        else:
            slide_duration = int(slide_suggested_mins)
            mins_assigned_so_far += slide_duration

        page_id = slide["objectId"]
        video_object_id = f"video_{uuid.uuid4().hex[:8]}"

        video_url = TIMER_VIDEOS[str(slide_duration)]
        if video_url:
            video_id = get_video_id(video_url)

            # 16:9 box ~480x270 pt, positioned near the top-left
            create_video = {
                "createVideo": {
                    "objectId": video_object_id,
                    "elementProperties": {
                        "pageObjectId": page_id,
                        "size": {
                            "width": {"magnitude": 480, "unit": "PT"},
                            "height": {"magnitude": 270, "unit": "PT"},
                        },
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": 80,
                            "translateY": 120,
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
                        # "start": 5,
                        # "end": 30,
                        # "mute": True,
                    },
                    "fields": "autoPlay",  # add start, end, mute if you set them
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

            # created = next(
            #     (
            #         r["createVideo"]["objectId"]
            #         for r in resp.get("replies", [])
            #         if "createVideo" in r
            #     ),
            #     None,
            # )

            print(f"\nInserted video id: {video_id} on slide {page_id}\n")
            if slide_in_range:
                print(
                    f"Time allotted for slides {range_start}-{range_end}: {range_total_mins} minute(s)"
                )
                print(
                    f"Time left in range: {range_total_mins - range_mins_assigned_so_far}"
                )
                print(f"Number of slides left in range: {range_end - slide_number}\n")

            print(f"Time left in presentation: {70 - mins_assigned_so_far}")
            print(
                f"Slides left in presentation: {get_total_slides_in_pres(slides) - slide_number}\n"
            )
            print_full_width_dotted_line()

            inserted_count += 1
            slide_suggested_mins = 0
            slide_duration = 0

    print(f"Added {inserted_count} video(s) to presentation {presentation_id}")


def delete_videos(svc, presentation_id):
    """
    Delete all video elements from a Google Slides presentation.

    """
    # Fetch only what's needed: pageElements with objectId and video
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


# ------------------------Command Line Interface------------------------
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
