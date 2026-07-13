#!/usr/bin/env python3
"""Slide 12 — USER FEEDBACK, continued (testimonial wall, template)."""
import deck_kit as K
from slides import _fb as FB


def render(prs):
    s = K.page("User feedback",
               "More voices from the pilot")
    FB.wall(s, [
        (K.NAVY, "Add the user’s feedback statement here. A Product Owner or branch "
                 "lead perspective works well on this slide.", "Username",
         "Title / Role · Team"),
        (K.TEAL_D, "Add the user’s feedback statement here. Highlight time saved, "
                   "accuracy, or customer reaction.", "Username", "Title / Role · Team"),
        (K.AMBER, "Add the user’s feedback statement here. Close on the impact for the "
                  "customer.", "Username", "Title / Role · Team"),
    ])
    return s
