#!/usr/bin/env python3
"""Slide 11 — USER FEEDBACK (testimonial wall, template). Username, title, statement
and a user-picture placeholder per card; replace the placeholders with real feedback.
"""
import deck_kit as K
from slides import _fb as FB


def render(prs):
    s = K.page("User feedback",
               "What our users say about OSG")
    FB.wall(s, [
        (K.TEAL_D, "Add the user’s feedback statement here. Share a real quote about "
                   "how OSG changed their day.", "Username", "Title / Role · Branch"),
        (K.AMBER, "Add the user’s feedback statement here. Keep it short, specific and "
                  "in their own words.", "Username", "Title / Role · Branch"),
        (K.GOLD, "Add the user’s feedback statement here. A number or a before/after "
                 "detail makes it land.", "Username", "Title / Role · Branch"),
    ])
    return s
