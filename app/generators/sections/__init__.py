"""
Module des générateurs de sections.
"""
from .sections import (
    generate_toc_section,
    generate_thanks_section,
    generate_abstract_section,
    generate_chapters,
    generate_annexes_section,
    get_chapter_hint,
)

__all__ = [
    'generate_toc_section',
    'generate_thanks_section',
    'generate_abstract_section',
    'generate_chapters',
    'generate_annexes_section',
    'get_chapter_hint',
]
