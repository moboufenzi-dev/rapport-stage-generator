"""
Module des générateurs de documents.
"""
from .report_generator import generate_report
from .covers import get_cover_generator, COVER_GENERATORS
from .sections import (
    generate_toc_section,
    generate_thanks_section,
    generate_abstract_section,
    generate_chapters,
    generate_annexes_section,
)

__all__ = [
    'generate_report',
    'get_cover_generator',
    'COVER_GENERATORS',
    'generate_toc_section',
    'generate_thanks_section',
    'generate_abstract_section',
    'generate_chapters',
    'generate_annexes_section',
]
