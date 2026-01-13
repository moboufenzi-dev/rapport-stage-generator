"""
Module des générateurs de pages de garde.
"""
from .covers import (
    generate_cover_classique,
    generate_cover_moderne,
    generate_cover_elegant,
    generate_cover_minimaliste,
    generate_cover_academique,
    generate_cover_geometrique,
    generate_cover_bicolore,
    generate_cover_pro,
    generate_cover_gradient,
    generate_cover_timeline,
    generate_cover_creative,
    generate_cover_luxe,
    COVER_GENERATORS,
    get_cover_generator,
)

__all__ = [
    'generate_cover_classique',
    'generate_cover_moderne',
    'generate_cover_elegant',
    'generate_cover_minimaliste',
    'generate_cover_academique',
    'generate_cover_geometrique',
    'generate_cover_bicolore',
    'generate_cover_pro',
    'generate_cover_gradient',
    'generate_cover_timeline',
    'generate_cover_creative',
    'generate_cover_luxe',
    'COVER_GENERATORS',
    'get_cover_generator',
]
