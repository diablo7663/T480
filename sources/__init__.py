"""Moroccan car-marketplace sources.

Every module in this package fetches one site and emits records in exactly the
shape scraper.py already writes for avito.ma, so merge.py can union all of them
and reuse the cleaning and Excel code unchanged.
"""

SOURCES = {}
