# Cleaning and matching policy

Trim whitespace and normalize product/location codes to uppercase. Deduplicate only exact normalized product-location-units-date records, retaining the first record and recording the duplicate. Do not silently drop unmatched or non-numeric rows.
