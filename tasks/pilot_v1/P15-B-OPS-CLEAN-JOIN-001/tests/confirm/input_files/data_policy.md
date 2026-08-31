# Cleaning and matching policy

Trim whitespace and normalize product and location codes to uppercase. Deduplicate only exact normalized product-location-units-date records, retain the first record, and preserve duplicate evidence. Route non-numeric units and unmatched master keys to Exceptions.
