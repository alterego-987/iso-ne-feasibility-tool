# Global constants for the ISO-NE Feasibility Tool

# List of zones considered part of the Boston area
BOSTON_ZONES = [1000, 1002, 1004, 1010, 1012, 1013, 1014, 1120, 1130]

# List of excluded plant substrings. If a bus name contains any of these, it will be skipped during redispatch.
EXCLUDED_PLANTS = [
    'SEABROOK',
    'MILLSTONE',
    'SUN',
    'NYNE',
    'NYPA',
    'NBNE'
]
