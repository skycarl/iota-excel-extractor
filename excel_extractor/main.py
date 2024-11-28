import sys
import pandas as pd
from tqdm import tqdm
from pathlib import Path
from datetime import time

VALID_GPS_FORMATS = [
    'deg min.mmm',
    'deg mm sec.ss',
    'deg.ddddd'
]
COORDS_DESIRED_FORMAT = 'deg.ddddd'
SUPPORTED_FORM_VERSION = 'V5.6.11'

def create_time_with_microseconds(hh, mm, ss):
    seconds_int = int(ss)
    microseconds = int((ss - seconds_int) * 1_000_000)
    return time(hh, mm, seconds_int, microseconds)

def extract_fields(file_path):
    try:
        df = pd.read_excel(file_path, sheet_name='DATA', header=None)
    except ValueError as e:
        if 'Worksheet named \'DATA\' not found' in str(e):
            return None
        else:
            raise

    form_version = df.iloc[1, 25]
    if form_version != SUPPORTED_FORM_VERSION:
        raise ValueError(f"Unsupported form version found in {file_path}: {form_version}. Expected {SUPPORTED_FORM_VERSION}")
    
    # Extract year, month, and day
    year = int(df.iloc[4, 3])
    month = df.iloc[4, 10]
    day = int(df.iloc[4, 15])
    
    # Convert month name to month number
    month_number = pd.to_datetime(month, format='%B').month
    
    # Assemble into ISO-8601 style date
    event_date = f"{year:04d}-{month_number:02d}-{day:02d}"

    asteroid_number = df.iloc[6, 4]
    asteroid_name = df.iloc[6, 10]
    star_catalog = df.iloc[6, 18]
    star_number = df.iloc[6, 23]
    predicted_h = int(df.iloc[4, 24])
    predicted_m = int(df.iloc[4, 26])
    predicted_s = int(df.iloc[4, 28])
    
    predicted_time_utc = time(predicted_h, predicted_m, predicted_s)

    # Get lat/long
    lat_coords_format = df.iloc[16, 4] # latitude coordinates format (one of VALID_GPS_FORMATS)
    lat_coords = str(df.iloc[17, 4]) # latitude coordinates
    lat_coords_ns = df.iloc[17, 9] # either `N` or `S` for north or south
    lat_coords += str(f' {lat_coords_ns}')
    
    long_coords_format = df.iloc[16, 13] # longitude coordinates format (one of VALID_GPS_FORMATS)
    long_coords = str(df.iloc[17, 13])# longitude coordinates
    long_coords_ew = df.iloc[17, 17] # either `E` or `W` for east or west
    long_coords += str(f' {long_coords_ew}')

    coords_datum = df.iloc[17, 26]

    if lat_coords_format not in VALID_GPS_FORMATS or long_coords_format not in VALID_GPS_FORMATS:
        raise ValueError(f'Invalid GPS coordinate format found in {file_path}; one of {VALID_GPS_FORMATS} is expected, but got lat: {lat_coords_format}, long: {long_coords_format}')
    
    gps_coords = normalize_coords(
        lat_coords,
        lat_coords_format,
        long_coords, 
        long_coords_format,
        COORDS_DESIRED_FORMAT
    )

    pos_neg = df.iloc[1, 0]
    observer = df.iloc[8, 3]
    email = df.iloc[8, 18]
    location = df.iloc[14, 4]

    return {
        "event_date": event_date,
        "asteroid_number": asteroid_number,
        "asteroid_name": asteroid_name,
        "star_catalog": star_catalog,
        "star_number": star_number,
        "predicted_time_utc": predicted_time_utc,
        "coords": f"{gps_coords['latitude']}, {gps_coords['longitude']}",
        "pos_neg": pos_neg,
        "observer": observer,
        "email": email,
        "location": location
    }

def normalize_coords(
    lat_coords,
    lat_coords_format,
    long_coords,
    long_coords_format,
    coords_desired_format
):
    def parse_coordinate(coord_str, coord_format):
        # Extract direction (N, S, E, W)
        direction = coord_str.strip()[-1].upper()
        coord_str = coord_str.strip()[:-1].strip()
        
        if coord_format == 'deg min.mmm':
            parts = coord_str.split()
            degrees = float(parts[0])
            minutes = float(parts[1])
            decimal_degrees = degrees + minutes / 60
        elif coord_format == 'deg mm sec.ss':
            parts = coord_str.split()
            degrees = float(parts[0])
            minutes = float(parts[1])
            seconds = float(parts[2])
            decimal_degrees = degrees + minutes / 60 + seconds / 3600
        elif coord_format == 'deg.ddddd':
            decimal_degrees = float(coord_str)
        else:
            raise ValueError(f"Unsupported coordinate format: {coord_format}")
        
        # Apply direction
        if direction in ['S', 'W']:
            decimal_degrees = -decimal_degrees
        elif direction not in ['N', 'E']:
            raise ValueError(f"Invalid direction: {direction}")
        
        return decimal_degrees

    def format_coordinate(decimal_degrees, coord_format, is_latitude):
        # Get the direction letter
        if decimal_degrees < 0:
            direction = 'S' if is_latitude else 'W'
            decimal_degrees = -decimal_degrees
        else:
            direction = 'N' if is_latitude else 'E'

        if coord_format == 'deg min.mmm':
            degrees = int(decimal_degrees)
            minutes = (decimal_degrees - degrees) * 60
            return f"{degrees} {minutes:.3f} {direction}"
        elif coord_format == 'deg mm sec.ss':
            degrees = int(decimal_degrees)
            minutes_full = (decimal_degrees - degrees) * 60
            minutes = int(minutes_full)
            seconds = (minutes_full - minutes) * 60
            return f"{degrees} {minutes} {seconds:.2f} {direction}"
        elif coord_format == 'deg.ddddd':
            return f"{decimal_degrees:.4f} {direction}"
        else:
            raise ValueError(f"Unsupported coordinate format: {coord_format}")

    # Parse input coordinates to decimal degrees
    lat_decimal = parse_coordinate(lat_coords, lat_coords_format)
    long_decimal = parse_coordinate(long_coords, long_coords_format)

    # Format coordinates to desired format
    lat_formatted = format_coordinate(lat_decimal, coords_desired_format, is_latitude=True)
    long_formatted = format_coordinate(long_decimal, coords_desired_format, is_latitude=False)

    gps_coords = {
        'latitude': lat_formatted,
        'longitude': long_formatted
    }
    return gps_coords

def main():
    if len(sys.argv) != 2:
        print("Usage: python main.py <source_dir>")
        sys.exit(1)

    source_dir = Path(sys.argv[1])

    if not source_dir.is_dir():
        print(f"Error: {source_dir} is not a valid directory")
        sys.exit(1)

    excel_files = list(source_dir.rglob('*.xlsx')) + list(source_dir.rglob('*.xls'))
    excel_files = [f for f in (list(source_dir.rglob('*.xlsx')) + list(source_dir.rglob('*.xls'))) if '__MACOSX' not in f.parts]

    num_files = len(excel_files)

    if num_files == 0:
        print("No Excel files found in the specified directory.")
        sys.exit(1)

    results = []
    errors = []

    for file in tqdm(excel_files, desc="Processing files", unit="file"):
        try:
            fields = extract_fields(file)
            if fields is not None:
                results.append(fields)
        except Exception as e:
            errors.append((file, str(e)))

    output_df = pd.DataFrame(results)
    output_df.to_excel("output.xlsx", index=False)

    if errors:
        print("\nErrors encountered:")
        for file, error in errors:
            print(f"File: {file}\nError: {error}\n")

if __name__ == "__main__":
    main()