#!/usr/bin/env python
import app
import sys

# Test data
rechnung_data = {
    'destination_en': 'Japan',
    'destination_de': 'Japan',
    'travel_start': '01.06.2026',
    'travel_end': '10.06.2026',
    'client_names': ['John Smith', 'Jane Smith'],
    'hotels': [
        {
            'check_in': '01.06.2026',
            'check_out': '03.06.2026',
            'nights': 2,
            'hotel_name': 'Imperial Hotel',
            'city': 'Tokyo',
            'room_type': 'Twin Room',
            'meal_plan_en': 'breakfast'
        },
        {
            'check_in': '03.06.2026',
            'check_out': '06.06.2026',
            'nights': 3,
            'hotel_name': 'Kyoto Central Hotel',
            'city': 'Kyoto',
            'room_type': 'Double Room',
            'meal_plan_en': 'halfboard'
        },
        {
            'check_in': '06.06.2026',
            'check_out': '10.06.2026',
            'nights': 4,
            'hotel_name': 'Grand Hyatt Osaka',
            'city': 'Osaka',
            'room_type': 'Deluxe Suite',
            'meal_plan_en': 'fullboard'
        }
    ]
}

dmc = {
    'name': 'Japan Destination Inc.',
    'contact_person': 'Yuki Tanaka',
    'mobile': '+81-3-1234-5678',
    'note': 'emergency'
}

try:
    docx_bytes = app.build_confirmation_docx(rechnung_data, dmc, 'Mr. Tanaka', '+81-90-1234-5678')

    # Save to test file
    out_path = 'test_confirmation.docx'
    with open(out_path, 'wb') as f:
        f.write(docx_bytes)

    print(f'[OK] Test confirmation document created: {out_path}')
    print(f'[OK] File size: {len(docx_bytes)} bytes')
    sys.exit(0)

except Exception as e:
    print(f'[ERROR] {e}')
    import traceback
    traceback.print_exc()
    sys.exit(1)
