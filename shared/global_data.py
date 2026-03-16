config = {
    # Existing fields
    "TEST_NO": None,
    "TEST_DATE": None,
    "PROJECT": None,  # Kept for backward compatibility if needed

    # New global fields
    "TEST_NAME": None,
    "REPORT_NO": None,
    "TEST_ID": None,
    "WO_NO": None,
    "OEM": None,
    "PROGRAM": None,
    "PURPOSE": None,

    # Seat Configuration
    "SEAT_COUNT": 1,

    # Per-Seat Dynamic Fields (Lists of length 5)
    "SMP_ID": [None, None, None, None, None],
    "TEST_SAMPLE": [None, None, None, None, None],

    # Photo module state
    "PHOTO_MODULE_SLOTS": {
        "photo_slot_1": "",
        "photo_slot_2": "",
        "photo_slot_3": "",
        "photo_slot_4": "",
    },
    "PHOTO_MODULE_SESSION_ID": None,
}
