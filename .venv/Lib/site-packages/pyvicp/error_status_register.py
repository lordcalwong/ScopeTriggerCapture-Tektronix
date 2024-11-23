EXR_LOOKUP = {
    #
    # execution error status register (response from "EXR?" query)
    #
    21: "Permission error. The command cannot be executed in local mode.",
    22: (
        "Environment error. The oscilloscope is not configured to correctly "
        "process a command. For instance, the oscilloscope cannot be set to "
        "RIS at a slow timebase."
    ),
    23: (
        "Option error. The command applies to an option which has not been "
        "installed."
    ),
    24: "Unresolved parsing error.",
    25: "Parameter error. Too many parameters specified.",
    26: "Non-implemented command.",
    27: "Parameter missing. A parameter was expected by the command.",
    30: (
        "Hex data error. A non-hexadecimal character has been detected in a "
        "hex data block."
    ),
    31: (
        "Waveform error. The amount of data received does not correspond to "
        "descriptor indicators."
    ),
    32: (
        "Waveform descriptor error. An invalid waveform descriptor has been "
        "detected."
    ),
    33: "Waveform text error. A corrupted waveform user text has been detected.",
    34: "Waveform time error. Invalid RIS or TRIG time data has been detected.",
    35: "Waveform data error. Invalid waveform data have been detected.",
    36: "Panel setup error. An invalid panel setup data block has been detected.",
    50: "No mass storage present when user attempted to access it.",
    51: "Mass storage not formatted when user attempted to access it.",
    53: (
        "Mass storage was write protected when user attempted to create a "
        "file, to delete a file, or to format the device."
    ),
    54: "Bad mass storage detected during formatting.",
    55: "Mass storage root directory full. Cannot add directory.",
    56: "Mass storage full when user attempted to write to it.",
    57: "Mass storage file sequence numbers exhausted (999 reached).",
    58: "Mass storage file not found.",
    59: "Requested directory not found.",
    61: "Mass storage filename not DOS compatible, or illegal filename.",
    62: "Cannot write on mass storage because filename already exists.",
}
