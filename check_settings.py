import xml.etree.ElementTree as ET
import glob
import csv
import os

# CHECKS_TO_PERFORM list contains all settings to check.
# 'name': A descriptive name for the setting.
# 'xpath': The path to find the setting in the XML file.
# 'expected': The value the setting should have.

CHECKS_TO_PERFORM = [
        # -- Phase 2 Voice Checked --
    {
        'name': "Trunking - Phase 2 Voice Capable",
        'xpath': ".//Recset[@Name='Trunking System']//Section[@Name='ASTRO 25']/Field[@Name='Phase 2 Voice Capable']",
        'expected': "True"
    },
    # -- Trunking: Channel ID 3 --
    {
        'name': "Identifier Enabled on Channel ID 3",
        'xpath': ".//Recset[@Name='Trunking System']//EmbeddedNode[@ReferenceKey='Channel ID 3']//Field[@Name='Identifier Enable']",
        'expected': "True"
    },
    {
        'name': "Channel ID 3 Base Frequency",
        'xpath': ".//Recset[@Name='Trunking System']//EmbeddedNode[@ReferenceKey='Channel ID 3']//Field[@Name='Base Frequency (MHz)']",
        'expected': "851.012500"
    },
    {
        'name': "Channel ID 3 Channel Spacing",
        'xpath': ".//Recset[@Name='Trunking System']//EmbeddedNode[@ReferenceKey='Channel ID 3']//Field[@Name='Channel Spacing (kHz)']",
        'expected': "12.500"
    },
    {
        'name': "Channel ID 3 Type",
        'xpath': ".//Recset[@Name='Trunking System']//EmbeddedNode[@ReferenceKey='Channel ID 3']//Field[@Name='Channel Type']",
        'expected': "TDMA"
    },
    {
        'name': "Channel ID 3 Transmit Offset",
        'xpath': ".//Recset[@Name='Trunking System']//EmbeddedNode[@ReferenceKey='Channel ID 3']//Field[@Name='Transmit Offset (MHz)']",
        'expected': "45.000000"
    },
    {
        'name': "Channel ID 3 Transmit Offset Sign",
        'xpath': ".//Recset[@Name='Trunking System']//EmbeddedNode[@ReferenceKey='Channel ID 3']//Field[@Name='Transmit Offset Sign']",
        'expected': "-"
    },

    # -- Trunking: Channel ID 4 --
    {
        'name': "Identifier Enabled on Channel ID 4",
        'xpath': ".//Recset[@Name='Trunking System']//EmbeddedNode[@ReferenceKey='Channel ID 4']//Field[@Name='Identifier Enable']",
        'expected': "True"
    },
    {
        'name': "Channel ID 4 Base Frequency",
        'xpath': ".//Recset[@Name='Trunking System']//EmbeddedNode[@ReferenceKey='Channel ID 4']//Field[@Name='Base Frequency (MHz)']",
        'expected': "762.006250"
    },
    {
        'name': "Channel ID 4 Channel Spacing",
        'xpath': ".//Recset[@Name='Trunking System']//EmbeddedNode[@ReferenceKey='Channel ID 4']//Field[@Name='Channel Spacing (kHz)']",
        'expected': "12.500"
    },
    {
        'name': "Channel ID 4 Type",
        'xpath': ".//Recset[@Name='Trunking System']//EmbeddedNode[@ReferenceKey='Channel ID 4']//Field[@Name='Channel Type']",
        'expected': "TDMA"
    },
    {
        'name': "Channel ID 4 Transmit Offset",
        'xpath': ".//Recset[@Name='Trunking System']//EmbeddedNode[@ReferenceKey='Channel ID 4']//Field[@Name='Transmit Offset (MHz)']",
        'expected': "30.000000"
    },
    {
        'name': "Channel ID 4 Transmit Offset Sign",
        'xpath': ".//Recset[@Name='Trunking System']//EmbeddedNode[@ReferenceKey='Channel ID 4']//Field[@Name='Transmit Offset Sign']",
        'expected': "+"
    },
    
    # -- Channel: 8CALL90 --
    {
        'name': "8TAC/8CALL90 - Rx Freq",
        'xpath': ".//Recset[@Name='Conventional Personality']/Node[@ReferenceKey='8TAC']//EmbeddedNode[@ReferenceKey='8CALL90']//Field[@Name='Rx / TA Frequency (MHz)']",
        'expected': "851.012500"
    },
    {
        'name': "8TAC/8CALL90 - Tx Freq",
        'xpath': ".//Recset[@Name='Conventional Personality']/Node[@ReferenceKey='8TAC']//EmbeddedNode[@ReferenceKey='8CALL90']//Field[@Name='Tx Frequency (MHz)']",
        'expected': "806.012500"
    },
    {
        'name': "8TAC/8CALL90 - Channel Spacing",
        'xpath': ".//Recset[@Name='Conventional Personality']/Node[@ReferenceKey='8TAC']//EmbeddedNode[@ReferenceKey='8CALL90']//Field[@Name='Tx Deviation / Channel Spacing']",
        'expected': "4 kHz / 20 kHz"
    },
    {
        'name': "8TAC/8CALL90 - Direct/Talkaround",
        'xpath': ".//Recset[@Name='Conventional Personality']/Node[@ReferenceKey='8TAC']//EmbeddedNode[@ReferenceKey='8CALL90']//Field[@Name='Direct / Talkaround']",
        'expected': "False"
    },
    
    # -- Channel: 8CALL90D --
    {
        'name': "8TAC/8CALL90D - Rx Freq",
        'xpath': ".//Recset[@Name='Conventional Personality']/Node[@ReferenceKey='8TAC']//EmbeddedNode[@ReferenceKey='8CALL90D']//Field[@Name='Rx / TA Frequency (MHz)']",
        'expected': "851.012500"
    },
    {
        'name': "8TAC/8CALL90D - Tx Freq",
        'xpath': ".//Recset[@Name='Conventional Personality']/Node[@ReferenceKey='8TAC']//EmbeddedNode[@ReferenceKey='8CALL90D']//Field[@Name='Tx Frequency (MHz)']",
        'expected': "851.012500"
    },
    {
        'name': "8TAC/8CALL90D - Channel Spacing",
        'xpath': ".//Recset[@Name='Conventional Personality']/Node[@ReferenceKey='8TAC']//EmbeddedNode[@ReferenceKey='8CALL90D']//Field[@Name='Tx Deviation / Channel Spacing']",
        'expected': "4 kHz / 20 kHz"
    },
    {
        'name': "8TAC/8CALL90D - Direct/Talkaround",
        'xpath': ".//Recset[@Name='Conventional Personality']/Node[@ReferenceKey='8TAC']//EmbeddedNode[@ReferenceKey='8CALL90D']//Field[@Name='Direct / Talkaround']",
        'expected': "True"
    },
    
    # -- Channel: 8TAC91 --
    {
        'name': "8TAC/8TAC91 - Rx Freq",
        'xpath': ".//Recset[@Name='Conventional Personality']/Node[@ReferenceKey='8TAC']//EmbeddedNode[@ReferenceKey='8TAC91']//Field[@Name='Rx / TA Frequency (MHz)']",
        'expected': "851.512500"
    },
    {
        'name': "8TAC/8TAC91 - Tx Freq",
        'xpath': ".//Recset[@Name='Conventional Personality']/Node[@ReferenceKey='8TAC']//EmbeddedNode[@ReferenceKey='8TAC91']//Field[@Name='Tx Frequency (MHz)']",
        'expected': "806.512500"
    },
    {
        'name': "8TAC/8TAC91 - Channel Spacing",
        'xpath': ".//Recset[@Name='Conventional Personality']/Node[@ReferenceKey='8TAC']//EmbeddedNode[@ReferenceKey='8TAC91']//Field[@Name='Tx Deviation / Channel Spacing']",
        'expected': "4 kHz / 20 kHz"
    },
    
    # -- Channel: 8TAC91D --
    {
        'name': "8TAC/8TAC91D - Rx Freq",
        'xpath': ".//Recset[@Name='Conventional Personality']/Node[@ReferenceKey='8TAC']//EmbeddedNode[@ReferenceKey='8TAC91D']//Field[@Name='Rx / TA Frequency (MHz)']",
        'expected': "851.512500"
    },
    {
        'name': "8TAC/8TAC91D - Tx Freq",
        'xpath': ".//Recset[@Name='Conventional Personality']/Node[@ReferenceKey='8TAC']//EmbeddedNode[@ReferenceKey='8TAC91D']//Field[@Name='Tx Frequency (MHz)']",
        'expected': "851.512500"
    },
    {
        'name': "8TAC/8TAC91D - Channel Spacing",
        'xpath': ".//Recset[@Name='Conventional Personality']/Node[@ReferenceKey='8TAC']//EmbeddedNode[@ReferenceKey='8TAC91D']//Field[@Name='Tx Deviation / Channel Spacing']",
        'expected': "4 kHz / 20 kHz"
    },

    # -- Zone "6-Z6-INTEROP", Channel "6-GW IO 6" --
    {
        'name': "Z6/GW IO 6 - Channel Type",
        'xpath': ".//Recset[@Name='Zone Channel Assignment']/Node[@ReferenceKey='6-Z6-INTEROP']//EmbeddedNode[@ReferenceKey='6-GW IO 6']//Field[@Name='Channel Type']",
        'expected': "Trk"
    },
    {
        'name': "Z6/GW IO 6 - Personality",
        'xpath': ".//Recset[@Name='Zone Channel Assignment']/Node[@ReferenceKey='6-Z6-INTEROP']//EmbeddedNode[@ReferenceKey='6-GW IO 6']//Field[@Name='Personality']",
        'expected': "027A - IO"
    },
    {
        'name': "Z6/GW IO 6 - Channel Name",
        'xpath': ".//Recset[@Name='Zone Channel Assignment']/Node[@ReferenceKey='6-Z6-INTEROP']//EmbeddedNode[@ReferenceKey='6-GW IO 6']//Field[@Name='Channel Name']",
        'expected': "GW IO 6"
    },
    {
        'name': "Z6/GW IO 6 - Top Display Channel",
        'xpath': ".//Recset[@Name='Zone Channel Assignment']/Node[@ReferenceKey='6-Z6-INTEROP']//EmbeddedNode[@ReferenceKey='6-GW IO 6']//Field[@Name='Top Display Channel']",
        'expected': "GW IO 6"
    },
    {
        'name': "Z6/GW IO 6 - Trunking Talkgroup",
        'xpath': ".//Recset[@Name='Zone Channel Assignment']/Node[@ReferenceKey='6-Z6-INTEROP']//EmbeddedNode[@ReferenceKey='6-GW IO 6']//Field[@Name='Trunking Talkgroup']",
        'expected': "IO 6"
    },
    {
        'name': "Z6/GW IO 6 - Active Channel",
        'xpath': ".//Recset[@Name='Zone Channel Assignment']/Node[@ReferenceKey='6-Z6-INTEROP']//EmbeddedNode[@ReferenceKey='6-GW IO 6']//Field[@Name='Active Channel']",
        'expected': "True"
    },

]

# Parses XML files and performs the above CHECKS_TO_PERFORM
def check_xml_file(filepath, writer):
    try:
        tree = ET.parse(filepath)
        root = tree.getroot()
        filename = os.path.basename(filepath)
        
        for check in CHECKS_TO_PERFORM:
            element = root.find(check['xpath'])
            
            # Check if the setting exists in the file
            if element is None:
                writer.writerow([filename, check['name'], "Setting Not Found", check['expected'], "N/A"])
                continue

            actual_value = element.text

            # Write to CSV where the settings do not match
            if actual_value != check['expected']:
                writer.writerow([filename, check['name'], "Incorrect Value", check['expected'], actual_value])

    # Write to CSV where cannot parse the XML file
    except ET.ParseError:
        writer.writerow([os.path.basename(filepath), "File Error", "Could not parse XML", "N/A", "N/A"])

def main():
    """Finds all XML files and generates a CSV report for any discrepancies."""
    xml_files = glob.glob('*.xml')
    if not xml_files:
        print("No .xml files found in this directory.")
        return

    print(f"Found {len(xml_files)} XML files. Checking settings...")

    with open('report.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        # Write the header row for the report
        writer.writerow(['Filename', 'Setting', 'Issue', 'Expected Value', 'Actual Value'])

        # Check each file
        for filepath in xml_files:
            check_xml_file(filepath, writer)

    print("Check complete. Please see report.csv for any issues.")

if __name__ == "__main__":
    main()