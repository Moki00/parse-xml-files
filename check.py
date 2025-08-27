import lxml.etree as ET
import glob
import csv
import os

CHECKS_TO_PERFORM = [
    # -- Phase 2 Voice Capable --
    {
        'group_name': 'Phase 2 Voice Capable',
        # CHANGED: Using contains() to find the node more flexibly
        'base_xpath': ".//Recset[@Name='Trunking System']/Node[contains(@ReferenceKey, 'GWINNETT')]/Section[@Name='ASTRO 25']",
        'fields': {
            'Phase 2 Voice Capable': 'True'
        }
    },
       
    # -- TDMA Channel ID 3 --
    {
        'group_name': 'Trunking System - Channel ID 3',
        'base_xpath': ".//Recset[@Name='Trunking System']//EmbeddedNode[@ReferenceKey='Channel ID 3']",
        'fields': {
            'Identifier Enable': 'True',
            'Base Frequency (MHz)': '851.012500',
            'Channel Spacing (kHz)': '12.500',
            'Channel Type': 'TDMA',
            'Transmit Offset (MHz)': '45.000000',
            'Transmit Offset Sign': '-'
        }
    },

    # -- TDMA Channel ID 4 --
    {
        'group_name': 'Trunking System - Channel ID 4',
        'base_xpath': ".//Recset[@Name='Trunking System']//EmbeddedNode[@ReferenceKey='Channel ID 4']",
        'fields': {
            'Identifier Enable': 'True',
            'Base Frequency (MHz)': '762.006250',
            'Channel Spacing (kHz)': '12.500',
            'Channel Type': 'TDMA',
            'Transmit Offset (MHz)': '30.000000',
            'Transmit Offset Sign': '+'
        }
    },

    # -- Channel: 8CALL90 --
    {
        'group_name': 'Channel: 8CALL90',
        'base_xpath': ".//Recset[@Name='Conventional Personality']//EmbeddedNode[@ReferenceKey='8CALL90']",
        'fields': {
            'Rx / TA Frequency (MHz)': '851.012500',
            'User Selectable PL (MPL)':'False',
            'Tx Squelch Type':'PL',
            'Tx DPL Code':'023',
            'Tx DPL Invert':'False',
            'Rx / TA Squelch Type':'PL',
            'Tx Frequency (MHz)':'806.012500',
            'Tx Network ID':'659',
            'Tx PL Code':'5A',
            'Tx PL Freq':'156.7',
            'Rx / TA  PL Code':'5A',
            'Rx / TA PL Freq':'156.7',
            'Rx / TA DPL Code':'023',
            'Rx / TA DPL Invert':'False',
            'Rx / TA  Network ID':'659',
            'Direct / Talkaround':'False',
            'Direct Squelch Type':'PL',
            'Direct PL Freq':'67.0',
            'Direct PL Code':'XZ',
            'ASTRO Talkgroup ID':'TG 1',
            'Tx Deviation / Channel Spacing': '4 kHz / 20 kHz',
            'Name':'8CALL90',
            'Direct Network ID':'659',
            'User Selectable PL [MPL]':'Disabled',
            'Direct Frequency (MHz)':'851.012500',
            'Direct DPL Code':'023',
            'Direct DPL Invert':'False',
        }
    },
        # -- Channel: 8CALL90Direct --
    {
        'group_name': 'Channel: 8CALL90Direct',
        'base_xpath': ".//Recset[@Name='Conventional Personality']//EmbeddedNode[@ReferenceKey='8CALL90D']",
        'fields': {
            'Rx / TA Frequency (MHz)': '851.012500',
            'User Selectable PL (MPL)':'False',
            'Tx Squelch Type':'PL',
            'Tx DPL Code':'023',
            'Tx DPL Invert':'False',
            'Rx / TA Squelch Type':'PL',
            'Tx Frequency (MHz)':'851.012500',
            'Tx Network ID':'659',
            'Tx PL Code':'5A',
            'Tx PL Freq':'156.7',
            'Rx / TA  PL Code':'5A',
            'Rx / TA PL Freq':'156.7',
            'Rx / TA DPL Code':'023',
            'Rx / TA DPL Invert':'False',
            'Rx / TA  Network ID':'659',
            'Direct / Talkaround':'True',
            'Direct Squelch Type':'PL',
            'Direct PL Freq':'67.0',
            'Direct PL Code':'XZ',
            'ASTRO Talkgroup ID':'TG 1',
            'Tx Deviation / Channel Spacing': '4 kHz / 20 kHz',
            'Name':'8CALL90D',
            'Direct Network ID':'659',
            'User Selectable PL [MPL]':'Disabled',
            'Direct Frequency (MHz)':'851.012500',
            'Direct DPL Code':'023',
            'Direct DPL Invert':'False',
        }
    },

            # -- Channel: 8TAC91Direct --
    {
        'group_name': 'Channel: 8TAC91Direct',
        'base_xpath': ".//Recset[@Name='Conventional Personality']//EmbeddedNode[@ReferenceKey='8TAC91D']",
        'fields': {
            'Rx / TA Frequency (MHz)': '851.512500',
            'User Selectable PL (MPL)':'False',
            'Tx Squelch Type':'PL',
            'Tx DPL Code':'023',
            'Tx DPL Invert':'False',
            'Rx / TA Squelch Type':'PL',
            'Tx Frequency (MHz)':'851.512500',
            'Tx Network ID':'659',
            'Tx PL Code':'5A',
            'Tx PL Freq':'156.7',
            'Rx / TA  PL Code':'5A',
            'Rx / TA PL Freq':'156.7',
            'Rx / TA DPL Code':'023',
            'Rx / TA DPL Invert':'False',
            'Rx / TA  Network ID':'659',
            'Direct / Talkaround':'True',
            'Direct Squelch Type':'PL',
            'Direct PL Freq':'67.0',
            'Direct PL Code':'XZ',
            'ASTRO Talkgroup ID':'TG 1',
            'Tx Deviation / Channel Spacing': '4 kHz / 20 kHz',
            'Name':'8TAC91D',
            'Direct Network ID':'659',
            'User Selectable PL [MPL]':'Disabled',
            'Direct Frequency (MHz)':'851.512500',
            'Direct DPL Code':'023',
            'Direct DPL Invert':'False',
        }
    },
 
    # -- GW IO 6 --
    {
        'group_name': 'INTEROP - GW IO 6',
        'base_xpath': ".//Recset[@Name='Zone Channel Assignment']/Node[contains(@ReferenceKey, 'INTEROP')]//EmbeddedNode[@ReferenceKey='6-GW IO 6']",
        'fields': {
            'Channel Type': 'Trk',
            'Personality': '027A - IO',
            'Channel Name': 'GW IO 6',
            'Top Display Channel': 'GW IO 6',
            'Trunking Talkgroup': 'IO 6',
            'Active Channel': 'True'
        }
    },
        # -- 8CALL90 --
    {
        'group_name': 'INTEROP - 8CALL90',
        'base_xpath': ".//Recset[@Name='Zone Channel Assignment']/Node[contains(@ReferenceKey, 'INTEROP')]//EmbeddedNode[@ReferenceKey='7-8CALL90']",
        'fields': {
            'Channel Type': 'Cnv',
            'Personality': '8TAC',
            'Channel Name': '8CALL90',
            'Top Display Channel': '8CALL90',
            'Active Channel': 'True'
        }
    },

]

"""CHECKS_TO_PERFORM run against a single XML file and writes any discrepancies to the CSV writer."""
def check_xml_file(filepath, writer):
    try:
        parser = ET.XMLParser(remove_blank_text=True, resolve_entities=False)
        tree = ET.parse(filepath, parser)
        root = tree.getroot()
        filename = os.path.basename(filepath)
        
        for group in CHECKS_TO_PERFORM:
            group_name = group['group_name']
            parent = None
            parents = root.xpath(group['base_xpath'])

            if not parents:
                writer.writerow([filename, group_name, "Parent Section Not Found", "N/A", "N/A"])
                continue

            for parent in parents:
                for field_name, expected_value in group['fields'].items():
                    field_elements = parent.xpath(f".//Field[@Name='{field_name}']")

                    if not field_elements:
                        writer.writerow([filename, f"{group_name} - {field_name}", "Setting Not Found", expected_value, "N/A"])
                        continue

                    actual_value = field_elements[0].text or ""

                    if actual_value != expected_value:
                        writer.writerow([filename, f"{group_name} - {field_name}", "Incorrect Value", expected_value, actual_value])
                        
    except ET.XMLSyntaxError:
        writer.writerow([os.path.basename(filepath), "File Error", "Could not parse XML", "N/A", "N/A"])

def main():
    """Finds all XML files and generates a CSV report for any discrepancies."""
    xml_files = glob.glob('*.xml')
    if not xml_files:
        print("No .xml files found in this directory.")
        return

    report_filename = 'report.csv'
    print(f"Found {len(xml_files)} XML files. Checking settings...")

    with open(report_filename, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['Filename', 'Setting', 'Issue', 'Expected Value', 'Actual Value'])

        for filepath in xml_files:
            check_xml_file(filepath, writer)

    print(f"Check complete. Please see {report_filename} for any issues.")

if __name__ == "__main__":
    main()