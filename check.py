import lxml.etree as ET
import glob
import csv
import os

CHECKS_TO_PERFORM = [
    # -- Phase 2 Voice Capable --
    {
        'group_name': 'Trunking System - General',
        'base_xpath': ".//Recset[@Name='Trunking System']/Node[@ReferenceKey='027A GWINNETT']/Section[@Name='ASTRO 25']",
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
        # -- Channel: 8CALL90D --
    {
        'group_name': 'Channel: 8CALL90D',
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
 
    # -- GW IO 6 --
    {
        'group_name': 'INTEROP - GW IO 6',
        # parent container that holds all the zones
        'parent_xpath': ".//Recset[@Name='Zone Channel Assignment']",
        # the node contains
        'node_search': {
            'tag': 'Node',
            'attribute': 'ReferenceKey',
            'contains': 'INTEROP'
        },
        # the path to the channel relative to the node
        'child_base_xpath': ".//EmbeddedNode[@ReferenceKey='6-GW IO 6']",
        # fields to check
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
        # 1. Find the parent container that holds all the zones
        'parent_xpath': ".//Recset[@Name='Zone Channel Assignment']",
        # 2. Describe the node we need to find within that container
        'node_search': {
            'tag': 'Node',
            'attribute': 'ReferenceKey',
            'contains': 'INTEROP'
        },
        # 3. Define the path to the channel *relative to the node we find*
        'child_base_xpath': ".//EmbeddedNode[@ReferenceKey='7-8CALL90']",
        # 4. List the fields to check inside
        'fields': {
            'Channel Type': 'Cnv',
            'Personality': '8TAC',
            'Channel Name': '8CALL90',
            'Top Display Channel': '8CALL90',
            'Active Channel': 'True'
        }
    },

]

def check_xml_file(filepath, writer):
    """Parses a single XML file and performs all grouped checks."""
    try:
        tree = ET.parse(filepath)
        root = tree.getroot()
        filename = os.path.basename(filepath)
        
        for group in CHECKS_TO_PERFORM:
            group_name = group['group_name']
            parent_element = None

            # Check if this is a dynamic search group
            if 'node_search' in group:
                search_container = root.find(group['parent_xpath'])
                if search_container is not None:
                    search_params = group['node_search']

                    for node in search_container.findall(search_params['tag']):
                        # Check if the attribute value contains the required text
                        if search_params['contains'] in node.get(search_params['attribute'], ''):
                            # We found our INTEROP zone, now find the specific channel within it
                            parent_element = node.find(group['child_base_xpath'])
                            break # Stop searching once we find the first match
            else:
                # This is a standard check with a direct path
                parent_element = root.find(group['base_xpath'])
            
            if parent_element is None:
                writer.writerow([filename, group_name, "Parent Section Not Found", "N/A", "N/A"])
                continue

            # Now, check each field within the found parent container
            for field_name, expected_value in group['fields'].items():
                field_element = parent_element.find(f".//Field[@Name='{field_name}']")
                
                if field_element is None:
                    writer.writerow([filename, f"{group_name} - {field_name}", "Setting Not Found", expected_value, "N/A"])
                    continue
                
                actual_value = field_element.text or "" # Handle empty fields

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