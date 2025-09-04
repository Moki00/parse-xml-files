import pandas as pd
import lxml.etree as ET
import glob
import os
from openpyxl.utils import get_column_letter

CHECKS_TO_PERFORM = [
    # -- Phase 2 Voice Capable --
    {
        'group_name': 'Phase 2 Voice Capable',
        'base_xpath': ".//Recset[@Name='Trunking System']/Node[contains(@ReferenceKey, 'GWINNETT')]/Section[@Name='ASTRO 25']",
        'context_node_name': 'Trunking System',
        'fields': {
            'Phase 2 Voice Capable': 'True'
        }
    },
       
    # -- TDMA Channel ID 3 --
    {
        'group_name': 'Trunking System - Channel ID 3',
        'base_xpath': ".//Recset[@Name='Trunking System']//EmbeddedNode[@ReferenceKey='Channel ID 3']",
        'context_node_name': 'Trunking System',
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
        'context_node_name': 'Trunking System',
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
        'context_node_name': 'Conventional Personality',
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
        'context_node_name': 'Conventional Personality',
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
        'context_node_name': 'Conventional Personality',
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

        # -- GW IO 1 --
    {
        'group_name': 'INTEROP - GW IO 1',
        'base_xpath': ".//Recset[@Name='Zone Channel Assignment']/Node[contains(@ReferenceKey, 'INTEROP')]//EmbeddedNode[@ReferenceKey='1-GW IO 1']",
        'context_node_name': 'Zone Channel Assignment',
        'fields': {
            'Channel Type': 'Trk',
            'Personality': '027A - IO',
            'Channel Name': 'GW IO 1',
            'Top Display Channel': 'GW IO 1',
            'Trunking Talkgroup': 'IO 1',
            'Active Channel': 'True'
        }
    },
        # -- GW IO 2 --
    {
        'group_name': 'INTEROP - GW IO 2',
        'base_xpath': ".//Recset[@Name='Zone Channel Assignment']/Node[contains(@ReferenceKey, 'INTEROP')]//EmbeddedNode[@ReferenceKey='2-GW IO 2']",
        'context_node_name': 'Zone Channel Assignment',
        'fields': {
            'Channel Type': 'Trk',
            'Personality': '027A - IO',
            'Channel Name': 'GW IO 2',
            'Top Display Channel': 'GW IO 2',
            'Trunking Talkgroup': 'IO 2',
            'Active Channel': 'True'
        }
    },
        # -- GW IO 3 --
    {
        'group_name': 'INTEROP - GW IO 3',
        'base_xpath': ".//Recset[@Name='Zone Channel Assignment']/Node[contains(@ReferenceKey, 'INTEROP')]//EmbeddedNode[@ReferenceKey='3-GW IO 3']",
        'context_node_name': 'Zone Channel Assignment',
        'fields': {
            'Channel Type': 'Trk',
            'Personality': '027A - IO',
            'Channel Name': 'GW IO 3',
            'Top Display Channel': 'GW IO 3',
            'Trunking Talkgroup': 'IO 3',
            'Active Channel': 'True'
        }
    },
        # -- GW IO 4 --
    {
        'group_name': 'INTEROP - GW IO 4',
        'base_xpath': ".//Recset[@Name='Zone Channel Assignment']/Node[contains(@ReferenceKey, 'INTEROP')]//EmbeddedNode[@ReferenceKey='4-GW IO 4']",
        'context_node_name': 'Zone Channel Assignment',
        'fields': {
            'Channel Type': 'Trk',
            'Personality': '027A - IO',
            'Channel Name': 'GW IO 4',
            'Top Display Channel': 'GW IO 4',
            'Trunking Talkgroup': 'IO 4',
            'Active Channel': 'True'
        }
    },

    # -- GW IO 5 --
    {
        'group_name': 'INTEROP - GW IO 5',
        'base_xpath': ".//Recset[@Name='Zone Channel Assignment']/Node[contains(@ReferenceKey, 'INTEROP')]//EmbeddedNode[@ReferenceKey='5-GW IO 5']",
        'context_node_name': 'Zone Channel Assignment',
        'fields': {
            'Channel Type': 'Trk',
            'Personality': '027A - IO',
            'Channel Name': 'GW IO 5',
            'Top Display Channel': 'GW IO 5',
            'Trunking Talkgroup': 'IO 5',
            'Active Channel': 'True'
        }
    },
 
    # -- GW IO 6 --
    {
        'group_name': 'INTEROP - GW IO 6',
        'base_xpath': ".//Recset[@Name='Zone Channel Assignment']/Node[contains(@ReferenceKey, 'INTEROP')]//EmbeddedNode[@ReferenceKey='6-GW IO 6']",
        'context_node_name': 'Zone Channel Assignment',
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
        'context_node_name': 'Zone Channel Assignment',
        'fields': {
            'Channel Type': 'Cnv',
            'Personality': '8TAC',
            'Channel Name': '8CALL90',
            'Top Display Channel': '8CALL90',
            'Active Channel': 'True'
        }
    },
]

def _extract_metadata(root):
    metadata={
        "alias": "",
        "gwinnett_id": 0
    }

    alias_xpath = ".//Recset[@Name='Radio Wide']//Field[@Name='User Information\\Radio Alias']"
    alias_elements = root.xpath(alias_xpath)
    if alias_elements and alias_elements[0].text:
        metadata["alias"] = alias_elements[0].text.strip()

    gwinnett_xpath = ".//Recset[@Name='Trunking System']/Node[@Name='Trunking System']/Section[@Name='General']/Field[@Name='Unit ID']"
    gwinnett_elements = root.xpath(gwinnett_xpath)
    if gwinnett_elements and gwinnett_elements[0].text:
        try:
            metadata["gwinnett_id"] = int(gwinnett_elements[0].text.strip())
        except (ValueError, TypeError):
            print(f"Warning: Could not convert Unit ID '{gwinnett_elements[0].text}' to an integer.")

    return metadata

def _process_check_group(root, group, metadata, serial, model, mobile_hh):
    error_rows = []
    group_name = group['group_name']
    parents = root.xpath(group['base_xpath'])
    if not parents:
        error_rows.append([serial, metadata['alias'], metadata['gwinnett_id'], "N/A", group_name, "N/A", "Section Missing", "N/A", "N/A", model, mobile_hh])
        return error_rows

    for parent in parents:
        system_context = "N/A"
        context_name = group.get('context_node_name') # Get the context to search for
        if context_name:
            context_xpath = f"ancestor::Node[@Name='{context_name}'][1]/@ReferenceKey"
            context_keys = parent.xpath(context_xpath)
            if context_keys:
                system_context = context_keys[0]

        for field_name, expected_value in group['fields'].items():
            if mobile_hh == 'Mobile' and field_name == 'Top Display Channel':
                continue # Skip this 'Top Display Channel': field for Mobile devices
            field_elements = parent.xpath(f".//Field[@Name='{field_name}']")

            if not field_elements:
                error_rows.append([serial, metadata['alias'], metadata['gwinnett_id'], system_context, group_name, field_name, "Setting Missing", expected_value, "N/A", model, mobile_hh])
                continue

            actual_value = field_elements[0].text or ""
            if actual_value != expected_value:
                error_rows.append([serial, metadata['alias'], metadata['gwinnett_id'], system_context, group_name, field_name, "Incorrect Value", expected_value, actual_value, model, mobile_hh])
                
    return error_rows

def _get_model_from_serial(serial):
    if serial.startswith('426'):
        return 4000, 'Handheld'
    elif serial.startswith('481'):
        return 6000, 'Handheld'
    elif serial.startswith('527'):
        return 6500, 'Mobile'
    elif serial.startswith('579'):
        return 8000, 'Handheld'
    elif serial.startswith('652'):
        return 8000, 'Mobile'
    elif serial.startswith('681'):
        return 8500, 'Mobile'
    elif serial.startswith('755'):
        return 6500, 'Handheld'
    elif serial.startswith('756'):
        return 6000, 'Handheld'
    elif serial.startswith('761'):
        return 7500, 'Mobile'
    else:
        return 0, 'Is Serial Correct?'

def _get_model_from_filename(serial):
    serial_upper = serial.upper()
    if '4000' in serial_upper:
        return 4000, 'Handheld'
    elif '6000' in serial_upper:
        return 6000, 'Handheld'
    elif '6500' in serial_upper:
        return 6500, 'Mobile'
    elif '8000' in serial_upper:
        return 8000, 'Handheld'
    elif '8500' in serial_upper:
        return 8500, 'Mobile'
    elif '7500' in serial_upper:
        return 7500, 'Mobile'
    else:
        return 0, 'Is Model in Filename?'

# Check XML file
def check_xml_file(filepath, report_rows):
    try:
        parser = ET.XMLParser(remove_blank_text=True, resolve_entities=False)
        tree = ET.parse(filepath, parser)
        root = tree.getroot()
        filename = os.path.basename(filepath)
        serial = filename.removesuffix('.xml')

        if len(serial)==10:
            model, mobile = _get_model_from_serial(serial)
        else:
            model, mobile = _get_model_from_filename(serial)

        metadata = _extract_metadata(root)
        all_discrepancies_in_file = []
        
        for group in CHECKS_TO_PERFORM:
            group_errors = _process_check_group(root, group, metadata, serial, model, mobile)
            if group_errors:
                all_discrepancies_in_file.extend(group_errors)
        if not all_discrepancies_in_file:
            success_row = [serial, metadata['alias'], metadata['gwinnett_id'], "OK", "OK", "OK", "OK", "OK", "OK", model, mobile]
            report_rows.append(success_row)
        else:
            report_rows.extend(all_discrepancies_in_file)

    except ET.XMLSyntaxError:
        report_rows.append([os.path.basename(filepath), "File Error", "Alias", "ID", "Sys", "Group","Could not parse XML", "value", "value", "model", "type"])

def main():
    print("Starting check...")
    xml_files = glob.glob('*.xml') # Find all XML files in folder
    total_files = len(xml_files)

    if not xml_files:
        print("No .xml files in this folder.")
        return

    report_filename = 'report.xlsx'
    print(f"Found {total_files} XML files. Checking settings...")

    report_rows = []

    # input each error in a new row
    for i, filepath in enumerate(xml_files):
        print(f"Processing file {i+1}/{total_files}: {os.path.basename(filepath)}")
        check_xml_file(filepath, report_rows)

    # Generate report
    with pd.ExcelWriter(report_filename, engine='openpyxl') as writer:
        if not report_rows: # No errors found
            sheet_name="No errors"
            if total_files == 1:
                message = 'All settings are correct in the XML file in this folder.'
            else:
                message = f'All settings are correct across all {total_files} files in this folder.'
            
            print(message)
            df = pd.DataFrame(columns=[message])
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            worksheet = writer.sheets[sheet_name]
            worksheet.column_dimensions['A'].width = len(message) + 5

        else: # Handle errors
            print(f"Total rows recorded: {len(report_rows)}")
            
            header = ['Serial', 'Alias', 'ID', 'Setting','Reference', 'Group','Problem', 'Expected', 'Actual', 'Model', 'Type']
            df = pd.DataFrame(report_rows, columns=header)
            sheet_name = "Report-On-XML-Files"
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            worksheet = writer.sheets[sheet_name]
            for col, column_title in enumerate(df.columns, 1):
                column_letter = get_column_letter(col)
                max_length = df[column_title].astype(str).map(len).max() # max length of content
                max_length = max(max_length, len(column_title)) + 1 # the column header may be longer
                worksheet.column_dimensions[column_letter].width = max_length # Set the column width

if __name__ == "__main__":
    main()