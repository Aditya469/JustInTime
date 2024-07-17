import os
import json

def parse_edi_to_json(edi_content):
    segments = edi_content.strip().split("'")
    json_data = {
        "interchange": {
            "header": {},
            "message": {
                "header": {},
                "beginning": {},
                "dateTimePeriod": "",
                "parties": [],
                "processingIndicator": "",
                "lineItems": []
            },
            "trailer": {}
        },
        "trailer": {}
    }

    current_line_item = None

    for segment in segments:
        if segment.strip():
            fields = segment.split("+")
            segment_tag = fields[0]

            if segment_tag == "UNB":
                json_data["interchange"]["header"] = {
                    "syntaxIdentifier": fields[1],
                    "sender": fields[2],
                    "recipient": fields[3],
                    "preparationDateTime": fields[4].split(":")[0],
                    "controlReference": fields[5].strip("'")
                }
            elif segment_tag == "UNH":
                json_data["interchange"]["message"]["header"] = {
                    "referenceNumber": fields[1],
                    "type": fields[2]
                }
            elif segment_tag == "BGM":
                json_data["interchange"]["message"]["beginning"] = {
                    "name": fields[1],
                    "number": fields[2],
                    "function": fields[3]
                }
            elif segment_tag == "DTM" and fields[1].startswith("137"):
                json_data["interchange"]["message"]["dateTimePeriod"] = fields[1].split(":")[1]
            elif segment_tag == "NAD":
                party = {
                    "qualifier": fields[1],
                    "identification": fields[2].split("::")[0],
                    "name": " ".join(fields[3:])
                }
                json_data["interchange"]["message"]["parties"].append(party)
            elif segment_tag == "GIS":
                json_data["interchange"]["message"]["processingIndicator"] = fields[1]
            elif segment_tag == "LIN":
                current_line_item = {
                    "number": fields[3],
                    "quantities": []
                }
                json_data["interchange"]["message"]["lineItems"].append(current_line_item)
            elif segment_tag == "PIA":
                current_line_item["productId"] = fields[3]
            elif segment_tag == "IMD":
                current_line_item["description"] = " ".join(fields[4:])
            elif segment_tag == "LOC":
                current_line_item["location"] = fields[2]
            elif segment_tag == "RFF":
                if fields[1] == "ON":
                    current_line_item["reference"] = fields[2]
                elif fields[1] == "AAK":
                    current_line_item["reference2"] = fields[2]
            elif segment_tag == "FTX":
                current_line_item["freeText"] = " ".join(fields[4:])
            elif segment_tag == "QTY":
                if fields[1].startswith("12"):
                    current_line_item["quantity"] = int(fields[1].split(":")[1])
                elif fields[1].startswith("1"):
                    quantity = {
                        "quantity": int(fields[1].split(":")[1]),
                        "dateTimePeriod": ""
                    }
                    current_line_item["quantities"].append(quantity)
            elif segment_tag == "DTM" and fields[1].startswith("2"):
                if current_line_item["quantities"]:
                    current_line_item["quantities"][-1]["dateTimePeriod"] = fields[1].split(":")[1]
                else:
                    current_line_item["dateTimePeriod2"] = fields[1].split(":")[1]
            elif segment_tag == "SCC":
                if current_line_item["quantities"]:
                    current_line_item["quantities"][-1]["deliveryPlanStatus"] = fields[1]
                    current_line_item["quantities"][-1]["deliveryRequirements"] = fields[2]
            elif segment_tag == "UNT":
                json_data["interchange"]["trailer"] = {
                    "segmentsCount": fields[1],
                    "referenceNumber": fields[2].strip("'")
                }
            elif segment_tag == "UNZ":
                json_data["trailer"] = {
                    "controlCount": fields[1],
                    "controlReference": fields[2].strip("'")
                }

    return json_data

def convert_edi_to_json(path_directory):
    edi_directory = path_directory
    json_directory = os.path.join(path_directory, "JSONconverted")

    if not os.path.exists(json_directory):
        os.makedirs(json_directory)

    for filename in os.listdir(edi_directory):
        if filename.endswith(".EDI") or filename.endswith(".txt"):
            edi_file_path = os.path.join(edi_directory, filename)
            json_file_path = os.path.join(json_directory, f"{os.path.splitext(filename)[0]}.json")

            with open(edi_file_path, "r") as edi_file:
                edi_content = edi_file.read()

            json_data = parse_edi_to_json(edi_content)

            with open(json_file_path, "w") as json_file:
                json.dump(json_data, json_file, indent=2)

            print(f"Converted {filename} to {os.path.basename(json_file_path)}")

# Specify the path directory containing the .EDI files
path_directory = "C:/Users/Rama/JustInTime/EDI Files"

# Call the function to convert EDI files to JSON
convert_edi_to_json(path_directory)
