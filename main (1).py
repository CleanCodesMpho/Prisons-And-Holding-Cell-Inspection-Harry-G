from fastapi import FastAPI, Request
from arcgis.gis import GIS
from arcgis.features import FeatureLayer
import os
import qrcode
from datetime import datetime

from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

# =========================
# APP INIT
# =========================
app = FastAPI()

# =========================
# AGOL AUTH
# =========================
AGOL_USERNAME = os.getenv("AGOL_USERNAME")
AGOL_PASSWORD = os.getenv("AGOL_PASSWORD")

if not AGOL_USERNAME or not AGOL_PASSWORD:
    raise Exception("AGOL credentials not set in environment variables")

gis = GIS("https://www.arcgis.com", AGOL_USERNAME, AGOL_PASSWORD)

# =========================
# FEATURE LAYER
# =========================
SURVEY_LAYER_URL = "https://services6.arcgis.com/345WScIubRHps95b/arcgis/rest/services/service_bd43c481ce0345febf5fc02b8ec3b09f/FeatureServer/FeatureServer/0"
layer = FeatureLayer(SURVEY_LAYER_URL, gis=gis)

# =========================
# TEMPLATE
# =========================
TEMPLATE_PATH = "accomodation_establishments_Report.docx"

if not os.path.exists(TEMPLATE_PATH):
    raise Exception(f"{TEMPLATE_PATH} not found in project root")

# =========================
# TEMP PAYLOAD STORAGE
# =========================
LAST_PAYLOAD = {}
LAST_ERROR = None

# =========================
# HEALTH CHECK
# =========================
@app.get("/")
def home():
    return {"status": "running"}

# =========================
# DEBUG ENDPOINT
# =========================
@app.get("/debug")
def debug():
    return {
        "template_exists": os.path.exists(TEMPLATE_PATH),
        "username_set": bool(AGOL_USERNAME),
        "password_set": bool(AGOL_PASSWORD),
        "layer_url": SURVEY_LAYER_URL
    }

# =========================
# LAST PAYLOAD
# =========================
@app.get("/last-payload")
def last_payload():
    return {
        "last_error": LAST_ERROR,
        "payload": LAST_PAYLOAD
    }

# =========================
# TEST QUERY
# =========================
@app.get("/test-query/{objectid}")
def test_query(objectid: int):
    result = layer.query(where=f"OBJECTID={objectid}", out_fields="*")
    return {
        "found": len(result.features),
        "attributes": result.features[0].attributes if result.features else None
    }

# =========================
# TEST UPDATE
# =========================
@app.get("/test-update/{objectid}")
def test_update(objectid: int):
    result = layer.edit_features(updates=[{
        "attributes": {
            "OBJECTID": objectid,
            "report_status": "test_ok",
            "report_url": "https://example.com/test.docx"
        }
    }])
    return {"edit_result": result}

# =========================
# HELPER: EXTRACT OBJECTID
# =========================
def extract_objectid(payload):
    if "submittedRecord" in payload:
        attrs = payload["submittedRecord"].get("attributes", {})
        if "OBJECTID" in attrs:
            return attrs["OBJECTID"]

    if "serverResponse" in payload:
        sr = payload["serverResponse"]
        if isinstance(sr, dict):
            if "objectId" in sr:
                return sr["objectId"]
            if "editResults" in sr and sr["editResults"]:
                first = sr["editResults"][0]
                if "objectId" in first:
                    return first["objectId"]

    if "feature" in payload:
        feature = payload["feature"]
        if isinstance(feature, dict):
            attrs = feature.get("attributes", {})
            if "OBJECTID" in attrs:
                return attrs["OBJECTID"]
            result = feature.get("result", {})
            if "objectId" in result:
                return result["objectId"]

    if "features" in payload and payload["features"]:
        first = payload["features"][0]
        attrs = first.get("attributes", {})
        if "OBJECTID" in attrs:
            return attrs["OBJECTID"]

    if isinstance(payload, dict):
        for key, value in payload.items():
            if key in ("OBJECTID", "objectId"):
                return value
            found = extract_objectid(value)
            if found is not None:
                return found

    if isinstance(payload, list):
        for item in payload:
            found = extract_objectid(item)
            if found is not None:
                return found

    return None

# =========================
# QR GENERATOR
# =========================
def generate_qr(url, path):
    img = qrcode.make(url)
    img.save(path)

# =========================
# UPLOAD REPORT TO AGOL
# =========================
def upload_report_to_agol(file_path, objectid):
    root_folder = gis.content.folders.get()

    item_properties = {
        "title": f"Report_{objectid}",
        "type": "Microsoft Word",
        "tags": ["survey123", "report", "automation"],
        "snippet": f"Automatically generated report for Survey123 submission {objectid}"
    }

    report_item = root_folder.add(
        item_properties=item_properties,
        file=file_path
    ).result()

    report_item.sharing.sharing_level = "EVERYONE"

    return f"https://www.arcgis.com/home/item.html?id={report_item.itemid}"

# =========================
# REPORT GENERATION
# =========================
def generate_report(attributes, objectid):
    os.makedirs("output", exist_ok=True)

    docx_file = os.path.join("output", f"report_{objectid}.docx")
    qr_file = os.path.join("output", f"qr_{objectid}.png")

    # Temporary QR target for first render
    temp_url = f"https://www.arcgis.com/home/item.html?id=temp-{objectid}"
    generate_qr(temp_url, qr_file)

    edit_date = attributes.get("EditDate")
    if edit_date:
        edit_date = datetime.fromtimestamp(edit_date / 1000).strftime("%Y-%m-%d %H:%M:%S")
    else:
        edit_date = "N/A"

    doc = DocxTemplate(TEMPLATE_PATH)
    qr_image = InlineImage(doc, qr_file, width=Mm(25))

    context = {
        "local_munic_name": attributes.get("local_munic_name", "N/A"),
        "site_name": attributes.get("site_name", "N/A"),
        "address": attributes.get("address", "N/A"),
        "site_type": attributes.get("site_type", "N/A"),
        "Name_of_owner": attributes.get("Name_of_owner", "N/A"),
        "Surname_of_Owner": attributes.get("Surname_of_Owner", "N/A"),
        "id_number": attributes.get("id_number", "N/A"),
        "registration_number_of_farm": attributes.get("registration_number_of_farm", "N/A"),
        "address_occupier": attributes.get("address_occupier", "N/A"),
        "Description": attributes.get("Description", "N/A"),
        "telephone_number": attributes.get("telephone_number", "N/A"),
        "email_address": attributes.get("email_address", "N/A"),
        "inspection_date": edit_date,
        "township_village": attributes.get("township_village", "N/A"),

        "number_of_males": attributes.get("number_of_males", "N/A"),
        "number_of_females": attributes.get("number_of_females", "N/A"),
        "total_number_of_animals": attributes.get("total_number_of_animals", "N/A"),
        "number_animals_lactation": attributes.get("number_animals_lactation", "N/A"),
        "annual_milk_production": attributes.get("annual_milk_production", "N/A"),
        "type_of_milking_parlour": attributes.get("type_of_milking_parlour", "N/A"),
        "milking_times_per_day": attributes.get("milking_times_per_day", "N/A"),
        "collection_frequency": attributes.get("collection_frequency", "N/A"),
        "milk_sample_frequency_collection": attributes.get("milk_sample_frequency_collection", "N/A"),
        "local_authority": attributes.get("local_authority", "N/A"),

        "borehole_": attributes.get("borehole_", "N/A"),
        "surface_water": attributes.get("surface_water", "N/A"),
        "electricity_supply_generator": attributes.get("electricity_supply_generator", "N/A"),
        "name_of_veterinarian": attributes.get("name_of_veterinarian", "N/A"),
        "frequency_of_visits": attributes.get("frequency_of_visits", "N/A"),
        "visit_date_1": attributes.get("visit_date_1", "N/A"),
        "visit_date_2": attributes.get("visit_date_2", "N/A"),
        "visit_date_3": attributes.get("visit_date_3", "N/A"),
        "visit_date_4": attributes.get("visit_date_4", "N/A"),
        "certificate_acceptibility": attributes.get("certificate_acceptibility", "N/A"),

        "certificate_acceptibility_coms": attributes.get("certificate_acceptibility_coms", "N/A"),
        "animal_examination_coms": attributes.get("animal_examination_coms", "N/A"),
        "animal_examination_registration": attributes.get("animal_examination_registration", "N/A"),
        "animal_health_certificates_tb": attributes.get("animal_health_certificates_tb", "N/A"),
        "animal_health_coms": attributes.get("animal_health_coms", "N/A"),
        "animal_health_certificates_bm": attributes.get("animal_health_certificates_bm", "N/A"),
        "animal_health_bm_coms": attributes.get("animal_health_bm_coms", "N/A"),
        "register_vaccinations": attributes.get("register_vaccinations", "N/A"),
        "register_vaccinations_coms": attributes.get("register_vaccinations_coms", "N/A"),
        "register_anti_parasitic": attributes.get("register_anti_parasitic", "N/A"),

        "register_anti_parasitic_coms": attributes.get("register_anti_parasitic_coms", "N/A"),
        "register_antibitiocs_used": attributes.get("register_antibitiocs_used", "N/A"),
        "register_antibitiocs_coms": attributes.get("register_antibitiocs_coms", "N/A"),
        "antibiotic_withdrawal_periods": attributes.get("antibiotic_withdrawal_periods", "N/A"),
        "antibiotic_withdrawal_coms": attributes.get("antibiotic_withdrawal_coms", "N/A"),
        "separate_milking": attributes.get("separate_milking", "N/A"),
        "separate_milking_coms": attributes.get("separate_milking_coms", "N/A"),
        "record_other_treatment": attributes.get("record_other_treatment", "N/A"),

        "other_treatment_coms": attributes.get("other_treatment_coms", "N/A"),
        "milking_bulk_tank": attributes.get("milking_bulk_tank", "N/A"),
        "milking_bulk_tank_coms": attributes.get("milking_bulk_tank_coms", "N/A"),
        "dairy_milk_tank_temperature": attributes.get("dairy_milk_tank_temperature", "N/A"),
        "dairy_milk_tank_temp_coms": attributes.get("dairy_milk_tank_temp_coms", "N/A"),
        "milk_collection_daily_second_base": attributes.get("milk_collection_daily_second_base", "N/A"),
        "milk_collection_daily_coms": attributes.get("milk_collection_daily_coms", "N/A"),
        "bulk_tank_sample": attributes.get("bulk_tank_sample", "N/A"),

        "bulk_tank_sample_coms": attributes.get("bulk_tank_sample_coms", "N/A"),
        "udder_washing_": attributes.get("udder_washing_", "N/A"),
        "udder_washing_coms": attributes.get("udder_washing_coms", "N/A"),
        "teat_dipping": attributes.get("teat_dipping", "N/A"),
        "teat_dipping_coms": attributes.get("teat_dipping_coms", "N/A"),
        "hygiene_cleaning_schedule": attributes.get("hygiene_cleaning_schedule", "N/A"),
        "hygiene_cleaning_schedule_coms": attributes.get("hygiene_cleaning_schedule_coms", "N/A"),
        "equipment_maintenance_programme": attributes.get("equipment_maintenance_programme", "N/A"),

        "equipment_maintenance_coms": attributes.get("equipment_maintenance_coms", "N/A"),
        "hygiene__schedule_milking_shed": attributes.get("hygiene__schedule_milking_shed", "N/A"),
        "hygiene_schedule_coms": attributes.get("hygiene_schedule_coms", "N/A"),
        "hand_washing_procedure": attributes.get("hand_washing_procedure", "N/A"),

         "hand_washing_procedure_coms": attributes.get("hand_washing_procedure_coms", "N/A"),
        "water_source_originates": attributes.get("water_source_originates", "N/A"),
        "water_source_origin_coms": attributes.get("water_source_origin_coms", "N/A"),
        "water_source_register": attributes.get("water_source_register", "N/A"),
        "water_source_register_coms": attributes.get("water_source_register_coms", "N/A"),
        "water_chemical_analysis": attributes.get("water_chemical_analysis", "N/A"),
        "water_chemical_analysis_coms": attributes.get("water_chemical_analysis_coms", "N/A"),
        "water_chlorine_analysis": attributes.get("water_chlorine_analysis", "N/A"),
        "water_chlorine_analysis_comms": attributes.get("water_chlorine_analysis_comms", "N/A"),
        "total_bacterial_counts": attributes.get("total_bacterial_counts", "N/A"),
        "total_bacterial_counts_coms": attributes.get("total_bacterial_counts_coms", "N/A"),
        "total_somatic_cellcount": attributes.get("total_somatic_cellcount", "N/A"),
        "total_somatic_cellcount_coms": attributes.get("total_somatic_cellcount_coms", "N/A"),
        "communicable_diseases": attributes.get("communicable_diseases", "N/A"),
        "communicable_diseases_coms": attributes.get("communicable_diseases_coms", "N/A"),
        "record_staff_sore_abscess_": attributes.get("record_staff_sore_abscess_", "N/A"),
        "record_staff_sore_abscess_coms": attributes.get("record_staff_sore_abscess_coms", "N/A"),
        "contract_with_buyer_processor": attributes.get("contract_with_buyer_processor", "N/A"),
        "contract_with_buyer_coms": attributes.get("contract_with_buyer_coms", "N/A"),
        "list_animals_id": attributes.get("list_animals_id", "N/A"),
        "list_animals_id_coms": attributes.get("list_animals_id_coms", "N/A"),
        "maintenance_programme_record": attributes.get("maintenance_programme_record", "N/A"),
        "maintenance_record_coms": attributes.get("maintenance_record_coms", "N/A"),
        "": attributes.get("", "N/A"),
        "": attributes.get("", "N/A"),
        "": attributes.get("", "N/A"),
        "": attributes.get("", "N/A"),
        "": attributes.get("", "N/A"),
        "": attributes.get("", "N/A"),
        "": attributes.get("", "N/A"),
        "": attributes.get("", "N/A"),

        "": attributes.get("", "N/A"),
        "": attributes.get("", "N/A"),
        "": attributes.get("", "N/A"),
        "": attributes.get("", "N/A"),
        "": attributes.get("", "N/A"),
        "": attributes.get("", "N/A"),
        "": attributes.get("", "N/A"),
        "": attributes.get("", "N/A"),
        "": attributes.get("", "N/A"),
        "": attributes.get("", "N/A"),
        "": attributes.get("", "N/A"),
        "": attributes.get("", "N/A"),
        "": attributes.get("", "N/A"),
        "": attributes.get("", "N/A"),
        "": attributes.get("", "N/A"),
        "": attributes.get("", "N/A"),
        "thermometer_": attributes.get("thermometer_", "N/A"),
        "comments56": attributes.get("comments56", "N/A"),
        "recording_bulk_milk": attributes.get("recording_bulk_milk", "N/A"),
        "comments57": attributes.get("comments57", "N/A"),
        "temperature_during_inspection": attributes.get("temperature_during_inspection", "N/A"),
        "comments58": attributes.get("comments58", "N/A"),

        "cooling_system_used": attributes.get("cooling_system_used", "N/A"),
        "comments59": attributes.get("comments59", "N/A"),
        "plate_cooler": attributes.get("plate_cooler", "N/A"),
        "comments60": attributes.get("comments60", "N/A"),
        "cooling_through_agitation": attributes.get("cooling_through_agitation", "N/A"),
        "comments61": attributes.get("comments61", "N/A"),
        "effective_agitation": attributes.get("effective_agitation", "N/A"),
        "comments62": attributes.get("comments62", "N/A"),

        "cleaning_equipment": attributes.get("cleaning_equipment", "N/A"),
        "comments63": attributes.get("comments63", "N/A"),
        "cleaning_programme": attributes.get("cleaning_programme", "N/A"),
        "comments64": attributes.get("comments64", "N/A"),
        "cip_cleaning2": attributes.get("cip_cleaning2", "N/A"),
        "comments65": attributes.get("comments65", "N/A"),
        "manual_cleaning_utensils2": attributes.get("manual_cleaning_utensils2", "N/A"),
        "comments66": attributes.get("comments66", "N/A"),
        "manual_cleaning_bulk_tank": attributes.get("manual_cleaning_bulk_tank", "N/A"),
        "comments67": attributes.get("comments67", "N/A"),
        "chemical_stored_dedicated_area": attributes.get("chemical_stored_dedicated_area", "N/A"),
        "comments68": attributes.get("comments68", "N/A"),
        "disposal_effluent_": attributes.get("disposal_effluent_", "N/A"),
        "comments69": attributes.get("comments69", "N/A"),
        
        "daily_provision_clean_clothing": attributes.get("daily_provision_clean_clothing", "N/A"),
        "comments70": attributes.get("comments70", "N/A"),
        "general_hygiene": attributes.get("general_hygiene", "N/A"),
        "comments71": attributes.get("comments71", "N/A"),
        "cuts_wounds_covered_": attributes.get("cuts_wounds_covered_", "N/A"),
        "comments72": attributes.get("comments72", "N/A"),
        "cleans_hands_arms_": attributes.get("cleans_hands_arms_", "N/A"),
        "comments73": attributes.get("comments73", "N/A"),
        "clean_tidy": attributes.get("clean_tidy", "N/A"),
        "comments74": attributes.get("comments74", "N/A"),
        "hand_washing_facilities_soap": attributes.get("hand_washing_facilities_soap", "N/A"),
        "comments75": attributes.get("comments75", "N/A"),

        "paper_towel_hand_drying": attributes.get("paper_towel_hand_drying", "N/A"),
        "comments76": attributes.get("comments76", "N/A"),
        "appropriate_restrooms": attributes.get("appropriate_restrooms", "N/A"),
        "comments77": attributes.get("comments77", "N/A"),
        "staff_boarding_facilities": attributes.get("staff_boarding_facilities", "N/A"),
        "comments78": attributes.get("comments78", "N/A"),
        "provision_protective_clothing": attributes.get("provision_protective_clothing", "N/A"),
        "comments79": attributes.get("comments79", "N/A"),
        "access_clean_water": attributes.get("access_clean_water", "N/A"),
        "comments80": attributes.get("comments80", "N/A"),

        "grass_cut_short": attributes.get("grass_cut_short", "N/A"),
        "comments81": attributes.get("comments81", "N/A"),
        "gutters_clean": attributes.get("gutters_clean", "N/A"),
        "comments83": attributes.get("comments83", "N/A"),
        "litter_waste_stored": attributes.get("litter_waste_stored", "N/A"),
        "comments82": attributes.get("comments82", "N/A"),
        "no_stagnat_water": attributes.get("no_stagnat_water", "N/A"),
        "comments84": attributes.get("comments84", "N/A"),
        "refuse_disposal": attributes.get("refuse_disposal", "N/A"),
        "comments85": attributes.get("comments85", "N/A"),

        "access_permitted_with_fence": attributes.get("access_permitted_with_fence", "N/A"),
        "comments86": attributes.get("comments86", "N/A"),
        "pest_proof": attributes.get("pest_proof", "N/A"),
        "comments87": attributes.get("comments87", "N/A"),
        "no_nesting_on_roof": attributes.get("no_nesting_on_roof", "N/A"),
        "comments88": attributes.get("comments88", "N/A"),
        "engineering_equipment": attributes.get("engineering_equipment", "N/A"),
        "comments89": attributes.get("comments89", "N/A"),
        "total_pop_registered": attributes.get("total_pop_registered", "N/A"),
        "total_unregistered_pop": attributes.get("total_unregistered_pop", "N/A"),
        
        "compliance_status": attributes.get("compliance_status", "N/A"),
        "recommedations_": attributes.get("recommedations_", "N/A"),
        "person_incharge": attributes.get("person_incharge", "N/A"),
        "_signature": attributes.get("_signature", "N/A"),
        "email_address": attributes.get("email_address", "N/A"),
        "additional_pictures": attributes.get("additional_pictures", "N/A"),
        "risk_rating": attributes.get("risk_rating", "N/A"),
        "action_taken": attributes.get("action_taken", "N/A"),
        "additional_pictures": attributes.get("additional_pictures", "N/A"),
        "_name_and_surname": attributes.get("_name_and_surname", "N/A"),
      
        "qr_code": qr_image
    }

    doc.render(context)
    doc.save(docx_file)

    real_url = upload_report_to_agol(docx_file, objectid)
    return real_url

# =========================
# UPDATE FEATURE
# =========================
def update_feature(objectid, url, status):
    result = layer.edit_features(updates=[{
        "attributes": {
            "OBJECTID": objectid,
            "report_url": url,
            "report_status": status
        }
    }])
    return result

# =========================
# WEBHOOK ENDPOINT
# =========================
@app.post("/webhook/survey123")
async def survey_webhook(request: Request):
    global LAST_PAYLOAD, LAST_ERROR

    payload = await request.json()
    LAST_PAYLOAD = payload
    LAST_ERROR = None
    objectid = None

    try:
        objectid = extract_objectid(payload)

        if objectid is None:
            LAST_ERROR = f"OBJECTID not found. Payload keys: {list(payload.keys()) if isinstance(payload, dict) else 'not a dict'}"
            return {
                "status": "failed",
                "error": LAST_ERROR
            }

        update_feature(objectid, "webhook_received", "received")

        result = layer.query(where=f"OBJECTID={objectid}", out_fields="*")

        if not result.features:
            update_feature(objectid, "query_failed", "failed")
            LAST_ERROR = f"No feature found for OBJECTID {objectid}"
            return {
                "status": "failed",
                "error": LAST_ERROR
            }

        attributes = result.features[0].attributes

        update_feature(objectid, "query_ok", "queried")

        report_url = generate_report(attributes, objectid)

        edit_result = update_feature(objectid, report_url, "completed")

        return {
            "status": "success",
            "objectid": objectid,
            "report_url": report_url,
            "edit_result": str(edit_result)
        }

    except Exception as e:
        LAST_ERROR = str(e)
        if objectid is not None:
            try:
                update_feature(objectid, f"ERROR: {str(e)}", "failed")
            except Exception:
                pass

        return {
            "status": "failed",
            "objectid": objectid,
            "error": str(e)
        }
