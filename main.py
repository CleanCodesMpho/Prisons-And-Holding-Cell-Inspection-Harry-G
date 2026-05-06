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
SURVEY_LAYER_URL = "https://services6.arcgis.com/345WScIubRHps95b/arcgis/rest/services/service_1a23300536014626b5f6fbcc21d3141b/FeatureServer/FeatureServer/0"
layer = FeatureLayer(SURVEY_LAYER_URL, gis=gis)

# =========================
# TEMPLATE
# =========================
TEMPLATE_PATH = "Prisons_and_Holding_cells_Report.docx"

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
        "MunicName": attributes.get("MunicName", "N/A"),
        "ward_no": attributes.get("ward_no", "N/A"),
        "address__": attributes.get("address__", "N/A"),
        "premise_namee": attributes.get("premise_name", "N/A"),
        "Name_Owner": attributes.get("Name_Ownerr", "N/A"),
        "Names": attributes.get("Names", "N/A"),
        "Description": attributes.get("Description", "N/A"),
        "Owner_email": attributes.get("Owner_email", "N/A"),
        "tell_no": attributes.get("tell_no", "N/A"),
        "facility_manager": attributes.get("facility_manager", "N/A"),
        "telephone_number": attributes.get("telephone_number", "N/A"),
        "Supporting_Docs": attributes.get("Supporting_Docs", "N/A"),
        "inspection_date": edit_date,
        "village_town": attributes.get("village_town", "N/A"),

        "health_certificate": attributes.get("health_certificate", "N/A"),
        "comment0": attributes.get("comment0", "N/A"),
        "certificate_acceptability": attributes.get("certificate_acceptability", "N/A"),
        "comment1": attributes.get("comment1", "N/A"),
        "certificate_competency": attributes.get("certificate_competency", "N/A"),
        "comment2": attributes.get("comment2", "N/A"),
        "pest_control": attributes.get("pest_control", "N/A"),
        "comment3": attributes.get("comment3", "N/A"),
        "health_care": attributes.get("health_care", "N/A"),
        "comment4": attributes.get("comment4", "N/A"),

        "ceilings": attributes.get("ceilings", "N/A"),
        "comment7": attributes.get("comment7", "N/A"),
        "internal_walls": attributes.get("internal_walls", "N/A"),
        "comment8": attributes.get("comment8", "N/A"),
        "floor_surfaces": attributes.get("floor_surfaces", "N/A"),
        "comment9": attributes.get("comment9", "N/A"),
        "adequate_vent": attributes.get("adequate_vent", "N/A"),
        "comment10": attributes.get("comment10", "N/A"),
        "inmates": attributes.get("inmates", "N/A"),
        "comment11": attributes.get("comment11", "N/A"),

        "showers_provided": attributes.get("showers_provided", "N/A"),
        "comment12": attributes.get("comment12", "N/A"),
        "washing_basins": attributes.get("washing_basins", "N/A"),
        "comment13": attributes.get("comment13", "N/A"),
        "wash_up": attributes.get("wash_up", "N/A"),
        "comment14": attributes.get("comment14", "N/A"),
        "cleanable_material": attributes.get("cleanable_material", "N/A"),
        "comment15": attributes.get("comment15", "N/A"),
        "good_repairs": attributes.get("good_repairs", "N/A"),
        "comment16": attributes.get("comment16", "N/A"),

        "facilities_adequate": attributes.get("facilities_adequate", "N/A"),
        "comment17": attributes.get("comment17", "N/A"),
        "personal_belongings": attributes.get("personal_belongings", "N/A"),
        "comment18": attributes.get("comment18", "N/A"),
        "hazardous_substance": attributes.get("hazardous_substance", "N/A"),
        "comment19": attributes.get("comment19", "N/A"),
        "above_floor": attributes.get("above_floor", "N/A"),
        "comment20": attributes.get("comment20", "N/A"),

        "every_week": attributes.get("every_week", "N/A"),
        "comment21": attributes.get("comment21", "N/A"),
        "linen_provided": attributes.get("linen_provided", "N/A"),
        "comment22": attributes.get("comment22", "N/A"),
        "food_items": attributes.get("food_items", "N/A"),
        "comment23": attributes.get("comment23", "N/A"),
        "refridgeration": attributes.get("refridgeration", "N/A"),
        "comment24": attributes.get("comment24", "N/A"),

        "laundry0": attributes.get("laundry0", "N/A"),
        "comment25": attributes.get("comment25", "N/A"),
        "internal_walls1": attributes.get("internal_walls1", "N/A"),
        "comment26": attributes.get("comment26", "N/A"),
        "clean": attributes.get("clean", "N/A"),
        "comment27": attributes.get("comment27", "N/A"),
        "proof_material": attributes.get("proof_material", "N/A"),
        "comment28": attributes.get("comment28", "N/A"),

        "illumination": attributes.get("illumination", "N/A"),
        "comment29": attributes.get("comment29", "N/A"),
        "drainage_systems": attributes.get("drainage_systems", "N/A"),
        "comment30": attributes.get("comment30", "N/A"),

        "clean_linen": attributes.get("clean_linen", "N/A"),
        "comment31": attributes.get("comment31", "N/A"),
        "vehicles_containers": attributes.get("vehicles_containers", "N/A"),
        "comment32": attributes.get("comment32", "N/A"),
        "Standard_op": attributes.get("Standard_op", "N/A"),
        "comment33": attributes.get("comment33", "N/A"),
        
        "premises": attributes.get("premises", "N/A"),
        "comment34": attributes.get("comment34", "N/A"),
        "cleaning_materials": attributes.get("cleaning_materials", "N/A"),
        "comment35": attributes.get("comment35", "N/A"),
        "the_facility": attributes.get("the_facility", "N/A"),
        "comment36": attributes.get("comment36", "N/A"),
        "cleaning_mater1": attributes.get("cleaning_mater1", "N/A"),
        "comment37": attributes.get("comment37", "N/A"),
        
        "water_avail": attributes.get("water_avail", "N/A"),
        "comment38": attributes.get("comment38", "N/A"),
        "comment39": attributes.get("comment39", "N/A"),
        "water_sources": attributes.get("water_sources", "N/A"),
        "comment40": attributes.get("comment40", "N/A"),
        "risk_management": attributes.get("risk_management", "N/A"),
        "comment41": attributes.get("comment41", "N/A"),
        "taps_and_pipes": attributes.get("taps_and_pipes", "N/A"),
        "comment42": attributes.get("comment42", "N/A"),
        "water_storage2": attributes.get("water_storage2", "N/A"),
        "comment43": attributes.get("comment43", "N/A"),
        "comply_kitchen": attributes.get("comply_kitchen", "N/A"),
        "comment44": attributes.get("comment44", "N/A"),
        "plan_in_place": attributes.get("plan_in_place", "N/A"),
        "comment45": attributes.get("comment45", "N/A"),
        "general_waste": attributes.get("general_waste", "N/A"),
        "comment46": attributes.get("comment46", "N/A"),

        "waste_cage": attributes.get("waste_cage", "N/A"),
        "comment47": attributes.get("comment47", "N/A"),
        "vector_control1": attributes.get("vector_control1", "N/A"),
        "comment48": attributes.get("comment48", "N/A"),
        "recomm": attributes.get("recomm", "N/A"),
        "risk_rating": attributes.get("risk_rating", "N/A"),
        "compliance": attributes.get("compliance", "N/A"),
        "EHP": attributes.get("EHP", "N/A"),
        "email": attributes.get("email", "N/A"),
        "contact": attributes.get("contact", "N/A"),
        "signature": attributes.get("signature", "N/A"),
        
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
