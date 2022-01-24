"""
author: evan fedorko, evanjfedorko@gmail.com
date: 6/2021

This script ggenerate PDF reports from data in Pandas DataFrames using jinja and pdfkit.
someone needed (very) simple documents recording the content of a survey but needed it for
~50 items so I made this to get the job done relatively quickly.

"""
import pandas as pan
from jinja2 import Environment, FileSystemLoader
import pdfkit
import os

# inputs
cwd = os.getcwd()
infile = cwd + r"\bbt_final_2.xls"


def cond(check, structure, vacant):
    if check == "Vacant lot":
        return vacant
    else:
        return structure


def descrip(descrip, other):
    if descrip == "Other (Describe)":
        return other
    else:
        return descrip


df = pan.read_excel(infile, sheet_name="sample", header=0)

for rec in df.index:
    objectid = str(df.objectid[rec])
    print("working on objectid " + objectid)
    address = str(df.SAMSAddres[rec]) + ", " + str(df.SAMSCity[rec]) + ", " + str(
        df.SAMSState[rec]) + ", " + str(df.SAMSZip[rec])  # df.FULLADDR[rec]
    struct_type = descrip(
        df.property_description[rec], df.property_description_other[rec])
    condition = cond(
        df.property_description[rec], df.overall_condition_of_structure[rec], df.overall_condition_of_vacant_lot[rec])
    image = df.ATT_NAME[rec]
    date = str(df.date_autofilled[rec]).split(' ')[0]
    occupied = df.does_the_property_appear_to_be_[rec]
    lot_size = str(round(df.acres_c[rec], 3)) + " acres"
    parcel_num = df.GISPID[rec]
    lot_desc = df.FullLegalD[rec]
    # .split is used to remove description after :, which is general rather than specific
    bldg_fr = str(df.condition_of_building_framestru[rec]).split(':')[0]
    windows = str(df.condition_of_windowsdoors[rec]).split(':')[0]
    porch = str(df.condition_of_porchentrance_[rec]).split(':')[0]
    roof = str(df.condition_of_roofgutterschimney[rec]).split(':')[0]
    siding = str(df.condition_of_sidingveneerpaint[rec]).split(':')[0]
    owner = df.FullOwnerN[rec]
    lon_lat = str(round(df.longitude[rec], 5)) + \
        ", " + str(round(df.latitude[rec], 5))
    own_address = df.FullOwnerA[rec]
    building_size = str(df.StructureA[rec]) + " sq. ft."
    hood = df.DistrictNa[rec]

    env = Environment(loader=FileSystemLoader(cwd))
    template = env.get_template("one_page.html")
    template_vars = {"address": address,
                     "type": struct_type,
                     "condition": condition,
                     "rating": "TBD",
                     "image": image,
                     "date": date,
                     "occupied": occupied,
                     "years": "Unknown",
                     "lot_size": lot_size,
                     "building_size": building_size,
                     "parcel_num": parcel_num,
                     "reg": "TBD",
                     "hood": hood,
                     "lot_desc": lot_desc,
                     "lon_lat": lon_lat,
                     "bldg_fr": bldg_fr,
                     "windows": windows,
                     "porch": porch,
                     "roof": roof,
                     "siding": siding,
                     "owner": owner,
                     "own_address": own_address}
    # "vacant_lot": vacant_lot}
    html_out = template.render(template_vars)
    html_name = cwd + r"\One_Page_" + \
        df.FULLADDR[rec] + "_" + objectid + ".html"
    with open(html_name, "w") as fh:
        fh.write(html_out)
    pdf_name = cwd + r"\One_Page_" + \
        df.FULLADDR[rec] + "_" + objectid + ".pdf"
    pdfkit.from_file(html_name, pdf_name)
