## an example of grabbing files from Sharepoint and loading them into memory, then moving to S3
## need credentials! i.e., app registration for the target sharepoint site
## this file doesn't contain all of the necessary components to run; just for illustration
## this code was written for PoC purposes for Assessor's redaction. This can also be done in C#
# useful link: https://github.com/vgrem/Office365-REST-Python-Client/tree/master/examples/sharepoint/files
# useful link: https://github.com/vgrem/Office365-REST-Python-Client/issues/94




# !pip install Office365-REST-Python-Client

import boto3

import os
import io
import json
import random
import time
import tempfile
import requests
import sys
from PIL import Image, ImageDraw, ExifTags, ImageColor, ImageFont, UnidentifiedImageError

import office365
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.files.file import File
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.permissions.permission_kind import PermissionKind
from office365.sharepoint.files.file_system_object_type import FileSystemObjectType
from office365.sharepoint.listitems.listitem import ListItem

# python modules
import utilities
import tools
import config
import credentials


## in this example, we have a .csv file in S3 with id's of files we want to grab from a Sharepoint site
## you can generate your desired id's using other methods to

# grab list of document id's from csv file
sharepoint_document_ids_df = tools.read_s3_csv(config.bucket, 'XXXXXXXXX.csv')
# generate id's
filtered_sharepoint_doc_ids = list(sharepoint_document_ids_df['XXXXXX'])


# set up credentials
client_credentials = ClientCredential(credentials.client_id, credentials.client_secret) # credentials stored in separate file

# set up configs for sharepoint api
test_team_site_url = 'https://XXX.sharepoint.com/sites/XXXXXX/XXXXXX'
ctx = ClientContext(test_team_site_url).with_credentials(client_credentials)
doc_lib = ctx.web.lists.get_by_title("XXXXX")
items = doc_lib.items.select(["FileSystemObjectType"]).expand(["File", "Folder"]).get().execute_query()



# generate list of paths of available files corresponding to whatever filter you're interested in
available_filepaths = [item.file.serverRelativeUrl for item in items if ((item.file_system_object_type == FileSystemObjectType.File) and (item.file.serverRelativeUrl.split('/')[-2] in filtered_sharepoint_doc_ids))]
# subset as necessary for your purposes (e.g. memory issues)
available_filepaths_subset = available_filepaths[0:150]




# read files into memory using Python Sharepoint API
# in this example, we're loading images files. but should be able to adapt this code to other file types stored in Sharepoint

objects_from_sharepoint = []
filepaths_from_sharepoint = []

for idx, path in enumerate(available_filepaths_subset):
    
    try:
        with open('meow.png', "wb") as localFile:
            response = File.open_binary(ctx, path) # this is the important part: opening the file in memory
            image = Image.open(io.BytesIO(response.content))
            rgb_image = image.convert('RGB') # make sure to convert to RGB (or make sure not RGBA)
        objects_from_sharepoint.append(rgb_image)
        filepaths_from_sharepoint.append(path)
        print(f'read file id {idx} of {len(available_filepaths_subset)} at path {path}')
    except UnidentifiedImageError as e:
        print(f'** WARNING: skipped file number {idx} at path {path} due to error: {e} **')
        continue



# move images to S3
destination_key = "XXXXX"

for idx, thing in enumerate(zip(filepaths_from_sharepoint, objects_from_sharepoint)):
    filename = os.path.join(thing[0].split('/')[-2], thing[0].split('/')[-1])
    PILimage = thing[1]
    in_mem_file = io.BytesIO()
    PILimage.save(in_mem_file, format='JPEG')
    in_mem_file.seek(0)
    
    # Upload image to s3
    key = f"{destination_key}/{filename}"
    s3_client = config.s3_client
    s3_client.upload_fileobj(in_mem_file, config.bucket, key)
    
    print(f'uploaded file number {idx} of {len(objects_from_sharepoint)} to {key}')