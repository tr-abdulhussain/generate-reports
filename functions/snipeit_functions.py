import aiohttp
import asyncio
from urllib.parse import quote
import os
import sys

ASSET_TAG_COLUMN = "Asset Tag"
ACTIONS_COLUMN = "Actions to be taken"

async def correct_asset_status(asset):
    failed_patch = []
    patched = []
    requests_processed = 0
    max_requests_per_minute = 100
    snipeit_link = "http://3.0.176.121"

    async with aiohttp.ClientSession() as session:
        for tag, action in zip(asset[ASSET_TAG_COLUMN], asset[ACTIONS_COLUMN]):
            if requests_processed >= max_requests_per_minute:
                print("Rate limit reached. Waiting for 60 seconds...")
                await asyncio.sleep(60)
                requests_processed = 0

            encoded_tag = quote(tag)
            get_id_url = f"{snipeit_link}/api/v1/hardware/bytag/{encoded_tag}"
            # api_key = os.environ["SNIPEIT_API_KEY"]
            api_key = os.environ["TEST_SNIPEIT_KEY"]

            headers = {
                "Authorization": f"Bearer {api_key}",
                "Accept": "application/json",
                "Content-Type": "application/json"
            }

            try:
                async with session.get(get_id_url, headers=headers) as get_response:
                    requests_processed += 1
                    if get_response.status == 200:
                        response_json = await get_response.json()
                        asset_id = response_json["id"]
                        status_id = response_json["status_label"]["id"]
                        if status_id == 2 and action == "Set as Spare":
                            payload = {
                                "status_id": status_id,
                                "name": ""
                            }
                            post_url = f"{snipeit_link}/api/v1/hardware/{asset_id}/checkin"
                            await session.post(post_url, headers=headers, json=payload)
                        else:
                            if action == "Set as Spare":
                                status_id = 2
                            if action == "Set as Assigned":
                                status_id = 4
                            payload = {
                                "status_id": status_id
                            }
                            patch_url = f"{snipeit_link}/api/v1/hardware/{asset_id}"
                            await session.patch(patch_url, headers=headers, json=payload)
                        requests_processed += 1
                        patched.append(tag)
                        print(f"Asset Tag {tag} has been {action}")
                    else:
                        failed_patch.append({"asset_tag": tag, "error": str(get_response.status)})
                        print(f"Failed to get data for {tag}, Status: {get_response.status}")
            except KeyError as e:
                failed_patch.append({"asset_tag": tag, "error": str(e)})
                print("Error processing asset:", tag)
    print("Failed Asset Updates: ", failed_patch)
    print("Updated Assets: ", patched)

    if failed_patch:
        print("Manually update assets that failed. Import a new CSV file after manually updating the assets.")
        sys.exit()
    return patched