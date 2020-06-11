import requests
import asyncio
from aiohttp import ClientSession
import datetime
import json

tempus = "https://tempus.xyz/api"
demo_end = "/demos/id"

async_request_increment = 100
limit = 100_000

uploaded_demos = []

async def parse(demo):
    if not demo:
        return

    if demo['url'] is None:
        return

    time = datetime.datetime.fromtimestamp(demo['date']).strftime('%d/%m/%Y')

    uploaded_demos.append({
        "link": 'https://tempus.xyz/demos/' + str(demo['id']),
        "date": time
    })

async def get_demo_async(id, session):
    try:
        r = await session.request(method='GET', url = f'{tempus}{demo_end}/{id}')
        r.raise_for_status()
    except:
        return
    
    parsed = await r.json()
    return parsed


async def get_demo_wrapper(id, session):
    demo = await get_demo_async(id, session)
    await parse(demo)

async def async_demo_wrapper(start, max):
    async with ClientSession() as session:
        await asyncio.gather(*[get_demo_wrapper(id, session) for id in range(start, max)])

def main():
    loop = asyncio.get_event_loop()

    current_id = 0
    current_max_id = async_request_increment

    for i in range(current_id, limit, async_request_increment):
        loop.run_until_complete(async_demo_wrapper(current_id, current_max_id))

        current_id += async_request_increment
        current_max_id += async_request_increment

    loop.close()


main()
with open('demos.txt', 'w') as f:
    f.write(json.dumps(uploaded_demos, indent=4))