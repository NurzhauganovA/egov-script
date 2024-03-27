import asyncio

import requests
import websockets


async def handler(websocket, path):
    print("Client connected")
    message = await websocket.recv()
    print(f"Received message: {message}")
    await websocket.send(f"Hello, {message}!")
    requests.post(
        url='https://idp.egov.kz/idp/eds-login.do/',
        data={'certificate': message},
        headers={'Content-Type': 'application/x-www-form-urlencoded'}
    )


start_server = websockets.serve(handler, "127.0.0.1", 13579)

asyncio.get_event_loop().run_until_complete(start_server)
asyncio.get_event_loop().run_forever()
