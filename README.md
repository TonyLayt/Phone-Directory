# Phone Directory

Desktop phone directory built with Python and PyQt5. An administration application manages contacts and departments, while a separate client displays the directory and downloads updates from a TCP server.

The interface is in Ukrainian.

## Screenshots

| Client | Administration |
| --- | --- |
| ![Client](screenshots/client.jpg) | ![Administration](screenshots/admin.jpg) |

All contacts shown in the screenshots are fictional.

## Features

- Contacts grouped by department
- Search with highlighted matches and navigation between results
- Adding, editing and deleting contacts and departments
- SQLite storage and JSON export
- Manual client updates over TCP
- Offline viewing of previously downloaded data
- Word export from the client
- Supporting Word documents for contacts and departments

## Running from source

The application targets Windows and requires Python 3.

From the repository root, install the dependencies:

```bat
python -m pip install PyQt5 python-docx
```

Create the local data folders if they do not already exist:

```bat
if not exist server\docdat mkdir server\docdat
if not exist client\docdat mkdir client\docdat
```

### Administration

From the repository root:

```bat
cd server
python main.py
```

The application creates `datetable.db` on first launch. Add and save departments and contacts to populate the directory. Saving records also generates `docdat/dict_tab.json` for client updates.

### TCP server

Open another terminal at the repository root:

```bat
cd server
python telephone_directory_server.py
```

Keep this terminal open while downloading updates.

The server button in the administration application expects a compiled `telephone_directory_server.exe`. When running from source, start the Python script directly as shown above.

### Client

Open another terminal at the repository root:

```bat
cd client
python main.py
```

Press the refresh button to download the directory. Previously downloaded records remain available when the server is offline.

The default connection is `localhost:8000`, with both applications running on the same computer.

## Local data

- `server/datetable.db` — administration database
- `server/docdat/` — JSON directory and documents distributed by the TCP server
- `client/docdat/` — downloaded files used by the client

Databases, local settings, directory records, documents and executables are excluded from Git. The demonstration records shown in the screenshots are not included in the repository.

Applications must be started from their respective folders because resource and data paths are relative.

## Network limitations

The TCP transfer protocol has no authentication or encryption. Keep the default localhost configuration for local demonstrations; the server is not intended to be exposed to the public internet.