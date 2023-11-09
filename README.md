
# bByMalware

This project is done with Python and is for educational purposes, please use it wisely.

The project helps create .doc files that load a macro which enumerates the:
- servicios
- IPs
- Hostname





## Usage

The project consists of 3 files:

- rdTeam.py: This file contains the creation of the document and loading of the macro.
- index.php: This file will be uploaded to the attacking server, which receives the data sent by the macro. 
- redteam.vb: This file contains the code of the macro.


## Usage/Examples
To be able to execute it, make sure to have the index.php file loaded on a server.

```javascript
python3 rdTeam.py
```

This will run the script which prompts for the following parameters:


Name of the .doc file
```javascript
-> Enter the name of the .doc file (for example: test.doc):
```

Server URL where index.php is hosted
```javascript
-> Enter the URL (for example: http://127.0.0.1/):
```
