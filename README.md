
# BbyMalware

Este proyecto esta realizado con python y es para fines didacticos, por favor usalo sabiamente.

El proyecto ayuda a crear archivos .doc en los cuales carga una macro que enumera los: 
- servicios
- IPs
- Hostname





## Usage

El proyecto se compone de 3 archivos

- rdTeam.py: este archivo contiene la creación del documento y carga de la macro  
- index.php: este archivo se cargara en el servidor atacante y este recibe los datos que envía la macro
- redteam.vb: este archivo contiene el código de la macro.


## Usage/Examples
Para poder ejecutarlo, asegurate de tener el archivo index.php cargado en un servidor.

```javascript
python3 rdTeam.py
```

Esto ejecutará el script el cual solicita los siguientes parametros:


Nombre del archivo .doc
```javascript
-> Enter the name of the .doc file (for example: test.doc):
```

URL del servidor donde se cargo en index.php
```javascript
-> Enter the URL (for example: http://127.0.0.1/):
```
