# Script certificados en router Teldat

Script para automatizar la generación y actualización de certificados en router Teldad [generar key y csr, y subir/actualizar cert] a través de un host bastion (equipo de salto). Además se puede obtener ciertos parámetros del equipo: caducidad certificados, files, etc.

Permite trabajar de manera unitaria o por lotes (lista de routers). Genera un report al finalizar con el resultado y errores de las operaciones efectuadas. 

Permite crear nuevas funcionalidades usando la misma base de una manera agil.

## Requisitos

Diseñado en VBScript. Requiere de SecureCRT para funcionar y Openssl para exportar a base64:

- [ SecureCRT ](https://www.vandyke.com/products/securecrt/)
- [ OpenSSL ](https://slproweb.com/products/Win32OpenSSL.html)

## Configuración

Antes de usar, es necesario modificar ciertas variables: 
```bash
user = <mi_usuario>
password = <mi_password>
host_gestion = <ip_host_gestion_o_bastion>
path_openssl = </directorio/openssl/ejecutable/openssl.exe>
```

## Utilización

1. Generar un fichero **"lote.txt"** siguiendo el siguiendo patrón: una línea por equipo (hostname y ip_gestión, separados por un espacio " "). ***Se puede omitir e introducir el hostname e ip del router en el trasnscurso de ejecución** 
```bash
RT_central 10.1.1.1
RT_sucursal2 123.123.123.2
RT_sucursal3 123.123.123.3
```
2. Lanzar el script desde SecureCRT > Script > ScriptSecureV3.vbs
3. Indicar directorio de trabajo
4. Seleccionar acción a realizar 
5. Comprobar errores generados en el report de operación.

## Licencia

Código distribuido bajo Licencia Apache 2.0
