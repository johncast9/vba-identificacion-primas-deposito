# identificacion-primas-deposito

# Identificación de Primas en Deposito

Macros en VBA para procesar información de primas en depósito, realizar cruces y generar referencias automáticas en Excel.

## Archivos

* Macro principal: procesa y consolida información
* Funciones de cruce: validan datos entre meses
* Módulos de referencia: identifican primas a partir de descripciones

## Uso

1. Configurar rutas en el código:

```vba
Const RUTA_ORIGEN = "C:\Ruta\Origen\"
Const RUTA_DESTINO = "C:\Ruta\Destino\"
```

2. Ejecutar:

```vba
EjecutarProcesoCompleto
```

## Funcionalidad

* Lectura de archivos más recientes
* Consolidación de primas
* Cruce de información entre periodos
* Identificación automática de referencias
* Procesamiento por reglas de negocio

## Requisitos

* Excel con macros habilitadas

## Notas

* No incluir rutas o datos sensibles
* Ajustar nombres de columnas según archivo
* El proceso depende de la estructura del Excel
