# identificacion-primas-deposito

# Identificación de Primas en Deposito

Macros en VBA para procesar primas en depósito, realizar cruces y generar referencias automáticas en Excel.

## Archivos

* Módulo principal: consolidación de primas
* Cruces: validación por monto y fecha
* Referenciación: identificación basada en texto

## Uso

1. Configurar rutas:

```vba
Const RUTA_ORIGEN = "C:\Ruta\Origen\"
Const RUTA_DESTINO = "C:\Ruta\Destino\"
Const RUTA_EC = "C:\Ruta\EstadosCuenta\"
```

2. Ejecutar:

```vba
EjecutarProcesoCompleto
```

## Funcionalidad

* Lectura de archivos más recientes
* Consolidación de registros
* Cruce por monto y fecha
* Generación de referencias automáticas
* Actualización de estados de cuenta

## Requisitos

* Excel con macros habilitadas

## Notas

* Ajustar nombres de columnas según archivo
* No incluir rutas o datos sensibles
