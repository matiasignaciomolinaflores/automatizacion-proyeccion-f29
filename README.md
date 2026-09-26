# Automatización de Proyección del Formulario 29

Proyecto desarrollado en Python para automatizar parte del proceso mensual de proyección del Formulario 29.

## Objetivo

Reducir tareas manuales repetitivas asociadas a la recopilación, validación y procesamiento de información tributaria.

## Funcionalidades principales

- Automatización del Registro de Compras y Ventas mediante Selenium.
- Obtención de IVA crédito e IVA débito fiscal.
- Descarga y procesamiento del Formulario 29 anterior.
- Validación de RUT y período tributario.
- Integración con Microsoft Excel.
- Actualización automática de valores de UTM.
- Manejo de errores y registro de ejecución.
- Generación y envío automático de correos mediante Gmail API.

## Tecnologías utilizadas

- Python
- Selenium
- OpenPyXL
- PyPDF
- Pandas
- Requests
- Gmail API
- OAuth 2.0

## Flujo general

Período → Clientes → Registro de Compras y Ventas → F29 anterior → Validación → Excel → Cálculo → Correo

## Contexto del proyecto

La automatización fue desarrollada a partir de una necesidad real de Contabilidades PVM, donde la elaboración mensual de proyecciones del Formulario 29 implicaba múltiples tareas repetitivas.

La solución permite integrar distintas herramientas dentro de un único flujo, reduciendo trabajo operativo y manteniendo controles de validación sobre la información procesada.

## Seguridad y privacidad

Este repositorio contiene una versión demostrativa del proyecto.

No se incluyen:

- credenciales;
- claves tributarias;
- tokens de acceso;
- información real de clientes;
- Formularios 29 reales;
- planillas con información confidencial.

## Autor

Proyecto desarrollado como parte de un trabajo de automatización y mejora de procesos empresariales.
