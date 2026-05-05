# Generador Informe Mensual – Servicio Apple (Ricoh / Banco de Chile)

## Backend
1) Instalar dependencias:
   python3 -m pip install -r requirements.txt

2) Ejecutar API:
   python3 -m uvicorn main:app --reload

API disponible en:
- http://127.0.0.1:8000
- Documentación: http://127.0.0.1:8000/docs

## Entradas (form-data)
- INC (xlsx)
- RITM (xlsx)
- INC_Abiertos (xlsx)
- RITM_Abiertos (xlsx)

## Salida
- Word .docx con portada Ricoh y nombre: Informe_Servicio_Apple_<mes>_<año>.docx

