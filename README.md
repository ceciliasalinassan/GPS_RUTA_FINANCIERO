# GPSruta Financiero PRO

Clave: 1234

Mejoras: deudas/facturas por pagar, ingresos separados, egresos separados, giro de clientes, categorías predeterminadas, dashboard mensual real, cambio automático de factura a pagada al ingresar pago.

Incluye guardado manual, respaldo, importar respaldo y eliminación de historial por mes/año.

Gráficos compactos y dashboard visualmente más profesional.

Carga masiva de clientes desde Excel/CSV incluida. Columnas: nombre, rut, giro, telefono, email, direccion, contacto.

Se agregó descarga de informe Excel con hojas: Resumen, Pagos, Ingresos, Egresos, Deudas, Facturas por pagar, Facturas por cobrar y Clientes.
Carga masiva acepta columna RAZON SOCIAL / RAZÓN SOCIAL.

Recordatorio actualizado: exige factura adjunta antes de enviar y usa texto solicitado por GPS-RUTA.

Facturas por cobrar: permite adjuntar factura y usa mensaje: BUENOS DIAS ESTIMADOS, SE ADJUNTA FACTURA DEL MES. SALUDOS CORDIALES. GPSRUTA.

Se agregó carga masiva de facturas por cobrar y vínculo de egreso con factura/deuda por pagar.

Recordatorio separado de Enviar factura. Deudas pendientes por defecto. Pie con números visibles.

Informe corregido: Facturas por cobrar exporta solo vencidas y por vencer.

Clientes simplificado: listado muestra solo nombre/razón social y RUT.

Fondo actualizado a estilo premium negro/dorado/verde neón.

Dashboard corregido: Estado facturas del mes muestra números visibles de inmediato.

Dashboard Futurista 3D agregado: estilo holográfico, neón, animaciones y panel financiero avanzado.

Correo cobranza configurado: gpsruta007@outlook.com. Reporte Estado de Facturas e IA de cobranza agregados.

Informe Excel corregido: hoja Deudas muestra solo pendientes/no pagadas.

Se agregó envío de correo para deudas/facturas por pagar con email del proveedor.

Correo de cobranza actualizado a gpsruta007@outlook.com.

Dashboard ajustado: números del resumen más pequeños y estilo futurista reforzado.

Dashboard mejorado con estilo futurista animado, holográfico y efectos neón.

Hora en formato 24 hrs. Cliente PREMIUM automático al estar al día. Riesgo: alto 4 vencidas, medio 3 vencidas, bajo 2 vencidas.

Dashboard 3D más futurista e IA Chat dentro del dashboard agregados.

Clientes compactos; edición sticky; IA y recordatorios a 15 días corregidos.


## Envío automático real Outlook

Correo configurado:
gpsruta007@outlook.com

En Vercel debes agregar estas variables de entorno:

OUTLOOK_USER=gpsruta007@outlook.com
OUTLOOK_APP_PASSWORD=contraseña_de_aplicación_outlook

Luego hacer redeploy.

El sistema incluye:
- /api/send-email.js
- Envío automático de recordatorios
- Envío automático de factura adjunta
- Envío automático de correo de deuda/proveedor
- CC automático al correo de cobranza

Se agregó Centro de Control IA: radar financiero, centro de alertas y actividad en tiempo real.

Facturas compactas con formulario sticky. IA agrega clientes con 2, 3 y 4+ facturas pendientes.

Facturas profesionales: carpetas por mes/año, filtros por estado y paginación de 50 registros.

Cobros/Recordatorios compactado con filtros, contador, listado con scroll y panel de acción fijo.

Historiales compactos: deudas, ingresos y egresos con filtros rápidos, scroll y panel de acción fijo.


## Versión Estable PRO
Se agregó protección contra errores, validaciones, advertencia de duplicados, mejoras de rendimiento visual y documentación de estabilización.

Filtro Carpeta mes/año actualizado: incluye opción TODAS por defecto.

Panel de vencimientos, filtro Vencida por defecto, categorías egresos actualizadas y block de tareas agregado.

Validación agregada: no permite clientes con RUT duplicado. Saldo neto negativo se muestra en rojo.

IA GPSRUTA mejorada: responde factura más antigua, 10 facturas más antiguas, deuda más antigua, facturas sobre 90 días y prioridad de cobranza.

IA renombrada a LUXURY GPSRUTA.

Agregado: auditoría automática, flujo de caja proyectado y calendario de cobranza.

Agregado: Centro de Alertas Ejecutivas, Top 10 Clientes y Cliente VIP automático.

Corrección aplicada: se agregaron los íconos faltantes CalendarCheck, ShieldAlert y ClipboardCheck para evitar el error visual del sistema.

Agregado: compromisos de pago, meta de cobranza mensual y gestión/historial de cobranza.

Versión recuperación estable: carga inicial robusta para evitar pantalla negra. Build no probado: Command '['npm', '--prefix', '/mnt/data/GPS_RUTA_RECUPERACION_ESTABLE', 'install', '--silent']' returned non-zero exit status 1.
