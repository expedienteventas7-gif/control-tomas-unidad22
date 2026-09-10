import React, { useState, useEffect, useMemo, useCallback } from "react";
import {
  LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, Cell,
} from "recharts";
import * as XLSX from "xlsx";

/* ---------------------------------------------------------
   Tokens
--------------------------------------------------------- */
const C = {
  bg: "#EEF0EE",
  surface: "#FFFFFF",
  ink: "#1B231F",
  inkMuted: "#5B6660",
  inkFaint: "#8B948E",
  line: "#DBDFD9",
  lineStrong: "#C3C9C1",
  primary: "#2B5C4E",
  primarySoft: "#E1EBE6",
  accent: "#C1622F",
  accentSoft: "#F6E4D8",
  warn: "#AD8A2E",
  warnSoft: "#F2ECD8",
  danger: "#A23E3E",
  dangerSoft: "#F3E1DF",
};

const FONT_HEAD = "'Space Grotesk', sans-serif";
const FONT_MONO = "'IBM Plex Mono', monospace";
const FONT_BODY = "'IBM Plex Sans', sans-serif";

/* ---------------------------------------------------------
   Domain constants
--------------------------------------------------------- */
const SUCURSALES = ["TUXTLA","TAPACHULA","SAN CRISTOBAL","OCOSINGO","PALENQUE","ISTMO","ORIENTE","PONIENTE","TENOSIQUE","COMITAN"];
const MARCAS = ["NISSAN","RENAULT","CHANGAN"];
const VALIDADORES = ["BLANCA","GIL","GILBERTO","BRAYAN","URIEL"];
const TIPOS_TOMA = ["COMPRA DIRECTA","SEMINUEVO X NUEVO","SEMINUEVO X SEMINUEVO"];
const FINANCIAMIENTO_OPTS = ["SI","NO","CREDITO NR"];
const ESTATUS_OPTS = ["EN PROCESO","COMPLETADO"];

const MESES = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"];

const STORAGE_KEY = "tomas-records-v1";
const USER_KEY = "tomas-user-name-v1"; // identidad personal, no compartida
const ALERT_DAYS = 5; // dias en proceso a partir de los cuales se marca una toma como atorada

const ERROR_CATALOG = [
  "Factura / refacturación",
  "Revisión mecánica / VIN",
  "RFC / datos fiscales",
  "Comprobante de domicilio",
  "INE / identificación",
  "Endoso pendiente",
  "Baja de placas",
  "Verificación QR / AMDA",
  "Alta en SIA (montos, claves)",
  "Formulario mal capturado",
  "Documentación societaria",
  "Otro",
];

const SEED_DATA = [{"id": "seed-1", "cliente": "HIGINIO MORGA BAUTISTA", "exp": "2066", "inv": "90222", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "", "financiamiento": "SI", "fSolicitud": "2026-01-08", "fRevision": "2026-01-08", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente endoso.", "observaciones": "", "primerPago": "2026-01-10", "segundoPago": "2026-01-14"}, {"id": "seed-2", "cliente": "JOSE GABRIEL ZUÑIGA RAMOS", "exp": "2067", "inv": "90223", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "OCOSINGO", "marca": "RENAULT", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-01-10", "fRevision": "2026-01-10", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente validación de la factura.", "observaciones": "VALIDACION REPUVE: NO SE CONSULTARON CORRECTAMENTE LAS AUTORIDADES. REVISION MECANICA: NUMERO DE VIN INCORRECTO, NO TIENE CALCA Y FOTO DEL VIN. RFC: TIPO DE PERSONA INCORRECTO.", "primerPago": "ODP 15/1/26", "segundoPago": ""}, {"id": "seed-3", "cliente": "CRISTHIAN OMAR VICENTE ALFARO", "exp": "2068", "inv": "90224", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-01-13", "fRevision": "2026-01-13", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "", "primerPago": "2026-01-14", "segundoPago": ""}, {"id": "seed-4", "cliente": "JOSE HIGINIO MARTINEZ DOMINGUEZ", "exp": "2069", "inv": "90225", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-01-15", "fRevision": "2026-01-15", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "", "primerPago": "2026-01-16", "segundoPago": ""}, {"id": "seed-5", "cliente": "FRANCISCO JAVIER SALGADO LEON", "exp": "2072", "inv": "92809", "tipoToma": "SEMINUEVO X SEMINUEVO", "sucursal": "PALENQUE", "marca": "NISSAN", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-01-21", "fRevision": "2026-01-21", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "DIRECCION: TIENE UNA T DE MAS Y EL CODIGO POSTAL ES INCORRECTO, FACTURA ORIGEN, NO TIENE ENDOSO", "primerPago": "ODP 27/1/26", "segundoPago": ""}, {"id": "seed-6", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2074", "inv": "90227", "tipoToma": "COMPRA DIRECTA", "sucursal": "ISTMO", "marca": "RENAULT", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-01-23", "fRevision": "2026-01-23", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "FACTURA DE ORIGEN: FECHA Y FOLIO DE REFACTURACION INCORRECTOS. REVISION MECANICA: NO VIENEN TODOS LOS CAMPOS LLENOS. PANTALLA DE ALTA: TIPO DE TOMA INCORRECTO. VERIFICACION CONSULT: FALTA LA FOTO DEL DISPOSITIVO. FOTO DE PARABRISAS: EL COLOR NO CORRESPONDE CON LA FACTURA. DEBERAN AGREGAR AUTORIZACION DEL GERENTE.", "primerPago": "2026-01-26", "segundoPago": ""}, {"id": "seed-7", "cliente": "JOSE IGNACIO RUBIO MORA", "exp": "2077", "inv": "92", "tipoToma": "COMPRA DIRECTA", "sucursal": "COMITAN", "marca": "CHANGAN", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-01-27", "fRevision": "2026-01-27", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "FACTURA ORIGEN: AGREGARAN LA REFACTURA. SUBMARCA Y DESCRIPCION INCORRECTOS. El numero de serie de la refactura es incorrecta. FOLIO FACTURA NUEVA INCORRECTA. REVISION MECANICA: FALTA FOTO DEL VIN. PANTALLA SIA: UBICACIÓN INCORRECTA, FAC PROV INCORRECTO, DATOS RECORTADOS.IMAGEN QR: NO TIENE LECTURA. ETIQUETA HOLOGRAFICA: NO ES EL CORRECTO.xml cambiara lo mismo que la lectura del qr", "primerPago": "2026-01-30", "segundoPago": ""}, {"id": "seed-8", "cliente": "BERNARDO BAUTISTA MALDONADO", "exp": "2073", "inv": "92810", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "NISSAN", "validador": "", "financiamiento": "SI", "fSolicitud": "2026-01-28", "fRevision": "2026-01-28", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente endoso.", "observaciones": "FACTURA: EXPEDIDA INCORRECTO. CREDENCIAL: NO TIENE ARCHIVO. CURP Y COMPROBANTE VENCIDOS. REVISION MECANICA: VIN REMARCADO Y FOTO NO ESTA.", "primerPago": "2026-01-30", "segundoPago": "2026-03-05"}, {"id": "seed-9", "cliente": "GUILLERMO VAZQUEZ DIAZ", "exp": "2078", "inv": "92817", "tipoToma": "COMPRA DIRECTA", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-01-28", "fRevision": "2026-01-29", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "REVISION MECANICA: LAS CALCAS DEL VIN NO SON LEGIBLES, NO ANEXAN FOTOGRAFIA DEL VIN", "primerPago": "2026-01-29", "segundoPago": ""}, {"id": "seed-10", "cliente": "LUIS DANIEL PEREZ GOMEZ", "exp": "2080", "inv": "92818", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-01-29", "fRevision": "2026-01-29", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente endoso. *Pendiente baja de placas.", "observaciones": "FACTURA DE ORIGEN: FECHA DE FACTURA INCORRECTA. REVISION MECANICA: LA CALCA NO ES DE LA UNIDAD. ETIQUETA: NO ES EL DOCUMENTO CORRECTO.", "primerPago": "ODP 5/2/26", "segundoPago": ""}, {"id": "seed-11", "cliente": "BEATRIZ ADRIANA DIAZ MONTEJO", "exp": "2081", "inv": "92816", "tipoToma": "COMPRA DIRECTA", "sucursal": "PALENQUE", "marca": "NISSAN", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-01-29", "fRevision": "2026-01-29", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "FACTURA: COLOR UNIDAD INCORRECTA. PANTALLA DE ALTA: DATOS RECORTADOS. RFC: COLONIA Y CODIGO POSTAL INCORRECTOS.", "primerPago": "2026-01-30", "segundoPago": ""}, {"id": "seed-12", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2082", "inv": "90229", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-01-30", "fRevision": "2026-01-30", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "DESCRIPCION Y SUBMARCA. IMAGEN QR AGREGARAN LEYENDA", "primerPago": "2026-01-30", "segundoPago": ""}, {"id": "seed-13", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2083", "inv": "90228", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-01-30", "fRevision": "2026-01-30", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "", "primerPago": "2026-01-30", "segundoPago": ""}, {"id": "seed-14", "cliente": "RICARDO GOMEZ SANTIZ", "exp": "2087", "inv": "92819", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-01-31", "fRevision": "2026-01-31", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "NO ANEXAN REVISION MECANICA", "primerPago": "", "segundoPago": ""}, {"id": "seed-15", "cliente": "JUAN CARLOS AVENDAÑO GOMEZ", "exp": "2084", "inv": "92820", "tipoToma": "SEMINUEVO X SEMINUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-01-31", "fRevision": "2026-01-31", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "No anexan la validacion ante la distribuidora", "primerPago": "ODP 5/2/26", "segundoPago": ""}, {"id": "seed-16", "cliente": "VIRIDIANA HERNANDEZ ROBLERO", "exp": "2089", "inv": "92821", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "NISSAN", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-01-31", "fRevision": "2026-01-31", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "FACTURA: FALTA REFACTURACIÓN, FOLIO INCORRECTO. NO AGREGAN CREDENCIAL, CURP Y COMPROBANTE DE DOMICILIO. REVISION MECANICA: CAMPOS SIN LLENAR. IMAGEN QR NO ES DE LA UNIDAD. CAMBIARAN REFACTURACIÓN", "primerPago": "", "segundoPago": ""}, {"id": "seed-17", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2086", "inv": "90230", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-02-02", "fRevision": "2026-02-02", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas.", "observaciones": "", "primerPago": "2026-02-03", "segundoPago": ""}, {"id": "seed-18", "cliente": "RICARDO SANTIZ GOMEZ", "exp": "2092", "inv": "92819", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-01-31", "fRevision": "2026-02-02", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas.", "observaciones": "", "primerPago": "ODP 12/02/26", "segundoPago": ""}, {"id": "seed-19", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2093", "inv": "90231", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-02-02", "fRevision": "2026-02-02", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas.", "observaciones": "", "primerPago": "2026-02-03", "segundoPago": ""}, {"id": "seed-20", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2094", "inv": "90232", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-02-03", "fRevision": "2026-02-03", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas.", "observaciones": "", "primerPago": "2026-02-04", "segundoPago": ""}, {"id": "seed-21", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2095", "inv": "90233", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-02-03", "fRevision": "2026-02-03", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas.", "observaciones": "COMPROBANTE DE DOMICILIO: CARGAN EL DOCUMENTO INCORRECTO", "primerPago": "2026-02-04", "segundoPago": ""}, {"id": "seed-22", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2097", "inv": "90234", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-02-03", "fRevision": "2026-02-03", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas.", "observaciones": "", "primerPago": "2026-02-04", "segundoPago": ""}, {"id": "seed-23", "cliente": "VIRIDIANA HERNANDEZ ROBLERO", "exp": "2089", "inv": "92821", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "NISSAN", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-02-03", "fRevision": "2026-02-03", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "FACTURA: FALTA REFACTURACIÓN, FOLIO INCORRECTO. NO AGREGAN CREDENCIAL, CURP Y COMPROBANTE DE DOMICILIO. REVISION MECANICA: CAMPOS SIN LLENAR. IMAGEN QR NO ES DE LA UNIDAD. CAMBIARAN REFACTURACIÓN", "primerPago": "2026-02-04", "segundoPago": ""}, {"id": "seed-24", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2098", "inv": "90235", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-02-04", "fRevision": "2026-02-05", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas.", "observaciones": "", "primerPago": "2026-02-05", "segundoPago": ""}, {"id": "seed-25", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2100", "inv": "90236", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-02-05", "fRevision": "2026-02-05", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas.", "observaciones": "", "primerPago": "2026-02-06", "segundoPago": ""}, {"id": "seed-26", "cliente": "GABRIELA TELLO COUTIÑO", "exp": "2096", "inv": "92822", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "NISSAN", "validador": "", "financiamiento": "SI", "fSolicitud": "2026-02-05", "fRevision": "2026-02-06", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente factura origen y endoso. *Pendiente validación de la factura. *Pendiente verificacion QR AMDA Y SAT", "observaciones": "REVISION MECANICA: DATOS ILEGIBLE, FOTO VIN ILEGIBLE. PANTALLA ALTA: CLAVE VEHICULAR INCORRECTA, MONTOS ALREVES.", "primerPago": "2026-12-09", "segundoPago": "2026-03-11"}, {"id": "seed-27", "cliente": "JOSE ELIUTH OVANDO CARBAJAL", "exp": "2099", "inv": "92823", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "NISSAN", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-02-06", "fRevision": "2026-02-07", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "RFC: MUNICIPIO INCORRECTO. CORREGIR DESCRIPCION Y BAJA DE PLACAS.", "primerPago": "2026-12-09", "segundoPago": ""}, {"id": "seed-28", "cliente": "MONICA SOTELO CORZO", "exp": "2101", "inv": "92824", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-02-09", "fRevision": "2026-02-09", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente firmas en PROFECO", "observaciones": "REVISION MECANICA: LA CALCA DEL VIN ESTA INCOMPLETA. PANTALLA DEL ALTA: TIPO DE TOMA DICE SI", "primerPago": "ODP 16/02/2026", "segundoPago": ""}, {"id": "seed-29", "cliente": "JORGE ANTONIO JIMENEZ ANZA", "exp": "2103", "inv": "90237", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-02-11", "fRevision": "2026-02-12", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "PANTALLA SIA: FAC PROV INCORRECTO, REVISION MECANICA: NO TIENE FIRMA Y NOMBRE DE QUIEN REVISA LA UNIDAD", "primerPago": "2026-02-17", "segundoPago": ""}, {"id": "seed-30", "cliente": "JENNIFER PEREZ ESCOBAR", "exp": "2102", "inv": "92825", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "NISSAN", "validador": "", "financiamiento": "SI", "fSolicitud": "2026-02-18", "fRevision": "2026-02-18", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "FACTURA: EN BAJA PUSIERON NO. REVISION MECANICA: EL VIN ESTA MAL ESCRITO, TIENE TACHADURA Y NO TIENE FOTO DEL VIN. ALTA AL INV: LOS MONTOS NO COINCIDEN.", "primerPago": "", "segundoPago": ""}, {"id": "seed-31", "cliente": "CHRISTIAN AUGUSTO GARCIA COUTIÑO", "exp": "2104", "inv": "92828", "tipoToma": "COMPRA DIRECTA", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-02-18", "fRevision": "2026-02-18", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas.", "observaciones": "FORMULARIO: DIRECCION MUNICIPIO DICE PICHUCALCO", "primerPago": "2026-02-19", "segundoPago": ""}, {"id": "seed-32", "cliente": "URIEL LANESTOSA SOLORZANO", "exp": "2063", "inv": "92827", "tipoToma": "COMPRA DIRECTA", "sucursal": "OCOSINGO", "marca": "NISSAN", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-02-18", "fRevision": "2026-02-19", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "", "primerPago": "2026-02-19", "segundoPago": ""}, {"id": "seed-33", "cliente": "JD DISTRIBUCIONES Y OPERACIONES INSTITUCIONALES SA DE CV", "exp": "2079", "inv": "92829", "tipoToma": "COMPRA DIRECTA", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-02-18", "fRevision": "2026-02-19", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "VERIFICACION QR: SON 3 LECTURAS Y SUBEN 2,CONSTANCIA: EN ELO FORMULARIO NO TIENE LA RACION SOCIAL EL NOMBRE DE LA MORAL, ACTA CONSTITUTIVA: DATOS DEL REGISTRO PUBLICO SON INCORRECTOS, VERIFICACION DEL CONSULT: ESTA ILEGIBLE.", "primerPago": "2026-02-23", "segundoPago": ""}, {"id": "seed-34", "cliente": "RAFAEL ELIAS SANCHEZ MONTEMAYOR", "exp": "2110", "inv": "92830", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-02-24", "fRevision": "2026-02-25", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas. *Pendiente QR de la computadora", "observaciones": "RFC: TIPO DE PERSONA INCORRECTO.", "primerPago": "ODP 3/03/2026", "segundoPago": ""}, {"id": "seed-35", "cliente": "ANA GABRIELA CISNEROS DOMINGUEZ", "exp": "2111", "inv": "90239", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "", "financiamiento": "SI", "fSolicitud": "2026-02-26", "fRevision": "2026-02-26", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas. *Pendiente endoso del cliente.", "observaciones": "FACTURA ORIGEN: COLOR INCORRECTO", "primerPago": "2026-02-26", "segundoPago": "2026-03-13"}, {"id": "seed-36", "cliente": "VILIULFO PEREZ GONZALEZ", "exp": "2112", "inv": "92831", "tipoToma": "COMPRA DIRECTA", "sucursal": "COMITAN", "marca": "NISSAN", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-02-26", "fRevision": "2026-02-26", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "", "primerPago": "2026-02-27", "segundoPago": ""}, {"id": "seed-37", "cliente": "SERGIO POTENCIANO GALVEZ", "exp": "2121", "inv": "90240", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "OCOSINGO", "marca": "RENAULT", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-02-26", "fRevision": "2026-02-27", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "NO TIENE ENDOSO, FORMULARIO DICE PFAE Y ES PF", "primerPago": "ODP 6/03/2026", "segundoPago": ""}, {"id": "seed-38", "cliente": "NORBERTO CARLOS ARRIAGA MUÑOZ", "exp": "2109", "inv": "90238", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "RENAULT", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-02-27", "fRevision": "2026-02-27", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "NO ANEXAN RE-FACTURA , FORMULARIO: FOLIOS DE FACTURA MAL, INE: FOLIO ES INCORRECTO, COMP DOMI: NO LO ANEXAN, VERIFICACION QR: FALTAN VERIFICACIONES, REVISION MECANICA: NO TIENE SELLO DEL DISTRIBUIDOR Y NO TIENE INV TAMPOCO FOTO Y CALCAL DEL VIN, PANTALA DE SIA: NO ANEXAN EL DOCUMENTO, ESTADO DE CUENTA BANCARIO: NO LO ANEXAN, RFC: NO ESTA VIGENTE", "primerPago": "2026-03-02", "segundoPago": ""}, {"id": "seed-39", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2106", "inv": "92815", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "NISSAN", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-02-27", "fRevision": "2026-02-27", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "FACTURA ORIGEN: BAJA DE PLACAS DICE NO Y FECHA COMPRA DICE 2024, VERIFICACION QR: AMDA TENDRAN QUE SUBIR AUTORIZACION YA QUE NO SE PUEDE LEER, ALTA SIA: MONTOS NO COINCIDEN, FAC PROV ESTA MAL, ESTADO DE CUENTA BANCARIO ANEXAN OTRO QUE NO ES, NO ANEXAN ETIQUETA HOLOGRAFICA.", "primerPago": "", "segundoPago": ""}, {"id": "seed-40", "cliente": "ENRIQUE HERNANDEZ MECIAS", "exp": "2122", "inv": "92833", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "OCOSINGO", "marca": "NISSAN", "validador": "", "financiamiento": "NO", "fSolicitud": "2026-02-28", "fRevision": "2026-02-28", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "", "primerPago": "ODP 03/03/2026", "segundoPago": ""}, {"id": "seed-41", "cliente": "CARLOS MARIO ROMAN TREJO", "exp": "2123", "inv": "92832", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-02", "fRevision": "2026-03-02", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "Factura origen: endoso, QR del motor: falta la foto", "observaciones": "", "primerPago": "ODP 17/03/2026", "segundoPago": ""}, {"id": "seed-42", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2129", "inv": "92812", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "NISSAN", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-03", "fRevision": "2026-03-03", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "FORMULARIO: BAJA DE PLACAS DICE NO, VALIDACION DE REPUVE: ANEXAN UNA AUTORIZACION QUE NO APLICA, REVISION MECANICA: HACE FALTA EL NUMERO DE INV, ALTA DE SIA: MONTOS INCORRECTOS Y FAC PROV INCORRECTO PANTALLA DEL ALTA: MONTO DE ENAJENACIÓN INCORRECTO", "primerPago": "", "segundoPago": ""}, {"id": "seed-43", "cliente": "ERIKA VIDAL PEREZ", "exp": "2130", "inv": "90241", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-04", "fRevision": "2026-03-04", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas. *Pendiente validación de la factura origen.", "observaciones": "", "primerPago": "2026-03-05", "segundoPago": ""}, {"id": "seed-44", "cliente": "BERENICE DE JESUS GARCIA CASTELAN", "exp": "2128", "inv": "92834", "tipoToma": "COMPRA DIRECTA", "sucursal": "TENOSIQUE", "marca": "NISSAN", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-06", "fRevision": "2026-03-06", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*No tendrá baja de placas.* No tendrá validación transunion", "observaciones": "FACTURA: SUBMARCA, DESCRIPCION Y BAJA INCORRECTOS. PANTALLA DE ALTA: UBICACIÓN Y FACT PROV. VERIFICACION CONSULT: NO TIENE DATOS DE LA UNIDAD", "primerPago": "2026-03-12", "segundoPago": ""}, {"id": "seed-45", "cliente": "ANGEL GABRIEL CRUZ GARCIA", "exp": "2124", "inv": "92835", "tipoToma": "COMPRA DIRECTA", "sucursal": "OCOSINGO", "marca": "NISSAN", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-07", "fRevision": "2026-03-07", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "PENDIENTE ENDOSO DEL CLIENTE", "primerPago": "", "segundoPago": ""}, {"id": "seed-46", "cliente": "RENE ZENTENO FALCONI", "exp": "2131", "inv": "90242", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "BLANCA", "financiamiento": "SI", "fSolicitud": "2026-03-10", "fRevision": "2026-03-10", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente factura origen y endoso.", "observaciones": "", "primerPago": "2026-03-11", "segundoPago": "2026-03-27"}, {"id": "seed-47", "cliente": "RAYMUNDO HERNANDEZ RODRIGUEZ", "exp": "2133", "inv": "90243", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-10", "fRevision": "2026-03-10", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "COMPROBANTE DE DOM INCORRECTO. PANTALLA DE ALTA: MONTO DE ADQUISICIÓN INCORRECTO", "primerPago": "2026-03-12", "segundoPago": ""}, {"id": "seed-48", "cliente": "GUILLERMO DE COSS GUZMAN", "exp": "2137", "inv": "90245", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-10", "fRevision": "2026-03-10", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "FACTURA: FECHA DE FACTURA INCORRECTA", "primerPago": "ODP 13/03/2026", "segundoPago": ""}, {"id": "seed-49", "cliente": "ALEJANDRA RODAS ARREVILLAGA", "exp": "2131", "inv": "90244", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "RENAULT", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-10", "fRevision": "2026-03-10", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "FACTURA: SUBMARCA, FOLIO REFACTURACION INCORRECTOS. COMPROBANTE DE DOM: DEBEN AGREGAR EL RECIBO DE LUZ O EDO DE CUENTA, LA CALLE ESTA MAL ESCRITA. VERIFICACION QR NO ES EL DOCUMENTO CORRECTO. ELIMINAR COMPROBANTE DE DOM MORAL", "primerPago": "2026-03-12", "segundoPago": ""}, {"id": "seed-50", "cliente": "MARIA EUGENIA MARTINEZ RODRIGUEZ", "exp": "2136", "inv": "92837", "tipoToma": "COMPRA DIRECTA", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "BLANCA", "financiamiento": "SI", "fSolicitud": "2026-03-10", "fRevision": "2026-03-11", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente factura origen y endoso. *Pendiente baja de placas", "observaciones": "PANTALLA SIA: APARTADOS RECORTADOS, MONTO DE ENAJENACION Y ADQUISICIÓN INVERTIDOS.", "primerPago": "2026-03-12", "segundoPago": "2026-04-01"}, {"id": "seed-51", "cliente": "MEGAMOTRIZ SA DE CV", "exp": "2118", "inv": "92826", "tipoToma": "COMPRA DIRECTA", "sucursal": "OCOSINGO", "marca": "NISSAN", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-11", "fRevision": "2026-03-11", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "", "primerPago": "2026-03-17", "segundoPago": ""}, {"id": "seed-52", "cliente": "MARIA MAGDALENA CRUZ FARRERA", "exp": "2140", "inv": "92838", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "TAPACHULA", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-03-12", "fRevision": "2026-03-12", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "TIPO DE TOMA INCORRECTA, EL CLIENTE NOS TIENE QUE FACTURAR", "primerPago": "", "segundoPago": ""}, {"id": "seed-53", "cliente": "RAMON ISRAEL VILLALOBOS CHAVEZ", "exp": "2142", "inv": "92840", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-12", "fRevision": "2026-03-13", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas.", "observaciones": "CREDENCIAL: EL NOMBRE DEL CLIENTE ESTA MAL ESCRITO", "primerPago": "", "segundoPago": ""}, {"id": "seed-54", "cliente": "YENI RAQUEL HERNANDEZ CRUZ", "exp": "2143", "inv": "92839", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "PALENQUE", "marca": "NISSAN", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-13", "fRevision": "2026-03-13", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas.", "observaciones": "FACTURA ORIGEN: SUBMARCA Y DESCRIPCION INCORRECTOS. REVISION MECANICA: NO TIENE INV. ALTA SIA: LA CLAVE VEHICULAR ES INCORRECTA.", "primerPago": "ODP 19/03/2026", "segundoPago": ""}, {"id": "seed-55", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2141", "inv": "90247", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "RENAULT", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-13", "fRevision": "2026-03-13", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas.", "observaciones": "FACTURA ORIGEN: FOLIO REFACTURACION INCORRECTA. VALIDACION DE REPUVE: CONSULTA INCORRECTA. REVISION MECANICA: CAMPOS VACIOS, VIN MAL ESCRITO, DEBERAN QUITAR LA HOJA DE CARTA FACTURA. ALTA AL INV: CAMPOS RECORTADOS. XML: NO ESTA COMPLETO EL DOCUMENTO. PODER NOTARIAL INCORRECTO. NO AGREGAN ACTA CONSTITUTIVA.", "primerPago": "2026-03-17", "segundoPago": ""}, {"id": "seed-56", "cliente": "ANDRES DIAZ GONZALEZ", "exp": "2139", "inv": "92841", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-13", "fRevision": "2026-03-13", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "PENDIENTE PROFECO", "observaciones": "BAJA DE PLACAS: AGREGARAN AUTORIZACION DE CARLOS. COMPROBANTE DE DOM VENCIDO. FALTA QR AMDA", "primerPago": "ODP 17/03/2026", "segundoPago": ""}, {"id": "seed-57", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2134", "inv": "92813", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "NISSAN", "validador": "GILBERTO", "financiamiento": "NO", "fSolicitud": "2026-03-13", "fRevision": "2026-03-13", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "FACTURA ORIGEN: FOLIOS DE FACTURA ESTAN INVERTIDOS, ALTA INV SIA: MONTO ENAJENACION INCORRECTO, Y VIN CARROCERIA MAL", "primerPago": "", "segundoPago": ""}, {"id": "seed-58", "cliente": "MIRNA FERNANDA MARTINEZ MARIN", "exp": "2148", "inv": "90250", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "ISTMO", "marca": "RENAULT", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-17", "fRevision": "2026-03-17", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "FACTURA DE ORIGEN: NUMERO MOTOR INCORRECTO. REVISION MECANICA: HOJA DE INSPECCION CON TACHADURAS Y CAMPOS EN BLANCO.", "primerPago": "ODP 23/03/2026", "segundoPago": ""}, {"id": "seed-59", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2144", "inv": "90246", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GILBERTO", "financiamiento": "NO", "fSolicitud": "2026-03-17", "fRevision": "2026-03-17", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "", "primerPago": "2026-03-18", "segundoPago": ""}, {"id": "seed-60", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2145", "inv": "90248", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-17", "fRevision": "2026-03-17", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas.", "observaciones": "", "primerPago": "2026-03-18", "segundoPago": ""}, {"id": "seed-61", "cliente": "MARIA GUADALUPE HERNANDEZ MARTINEZ", "exp": "2149", "inv": "92842", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "GILBERTO", "financiamiento": "NO", "fSolicitud": "2026-03-17", "fRevision": "2026-03-18", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "REVISION MECANICA: NO SUBEN TODOS LOS DOC, FACTURA: NO ANEXAN LA DOCUMENTACION OFICIAL DEL CLIENTE.", "primerPago": "ODP 23/03/2026", "segundoPago": ""}, {"id": "seed-62", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2150", "inv": "90251", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-03-19", "fRevision": "2026-03-20", "estatus": "EN PROCESO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "FACTURA: ACTIVIDAD ES INCORRECTA, TRANSUNION: NO ES EL DE LA UNIDAD, VERIFICACION QR: NO ANEXAN LOS qr DE LAS FACTURAS QUE ANEXAN", "primerPago": "2026-03-23", "segundoPago": ""}, {"id": "seed-63", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2147", "inv": "90249", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "RENAULT", "validador": "URIEL", "financiamiento": "NO", "fSolicitud": "2026-03-19", "fRevision": "2026-03-20", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "Verificacion de QR no agregaron la lctura de la factura origen, revisión mecanica no tiene nombre y firma del tecnico y las calcas  no es de la unidad, pantalla de alta de inventario la clave vehicular, monto de compra y venta.", "primerPago": "2026-03-23", "segundoPago": ""}, {"id": "seed-64", "cliente": "ARTURO CARRILLO SOLIS", "exp": "2151", "inv": "92843", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "NISSAN", "validador": "BLANCA", "financiamiento": "SI", "fSolicitud": "2026-03-20", "fRevision": "2026-03-20", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente factura origen. *Pendiente validación ante el distribuidor. *Pendiente qr factura origen.", "observaciones": "FACTURA ORIGEN: FOLIO REFACTURACION INCORRECTA. PANTALLA DEL ALTA: FAC PROV INCORRECTO", "primerPago": "2026-03-23", "segundoPago": "2026-04-22"}, {"id": "seed-65", "cliente": "BANY ARREVILLAGA VASQUEZ", "exp": "2132", "inv": "92836", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "NISSAN", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-20", "fRevision": "2026-03-20", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "PANTALLA SIA: EL NUMERO DE MOTOR ESTA MAL EN EL FORMULARIO", "primerPago": "2026-03-23", "segundoPago": ""}, {"id": "seed-66", "cliente": "PABLO LOPEZ DIAZ", "exp": "2155", "inv": "92844", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "GILBERTO", "financiamiento": "NO", "fSolicitud": "2026-03-20", "fRevision": "2026-03-20", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente firmas en PROFECO", "observaciones": "", "primerPago": "ODP 26/03/2026", "segundoPago": ""}, {"id": "seed-67", "cliente": "MARTHA ELBA GARCIA MORALES", "exp": "2152", "inv": "90252", "tipoToma": "COMPRA DIRECTA", "sucursal": "ISTMO", "marca": "RENAULT", "validador": "BLANCA", "financiamiento": "SI", "fSolicitud": "2026-03-23", "fRevision": "2026-03-23", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente factura origen y endoso.", "observaciones": "CURP Y COMPROBANTE VENCIDOS.", "primerPago": "2026-03-24", "segundoPago": "2026-03-26"}, {"id": "seed-68", "cliente": "JOSE MANUEL CAMACHO RAMOS", "exp": "2158", "inv": "92845", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-25", "fRevision": "2026-03-25", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente endoso.", "observaciones": "REVISION MECANICA: NO MARCARON LAS CASILLAS DE LA HOJA DE CALCAS. RFC: TIPO DE PERSONA ES INCORRECTA", "primerPago": "ODP 30/03/2026", "segundoPago": ""}, {"id": "seed-69", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2157", "inv": "90253", "tipoToma": "COMPRA DIRECTA", "sucursal": "OCOSINGO", "marca": "RENAULT", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-25", "fRevision": "2026-03-25", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas.", "observaciones": "FACTURA DE ORIGEN: DESCRIPCION INCORRECTA. XML: NO CARGARON EL ARCHIVO. VERIFICACION CONSULT: ILEGIBLE", "primerPago": "2026-03-16", "segundoPago": ""}, {"id": "seed-70", "cliente": "LETICIA MINERVA ZUÑIGA RAMIREZ", "exp": "2159", "inv": "92846", "tipoToma": "COMPRA DIRECTA", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "GILBERTO", "financiamiento": "NO", "fSolicitud": "2026-03-28", "fRevision": "2026-03-28", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "", "primerPago": "2026-03-30", "segundoPago": ""}, {"id": "seed-71", "cliente": "JOSE ANGEL ORTEGA MORENO", "exp": "2165", "inv": "92847", "tipoToma": "COMPRA DIRECTA", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "GILBERTO", "financiamiento": "NO", "fSolicitud": "2026-03-28", "fRevision": "2026-03-28", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "PENDIENTE ENDOSO FACTURA AMDA", "observaciones": "FECHA FACTURA EN FORMULARIO ES INCORRECTA", "primerPago": "2026-03-30", "segundoPago": ""}, {"id": "seed-72", "cliente": "LOURDES HERNANDEZ BOLAINA", "exp": "2163", "inv": "92848", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "OCOSINGO", "marca": "NISSAN", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-30", "fRevision": "2026-03-30", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "Pendiente firmas PROFECO", "observaciones": "SOL DE PAGO: NUMERO DE VIN INCORRECTO", "primerPago": "ODP 31/03/2026", "segundoPago": ""}, {"id": "seed-73", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2125", "inv": "92811", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "NISSAN", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-30", "fRevision": "2026-03-30", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas.", "observaciones": "FACTURA ORIGEN: SUBMARCA Y DESCRIPCION INCORRECTOS. PANTALLA DE ALTA: FACT PROV, NCI, MONTO DE ADQUISICIÓN, VALOR GUIA EBC INCORRECTOS.", "primerPago": "2026-01-29", "segundoPago": ""}, {"id": "seed-74", "cliente": "ERICKA ANAHI JIMENEZ PENAGOS", "exp": "2160", "inv": "92849", "tipoToma": "COMPRA DIRECTA", "sucursal": "OCOSINGO", "marca": "NISSAN", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-30", "fRevision": "2026-03-30", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-75", "cliente": "MONSERRATH SANCHEZ RODRIGUEZ", "exp": "2168", "inv": "90255", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "TAPACHULA", "marca": "RENAULT", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-30", "fRevision": "2026-03-30", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente validación de la factura.", "observaciones": "FACTURA: SUBMARCA Y DESCRIPCION INCORRECTOS. RFC: NUMERO EXT INCORRECTO. IMAGEN QR: FALTA FOTO.", "primerPago": "2026-03-31", "segundoPago": ""}, {"id": "seed-76", "cliente": "DANIEL SAMAYOA PENAGOS", "exp": "2169", "inv": "90256", "tipoToma": "COMPRA DIRECTA", "sucursal": "OCOSINGO", "marca": "RENAULT", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-03-31", "fRevision": "2026-03-31", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente endoso.", "observaciones": "", "primerPago": "2026-03-31", "segundoPago": ""}, {"id": "seed-77", "cliente": "YADIRA SOLIS PACHECO", "exp": "2173", "inv": "90259", "tipoToma": "COMPRA DIRECTA", "sucursal": "ISTMO", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-04-11", "fRevision": "2026-04-11", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "ALTA INV SIA: FAC PROV INCORRECTO", "primerPago": "2026-04-14", "segundoPago": ""}, {"id": "seed-78", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2126", "inv": "92814", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-04-13", "fRevision": "2026-04-13", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "ALTA INV SIA: FAC PROC INCORRECTO, MONTOS INCORRECTOS, ETIQUETA HOLOGRAFICA: NO SUBEN LA IMAGEN CORRECTA", "primerPago": "2026-04-15", "segundoPago": ""}, {"id": "seed-79", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2171", "inv": "90257", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "RENAULT", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-04-14", "fRevision": "2026-04-14", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "FACTURA: NO AGREGAN REFACTURACION. REPUVE: NO SE CONSULTARON TODAS LAS ENTIDADES. NO AGREGAN QR DE LA REFACTURACION. REVISION MECANICA: EL VIN TIENE DIGITOS ILEGIBLES. PANTALLA SIA: PRECIO VENTA Y ADQUISICIÓN INCORRECTOS. RFC Y COMPROBANTE VENCIDOS.", "primerPago": "2026-04-15", "segundoPago": ""}, {"id": "seed-80", "cliente": "CARLOS GONE ALATORRE", "exp": "2185", "inv": "90260", "tipoToma": "COMPRA DIRECTA", "sucursal": "ISTMO", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-04-15", "fRevision": "2026-04-15", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-81", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2170", "inv": "90258", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-04-15", "fRevision": "2026-04-15", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "FACTURA: TIENE MONTO DE GARANTIZADO Y NO APLICA, ALTA INV: MONTO DE ADQUICISION INCORRECTO", "primerPago": "", "segundoPago": ""}, {"id": "seed-82", "cliente": "JORGE LUIS GOMEZ DELGADO", "exp": "2174", "inv": "92850", "tipoToma": "SEMINUEVO X SEMINUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-04-16", "fRevision": "2026-04-16", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "FOLIO FACTURA NO TIENE GUION -", "primerPago": "ODP 22/04/2026", "segundoPago": ""}, {"id": "seed-83", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2188", "inv": "90258", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-04-17", "fRevision": "2026-04-17", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "VERIFICACION QR: NO ANEXA LECTURA AMDA, PANTALLA INV SIA MONTO ENAJENACION INCORRECTO, NUMERO DE MOTOR EN MIUSCULAS", "primerPago": "pendiente mandar a firma", "segundoPago": ""}, {"id": "seed-84", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2187", "inv": "90261", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "BLANCA", "financiamiento": "CREDITO NR", "fSolicitud": "2026-04-17", "fRevision": "2026-04-17", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-85", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2189", "inv": "90262", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "CREDITO NR", "fSolicitud": "2026-04-21", "fRevision": "2026-04-21", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-86", "cliente": "JOSE RODRIGO TOALA PEREZ", "exp": "2172", "inv": "92851", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-04-21", "fRevision": "2026-04-21", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente factura origen y endoso.\n*Pendiente validación de la factura origen.\n*Pendiente lectura qr AMDA Y SAT.", "observaciones": "FACTURA: SUBMARCA, DESCRIPCION, BAJA DE PLACAS, FOLIO REFACTURACION INCORRECTOS. CURP: MAL ESCRITA. REVISION MECANICA: TIENE CAMPOS EN BLANCO. PANTALLA DE SIA: MONTOS CAPTURADOS INCORRECTOS. RFC: DIRECCION INCORRECTA.", "primerPago": "2026-04-22", "segundoPago": ""}, {"id": "seed-87", "cliente": "MELISSA MUÑOZ CAMACHO", "exp": "2193", "inv": "90263", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-04-21", "fRevision": "2026-04-21", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas, y validacion factura origen", "observaciones": "", "primerPago": "2026-04-22", "segundoPago": ""}, {"id": "seed-88", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2199", "inv": "90258", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "RENAULT", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-04-22", "fRevision": "2026-04-22", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas.", "observaciones": "", "primerPago": "2026-04-24", "segundoPago": ""}, {"id": "seed-89", "cliente": "NORMA MAGALY GOMEZ LARA", "exp": "2194", "inv": "92852", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-04-24", "fRevision": "2026-04-24", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas, pendiente endoso", "observaciones": "INE: ILEGIBLE", "primerPago": "", "segundoPago": ""}, {"id": "seed-90", "cliente": "SERGIO DARINEL HERNANDEZ TORRES", "exp": "2201", "inv": "90267", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "SI", "fSolicitud": "2026-04-24", "fRevision": "2026-04-24", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente factura origen y endoso.", "observaciones": "", "primerPago": "2026-04-27", "segundoPago": "2026-05-06"}, {"id": "seed-91", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2195", "inv": "90264", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "BLANCA", "financiamiento": "CREDITO NR", "fSolicitud": "2026-04-24", "fRevision": "2026-04-24", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas.", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-92", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2200", "inv": "90266", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "CREDITO NR", "fSolicitud": "2026-04-24", "fRevision": "2026-04-24", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*pendiente baja de placas", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-93", "cliente": "CANDELARIA GOMEZ PEREZ", "exp": "2203", "inv": "90268", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-04-24", "fRevision": "2026-04-24", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "RFC: ACTIVIDAD ECONOMICA INCORRECTA", "primerPago": "2026-04-28", "segundoPago": ""}, {"id": "seed-94", "cliente": "JULIO ADRIAN TRUJILLO FLORES", "exp": "2198", "inv": "92853", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-04-24", "fRevision": "2026-04-24", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas y endoso", "observaciones": "FORMULARIO: NO TIENE INVENTARIO", "primerPago": "", "segundoPago": ""}, {"id": "seed-95", "cliente": "GUILEBALDO BULFRANO LOPEZ PEREZ", "exp": "2204", "inv": "90269", "tipoToma": "COMPRA DIRECTA", "sucursal": "OCOSINGO", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-04-25", "fRevision": "2026-04-25", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "NO CUENTA CON CAMBIO Y ACUSE DE CAMBIO DE ROL, ESTADO DE CUENTA VENCIDO", "primerPago": "2026-04-28", "segundoPago": ""}, {"id": "seed-96", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2206", "inv": "90270", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "CREDITO NR", "fSolicitud": "2026-04-28", "fRevision": "2026-04-28", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "PODER NOTARIAL, ES OTRO", "primerPago": "", "segundoPago": ""}, {"id": "seed-97", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2208", "inv": "90271", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "CREDITO NR", "fSolicitud": "2026-04-28", "fRevision": "2026-04-28", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "FORMULARIO: TIENE UN PESO DE MAS", "primerPago": "", "segundoPago": ""}, {"id": "seed-98", "cliente": "GUADALUPE BEATRIZ GONZALEZ MARTINEZ", "exp": "2192", "inv": "92854", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "OCOSINGO", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-04-28", "fRevision": "2026-04-28", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "", "primerPago": "ODP 5/05/2026", "segundoPago": ""}, {"id": "seed-99", "cliente": "TIBURCIO GALINDO RUIZ JIMENEZ", "exp": "2205", "inv": "92855", "tipoToma": "COMPRA DIRECTA", "sucursal": "OCOSINGO", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-04-28", "fRevision": "2026-04-29", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "", "primerPago": "2026-04-30", "segundoPago": ""}, {"id": "seed-100", "cliente": "CLAUDIA DEL ROSARIO JIMENEZ ESPINOSA", "exp": "2209", "inv": "92856", "tipoToma": "COMPRA DIRECTA", "sucursal": "PONIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "SI", "fSolicitud": "2026-04-29", "fRevision": "2026-04-29", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "muchas correcciones en el chat", "primerPago": "2026-05-04", "segundoPago": "2026-05-13"}, {"id": "seed-101", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2210", "inv": "90272", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "CREDITO NR", "fSolicitud": "2026-04-30", "fRevision": "2026-04-20", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas.", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-102", "cliente": "ALSOL CONTIGO SA DE CV, SOFOM ENR", "exp": "2207", "inv": "92857", "tipoToma": "COMPRA DIRECTA", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-05-02", "fRevision": "2026-05-02", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "PANTALLA DEL ALTA: MONTO DE ENAJENACION INCORRECTA. NUMERO DE PODER INCORRECTO.  LA REFACTURACION ES INCORRECTA: MODIFICARAN TODOS LOS APARTADOS DONDE SE REFLEJE EL DATO", "primerPago": "2026-05-12", "segundoPago": ""}, {"id": "seed-103", "cliente": "MARIA GUADALUPE MAYORGA ABARCA", "exp": "2190", "inv": "92858", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "NISSAN", "validador": "BLANCA", "financiamiento": "NO", "fSolicitud": "2026-05-04", "fRevision": "2026-05-04", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "FACTURA NIMEX: FOLIO DE REFACTURACION INCORRECTO", "primerPago": "2026-05-05", "segundoPago": ""}, {"id": "seed-104", "cliente": "YENY BERENICE GONZALEZ HILERIO", "exp": "2197", "inv": "90265", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-06", "fRevision": "2026-05-06", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*pendiente de endoso", "observaciones": "Formulario: fechas incorrectas, numero de motor incorrecto, factura origen: falta de endoso, alta en sia: monto de adquisicion incorrecto, constancia: direccion en el formulario mal capturada.", "primerPago": "2026-05-07", "segundoPago": ""}, {"id": "seed-105", "cliente": "PEDRO AGUILAR ALABAT", "exp": "2214", "inv": "92865", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-07", "fRevision": "2026-05-07", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "NO TIENE ENDOSO, NO TIENE NUMERO DE INV, SUBMARCA INCOMPLETA, FOTO DE VIN NO LEGIBLE, MONTO DE ENAJENACION INCORRECTO", "primerPago": "2026-05-08", "segundoPago": ""}, {"id": "seed-106", "cliente": "JUAN OLMOS CASAS", "exp": "2217", "inv": "92867", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-09", "fRevision": "2026-05-09", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "FACTURA NIMEX: NO ANEXAN ENDOSO, EN EL FORMULARIO NUMERO DE VIN MAL, NUMERO DE MOTOR MAL, BAJA DE PLACAS DICE NO, REVISION MECANICA: NUMERO DE INVENTARIO MAL, NO SUBEN LA FOTO DEL VIN, ALTA DE SIA: INV MAL, NO ANEXAN CLAVE VEHICULAR, MOTOR MAL", "primerPago": "", "segundoPago": ""}, {"id": "seed-107", "cliente": "ORANA FLORES CHAVEZ", "exp": "2219", "inv": "92866", "tipoToma": "COMPRA DIRECTA", "sucursal": "PONIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "SI", "fSolicitud": "2026-05-09", "fRevision": "2026-05-09", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas, endoso, validacion ante distribuidor, verificacion QR", "observaciones": "FACTURA: FOLIO REFACTURA INCORRECTO, REVISION MECANICA: MODELO INCORRECTONO FIRMAN LOS INVOLUCRADOS, PANTALLA INV SIA: MONTOS AL REVEZ, Y NO TIENE NCI", "primerPago": "2026-05-12", "segundoPago": ""}, {"id": "seed-108", "cliente": "ALICIA LOPEZ GUZMAN", "exp": "2211", "inv": "92868", "tipoToma": "COMPRA DIRECTA", "sucursal": "COMITAN", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-11", "fRevision": "2026-05-11", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "REVISION MECANICA: NO TIENE LA FOTO", "primerPago": "2026-05-11", "segundoPago": ""}, {"id": "seed-109", "cliente": "ABEL RODRIGUEZ ZALDIVAR", "exp": "2222", "inv": "90274", "tipoToma": "COMPRA DIRECTA", "sucursal": "ISTMO", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-11", "fRevision": "2026-05-11", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "", "primerPago": "2026-05-12", "segundoPago": ""}, {"id": "seed-110", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2223", "inv": "90273", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "CREDITO NR", "fSolicitud": "2026-05-11", "fRevision": "2026-05-11", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "NO TIENE RFC", "primerPago": "", "segundoPago": ""}, {"id": "seed-111", "cliente": "ARGENI CORTEZ FRANCO", "exp": "2224", "inv": "90275", "tipoToma": "COMPRA DIRECTA", "sucursal": "ISTMO", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-12", "fRevision": "2026-05-12", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "QR: NO ANEXAN EL QR DE LA REFACTURA", "primerPago": "2026-05-13", "segundoPago": ""}, {"id": "seed-112", "cliente": "NORA PATRICIA NUÑEZ ALIAS", "exp": "2225", "inv": "90277", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-13", "fRevision": "2026-05-13", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "PASAPORTE: NO ANEXAN COMPLETO", "primerPago": "", "segundoPago": ""}, {"id": "seed-113", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2226", "inv": "90279", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "CREDITO NR", "fSolicitud": "2026-05-14", "fRevision": "2026-05-14", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "NO ANEXAN ACTA CONSTITUTIVA", "primerPago": "", "segundoPago": ""}, {"id": "seed-114", "cliente": "JOSE JUAN LOPEZ MANUEL", "exp": "2227", "inv": "90280", "tipoToma": "COMPRA DIRECTA", "sucursal": "ISTMO", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-14", "fRevision": "2026-05-14", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "NO ANEXAN RFC", "primerPago": "2026-05-15", "segundoPago": ""}, {"id": "seed-115", "cliente": "HEBER ALVAREZ LOPEZ", "exp": "2228", "inv": "90278", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-15", "fRevision": "2026-05-15", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "INE: FOLIO EN FORMULARIO ESTA MAL, RFC: DIRECCION MAL", "primerPago": "", "segundoPago": ""}, {"id": "seed-116", "cliente": "GABRIEL LOPEZ GOMEZ", "exp": "2231", "inv": "92870", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "PALENQUE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-19", "fRevision": "2026-05-19", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "ALTA INV: NCI INCOMPLETO, RFC: DIRECCION INCOMPLETA", "primerPago": "", "segundoPago": ""}, {"id": "seed-117", "cliente": "ESTHER GALLEGOS LOPEZ", "exp": "2230", "inv": "90281", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "SI", "fSolicitud": "2026-05-19", "fRevision": "2026-05-20", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente endoso", "observaciones": "Formulario; Estan mal los datos del rfc y direccion", "primerPago": "2026-05-21", "segundoPago": "2026-05-26"}, {"id": "seed-118", "cliente": "SEBASTIAN MORENO VAZQUEZ", "exp": "2234", "inv": "90283", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "URIEL", "financiamiento": "SI", "fSolicitud": "2026-05-19", "fRevision": "2026-05-20", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "* Pendiente endoso", "observaciones": "", "primerPago": "2026-05-21", "segundoPago": "2026-05-27"}, {"id": "seed-119", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2232", "inv": "92862", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-19", "fRevision": "2026-05-20", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Baja de placa", "observaciones": "FACTURA ORIGEN: El formulario esta mal, VERIFICACION QR: Hace falta la consulta de la refactura  PANTALLA DE ALTA: El monto de precio de libro venta es incorrecto.", "primerPago": "2026-05-20", "segundoPago": ""}, {"id": "seed-120", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2233", "inv": "90282", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "BRAYAN", "financiamiento": "CREDITO NR", "fSolicitud": "2026-05-21", "fRevision": "2026-05-21", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-121", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2236", "inv": "92860", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-22", "fRevision": "2026-05-22", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "ALTA INV: CLAVE VEHICULAR INCORRECTA, VERIFICACION CONSULT: VIN NO COINCIDE", "primerPago": "", "segundoPago": ""}, {"id": "seed-122", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2229", "inv": "90276", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "CREDITO NR", "fSolicitud": "2026-05-23", "fRevision": "2026-05-23", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "verificacion consult: no tiene algo relacionado a la unidad", "primerPago": "", "segundoPago": ""}, {"id": "seed-123", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2237", "inv": "92859", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-23", "fRevision": "2026-05-23", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-124", "cliente": "MARIA ELENA GARCIA MENDEZ", "exp": "2238", "inv": "90285", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-23", "fRevision": "2026-05-23", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "factura origen: el folio de refactura en formulario esta mal", "primerPago": "", "segundoPago": ""}, {"id": "seed-125", "cliente": "LUCIA GUADALUPE LOPEZ SANTOS", "exp": "2239", "inv": "90284", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-23", "fRevision": "2026-05-23", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-126", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2240", "inv": "92861", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-23", "fRevision": "2026-05-25", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "COMPROBANTE DE DOMICILIO NO ESTA VIGENTE", "primerPago": "", "segundoPago": ""}, {"id": "seed-127", "cliente": "RAMON CRUZ AGUILAR", "exp": "2241", "inv": "92872", "tipoToma": "COMPRA DIRECTA", "sucursal": "COMITAN", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-25", "fRevision": "2026-05-25", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "", "primerPago": "2026-05-27", "segundoPago": ""}, {"id": "seed-128", "cliente": "NORMA GRISEL MALDONADO LOPEZ", "exp": "2235", "inv": "92871", "tipoToma": "COMPRA DIRECTA", "sucursal": "OCOSINGO", "marca": "NISSAN", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-05-26", "fRevision": "2026-05-26", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "", "primerPago": "2026-05-27", "segundoPago": ""}, {"id": "seed-129", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2242", "inv": "92863", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-05-26", "fRevision": "2026-05-26", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "EL RFC ESTA VENCIDO", "primerPago": "", "segundoPago": ""}, {"id": "seed-130", "cliente": "MARIA VAZQUEZ TORRES", "exp": "2245", "inv": "92873", "tipoToma": "SEMINUEVO X SEMINUEVO", "sucursal": "PALENQUE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-27", "fRevision": "2026-05-27", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-131", "cliente": "JESUS ANTONIO CASTRO GUILLEN", "exp": "2244", "inv": "92874", "tipoToma": "COMPRA DIRECTA", "sucursal": "COMITAN", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-27", "fRevision": "2026-05-27", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*PENDIENTE BAJA DE PLACAS, FACTURA ORIGEN Y ENDOSO, QR FACTURA, VALIDACION ANTE LA DISTRIBUIDORA", "observaciones": "No de folio esta mal", "primerPago": "2026-05-28", "segundoPago": ""}, {"id": "seed-132", "cliente": "SILVIA ORDOÑEZ RUIZ", "exp": "2246", "inv": "90286", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-05-28", "fRevision": "2026-05-28", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente endoso", "observaciones": "INE ilegible", "primerPago": "2026-06-01", "segundoPago": "2026-06-08"}, {"id": "seed-133", "cliente": "MONICA BECERRA HERNANDEZ", "exp": "2249", "inv": "92876", "tipoToma": "COMPRA DIRECTA", "sucursal": "PALENQUE", "marca": "NISSAN", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-05-28", "fRevision": "2026-05-28", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "Factura: No endosada. ALTA DE INVENTARIO: El precio de venta NSG inicial incorrecto y el precio de libro venta incorrecto. Acuse de cambio de rol (en espera)", "primerPago": "2026-05-29", "segundoPago": ""}, {"id": "seed-134", "cliente": "CRISTOBAL MENESES ARCOS", "exp": "3250", "inv": "92877", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "PALENQUE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-28", "fRevision": "2026-05-28", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "NO TIENE ENDOSO, RFC: LA DIRECCION EN EL FORMULARIO ESTA INCOMPLETA", "primerPago": "", "segundoPago": ""}, {"id": "seed-135", "cliente": "MARCO ANTONIO SANCHEZ DAZA", "exp": "3249", "inv": "90287", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-29", "fRevision": "2026-05-29", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "RFC: DIRECCION INCOMPLETA Y SUBMARCA INCOMPLETA", "primerPago": "2026-05-29", "segundoPago": ""}, {"id": "seed-136", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "2247", "inv": "92864", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-29", "fRevision": "2026-05-29", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "FORMULARIO: NUMERO DE TELEFONO INCOMPLETO", "primerPago": "2026-05-29", "segundoPago": ""}, {"id": "seed-137", "cliente": "PAULINA GOMEZ LOPEZ", "exp": "3252", "inv": "92878", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "PALENQUE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-29", "fRevision": "2026-05-29", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "REVISION MECANICA: FOTOGRAFIA VIN INCORRECTA, PANTALLA ALTA INV: MONTO DE ADQUISION INCORRECTO, IMAGEN QR DE LA COMPU: ES DE OTRA UNIDAD, ETIQUETA HOLOGRAFICA: ES DE OTRA UNIDAD, VERIFICACION CONSULT: ES DE OTRA UNIDAD", "primerPago": "", "segundoPago": ""}, {"id": "seed-138", "cliente": "CANDIDA BELLANERY TREJO RUIZ", "exp": "2248", "inv": "92875", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "OCOSINGO", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-29", "fRevision": "2026-05-29", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "INE: FOLIO INCORRECTO EN FORMULARIO", "primerPago": "", "segundoPago": ""}, {"id": "seed-139", "cliente": "ROBERTO HERNANDEZ ZAVALA", "exp": "3253", "inv": "90288", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "SI", "fSolicitud": "2026-05-29", "fRevision": "2026-05-29", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "*Pendiente factura origen y endoso", "observaciones": "", "primerPago": "2026-06-02", "segundoPago": "2026-06-04"}, {"id": "seed-140", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "3254", "inv": "90289", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "BRAYAN", "financiamiento": "CREDITO NR", "fSolicitud": "2026-05-30", "fRevision": "2026-05-30", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-141", "cliente": "ROSALBA ESTRADA GONZALEZ", "exp": "3251", "inv": "92880", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-30", "fRevision": "2026-05-30", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "*REVISION MECANICA: VIN INCORRECTO, PANTALLA INV; FAC PROV INCORRECTO", "primerPago": "", "segundoPago": ""}, {"id": "seed-142", "cliente": "ANA LUZ MENDEZ SAMUDIO", "exp": "3258", "inv": "92879", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "PALENQUE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-30", "fRevision": "2026-05-30", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "REVISION MECANICA: NO TIENE CALCA DEL VIN", "primerPago": "", "segundoPago": ""}, {"id": "seed-143", "cliente": "ANITA PEREZ CRUZ", "exp": "3256", "inv": "92881", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-05-30", "fRevision": "2026-05-30", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "*Etiqueta holografica no tiene", "primerPago": "", "segundoPago": ""}, {"id": "seed-144", "cliente": "OFELIA VELASCO AGUILAR", "exp": "3259", "inv": "92882", "tipoToma": "COMPRA DIRECTA", "sucursal": "COMITAN", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-06-03", "fRevision": "2026-06-03", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "NO TIENE RFC", "observaciones": "*Pendiente baja de placas", "primerPago": "2026-06-04", "segundoPago": ""}, {"id": "seed-145", "cliente": "DANIELA DEL BARCO REINOS", "exp": "3260", "inv": "90290", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "SI", "fSolicitud": "2026-06-04", "fRevision": "2026-06-04", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*pendiente endoso", "primerPago": "2026-06-05", "segundoPago": ""}, {"id": "seed-146", "cliente": "LUIS ALBERTO ZUÑIGA GOMEZ", "exp": "3261", "inv": "92883", "tipoToma": "COMPRA DIRECTA", "sucursal": "COMITAN", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-06-05", "fRevision": "2026-06-05", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "", "observaciones": "COMPROBANTE DOMICILIO INCORRECTO", "primerPago": "2026-06-08", "segundoPago": ""}, {"id": "seed-147", "cliente": "DEYANIRA YANET MORALES LUNA", "exp": "3263", "inv": "92884", "tipoToma": "COMPRA DIRECTA", "sucursal": "COMITAN", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-06-08", "fRevision": "2026-06-08", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*Pendiente baja de placas", "primerPago": "2026-06-08", "segundoPago": ""}, {"id": "seed-148", "cliente": "CONCEPCION PEREZ CRUZ", "exp": "3262", "inv": "92885", "tipoToma": "COMPRA DIRECTA", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-06-08", "fRevision": "2026-06-08", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "ALTA SIA: MONTO ENAJENACION INCORRECTA", "observaciones": "*PENDIENTE ENDOSO Y BAJA DE PLACAS", "primerPago": "2026-06-09", "segundoPago": ""}, {"id": "seed-149", "cliente": "LUIS ALBERTO ZUÑIGA GOMEZ", "exp": "3264", "inv": "92891", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-06-15", "fRevision": "2026-06-16", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "BAJA DE PLACAS", "observaciones": "ALTA SIA: DOCUMENTO INCOMPLETO, COMPROBANTE DOMICILIO: NO ES DE LUZ", "primerPago": "", "segundoPago": ""}, {"id": "seed-150", "cliente": "EMILIO SANTIZ LOPEZ", "exp": "3268", "inv": "92892", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "OCOSINGO", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-06-16", "fRevision": "2026-06-16", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-151", "cliente": "JOEL ROBLES CIGARROA", "exp": "3270", "inv": "92893", "tipoToma": "COMPRA DIRECTA", "sucursal": "PALENQUE", "marca": "NISSAN", "validador": "BRAYAN", "financiamiento": "SI", "fSolicitud": "2026-06-16", "fRevision": "2026-06-16", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "REVISION MECANICA:  NO TIENE NUMERO DE INVENTARIO", "observaciones": "*Pendiente baja de placas, pendiente endoso", "primerPago": "2026-06-18", "segundoPago": ""}, {"id": "seed-152", "cliente": "JOSE LUIS LOPEZ GORDILLO", "exp": "3269", "inv": "92895", "tipoToma": "SEMINUEVO X SEMINUEVO", "sucursal": "COMITAN", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-06-18", "fRevision": "2026-06-18", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "PANTALLA ALTA INV: MONTO DE ADQUISICION ES INCORRECTO", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-153", "cliente": "ANTONIA DE JESUS CANCINO MUÑOZ", "exp": "3274", "inv": "90292", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-06-19", "fRevision": "2026-06-19", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "RFC: DOCUMENTO VENCIDO", "observaciones": "", "primerPago": "2026-06-22", "segundoPago": ""}, {"id": "seed-154", "cliente": "EDILBERTO PEREZ PEREZ", "exp": "3272", "inv": "2894", "tipoToma": "SEMINUEVO X SEMINUEVO", "sucursal": "PALENQUE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-06-23", "fRevision": "2026-06-23", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*Pendiente baja de placas", "primerPago": "", "segundoPago": ""}, {"id": "seed-155", "cliente": "CIELO IVON DIAZ VAZQUEZ", "exp": "3290", "inv": "92896", "tipoToma": "COMPRA DIRECTA", "sucursal": "PONIENTE", "marca": "NISSAN", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-06-24", "fRevision": "2026-06-24", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "Alta de inventario en SIA fac prov incorrecto, QR de factura orifen no estan|", "observaciones": "*Pendiente bajda de placas y endoso", "primerPago": "2026-06-26", "segundoPago": ""}, {"id": "seed-156", "cliente": "MARIA ISABEL RAMOS CRUZ", "exp": "3280", "inv": "92905", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "OCOSINGO", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-06-25", "fRevision": "2026-06-25", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "ALTA INV SIA: EL TIPO DE UNIDAD DICE CHASIS Y ES CHA Y EN FORMULARIO TAMBIEN", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-157", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "3292", "inv": "90293", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "CREDITO NR", "fSolicitud": "2026-06-25", "fRevision": "2026-06-25", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "FACTURA ORIGEN: SUBEN EL DISTRINET", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-158", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "3297", "inv": "90294", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "CREDITO NR", "fSolicitud": "2026-06-25", "fRevision": "2026-06-26", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "BAJA DE PLACAS: NO FUE ENPLACADA Y SE PIDE AUTORIZACION", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-159", "cliente": "QUALITAS COMPAÑÍA DE SEGUROS SA DE CV", "exp": "3287", "inv": "92902", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-06-25", "fRevision": "2026-06-25", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "", "primerPago": "PAGADO", "segundoPago": ""}, {"id": "seed-160", "cliente": "CARLOS DE LA CRUZ GUTIERREZ", "exp": "3273", "inv": "92906", "tipoToma": "COMPRA DIRECTA", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-06-25", "fRevision": "2026-06-25", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "", "primerPago": "2026-06-26", "segundoPago": ""}, {"id": "seed-161", "cliente": "DARIANA ANZUETO CRUZ", "exp": "3302", "inv": "92908", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-06-26", "fRevision": "2026-06-26", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "", "primerPago": "2026-06-29", "segundoPago": ""}, {"id": "seed-162", "cliente": "QUALITAS COMPAÑÍA DE SEGUROS SA DE CV", "exp": "3275", "inv": "92898", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "NISSAN", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-06-26", "fRevision": "2026-06-27", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "Repuve no consuktado de forma correcta", "primerPago": "PAGADO", "segundoPago": ""}, {"id": "seed-163", "cliente": "QUALITAS COMPAÑÍA DE SEGUROS SA DE CV", "exp": "3279", "inv": "92903", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "NISSAN", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-06-26", "fRevision": "2026-06-27", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "Revision mecanica con tachaduras y el modelo diferente, alta de inventario el color esta mal, consult modelo erroneo.", "primerPago": "PAGADO", "segundoPago": ""}, {"id": "seed-164", "cliente": "QUALITAS COMPAÑÍA DE SEGUROS SA DE CV", "exp": "3285", "inv": "92900", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "NISSAN", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-06-26", "fRevision": "2026-06-27", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "Alta de inventario, color y precio es incorrecto", "primerPago": "PAGADO", "segundoPago": ""}, {"id": "seed-165", "cliente": "QUALITAS COMPAÑÍA DE SEGUROS SA DE CV", "exp": "3286", "inv": "92901", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "NISSAN", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-06-26", "fRevision": "2026-06-27", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "Consult con modelo distinto", "primerPago": "PAGADO", "segundoPago": ""}, {"id": "seed-166", "cliente": "QUALITAS COMPAÑÍA DE SEGUROS SA DE CV", "exp": "3288", "inv": "92904", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "NISSAN", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-06-26", "fRevision": "2026-06-27", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "CONSULT con modelo distinto, revision mecanica modelo y color incorrecto, alta de inventario color y precio incorrecto, falta autorizacion de baja de placas", "primerPago": "PAGADO", "segundoPago": ""}, {"id": "seed-167", "cliente": "AKEMY NOEMI ALVARADO MOSQUEDA", "exp": "3300", "inv": "92907", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "SI", "fSolicitud": "2026-06-27", "fRevision": "2026-06-29", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "Formulario vin incorrecto, alta en sia: no se visualiza el vin", "observaciones": "*Pendiente factura origen y endoso, validacion factura origen, baja de placas", "primerPago": "2026-06-29", "segundoPago": "2026-07-09"}, {"id": "seed-168", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "3304", "inv": "92888", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-06-29", "fRevision": "2026-06-29", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "", "primerPago": "PAGADO", "segundoPago": ""}, {"id": "seed-169", "cliente": "ELISEO HERNANDEZ BURGUETE", "exp": "3265", "inv": "90291", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "SI", "fSolicitud": "2026-06-29", "fRevision": "2026-06-29", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "Factura origen: formulario, descripcion y submarca incorrectos, e-mail y numero de telefono son incorrectos.", "observaciones": "", "primerPago": "2026-07-01", "segundoPago": "2026-07-24"}, {"id": "seed-170", "cliente": "QUALITAS COMPAÑÍA DE SEGUROS SA DE CV", "exp": "3284", "inv": "92899", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-06-29", "fRevision": "2026-06-29", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "CONSULT: MENCIONA MODELO 2018 Y ES MODELO 2025, FORMULARIO: EXPEDIDA ESTA MAL DEBE SER NISSAN MEXICANA", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-171", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "3307", "inv": "92890", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-06-29", "fRevision": "2026-06-29", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*Pendiente baja de placas", "primerPago": "PAGADO", "segundoPago": ""}, {"id": "seed-172", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "3308", "inv": "92889", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-06-29", "fRevision": "2026-06-29", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*Pendiente baja de placas", "primerPago": "PAGADO", "segundoPago": ""}, {"id": "seed-173", "cliente": "GUILLERMINA LIEVANO LARA", "exp": "3303", "inv": "92914", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-06-30", "fRevision": "2026-06-30", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "REVISION MECANICA: CALCA NO LEGIBLE, CONSULT: MODELO INCORRECTO", "observaciones": "*Pendiente baja de placas", "primerPago": "", "segundoPago": ""}, {"id": "seed-174", "cliente": "MARIA DE LOURDES GUILLEN DE LEON", "exp": "3306", "inv": "92915", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "COMITAN", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-06-30", "fRevision": "2026-06-30", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "*Pendiente baja de placas, validacion factura ante el distribuidor", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-175", "cliente": "SONIA DEL CARMEN MARTINEZ VAZQUEZ", "exp": "3278", "inv": "92913", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "TENOSIQUE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-06-30", "fRevision": "2026-06-30", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente endoso", "observaciones": "RFC: FORMULARIO CALLE NO COMPLETA", "primerPago": "", "segundoPago": ""}, {"id": "seed-176", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "3309", "inv": "90295", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "CREDITO NR", "fSolicitud": "2026-06-30", "fRevision": "2026-06-30", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-177", "cliente": "ALMA ROSA ALVARADO MORALES", "exp": "3311", "inv": "92917", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "COMITAN", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-07-02", "fRevision": "2026-07-02", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-178", "cliente": "DANIELA DEL CARMEN PENAGOS SOLIS", "exp": "3293", "inv": "92916", "tipoToma": "COMPRA DIRECTA", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-07-02", "fRevision": "2026-07-02", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "RFC: Esta mal escrito en el formulario, REVISION MECANICA: No menciona numero de INV y no es la fotografia del vin la correcta. CONSULT: No menciona el modelo correcto", "primerPago": "2026-07-04", "segundoPago": ""}, {"id": "seed-179", "cliente": "HUMBERTO DIAZ VAZQUEZ", "exp": "3301", "inv": "92918", "tipoToma": "COMPRA DIRECTA", "sucursal": "OCOSINGO", "marca": "NISSAN", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-07-06", "fRevision": "2026-07-06", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "El nombre de quien emite la factura esta mal ", "primerPago": "2026-07-08", "segundoPago": ""}, {"id": "seed-180", "cliente": "JULLIETH AMAIRANI ARREOLA GONZALEZ", "exp": "3316", "inv": "90297", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-07-08", "fRevision": "2026-07-08", "estatus": "EN PROCESO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "Validacion fac origen: no anexan el correo que es, estado de cuenta bancario: esta vencido", "primerPago": "2026-07-10", "segundoPago": ""}, {"id": "seed-181", "cliente": "ZOILA COELLO MONTES DE OCA", "exp": "3315", "inv": "90296", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-07-08", "fRevision": "2026-07-09", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-182", "cliente": "WILBERT CARPIO CORONEL", "exp": "3266", "inv": "92897", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "SI", "fSolicitud": "2026-06-23", "fRevision": "2026-06-23", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Factura origen", "observaciones": "", "primerPago": "2026-06-24", "segundoPago": "2026-07-09"}, {"id": "seed-183", "cliente": "NICOLAS PEREZ DOMINGUEZ", "exp": "3312", "inv": "92919", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "OCOSINGO", "marca": "NISSAN", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-07-10", "fRevision": "2026-07-11", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Banco virtual, Profeco y Factura original", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-184", "cliente": "ORALIA DE JESUS CORDERO CANCINO", "exp": "3313", "inv": "92920", "tipoToma": "COMPRA DIRECTA", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-07-11", "fRevision": "2026-07-13", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente endoso, *Pendiente baja de placas", "observaciones": "", "primerPago": "2026-07-14", "segundoPago": ""}, {"id": "seed-185", "cliente": "CONCEPCION DEL CARMEN GUTIERREZ ORTIZ", "exp": "3317", "inv": "92921", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-07-14", "fRevision": "2026-07-14", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "Formulario: tiene un 5 de mas en descripcion de la unidad", "primerPago": "2026-07-15", "segundoPago": ""}, {"id": "seed-186", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "3322", "inv": "92912", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-07-15", "fRevision": "2026-07-15", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "", "primerPago": "2026-07-15", "segundoPago": ""}, {"id": "seed-187", "cliente": "GERSON SANTIAGO CASTELLANO", "exp": "3318", "inv": "92902", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-07-17", "fRevision": "2026-07-17", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Factura de origen, *Baja de placas, *Verificacion QR, *Validacion de fac ante la distribuidora, *Banco virtual, *Profeco *Factura original", "observaciones": "Fecha de factura, CURP mal escrita, Direccion en el RFC", "primerPago": "", "segundoPago": ""}, {"id": "seed-188", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "3324", "inv": "92911", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-07-17", "fRevision": "2026-07-17", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "* Baja de placas", "observaciones": "*Revision mecanica sin firmas de servicio", "primerPago": "2026-07-17", "segundoPago": ""}, {"id": "seed-189", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "3326", "inv": "92910", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-07-17", "fRevision": "2026-07-17", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "* Baja de placas", "observaciones": "* La alta en sia es distinta a la factura", "primerPago": "2026-07-17", "segundoPago": ""}, {"id": "seed-190", "cliente": "ROSA AURORA TOLEDO LOPEZ", "exp": "3327", "inv": "92925", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-07-21", "fRevision": "2026-07-21", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "*Verificacion QR, la lectura dice estado de cfdi cancelado", "primerPago": "", "segundoPago": ""}, {"id": "seed-191", "cliente": "GERSON SANTIAGO CASTELLANO", "exp": "3329", "inv": "92922", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-07-21", "fRevision": "2026-07-21", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Factura de origen, *Baja de placas, *Verificacion QR, *Validacion de fac ante la distribuidora, *Banco virtual, *Profeco *Factura original", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-192", "cliente": "CLAUDIA CECILIA HERNANDEZ HERNANDEZ", "exp": "3330", "inv": "90298", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-07-22", "fRevision": "2026-07-22", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-193", "cliente": "LUCIA DEL CARMEN VELAZQUEZ HERNANDEZ", "exp": "3323", "inv": "92929", "tipoToma": "COMPRA DIRECTA", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-07-24", "fRevision": "2026-07-24", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "FACTURA: NO TIENE ENDOSO, REVISION MECANICA: NO TIENE NUMERO DE INV, ALTA SIA: NO SE REFLEJAN LOS MONTOS DE ENAJENACION Y TOTAL", "primerPago": "2026-07-27", "segundoPago": ""}, {"id": "seed-194", "cliente": "CONSUELO LURIA LOPEZ", "exp": "3328", "inv": "92927", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "COMITAN", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-07-24", "fRevision": "2026-07-24", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "FORMULARIO: ESTA MAL LA DESCRIPCION Y SUBMARCA, BAJA DE PLACAS DICE \"SI\"", "primerPago": "", "segundoPago": ""}, {"id": "seed-195", "cliente": "CANDIDO GUTIERREZ GOMEZ", "exp": "3333", "inv": "92930", "tipoToma": "COMPRA DIRECTA", "sucursal": "PALENQUE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-07-25", "fRevision": "2026-07-27", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*Pendiente baja de placas", "observaciones": "* NO TIENE RFC, FACTURA: VIN INCOMPLETO FALTA UNA C, FORMULARIO, FOLIO REFACTURA ES INCORRECTO, MONTO DE VENTA DICE 3 MILLONES, INE: FOLIO INCOMPLETO, REVISION MECANICA: NO TIENE NUMERO DE INV, ALTA SIA: NUMERO INV INCORRECTO, XML: VIN INCOMPLETO, ESTADO DE CUENTA: VIGENCIA EXPIRO, ", "primerPago": "2026-07-29", "segundoPago": ""}, {"id": "seed-196", "cliente": "JOSE JUAN GUZMAN PEREZ", "exp": "3343", "inv": "92928", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "TENOSIQUE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-07-27", "fRevision": "2026-07-27", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "FACTURA ORIGEN, CREDENCIAL, CURP, VALIDACION, VERIFICACION QR, REVISION MECANICA, ALTA SIA, ", "primerPago": "", "segundoPago": ""}, {"id": "seed-197", "cliente": "JOSE JUAN GUZMAN PEREZ", "exp": "3332", "inv": "92926", "tipoToma": "COMPRA DIRECTA", "sucursal": "TENOSIQUE", "marca": "NISSAN", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-07-27", "fRevision": "2026-07-27", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "FACTURA DE ORIGEN, REVISION MECANICA, PANTALLA DE ALTA EN SIA, VALIDACION DE QR", "primerPago": "2026-07-30", "segundoPago": ""}, {"id": "seed-198", "cliente": "DAVID FRANCISCO ARGUELLO GARCIA", "exp": "3325", "inv": "92931", "tipoToma": "COMPRA DIRECTA", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-07-29", "fRevision": "2026-07-29", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "", "primerPago": "2026-07-29", "segundoPago": ""}, {"id": "seed-199", "cliente": "EMPERATRIZ OCAÑA ESCOBAR", "exp": "3335", "inv": "92932", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "SI", "fSolicitud": "2026-07-29", "fRevision": "2026-07-29", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "correo electronico incorrecto", "observaciones": "*Pendiente validacion de factura, factura origen, verificacion QR", "primerPago": "2026-07-30", "segundoPago": ""}, {"id": "seed-200", "cliente": "CAROLINA VELAZQUEZ ESQUINCA", "exp": "3337", "inv": "90300", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-07-29", "fRevision": "2026-07-29", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "ALTA SIA: MONTO TOTAL INCORRECTO", "observaciones": "*Pendiente baja de placas", "primerPago": "", "segundoPago": ""}, {"id": "seed-201", "cliente": "JORGE ANTONIO AGUILAR CUESTA", "exp": "3336", "inv": "90299", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "SI", "fSolicitud": "2026-07-29", "fRevision": "2026-07-29", "estatus": "EN PROCESO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*Pendiente factura origen.", "primerPago": "2026-07-30", "segundoPago": "2026-08-10"}, {"id": "seed-202", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "3338", "inv": "90301", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-07-30", "fRevision": "2026-07-30", "estatus": "EN PROCESO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*Pendiente baja de placas", "primerPago": "", "segundoPago": ""}, {"id": "seed-203", "cliente": "MARIELA ARGUETA MEGCHUN", "exp": "3339", "inv": "90302", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "TAPACHULA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-07-30", "fRevision": "2026-07-30", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "*ALTA SIA: NO SE MIRA COMPLETO, RFC: DIRECCION MUNICIPIO DICE PICHUCALCO, FACTURA ORIGEN NO TIENE ENDOSO", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-204", "cliente": "JAIME JIMENEZ LOPEZ", "exp": "3334", "inv": "92934", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "OCOSINGO", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-07-31", "fRevision": "2026-07-31", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "*FORMULARIO: TIPO DE PERSONA ES INCORRECTA", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-205", "cliente": "ANA MARIA CRUZ LOPEZ", "exp": "3340", "inv": "92933", "tipoToma": "COMPRA DIRECTA", "sucursal": "PONIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-07-31", "fRevision": "2026-07-31", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*ALTA SIA: NUMERO MOTOR INCORRECTO, FORMULARIO DIRECCION MAL", "observaciones": "*PENDIENTE BAJA DE PLACAS, ENDOSO, VERIFICACION QR, VALIDACION", "primerPago": "2026-08-03", "segundoPago": ""}, {"id": "seed-206", "cliente": "RAUL HERNANDEZ HERNANDEZ", "exp": "3341", "inv": "90303", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "TAPACHULA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-07-31", "fRevision": "2026-07-31", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*RFC: EN FORMULARIO TIPO DE PERSONA INCORRECTA", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "", "segundoPago": ""}, {"id": "seed-207", "cliente": "ESMERALDA SANDOVAL CRUZ", "exp": "3342", "inv": "90304", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "TAPACHULA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-07-31", "fRevision": "2026-07-31", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "", "segundoPago": ""}, {"id": "seed-208", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "3345", "inv": "90305", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-05", "fRevision": "2026-08-05", "estatus": "EN PROCESO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "2026-08-06", "segundoPago": ""}, {"id": "seed-209", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "3346", "inv": "90306", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-05", "fRevision": "2026-08-05", "estatus": "EN PROCESO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "2026-08-06", "segundoPago": ""}, {"id": "seed-210", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "3348", "inv": "90308", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-05", "fRevision": "2026-08-05", "estatus": "EN PROCESO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "2026-08-10", "segundoPago": ""}, {"id": "seed-211", "cliente": "ELIET IDUVINA ALEGRIA MORALES", "exp": "3351", "inv": "92935", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-06", "fRevision": "2026-08-06", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "*FACTURA ORIGEN: NOMBRE DEL CLIENTE ES INCORRECTO", "observaciones": "", "primerPago": "2026-08-10", "segundoPago": ""}, {"id": "seed-212", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "3350", "inv": "90310", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-06", "fRevision": "2026-08-06", "estatus": "EN PROCESO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "2026-08-11", "segundoPago": ""}, {"id": "seed-213", "cliente": "OMAR ALEJANDRO LAYNEZ CONTRERAS", "exp": "3353", "inv": "92936", "tipoToma": "COMPRA DIRECTA", "sucursal": "PALENQUE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-07", "fRevision": "2026-08-07", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*FACTURA ORIG: DICE USADA Y NO TIENE ENDOSO, INE: NUMERO DE FOLIO ES INCORRECTO, VALIDAR QUE SUBAN LA FACTURA QUE DICE NUEVO Y LO QUE CONLLEVA", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "2026-08-11", "segundoPago": ""}, {"id": "seed-214", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "3358", "inv": "90311", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-08", "fRevision": "2026-08-08", "estatus": "EN PROCESO", "docsPend": true, "detalleDocsPend": "*RFC: TENIA OTRO DOC, TRANSUNION: DECIA SINIESTRADO", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "2026-08-10", "segundoPago": ""}, {"id": "seed-215", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "3354", "inv": "90309", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-08-07", "fRevision": "2026-08-07", "estatus": "EN PROCESO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "2026-08-10", "segundoPago": ""}, {"id": "seed-216", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "3356", "inv": "90312", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-10", "fRevision": "2026-08-10", "estatus": "EN PROCESO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "2026-08-10", "segundoPago": ""}, {"id": "seed-217", "cliente": "LUCERO LOPEZ PEREZ", "exp": "3357", "inv": "92937", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-10", "fRevision": "2026-08-10", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "VALIDACION ANTE DISTRIBUIDOR E IMAGEN QR DE LA COMPUTADORA: ANEXAN OTROS DOCUMENTOS QUE NO SON", "observaciones": "", "primerPago": "2026-08-12", "segundoPago": ""}, {"id": "seed-218", "cliente": "JAVIER SANTIZ GOMEZ", "exp": "3355", "inv": "92939", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "OCOSINGO", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-11", "fRevision": "2026-08-11", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-219", "cliente": "YARIBETH REYES GALLARDO", "exp": "3359", "inv": "92940", "tipoToma": "COMPRA DIRECTA", "sucursal": "PONIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "SI", "fSolicitud": "2026-08-11", "fRevision": "2026-08-11", "estatus": "EN PROCESO", "docsPend": false, "detalleDocsPend": "ALTA SIA: MONTOS MAL, REVISION MECANICA: INV MAL, FORMULARIO: DESCRIPCION Y SUBMARCA MAL.", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "2026-08-12", "segundoPago": ""}, {"id": "seed-220", "cliente": "LILIANA DOMINGUEZ LAZOS", "exp": "3366", "inv": "90314", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-11", "fRevision": "2026-08-11", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "", "segundoPago": ""}, {"id": "seed-221", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "3368", "inv": "90315", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-12", "fRevision": "2026-08-12", "estatus": "EN PROCESO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "2026-08-12", "segundoPago": ""}, {"id": "seed-222", "cliente": "ERIN GERONIMO MARIN", "exp": "3349", "inv": "90307", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "TAPACHULA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-12", "fRevision": "2026-08-13", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*FORMULARIO EN GENERAL", "observaciones": "*PENDIENTE ENDOSO", "primerPago": "", "segundoPago": ""}, {"id": "seed-223", "cliente": "ONEIDA PEREZ CASTELLANO", "exp": "3369", "inv": "92959", "tipoToma": "COMPRA DIRECTA", "sucursal": "ORIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-14", "fRevision": "2026-08-14", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "*FORMULARIO, COLOR DIRECCION, REVISION MECANICA: ANEXARAN CORREO, ESTADO DE CUENTA VENCIDO.", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "", "segundoPago": ""}, {"id": "seed-224", "cliente": "JANI DEL CARMEN RAMIREZ SUAREZ", "exp": "3374", "inv": "92960", "tipoToma": "COMPRA DIRECTA", "sucursal": "COMITAN", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-17", "fRevision": "2026-08-17", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "2026-08-18", "segundoPago": ""}, {"id": "seed-225", "cliente": "ALEJANDRINA ANTILLON ARAUJO", "exp": "3352", "inv": "92961", "tipoToma": "COMPRA DIRECTA", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-18", "fRevision": "2026-08-18", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "FACTURA NO TIENE ENDOSO, FORMULARIO: EXPEDIDA, FECHA FACTURa, FOLIO FACTURA, SUBMARCA Y DESCRIPCION, ALTA SIA: FACPROV INCORECTO", "observaciones": "", "primerPago": "2026-08-21", "segundoPago": ""}, {"id": "seed-226", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "3367", "inv": "90313", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-18", "fRevision": "2026-08-18", "estatus": "EN PROCESO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "2026-08-19", "segundoPago": ""}, {"id": "seed-227", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "3390", "inv": "90317", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-18", "fRevision": "2026-08-18", "estatus": "EN PROCESO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "2026-08-19", "segundoPago": ""}, {"id": "seed-228", "cliente": "ARMANDO ESTEBAN CORZO VILLANUEVA", "exp": "3392", "inv": "90318", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-19", "fRevision": "2026-08-19", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-229", "cliente": "RICARDO OBED CRUZ SANTOS", "exp": "3393", "inv": "92962", "tipoToma": "COMPRA DIRECTA", "sucursal": "PONIENTE", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-19", "fRevision": "2026-08-20", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "*COMPROBANTE DOM VENCIDO, FORMULARIO CODIGO POSTAL MAL", "observaciones": "", "primerPago": "2026-08-21", "segundoPago": ""}, {"id": "seed-230", "cliente": "JESUS EMILIO OSEGUERA CORDERO", "exp": "3375", "inv": "92964", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-08-20", "fRevision": "2026-08-20", "estatus": "EN PROCESO", "docsPend": false, "detalleDocsPend": "", "observaciones": "*CALCA DE MOTOR ILEGIBLE EN LA REVISION MECANICA, * VERIFICACION DE CONSULT EL MODELO ES INCORRECTO", "primerPago": "", "segundoPago": ""}, {"id": "seed-231", "cliente": "CARLOS ENRIQUE ZENTENO ZENTENO", "exp": "3388", "inv": "90316", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-20", "fRevision": "2026-08-20", "estatus": "EN PROCESO", "docsPend": false, "detalleDocsPend": "*COMPROBANTE DE DOMICILIO: NO ES DE LUZ, ALTA SIA: NO DESPLIEGA EL APARTADO DE DATOS TOMA", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-232", "cliente": "MARIA DE LOURDES GUZMAN VAZQUEZ", "exp": "3400", "inv": "92965", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "PALENQUE", "marca": "NISSAN", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-08-22", "fRevision": "2026-08-24", "estatus": "EN PROCESO", "docsPend": true, "detalleDocsPend": "*BAJA DE PLCAS* ENDOSO DE FACTURA", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-233", "cliente": "OLGA VARONA HERNANDEZ", "exp": "3391", "inv": "92966", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "COMITAN", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-25", "fRevision": "2026-08-25", "estatus": "EN PROCESO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "", "segundoPago": ""}, {"id": "seed-234", "cliente": "LILIANA DEL CARMEN CORZO HERNANDEZ", "exp": "3404", "inv": "90319", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-25", "fRevision": "2026-08-25", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "2026-08-26", "segundoPago": ""}, {"id": "seed-235", "cliente": "MANUEL DE JESUS CENICEROS MENDEZ", "exp": "3389", "inv": "92980", "tipoToma": "COMPRA DIRECTA", "sucursal": "COMITAN", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-26", "fRevision": "2026-08-26", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "2026-08-26", "segundoPago": ""}, {"id": "seed-236", "cliente": "IVAN LOPEZ KRAMSKY", "exp": "3398", "inv": "94", "tipoToma": "COMPRA DIRECTA", "sucursal": "SAN CRISTOBAL", "marca": "CHANGAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-26", "fRevision": "2026-08-26", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "Validar toda la toma desde cero", "observaciones": "", "primerPago": "2026-08-28", "segundoPago": ""}, {"id": "seed-237", "cliente": "JORGE ALEJANDRO ANDRADE ROBLERO", "exp": "29282", "inv": "3401", "tipoToma": "COMPRA DIRECTA", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-27", "fRevision": "2026-08-26", "estatus": "COMPLETADO", "docsPend": false, "detalleDocsPend": "NUMERO DE INV MAL, ANEXAR LA FOTOGRAFIA DEL VIN, ANEXAR LA LECTURA QR", "observaciones": "", "primerPago": "2026-08-29", "segundoPago": ""}, {"id": "seed-238", "cliente": "FORZA ARRENDADORA AUTOMOTRIZ SA DE CV", "exp": "92947", "inv": "3361", "tipoToma": "COMPRA DIRECTA", "sucursal": "COMITAN", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-27", "fRevision": "2026-08-27", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "NO TIENE LAS SIGLAS, COMP DOM NO ES EL CORRECTO, ALTA SIA: MONTOS INCORRECTOS", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "", "segundoPago": ""}, {"id": "seed-239", "cliente": "FORZA ARRENDADORA AUTOMOTRIZ SA DE CV", "exp": "92946", "inv": "3364", "tipoToma": "COMPRA DIRECTA", "sucursal": "COMITAN", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-27", "fRevision": "2026-08-27", "estatus": "COMPLETADO", "docsPend": true, "detalleDocsPend": "REVISION MEC: NO SE VISUALIZA EL ULTIMO NUMERO DEL VIN", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "", "segundoPago": ""}, {"id": "seed-240", "cliente": "NR FINANCE MEXICO S.A. DE C.V.", "exp": "90320", "inv": "3410", "tipoToma": "COMPRA DIRECTA", "sucursal": "TUXTLA", "marca": "RENAULT", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-28", "fRevision": "2026-08-28", "estatus": "EN PROCESO", "docsPend": true, "detalleDocsPend": "FORMULARIO: FOLIO REFACTURA INCORRECTO", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "", "segundoPago": ""}, {"id": "seed-241", "cliente": "JESUS MANUEL GARCIA ARGUELLO", "exp": "92981", "inv": "3399", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-28", "fRevision": "2026-08-28", "estatus": "EN PROCESO", "docsPend": true, "detalleDocsPend": "REVISION MEC: NO TIENE FOTO DEL VIN, TIPO DE PERSONA INCORRECTO, COMPROBANTE DOM: NO ANEXAN EL RECIBO DE LUZ", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "", "segundoPago": ""}, {"id": "seed-242", "cliente": "KARLA GABRIELA SANCHEZ MARTINEZ", "exp": "92983", "inv": "3402", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "COMITAN", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-29", "fRevision": "2026-08-29", "estatus": "EN PROCESO", "docsPend": true, "detalleDocsPend": "", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "", "segundoPago": ""}, {"id": "seed-243", "cliente": "ANA KAREN ROMAN VELASCO", "exp": "3409", "inv": "92985", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "SAN CRISTOBAL", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-29", "fRevision": "2026-08-29", "estatus": "EN PROCESO", "docsPend": true, "detalleDocsPend": "Formulario", "observaciones": "*PENDIENTE BAJA DE PLACAS", "primerPago": "", "segundoPago": ""}, {"id": "seed-244", "cliente": "OSBER BENJAMIN HERNANDEZ SILVANO", "exp": "3405", "inv": "92984", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "OCOSINGO", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-29", "fRevision": "2026-08-29", "estatus": "EN PROCESO", "docsPend": false, "detalleDocsPend": "", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-245", "cliente": "BEATRIZ DIAZ VAZQUEZ", "exp": "3408", "inv": "95", "tipoToma": "COMPRA DIRECTA", "sucursal": "TAPACHULA", "marca": "CHANGAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-31", "fRevision": "2026-08-31", "estatus": "EN PROCESO", "docsPend": false, "detalleDocsPend": "FORMULARIO", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-246", "cliente": "SAUL MENDEZ DIAZ", "exp": "3412", "inv": "92986", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "OCOSINGO", "marca": "NISSAN", "validador": "GIL", "financiamiento": "NO", "fSolicitud": "2026-08-31", "fRevision": "2026-08-31", "estatus": "EN PROCESO", "docsPend": false, "detalleDocsPend": "Formulario, etiqueta holografica no la anexan, no anexan el apartado de indentificacion del cliente", "observaciones": "", "primerPago": "", "segundoPago": ""}, {"id": "seed-247", "cliente": "LAZARO MIGUEL TORIJA PEREZ", "exp": "3413", "inv": "92987", "tipoToma": "SEMINUEVO X NUEVO", "sucursal": "PALENQUE", "marca": "NISSAN", "validador": "BRAYAN", "financiamiento": "NO", "fSolicitud": "2026-08-31", "fRevision": "2026-08-31", "estatus": "EN PROCESO", "docsPend": true, "detalleDocsPend": "*BAJA DE PLACAS *VALIDACION DE FACTURA", "observaciones": "*EL RFC MENCIONA QUE EL CLIENTE TIENE ACTIVIDAD EMPRESARIAL, ADEMAS LA CONSTANCIA NO MENCIONA LA ACTIVIDAD", "primerPago": "", "segundoPago": ""}];

/* ---------------------------------------------------------
   Helpers
--------------------------------------------------------- */
function uid() {
  return "t-" + Date.now().toString(36) + "-" + Math.random().toString(36).slice(2, 8);
}

function mesDe(fechaISO) {
  if (!fechaISO) return null;
  const d = new Date(fechaISO + "T00:00:00");
  if (isNaN(d.getTime())) return null;
  return MESES[d.getMonth()];
}

function anioMesDe(fechaISO) {
  if (!fechaISO) return null;
  return fechaISO.slice(0, 7);
}

function diasEnProceso(rec) {
  if (!rec.fSolicitud) return null;
  const start = new Date(rec.fSolicitud + "T00:00:00");
  if (isNaN(start.getTime())) return null;
  const endStr = rec.estatus === "COMPLETADO" && rec.fRevision ? rec.fRevision : null;
  const end = endStr ? new Date(endStr + "T00:00:00") : new Date();
  const diff = Math.round((end - start) / 86400000);
  return diff < 0 ? 0 : diff;
}

function tieneObs(rec) {
  return !!(rec.observaciones && rec.observaciones.trim().length > 0);
}

function fmtFecha(iso) {
  if (!iso) return "—";
  const d = new Date(iso + "T00:00:00");
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleDateString("es-MX", { day: "2-digit", month: "short", year: "numeric" });
}

function pct(n, total) {
  if (!total) return "0%";
  return Math.round((n / total) * 100) + "%";
}

function normalizeRecord(r) {
  return {
    errorTags: [],
    history: [],
    ...r,
  };
}

function emptyForm() {
  return {
    id: null,
    cliente: "",
    exp: "",
    inv: "",
    tipoToma: TIPOS_TOMA[0],
    sucursal: SUCURSALES[0],
    marca: MARCAS[0],
    validador: "",
    financiamiento: "NO",
    fSolicitud: new Date().toISOString().slice(0, 10),
    fRevision: "",
    estatus: "EN PROCESO",
    docsPend: false,
    detalleDocsPend: "",
    observaciones: "",
    primerPago: "",
    segundoPago: "",
    errorTags: [],
    history: [],
  };
}

/* ---------------------------------------------------------
   Small UI atoms
--------------------------------------------------------- */
function Field({ label, children, span }) {
  return (
    <label style={{ display: "flex", flexDirection: "column", gap: 6, gridColumn: span ? "1 / -1" : undefined }}>
      <span style={{ fontFamily: FONT_HEAD, fontSize: 12, color: C.inkMuted, letterSpacing: 0.2 }}>{label}</span>
      {children}
    </label>
  );
}

const inputStyle = {
  fontFamily: FONT_BODY,
  fontSize: 14,
  color: C.ink,
  background: C.surface,
  border: `1px solid ${C.line}`,
  borderRadius: 4,
  padding: "9px 10px",
  outline: "none",
};

function TextInput(props) {
  return <input {...props} style={{ ...inputStyle, ...(props.style || {}) }} />;
}
function Select({ options, ...props }) {
  return (
    <select {...props} style={{ ...inputStyle, ...(props.style || {}) }}>
      {options.map((o) => (
        <option key={o} value={o}>{o}</option>
      ))}
    </select>
  );
}
function TextArea(props) {
  return <textarea {...props} style={{ ...inputStyle, resize: "vertical", minHeight: 70, fontFamily: FONT_BODY, ...(props.style || {}) }} />;
}

function Pill({ children, tone = "muted" }) {
  const tones = {
    green: { bg: C.primarySoft, fg: C.primary },
    orange: { bg: C.accentSoft, fg: C.accent },
    gold: { bg: C.warnSoft, fg: C.warn },
    muted: { bg: "#E7E9E5", fg: C.inkMuted },
  };
  const t = tones[tone] || tones.muted;
  return (
    <span style={{
      display: "inline-flex", alignItems: "center", gap: 5,
      background: t.bg, color: t.fg, fontFamily: FONT_HEAD, fontSize: 11.5,
      padding: "3px 8px", borderRadius: 3, whiteSpace: "nowrap",
    }}>{children}</span>
  );
}

function Btn({ children, variant = "primary", ...props }) {
  const styles = {
    primary: { background: C.primary, color: "#fff", border: `1px solid ${C.primary}` },
    ghost: { background: "transparent", color: C.ink, border: `1px solid ${C.line}` },
    danger: { background: "transparent", color: C.danger, border: `1px solid ${C.dangerSoft}` },
  };
  return (
    <button
      {...props}
      style={{
        fontFamily: FONT_HEAD, fontSize: 13, cursor: "pointer",
        padding: "9px 14px", borderRadius: 4, ...styles[variant],
        ...(props.style || {}),
      }}
    >{children}</button>
  );
}

/* ---------------------------------------------------------
   Drawer: crear / editar registro
--------------------------------------------------------- */
function RecordDrawer({ open, initial, onClose, onSave, onDelete }) {
  const [form, setForm] = useState(initial || emptyForm());

  useEffect(() => {
    setForm(initial || emptyForm());
  }, [initial, open]);

  if (!open) return null;

  const set = (k) => (e) => {
    const v = e && e.target ? (e.target.type === "checkbox" ? e.target.checked : e.target.value) : e;
    setForm((f) => ({ ...f, [k]: v }));
  };

  const isEdit = !!form.id;
  const canSave = form.cliente.trim().length > 0 && form.fSolicitud;

  return (
    <div style={{
      position: "fixed", inset: 0, background: "rgba(27,35,31,0.35)",
      display: "flex", justifyContent: "flex-end", zIndex: 50,
    }} onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div style={{
        width: "min(560px, 100%)", height: "100%", background: C.bg,
        borderLeft: `1px solid ${C.lineStrong}`, overflowY: "auto",
        display: "flex", flexDirection: "column",
      }}>
        <div style={{
          padding: "18px 24px", borderBottom: `1px solid ${C.line}`,
          display: "flex", justifyContent: "space-between", alignItems: "center",
          background: C.surface, position: "sticky", top: 0, zIndex: 1,
        }}>
          <div>
            <div style={{ fontFamily: FONT_HEAD, fontSize: 18, color: C.ink }}>
              {isEdit ? "Editar toma" : "Nueva toma"}
            </div>
            {isEdit && (
              <div style={{ fontFamily: FONT_MONO, fontSize: 12, color: C.inkFaint }}>
                Exp. {form.exp || "s/n"} · {mesDe(form.fSolicitud) || "sin fecha"}
              </div>
            )}
          </div>
          <button onClick={onClose} style={{
            border: "none", background: "transparent", fontSize: 20, cursor: "pointer", color: C.inkMuted,
          }}>×</button>
        </div>

        <div style={{ padding: 24, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, flex: 1 }}>
          <Field label="Cliente" span>
            <TextInput value={form.cliente} onChange={set("cliente")} placeholder="Nombre completo" />
          </Field>

          <Field label="Expediente (EXP)">
            <TextInput value={form.exp} onChange={set("exp")} placeholder="2066" />
          </Field>
          <Field label="Inventario (INV)">
            <TextInput value={form.inv} onChange={set("inv")} placeholder="90222" />
          </Field>

          <Field label="Tipo de toma">
            <Select options={TIPOS_TOMA} value={form.tipoToma} onChange={set("tipoToma")} />
          </Field>
          <Field label="Marca">
            <Select options={MARCAS} value={form.marca} onChange={set("marca")} />
          </Field>

          <Field label="Sucursal / agencia">
            <Select options={SUCURSALES} value={form.sucursal} onChange={set("sucursal")} />
          </Field>
          <Field label="Validador">
            <Select options={["", ...VALIDADORES]} value={form.validador} onChange={set("validador")} />
          </Field>

          <Field label="Financiamiento">
            <Select options={FINANCIAMIENTO_OPTS} value={form.financiamiento} onChange={set("financiamiento")} />
          </Field>
          <Field label="Estatus">
            <Select options={ESTATUS_OPTS} value={form.estatus} onChange={set("estatus")} />
          </Field>

          <Field label="Fecha de solicitud">
            <TextInput type="date" value={form.fSolicitud} onChange={set("fSolicitud")} />
          </Field>
          <Field label="Fecha de revisión">
            <TextInput type="date" value={form.fRevision} onChange={set("fRevision")} />
          </Field>

          <Field label="1er pago">
            <TextInput value={form.primerPago} onChange={set("primerPago")} placeholder="Fecha o folio ODP" />
          </Field>
          <Field label="2do pago">
            <TextInput value={form.segundoPago} onChange={set("segundoPago")} placeholder="Fecha o folio ODP" />
          </Field>

          <label style={{ display: "flex", alignItems: "center", gap: 8, gridColumn: "1 / -1", marginTop: 4 }}>
            <input type="checkbox" checked={form.docsPend} onChange={set("docsPend")} />
            <span style={{ fontFamily: FONT_HEAD, fontSize: 13, color: C.ink }}>Tiene documentos pendientes</span>
          </label>

          {form.docsPend && (
            <Field label="Detalle de documentos pendientes" span>
              <TextArea value={form.detalleDocsPend} onChange={set("detalleDocsPend")} placeholder="*Pendiente endoso, baja de placas..." />
            </Field>
          )}

          <Field label="Tipo(s) de error encontrado" span>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
              {ERROR_CATALOG.map((tag) => {
                const active = (form.errorTags || []).includes(tag);
                return (
                  <button
                    key={tag}
                    type="button"
                    onClick={() => {
                      setForm((f) => ({
                        ...f,
                        errorTags: active
                          ? (f.errorTags || []).filter((t) => t !== tag)
                          : [...(f.errorTags || []), tag],
                      }));
                    }}
                    style={{
                      fontFamily: FONT_HEAD, fontSize: 11.5, padding: "5px 10px", borderRadius: 3, cursor: "pointer",
                      background: active ? C.accent : C.surface, color: active ? "#fff" : C.ink,
                      border: `1px solid ${active ? C.accent : C.line}`,
                    }}
                  >{tag}</button>
                );
              })}
            </div>
          </Field>

          <Field label="Observaciones / errores encontrados en la toma" span>
            <TextArea
              value={form.observaciones}
              onChange={set("observaciones")}
              placeholder="Describe los errores encontrados, campo por campo (ej. REVISIÓN MECÁNICA: VIN incorrecto...)"
              style={{ minHeight: 130 }}
            />
          </Field>

          {isEdit && form.history && form.history.length > 0 && (
            <div style={{ gridColumn: "1 / -1", borderTop: `1px solid ${C.line}`, paddingTop: 12, marginTop: 4 }}>
              <div style={{ fontFamily: FONT_HEAD, fontSize: 12, color: C.inkMuted, marginBottom: 8 }}>Historial de cambios</div>
              <div style={{ display: "flex", flexDirection: "column", gap: 6, maxHeight: 140, overflowY: "auto" }}>
                {[...form.history].reverse().map((h, i) => (
                  <div key={i} style={{ display: "flex", justifyContent: "space-between", gap: 10, fontFamily: FONT_BODY, fontSize: 12, color: C.inkMuted }}>
                    <span><strong style={{ color: C.ink }}>{h.by || "Sin nombre"}</strong> · {h.action}</span>
                    <span style={{ fontFamily: FONT_MONO, whiteSpace: "nowrap" }}>
                      {new Date(h.ts).toLocaleString("es-MX", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" })}
                    </span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>

        <div style={{
          padding: "16px 24px", borderTop: `1px solid ${C.line}`, background: C.surface,
          display: "flex", justifyContent: "space-between", position: "sticky", bottom: 0,
        }}>
          {isEdit ? (
            <Btn variant="danger" onClick={() => { if (onDelete) onDelete(form.id); }}>Eliminar</Btn>
          ) : <span />}
          <div style={{ display: "flex", gap: 8 }}>
            <Btn variant="ghost" onClick={onClose}>Cancelar</Btn>
            <Btn
              variant="primary"
              disabled={!canSave}
              style={{ opacity: canSave ? 1 : 0.5 }}
              onClick={() => canSave && onSave(form)}
            >
              {isEdit ? "Guardar cambios" : "Registrar toma"}
            </Btn>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Dashboard
--------------------------------------------------------- */
function Kpi({ label, value, tone }) {
  const fg = tone === "accent" ? C.accent : tone === "warn" ? C.warn : C.ink;
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 2, minWidth: 120 }}>
      <div style={{ fontFamily: FONT_MONO, fontSize: 34, color: fg, lineHeight: 1 }}>{value}</div>
      <div style={{ fontFamily: FONT_HEAD, fontSize: 12, color: C.inkMuted }}>{label}</div>
    </div>
  );
}

function Panel({ title, sub, children, style }) {
  return (
    <div style={{
      background: C.surface, border: `1px solid ${C.line}`, borderRadius: 6,
      padding: "16px 18px", display: "flex", flexDirection: "column", gap: 10, ...style,
    }}>
      <div>
        <div style={{ fontFamily: FONT_HEAD, fontSize: 14, color: C.ink }}>{title}</div>
        {sub && <div style={{ fontFamily: FONT_BODY, fontSize: 12, color: C.inkFaint }}>{sub}</div>}
      </div>
      {children}
    </div>
  );
}

function RankBar({ rows, maxKey = "total" }) {
  const max = Math.max(1, ...rows.map((r) => r[maxKey]));
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
      {rows.map((r) => (
        <div key={r.name} style={{ display: "grid", gridTemplateColumns: "100px 1fr 40px", alignItems: "center", gap: 8 }}>
          <span style={{ fontFamily: FONT_BODY, fontSize: 12.5, color: C.ink, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{r.name}</span>
          <div style={{ background: C.bg, borderRadius: 3, height: 10, overflow: "hidden" }}>
            <div style={{ width: `${(r[maxKey] / max) * 100}%`, height: "100%", background: r.color || C.primary, borderRadius: 3 }} />
          </div>
          <span style={{ fontFamily: FONT_MONO, fontSize: 12, color: C.inkMuted, textAlign: "right" }}>{r[maxKey]}</span>
        </div>
      ))}
    </div>
  );
}

function Dashboard({ records, onOpen, isMobile }) {
  const stats = useMemo(() => {
    const total = records.length;
    const completadas = records.filter((r) => r.estatus === "COMPLETADO").length;
    const docsPend = records.filter((r) => r.docsPend).length;
    const conObs = records.filter(tieneObs).length;

    const porMes = MESES.map((m) => {
      const rs = records.filter((r) => mesDe(r.fSolicitud) === m);
      return { mes: m.slice(0, 3), total: rs.length, completadas: rs.filter((r) => r.estatus === "COMPLETADO").length };
    });

    const groupBy = (keyFn) => {
      const map = new Map();
      records.forEach((r) => {
        const k = keyFn(r) || "SIN DATO";
        if (!map.has(k)) map.set(k, { name: k, total: 0, completadas: 0 });
        const g = map.get(k);
        g.total += 1;
        if (r.estatus === "COMPLETADO") g.completadas += 1;
      });
      return Array.from(map.values()).sort((a, b) => b.total - a.total);
    };

    const porMarca = groupBy((r) => r.marca);
    const porSucursal = groupBy((r) => r.sucursal);
    const porValidador = groupBy((r) => r.validador).filter((v) => v.name !== "SIN DATO");
    const porTipo = groupBy((r) => r.tipoToma);

    const diasPromedio = (() => {
      const completos = records.filter((r) => r.estatus === "COMPLETADO");
      if (!completos.length) return 0;
      const sum = completos.reduce((a, r) => a + (diasEnProceso(r) || 0), 0);
      return Math.round(sum / completos.length);
    })();

    const atoradas = records
      .filter((r) => r.estatus === "EN PROCESO")
      .map((r) => ({ ...r, dias: diasEnProceso(r) || 0 }))
      .filter((r) => r.dias > ALERT_DAYS)
      .sort((a, b) => b.dias - a.dias);

    const errorFrequency = (() => {
      const map = new Map();
      records.forEach((r) => (r.errorTags || []).forEach((tag) => map.set(tag, (map.get(tag) || 0) + 1)));
      return Array.from(map.entries())
        .map(([name, total]) => ({ name, total }))
        .sort((a, b) => b.total - a.total);
    })();

    const avgDiasPorSucursal = (() => {
      const map = new Map();
      records.filter((r) => r.estatus === "COMPLETADO").forEach((r) => {
        const k = r.sucursal || "SIN DATO";
        if (!map.has(k)) map.set(k, { name: k, sum: 0, n: 0 });
        const g = map.get(k);
        g.sum += diasEnProceso(r) || 0;
        g.n += 1;
      });
      return Array.from(map.values())
        .map((g) => ({ name: g.name, total: g.n ? Math.round(g.sum / g.n) : 0 }))
        .sort((a, b) => b.total - a.total);
    })();

    return {
      total, completadas, docsPend, conObs, porMes, porMarca, porSucursal, porValidador, porTipo, diasPromedio,
      atoradas, errorFrequency, avgDiasPorSucursal,
    };
  }, [records]);

  const colorFor = (i) => [C.primary, C.accent, C.warn, "#6B7A72", "#8C5A3C"][i % 5];

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
      {stats.atoradas.length > 0 && (
        <div style={{ background: C.warnSoft, border: `1px solid ${C.warn}`, borderRadius: 6, padding: "14px 18px", display: "flex", flexDirection: "column", gap: 8 }}>
          <div style={{ fontFamily: FONT_HEAD, fontSize: 13.5, color: C.warn }}>
            {stats.atoradas.length} toma{stats.atoradas.length === 1 ? "" : "s"} atorada{stats.atoradas.length === 1 ? "" : "s"} (más de {ALERT_DAYS} días en proceso)
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            {stats.atoradas.slice(0, 6).map((r) => (
              <div
                key={r.id}
                onClick={() => onOpen && onOpen(r)}
                style={{ display: "flex", justifyContent: "space-between", gap: 10, fontFamily: FONT_BODY, fontSize: 12.5, color: C.ink, cursor: onOpen ? "pointer" : "default" }}
              >
                <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.cliente} · {r.sucursal}</span>
                <span style={{ fontFamily: FONT_MONO, color: C.warn, whiteSpace: "nowrap" }}>{r.dias} días</span>
              </div>
            ))}
            {stats.atoradas.length > 6 && (
              <div style={{ fontFamily: FONT_BODY, fontSize: 11.5, color: C.inkFaint }}>y {stats.atoradas.length - 6} más…</div>
            )}
          </div>
        </div>
      )}

      <Panel title="Resumen general">
        <div style={{ display: "flex", gap: 40, flexWrap: "wrap" }}>
          <Kpi label="Total de tomas" value={stats.total} />
          <Kpi label="Completadas" value={stats.completadas} />
          <Kpi label="Docs. pendientes" value={stats.docsPend} tone="warn" />
          <Kpi label="Con observaciones" value={stats.conObs} tone="accent" />
          <Kpi label="Días promedio de proceso" value={stats.diasPromedio} />
        </div>
      </Panel>

      <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "1.4fr 1fr", gap: 18 }}>
        <Panel title="Tendencia mensual" sub="Solicitadas vs. completadas">
          <div style={{ width: "100%", height: 220 }}>
            <ResponsiveContainer>
              <LineChart data={stats.porMes} margin={{ left: -20, right: 10 }}>
                <CartesianGrid stroke={C.line} vertical={false} />
                <XAxis dataKey="mes" tick={{ fontFamily: FONT_MONO, fontSize: 11, fill: C.inkMuted }} axisLine={{ stroke: C.line }} tickLine={false} />
                <YAxis tick={{ fontFamily: FONT_MONO, fontSize: 11, fill: C.inkMuted }} axisLine={false} tickLine={false} />
                <Tooltip contentStyle={{ fontFamily: FONT_BODY, fontSize: 12, border: `1px solid ${C.line}`, borderRadius: 4 }} />
                <Line type="monotone" dataKey="total" name="Solicitadas" stroke={C.inkFaint} strokeWidth={2} dot={false} />
                <Line type="monotone" dataKey="completadas" name="Completadas" stroke={C.primary} strokeWidth={2.5} dot={false} />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </Panel>

        <Panel title="Por marca" sub={`${stats.total} tomas en total`}>
          <RankBar rows={stats.porMarca.map((r, i) => ({ ...r, color: colorFor(i) }))} />
        </Panel>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "1fr 1fr", gap: 18 }}>
        <Panel title="Por sucursal / agencia">
          <RankBar rows={stats.porSucursal.map((r, i) => ({ ...r, color: colorFor(i) }))} />
        </Panel>
        <Panel title="Por validador">
          <RankBar rows={stats.porValidador.map((r, i) => ({ ...r, color: colorFor(i) }))} />
        </Panel>
      </div>

      <Panel title="Por tipo de toma">
        <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "repeat(3, 1fr)", gap: 14 }}>
          {stats.porTipo.map((t, i) => (
            <div key={t.name} style={{ border: `1px solid ${C.line}`, borderRadius: 4, padding: 12 }}>
              <div style={{ fontFamily: FONT_HEAD, fontSize: 12.5, color: C.inkMuted }}>{t.name}</div>
              <div style={{ fontFamily: FONT_MONO, fontSize: 26, color: colorFor(i) }}>{t.total}</div>
              <div style={{ fontFamily: FONT_BODY, fontSize: 11.5, color: C.inkFaint }}>{pct(t.completadas, t.total)} completadas</div>
            </div>
          ))}
        </div>
      </Panel>

      <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "1fr 1fr", gap: 18 }}>
        <Panel title="Errores más frecuentes" sub="Según el tipo de error marcado en cada toma">
          {stats.errorFrequency.length ? (
            <RankBar rows={stats.errorFrequency.map((r, i) => ({ ...r, color: colorFor(i) }))} />
          ) : (
            <div style={{ fontFamily: FONT_BODY, fontSize: 12.5, color: C.inkFaint }}>
              Aún no hay tipos de error marcados. Selecciónalos al registrar o editar una toma.
            </div>
          )}
        </Panel>
        <Panel title="Días promedio de proceso por sucursal" sub="Solo tomas completadas">
          <RankBar rows={stats.avgDiasPorSucursal.map((r, i) => ({ ...r, color: colorFor(i) }))} />
        </Panel>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Tabla de registro
--------------------------------------------------------- */
function RegistroTable({ records, onOpen, isMobile }) {
  if (!records.length) {
    return (
      <div style={{
        border: `1px dashed ${C.lineStrong}`, borderRadius: 6, padding: "48px 24px",
        textAlign: "center", color: C.inkMuted, fontFamily: FONT_BODY, fontSize: 14,
      }}>
        No hay tomas que coincidan con los filtros. Ajusta los filtros o registra una toma nueva.
      </div>
    );
  }

  if (isMobile) {
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {records.map((r) => (
          <div
            key={r.id}
            onClick={() => onOpen(r)}
            style={{
              background: C.surface, border: `1px solid ${C.line}`, borderRadius: 6, padding: 14, cursor: "pointer",
              display: "flex", flexDirection: "column", gap: 6,
            }}
          >
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
              <div style={{ fontFamily: FONT_HEAD, fontSize: 14, color: C.ink }}>{r.cliente}</div>
              <span style={{ fontFamily: FONT_MONO, fontSize: 11, color: C.inkFaint, whiteSpace: "nowrap" }}>{fmtFecha(r.fSolicitud)}</span>
            </div>
            <div style={{ fontFamily: FONT_BODY, fontSize: 12.5, color: C.inkMuted }}>
              {r.tipoToma} · {r.marca} · {r.sucursal}{r.validador ? ` · ${r.validador}` : ""}
            </div>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
              <Pill tone={r.estatus === "COMPLETADO" ? "green" : "gold"}>{r.estatus}</Pill>
              {r.docsPend && <Pill tone="gold">Docs. pendientes</Pill>}
              {tieneObs(r) && <Pill tone="orange">Con obs.</Pill>}
            </div>
          </div>
        ))}
      </div>
    );
  }

  const th = { textAlign: "left", fontFamily: FONT_HEAD, fontSize: 11.5, color: C.inkMuted, padding: "10px 12px", borderBottom: `1px solid ${C.lineStrong}`, whiteSpace: "nowrap" };
  const td = { padding: "10px 12px", borderBottom: `1px solid ${C.line}`, fontFamily: FONT_BODY, fontSize: 13, color: C.ink, verticalAlign: "top" };

  return (
    <div style={{ overflowX: "auto", border: `1px solid ${C.line}`, borderRadius: 6, background: C.surface }}>
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead>
          <tr>
            <th style={th}>Fecha</th>
            <th style={th}>Exp.</th>
            <th style={th}>Cliente</th>
            <th style={th}>Tipo</th>
            <th style={th}>Marca</th>
            <th style={th}>Sucursal</th>
            <th style={th}>Validador</th>
            <th style={th}>Estatus</th>
            <th style={th}>Docs.</th>
            <th style={th}>Obs.</th>
          </tr>
        </thead>
        <tbody>
          {records.map((r) => (
            <tr
              key={r.id}
              onClick={() => onOpen(r)}
              style={{ cursor: "pointer" }}
              onMouseEnter={(e) => (e.currentTarget.style.background = C.bg)}
              onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
            >
              <td style={{ ...td, fontFamily: FONT_MONO, fontSize: 12 }}>{fmtFecha(r.fSolicitud)}</td>
              <td style={{ ...td, fontFamily: FONT_MONO, fontSize: 12 }}>{r.exp || "—"}</td>
              <td style={{ ...td, fontFamily: FONT_HEAD, fontSize: 13 }}>{r.cliente}</td>
              <td style={td}>{r.tipoToma}</td>
              <td style={td}>{r.marca}</td>
              <td style={td}>{r.sucursal}</td>
              <td style={td}>{r.validador || "—"}</td>
              <td style={td}>
                <Pill tone={r.estatus === "COMPLETADO" ? "green" : "gold"}>{r.estatus}</Pill>
              </td>
              <td style={td}>{r.docsPend ? <Pill tone="gold">Pendiente</Pill> : <Pill>Al día</Pill>}</td>
              <td style={td}>{tieneObs(r) ? <Pill tone="orange">Con obs.</Pill> : <Pill>Limpio</Pill>}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/* ---------------------------------------------------------
   App principal
--------------------------------------------------------- */
export default function TomasApp() {
  const [records, setRecords] = useState([]);
  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState("dashboard");
  const [filters, setFilters] = useState({ mes: "", sucursal: "", marca: "", validador: "", estatus: "", q: "" });
  const [drawer, setDrawer] = useState({ open: false, record: null });
  const [saveError, setSaveError] = useState(false);
  const [lastSynced, setLastSynced] = useState(null);
  const [syncing, setSyncing] = useState(false);
  const [conflict, setConflict] = useState(null); // { form, currentOnServer }
  const [currentUser, setCurrentUser] = useState("");
  const [askingName, setAskingName] = useState(false);
  const [userDraft, setUserDraft] = useState("");
  const [isMobile, setIsMobile] = useState(typeof window !== "undefined" ? window.innerWidth < 720 : false);

  // Responsivo: cambia a vista de celular por debajo de 720px.
  useEffect(() => {
    const onResize = () => setIsMobile(window.innerWidth < 720);
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);

  // Identidad personal (para el historial de cambios) — no se comparte entre los dos.
  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get(USER_KEY, false);
        if (res && res.value) setCurrentUser(res.value);
        else setAskingName(true);
      } catch (e) {
        setAskingName(true);
      }
    })();
  }, []);

  const saveUserName = useCallback(async () => {
    const name = userDraft.trim();
    if (!name) return;
    setCurrentUser(name);
    setAskingName(false);
    setUserDraft("");
    try { await window.storage.set(USER_KEY, name, false); } catch (e) { /* noop */ }
  }, [userDraft]);

  // Trae la versión más reciente del storage compartido y la aplica al estado local.
  const refreshFromStorage = useCallback(async () => {
    setSyncing(true);
    try {
      const res = await window.storage.get(STORAGE_KEY, true);
      const data = res && res.value ? JSON.parse(res.value).map(normalizeRecord) : [];
      setRecords(data);
      setLastSynced(Date.now());
    } catch (e) {
      /* si falla, se conserva lo que ya está en memoria */
    } finally {
      setSyncing(false);
    }
  }, []);

  // Carga inicial desde storage compartido; si no existe, siembra con el histórico del Excel.
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const res = await window.storage.get(STORAGE_KEY, true);
        const data = res && res.value ? JSON.parse(res.value).map(normalizeRecord) : [];
        if (!cancelled) { setRecords(data); setLastSynced(Date.now()); }
      } catch (e) {
        const seeded = SEED_DATA.map(normalizeRecord);
        if (!cancelled) { setRecords(seeded); setLastSynced(Date.now()); }
        try { await window.storage.set(STORAGE_KEY, JSON.stringify(seeded), true); } catch (e2) { /* noop */ }
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, []);

  // Refresco automático: cada 25s y al volver a la pestaña, siempre que no haya un formulario abierto
  // (para no pisar lo que la persona está escribiendo en ese momento).
  useEffect(() => {
    const interval = setInterval(() => { if (!drawer.open) refreshFromStorage(); }, 25000);
    const onFocus = () => { if (!drawer.open) refreshFromStorage(); };
    window.addEventListener("focus", onFocus);
    document.addEventListener("visibilitychange", onFocus);
    return () => {
      clearInterval(interval);
      window.removeEventListener("focus", onFocus);
      document.removeEventListener("visibilitychange", onFocus);
    };
  }, [drawer.open, refreshFromStorage]);

  // Aplica un cambio contra la copia más reciente del storage (no contra el estado local,
  // que puede estar desactualizado si la otra persona guardó algo mientras tanto).
  const applyChange = useCallback(async (mutator) => {
    try {
      const res = await window.storage.get(STORAGE_KEY, true);
      const fresh = res && res.value ? JSON.parse(res.value).map(normalizeRecord) : [];
      const next = mutator(fresh);
      await window.storage.set(STORAGE_KEY, JSON.stringify(next), true);
      setRecords(next);
      setLastSynced(Date.now());
      setSaveError(false);
      return { ok: true, fresh };
    } catch (e) {
      setSaveError(true);
      return { ok: false };
    }
  }, []);

  const handleSave = useCallback(async (form) => {
    const who = currentUser || "Sin nombre";
    if (form.id) {
      // ¿Alguien más cambió este mismo registro desde que se abrió el formulario?
      const res = await window.storage.get(STORAGE_KEY, true).catch(() => null);
      const fresh = res && res.value ? JSON.parse(res.value).map(normalizeRecord) : records;
      const onServer = fresh.find((r) => r.id === form.id);
      const original = drawer.record;
      if (onServer && original && JSON.stringify(onServer) !== JSON.stringify(original)) {
        setConflict({ form, currentOnServer: onServer });
        return;
      }
      const updated = { ...form, history: [...(form.history || []), { ts: Date.now(), by: who, action: "Editado" }] };
      await applyChange((data) => data.map((r) => (r.id === form.id ? updated : r)));
    } else {
      const created = { ...form, id: uid(), history: [{ ts: Date.now(), by: who, action: "Creado" }] };
      await applyChange((data) => [created, ...data]);
    }
    setDrawer({ open: false, record: null });
  }, [drawer.record, records, applyChange, currentUser]);

  const resolveConflict = useCallback(async (overwrite) => {
    if (overwrite && conflict) {
      const who = currentUser || "Sin nombre";
      const updated = {
        ...conflict.form,
        history: [...(conflict.form.history || []), { ts: Date.now(), by: who, action: "Editado (sobrescribió cambio de alguien más)" }],
      };
      await applyChange((data) => data.map((r) => (r.id === conflict.form.id ? updated : r)));
      setDrawer({ open: false, record: null });
    } else if (conflict) {
      // Descartar mi edición y mostrar la versión actual en el formulario para revisarla.
      setDrawer({ open: true, record: conflict.currentOnServer });
    }
    setConflict(null);
  }, [conflict, applyChange, currentUser]);

  const handleDelete = useCallback(async (id) => {
    await applyChange((data) => data.filter((r) => r.id !== id));
    setDrawer({ open: false, record: null });
  }, [applyChange]);

  const validadoresDisponibles = useMemo(() => {
    const fromData = new Set(records.map((r) => r.validador).filter(Boolean));
    VALIDADORES.forEach((v) => fromData.add(v));
    return Array.from(fromData).sort();
  }, [records]);

  const filtered = useMemo(() => {
    const q = filters.q.trim().toUpperCase();
    return records.filter((r) => {
      if (filters.mes && mesDe(r.fSolicitud) !== filters.mes) return false;
      if (filters.sucursal && r.sucursal !== filters.sucursal) return false;
      if (filters.marca && r.marca !== filters.marca) return false;
      if (filters.validador && r.validador !== filters.validador) return false;
      if (filters.estatus && r.estatus !== filters.estatus) return false;
      if (q && !(r.cliente || "").toUpperCase().includes(q) && !(r.exp || "").toUpperCase().includes(q)) return false;
      return true;
    }).sort((a, b) => (b.fSolicitud || "").localeCompare(a.fSolicitud || ""));
  }, [records, filters]);

  const setF = (k) => (e) => setFilters((f) => ({ ...f, [k]: e.target.value }));
  const clearFilters = () => setFilters({ mes: "", sucursal: "", marca: "", validador: "", estatus: "", q: "" });
  const hasFilters = Object.values(filters).some(Boolean);

  // Fuentes + hoja de estilo para impresión (Exportar PDF usa el diálogo de impresión del navegador).
  useEffect(() => {
    const styleId = "tomas-app-fonts";
    if (!document.getElementById(styleId)) {
      const link = document.createElement("link");
      link.id = styleId;
      link.rel = "stylesheet";
      link.href = "https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500&family=IBM+Plex+Mono:wght@400;500&display=swap";
      document.head.appendChild(link);
    }
    const printStyleId = "tomas-app-print";
    if (!document.getElementById(printStyleId)) {
      const style = document.createElement("style");
      style.id = printStyleId;
      style.textContent = `
        .tomas-print-only { display: none; }
        @media print {
          .tomas-no-print { display: none !important; }
          .tomas-print-only { display: block !important; }
        }
      `;
      document.head.appendChild(style);
    }
  }, []);

  const exportExcel = useCallback(() => {
    const rows = filtered.map((r) => ({
      "Fecha solicitud": r.fSolicitud || "",
      "Fecha revisión": r.fRevision || "",
      "Mes": mesDe(r.fSolicitud) || "",
      "Expediente": r.exp || "",
      "Inventario": r.inv || "",
      "Cliente": r.cliente || "",
      "Tipo de toma": r.tipoToma || "",
      "Sucursal": r.sucursal || "",
      "Marca": r.marca || "",
      "Validador": r.validador || "",
      "Financiamiento": r.financiamiento || "",
      "Estatus": r.estatus || "",
      "Docs. pendientes": r.docsPend ? "SI" : "NO",
      "Detalle docs. pendientes": r.detalleDocsPend || "",
      "Tipos de error": (r.errorTags || []).join("; "),
      "Observaciones": r.observaciones || "",
      "1er pago": r.primerPago || "",
      "2do pago": r.segundoPago || "",
      "Días en proceso": diasEnProceso(r),
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Tomas");
    XLSX.writeFile(wb, `tomas-unidad-${new Date().toISOString().slice(0, 10)}.xlsx`);
  }, [filtered]);

  const exportPDF = useCallback(() => {
    window.print();
  }, []);

  if (loading) {
    return (
      <div style={{ padding: 40, fontFamily: FONT_BODY, color: C.inkMuted }}>Cargando registros…</div>
    );
  }

  const selectFilterStyle = { ...inputStyle, padding: "7px 8px", fontSize: 12.5 };

  return (
    <div style={{ background: C.bg, minHeight: "100%", fontFamily: FONT_BODY, color: C.ink }}>
      <div className="tomas-no-print" style={{ maxWidth: 1180, margin: "0 auto", padding: "24px 20px 60px" }}>

        {/* Header */}
        <div style={{
          display: "flex", justifyContent: "space-between", alignItems: isMobile ? "flex-start" : "flex-end",
          marginBottom: 20, flexWrap: "wrap", gap: 12, flexDirection: isMobile ? "column" : "row",
        }}>
          <div>
            <div style={{ fontFamily: FONT_HEAD, fontSize: 22, fontWeight: 600, color: C.ink }}>Control de tomas de unidad</div>
            <div style={{ fontFamily: FONT_BODY, fontSize: 13, color: C.inkMuted }}>Validación diaria de expedientes · Chiapas 2026</div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 6, fontFamily: FONT_MONO, fontSize: 11.5, color: C.inkFaint }}>
              <span style={{ width: 6, height: 6, borderRadius: "50%", background: syncing ? C.warn : C.primary, display: "inline-block" }} />
              {syncing ? "Sincronizando…" : lastSynced ? `Actualizado ${new Date(lastSynced).toLocaleTimeString("es-MX", { hour: "2-digit", minute: "2-digit" })}` : ""}
              <button
                onClick={refreshFromStorage}
                title="Actualizar ahora"
                style={{ border: "none", background: "transparent", cursor: "pointer", color: C.inkMuted, fontFamily: FONT_HEAD, fontSize: 11.5, textDecoration: "underline", padding: 0 }}
              >actualizar</button>
            </div>
            <div style={{ fontFamily: FONT_BODY, fontSize: 11.5, color: C.inkFaint }}>
              {currentUser ? `Tú: ${currentUser}` : "Sin identificar"}{" "}
              <button
                onClick={() => { setUserDraft(currentUser); setAskingName(true); }}
                style={{ border: "none", background: "transparent", cursor: "pointer", color: C.inkMuted, fontFamily: FONT_HEAD, fontSize: 11.5, textDecoration: "underline", padding: 0 }}
              >cambiar</button>
            </div>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button
              onClick={() => setTab("dashboard")}
              style={{
                fontFamily: FONT_HEAD, fontSize: 13, padding: "8px 16px", borderRadius: 4, cursor: "pointer",
                background: tab === "dashboard" ? C.primary : "transparent",
                color: tab === "dashboard" ? "#fff" : C.ink,
                border: `1px solid ${tab === "dashboard" ? C.primary : C.line}`,
              }}
            >Dashboard</button>
            <button
              onClick={() => setTab("registro")}
              style={{
                fontFamily: FONT_HEAD, fontSize: 13, padding: "8px 16px", borderRadius: 4, cursor: "pointer",
                background: tab === "registro" ? C.primary : "transparent",
                color: tab === "registro" ? "#fff" : C.ink,
                border: `1px solid ${tab === "registro" ? C.primary : C.line}`,
              }}
            >Registro ({records.length})</button>
          </div>
        </div>

        {askingName && (
          <div style={{
            marginBottom: 14, padding: "10px 14px", background: C.primarySoft, borderRadius: 4,
            display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap",
          }}>
            <span style={{ fontFamily: FONT_BODY, fontSize: 13, color: C.ink }}>¿Cómo te llamas? Se usa para el historial de cambios de cada toma.</span>
            <TextInput
              value={userDraft}
              onChange={(e) => setUserDraft(e.target.value)}
              placeholder="Tu nombre"
              style={{ maxWidth: 180 }}
              onKeyDown={(e) => { if (e.key === "Enter") saveUserName(); }}
            />
            <Btn onClick={saveUserName}>Guardar</Btn>
          </div>
        )}

        {saveError && (
          <div style={{ marginBottom: 14, padding: "10px 14px", background: C.dangerSoft, color: C.danger, borderRadius: 4, fontFamily: FONT_BODY, fontSize: 13 }}>
            No se pudo guardar el último cambio. Revisa tu conexión e inténtalo de nuevo.
          </div>
        )}

        {/* Filtros */}
        <div style={{
          display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center",
          background: C.surface, border: `1px solid ${C.line}`, borderRadius: 6, padding: 12, marginBottom: 20,
        }}>
          <TextInput
            placeholder="Buscar cliente o expediente…"
            value={filters.q}
            onChange={setF("q")}
            style={{ ...selectFilterStyle, minWidth: 200, flex: 1 }}
          />
          <select style={selectFilterStyle} value={filters.mes} onChange={setF("mes")}>
            <option value="">Todos los meses</option>
            {MESES.map((m) => <option key={m} value={m}>{m}</option>)}
          </select>
          <select style={selectFilterStyle} value={filters.sucursal} onChange={setF("sucursal")}>
            <option value="">Todas las sucursales</option>
            {SUCURSALES.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          <select style={selectFilterStyle} value={filters.marca} onChange={setF("marca")}>
            <option value="">Todas las marcas</option>
            {MARCAS.map((m) => <option key={m} value={m}>{m}</option>)}
          </select>
          <select style={selectFilterStyle} value={filters.validador} onChange={setF("validador")}>
            <option value="">Todos los validadores</option>
            {validadoresDisponibles.map((v) => <option key={v} value={v}>{v}</option>)}
          </select>
          <select style={selectFilterStyle} value={filters.estatus} onChange={setF("estatus")}>
            <option value="">Todos los estatus</option>
            {ESTATUS_OPTS.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          {hasFilters && (
            <button onClick={clearFilters} style={{ fontFamily: FONT_HEAD, fontSize: 12, color: C.inkMuted, background: "transparent", border: "none", cursor: "pointer", textDecoration: "underline" }}>
              Limpiar filtros
            </button>
          )}
          <div style={{ marginLeft: "auto", display: "flex", gap: 8, flexWrap: "wrap" }}>
            <Btn variant="ghost" onClick={exportExcel}>Exportar Excel</Btn>
            <Btn variant="ghost" onClick={exportPDF}>Exportar PDF</Btn>
            <Btn onClick={() => setDrawer({ open: true, record: null })}>+ Nueva toma</Btn>
          </div>
        </div>

        <div style={{ fontFamily: FONT_MONO, fontSize: 12, color: C.inkFaint, marginBottom: 12 }}>
          {filtered.length} de {records.length} registros
        </div>

        {tab === "dashboard" ? (
          <Dashboard records={filtered} onOpen={(r) => setDrawer({ open: true, record: r })} isMobile={isMobile} />
        ) : (
          <RegistroTable records={filtered} onOpen={(r) => setDrawer({ open: true, record: r })} isMobile={isMobile} />
        )}
      </div>

      {/* Vista imprimible: solo aparece al usar Exportar PDF (diálogo de impresión) */}
      <div className="tomas-print-only" style={{ padding: 24 }}>
        <div style={{ fontFamily: FONT_HEAD, fontSize: 18, marginBottom: 2 }}>Control de tomas de unidad</div>
        <div style={{ fontFamily: FONT_BODY, fontSize: 11, color: "#444", marginBottom: 14 }}>
          Generado {new Date().toLocaleString("es-MX")} · {filtered.length} registros{hasFilters ? " (con filtros aplicados)" : ""}
        </div>
        <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: FONT_BODY, fontSize: 10 }}>
          <thead>
            <tr>
              {["Fecha", "Exp.", "Cliente", "Tipo", "Sucursal", "Marca", "Validador", "Estatus", "Docs.", "Observaciones"].map((h) => (
                <th key={h} style={{ textAlign: "left", borderBottom: "1px solid #999", padding: "4px 6px" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map((r) => (
              <tr key={r.id}>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #ddd" }}>{fmtFecha(r.fSolicitud)}</td>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #ddd" }}>{r.exp || "—"}</td>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #ddd" }}>{r.cliente}</td>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #ddd" }}>{r.tipoToma}</td>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #ddd" }}>{r.sucursal}</td>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #ddd" }}>{r.marca}</td>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #ddd" }}>{r.validador || "—"}</td>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #ddd" }}>{r.estatus}</td>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #ddd" }}>{r.docsPend ? "Pendiente" : "Al día"}</td>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #ddd" }}>{r.observaciones}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="tomas-no-print">
        <RecordDrawer
          open={drawer.open}
          initial={drawer.record}
          onClose={() => setDrawer({ open: false, record: null })}
          onSave={handleSave}
          onDelete={handleDelete}
        />

        {conflict && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(27,35,31,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 60, padding: 20 }}>
            <div style={{ background: C.surface, borderRadius: 6, border: `1px solid ${C.lineStrong}`, padding: 22, maxWidth: 420, display: "flex", flexDirection: "column", gap: 12 }}>
              <div style={{ fontFamily: FONT_HEAD, fontSize: 16, color: C.ink }}>Este registro cambió mientras lo tenías abierto</div>
              <div style={{ fontFamily: FONT_BODY, fontSize: 13.5, color: C.inkMuted, lineHeight: 1.5 }}>
                {conflict.currentOnServer.cliente} — alguien más guardó una actualización de esta toma
                después de que abriste el formulario. ¿Qué quieres hacer?
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 8, marginTop: 4 }}>
                <Btn variant="ghost" onClick={() => resolveConflict(false)}>Ver la versión más reciente (descartar mi edición)</Btn>
                <Btn variant="danger" onClick={() => resolveConflict(true)}>Sobrescribir con mi cambio de todos modos</Btn>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
