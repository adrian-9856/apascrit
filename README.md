# Sistema de Seguimiento de Prácticas v2.8

## 🆕 ¿Qué hay de nuevo en v2.8?

### Nuevas Columnas de Etapas (Registro Detallado)

Se agregaron **7 columnas nuevas** entre "Formación" y "Etapa Actual" para registrar información detallada de cada llamada según la etapa del participante:

| # | Columna | Descripción | Color |
|---|---------|-------------|-------|
| 5 | **Aliados** | Registro de conexiones y contactos con aliados | 🔵 Azul claro |
| 6 | **Plataformas** | Información sobre uso de plataformas de empleo | 🟣 Morado claro |
| 7 | **Conexión laboral** | Detalles de conexiones laborales establecidas | 🟢 Verde claro |
| 8 | **Por su cuenta** | Información de búsqueda independiente de empleo | 🟠 Naranja claro |
| 9 | **No busca trabajar** | Registro si el participante no busca empleo | ⚪ Gris claro |
| 10 | **Empleado** | Información de empleos conseguidos | 🟢 Verde oscuro claro |
| 11 | **No terminó la formación** | Registro de participantes que no completaron | 🔴 Rojo claro |

### 🎨 Colores Distintivos

Cada columna de etapa tiene un color único para facilitar su identificación visual:

- **🔵 Aliados** - Azul claro (#B3D9FF)
- **🟣 Plataformas** - Morado claro (#D9C3FF)
- **🟢 Conexión laboral** - Verde claro (#B3FFB3)
- **🟠 Por su cuenta** - Naranja claro (#FFD9B3)
- **⚪ No busca trabajar** - Gris claro (#E0E0E0)
- **🟢 Empleado** - Verde oscuro claro (#C5E8C5)
- **🔴 No terminó la formación** - Rojo claro (#FFCCCC)

Los colores se aplican automáticamente:
- En la **hoja principal** (📋 Seguimiento General)
- En todas las **hojas de llamadas** (📞 Llamada 1-5)
- En la hoja de **✅ Finalizados**

### Nueva Estructura de Columnas

**ANTES (v2.7):**
```
1. Creamos ID
2. Nombre Completo
3. Teléfono
4. Formación
5. Etapa Actual          ← Era columna 5
6. Resultados Obtenidos  ← Era columna 6
7. Documentos Faltantes  ← Era columna 7
8. Fecha Original        ← Era columna 8
9. Notas/Última Llamada  ← Era columna 9
10. Total Llamadas       ← Era columna 10
11. Procesar             ← Era columna 11
```

**AHORA (v2.8):**
```
1. Creamos ID
2. Nombre Completo
3. Teléfono
4. Formación
--- 7 COLUMNAS NUEVAS ---
5. Aliados                        ← NUEVA
6. Plataformas                    ← NUEVA
7. Conexión laboral               ← NUEVA
8. Por su cuenta                  ← NUEVA
9. No busca trabajar              ← NUEVA
10. Empleado                      ← NUEVA
11. No terminó la formación       ← NUEVA
--- COLUMNAS ORIGINALES ---
12. Etapa Actual          ← Ahora columna 12 (era 5)
13. Resultados Obtenidos  ← Ahora columna 13 (era 6)
14. Documentos Faltantes  ← Ahora columna 14 (era 7)
15. Fecha Original        ← Ahora columna 15 (era 8)
16. Notas/Última Llamada  ← Ahora columna 16 (era 9)
17. Total Llamadas        ← Ahora columna 17 (era 10)
18. Procesar              ← Ahora columna 18 (era 11)
```

---

## 🎯 Funcionalidades Principales

### 1. Escritura Automática en Columnas de Etapas

Cuando se procesa una llamada:
- El sistema detecta la "Etapa Actual" del participante
- Escribe automáticamente en la columna correspondiente
- Agrega fecha, hora y número de llamada

**Ejemplo:**
```
Participante: Juan Pérez
Etapa Actual: "Aliados"
Al procesar Llamada 1 → Escribe en columna "Aliados":
  "19/11/2025 14:30 - Llamada 1"

Al procesar Llamada 2 → Agrega en columna "Aliados":
  "19/11/2025 14:30 - Llamada 1 | 20/11/2025 10:15 - Llamada 2"
```

### 2. Función de Actualización SIN Pérdida de Datos

**`🔄 ACTUALIZAR Estructura de Tabla`**

Esta función permite actualizar hojas existentes agregando las 7 columnas nuevas **sin borrar ningún dato**.

#### ¿Cuándo usar?
- Cuando tienes datos en la versión anterior (v2.7 o anterior)
- Quieres agregar las nuevas columnas de etapas
- No quieres perder información existente

#### ¿Cómo funciona?
1. Detecta si las columnas nuevas ya existen
2. Inserta 7 columnas después de "Formación" (columna 4)
3. Ajusta todos los encabezados y formatos
4. Reconfigura validaciones y checkboxes
5. Mantiene TODOS los datos existentes

---

## 📋 Guía de Uso

### Para Instalaciones NUEVAS

1. **Configurar sistema**
   - Menú: `🔧 Configuración` → `⚙️ Configuración Inicial`
   - Esto crea todas las hojas con la estructura v2.8

2. **Configurar Sheet externo** (opcional)
   - Menú: `🔧 Configuración` → `🔗 Configurar Google Sheet Externo`
   - Pega la URL completa del Google Sheet

3. **Importar datos**
   - Menú: `📥 Importación` → `📥 Importar Datos 2024/2025`
   - O usa: `📥 Importar Desde Otra Hoja`

4. **Activar procesamiento automático**
   - Menú: `⚡ ACTIVAR Procesamiento Automático`

5. **Procesar llamadas**
   - Marca el checkbox en la columna "Procesar"
   - Espera 1-2 segundos entre cada checkbox
   - El sistema registrará automáticamente en las columnas de etapas

### Para Actualizar desde v2.7 (o anterior)

1. **IMPORTANTE: Crear respaldo primero**
   - Menú: `🛠️ Mantenimiento` → `💾 Crear Respaldo`
   - Guarda la URL del respaldo

2. **Actualizar estructura**
   - Menú: `🔧 Configuración` → `🔄 ACTUALIZAR Estructura de Tabla`
   - Confirma la actualización
   - El sistema insertará las 7 columnas nuevas
   - ✅ NO se perderán datos

3. **Verificar actualización**
   - Revisa que las columnas 5-11 sean las nuevas etapas
   - Verifica que los datos existentes estén intactos
   - Las columnas 12-18 deben tener los datos originales

4. **Reactivar procesamiento**
   - Si tenías procesamiento automático activo:
     - Desactiva: `🔴 DESACTIVAR Procesamiento Automático`
     - Espera 5 segundos
     - Activa: `⚡ ACTIVAR Procesamiento Automático`

---

## 🔧 Cambios Técnicos

### Mapeo de Columnas (Constante `COLUMNAS`)

```javascript
const COLUMNAS = {
  ID: 1,
  NOMBRE: 2,
  TELEFONO: 3,
  FORMACION: 4,
  // NUEVAS COLUMNAS DE ETAPAS (5-11)
  ALIADOS: 5,
  PLATAFORMAS: 6,
  CONEXION_LABORAL: 7,
  POR_SU_CUENTA: 8,
  NO_BUSCA_TRABAJAR: 9,
  EMPLEADO: 10,
  NO_TERMINO_FORMACION: 11,
  // COLUMNAS ORIGINALES (ahora 12-18)
  ETAPA_ACTUAL: 12,
  RESULTADOS: 13,
  DOCUMENTOS: 14,
  FECHA: 15,
  NOTAS: 16,
  TOTAL_LLAMADAS: 17,
  PROCESAR: 18
};
```

### Función Clave: `determinarColumnaEtapa()`

Mapea la etapa actual a la columna correspondiente:

```javascript
function determinarColumnaEtapa(etapa, numeroLlamada) {
  const mapeoEtapas = {
    'aliados': { indice: 0, nombre: 'Aliados' },
    'plataformas': { indice: 1, nombre: 'Plataformas' },
    'conexion laboral': { indice: 2, nombre: 'Conexión Laboral' },
    'por su cuenta': { indice: 3, nombre: 'Por su cuenta' },
    'no busca trabajar': { indice: 4, nombre: 'No busca trabajar' },
    'empleado': { indice: 5, nombre: 'Empleado' },
    'no termino': { indice: 6, nombre: 'No terminó la formación' }
  };
  // Retorna el índice de columna según la etapa
}
```

### Procesamiento de Llamadas Actualizado

```javascript
function procesarLlamadaOptimizada(hojaGeneral, fila) {
  // ... obtener datos ...

  // Determinar columna según etapa
  const infoEtapa = determinarColumnaEtapa(etapaActual, nuevasLlamadas);

  // Escribir en la columna correspondiente
  if (infoEtapa.indice !== -1) {
    const fechaFormateada = Utilities.formatDate(fechaActual, ...);
    const nuevoTexto = `${fechaFormateada} - Llamada ${nuevasLlamadas}`;
    columnasEtapas[infoEtapa.indice] = nuevoTexto;
  }

  // Actualizar en hoja destino y hoja actual
  // ...
}
```

---

## ⚠️ Solución de Problemas

### "No veo las columnas nuevas"
**Solución:**
1. Ejecuta: `🔄 ACTUALIZAR Estructura de Tabla`
2. Si no aparece el menú, recarga la página
3. Espera unos segundos y vuelve a abrir el menú

### "Perdí mis datos al actualizar"
**Solución:**
1. Si creaste respaldo: Restaura desde el respaldo
2. Si no: Usa `Extensiones` → `Apps Script` → `Ejecuciones` → Ver errores
3. Contacta soporte con el error específico

### "Las columnas de etapas no se llenan"
**Solución:**
1. Verifica que "Etapa Actual" tenga un valor válido
2. Las etapas válidas son:
   - Aliados, Plataformas, Conexión Laboral, Por su cuenta,
   - No busca trabajar, Empleado, No termino la formación
3. Procesa una llamada de prueba: `🧪 Procesar Una Llamada (Prueba)`

### "Doble procesamiento de llamadas"
**Solución:**
1. Ejecuta: `🔍 Diagnosticar Doble Procesamiento`
2. Si hay bloqueos: `🧹 Limpiar Bloqueos de Procesamiento`
3. Si persiste:
   - Desactiva procesamiento automático
   - Espera 5 segundos
   - Reactiva procesamiento automático

---

## 📊 Diferencias con v2.7

| Característica | v2.7 | v2.8 |
|----------------|------|------|
| Columnas totales | 11 | 18 |
| Columnas de etapas | ❌ No | ✅ 7 columnas |
| Registro detallado | ❌ Solo en Notas | ✅ Columnas específicas |
| Actualización sin pérdida | ❌ No disponible | ✅ Función dedicada |
| Mapeo de columnas | Hardcoded | Constante `COLUMNAS` |
| Formaciones | 30+ | 33+ (incluye SAC, Ofimática, Análisis de datos) |

---

## 🎓 Formaciones Soportadas

- **Barista**: I, II, III, IV
- **Barismo**: 1-10
- **Gastronomía**: I-V (o 1-5)
- **Food Manager**
- **Panadería**
- **Repostería**
- **Sommelier**
- **🆕 Análisis de datos E-commerce**
- **🆕 SAC (Servicio al Cliente)**
- **🆕 Ofimática**
- **Otra**

---

## 💡 Mejores Prácticas

1. **Respaldos regulares**
   - Crea respaldo antes de actualizaciones
   - Guarda URL del respaldo en lugar seguro

2. **Procesamiento automático**
   - Espera 1-2 segundos entre checkboxes
   - Evita marcar múltiples checkboxes simultáneamente

3. **Etapas claras**
   - Mantén "Etapa Actual" actualizada
   - Usa valores exactos del desplegable

4. **Importación de datos**
   - Verifica estructura antes de importar
   - Usa nombres de hoja exactos

5. **Mantenimiento**
   - Ejecuta diagnóstico mensualmente
   - Limpia datos vacíos regularmente

---

## 🔒 Seguridad y Privacidad

- Todos los datos permanecen en tu Google Sheet
- No se envía información a servidores externos
- El código se ejecuta en tu cuenta de Google
- Solo tú tienes acceso a los datos

---

## 📞 Soporte

### Diagnóstico Automático
Usa: `🔍 Diagnosticar Sistema` para análisis completo

### Reportes
Genera: `📈 Generar Reporte Completo` para estadísticas

### Configuración Actual
Revisa: `⚙️ Ver Configuración Actual` para estado del sistema

---

## 📝 Changelog v2.8

### Añadido
- ✅ 7 columnas nuevas de etapas (5-11)
- ✅ **Colores distintivos** para cada columna de etapa
- ✅ Constante `COLORES_ETAPAS` con paleta de colores
- ✅ Función `actualizarEstructuraTabla()` sin pérdida de datos
- ✅ Función `aplicarColoresColumnasEtapas()` para formato visual
- ✅ Escritura automática en columnas según etapa
- ✅ Constante `COLUMNAS` para mapeo centralizado
- ✅ Función `determinarColumnaEtapa()` para lógica de escritura
- ✅ Formaciones: Análisis de datos E-commerce, SAC, Ofimática

### Modificado
- 🔄 Todas las referencias de columnas actualizadas
- 🔄 `procesarLlamadaOptimizada()` con lógica de etapas
- 🔄 `configurarHojasCorregido()` con estructura v2.8
- 🔄 Importación de datos con 18 columnas
- 🔄 Formato y anchos de columnas
- 🔄 Hojas de llamadas con 17 columnas

### Mantenido
- ✅ Sistema anti-doble procesamiento
- ✅ Procesamiento automático/manual
- ✅ Importación desde sheets externos
- ✅ Gestión de documentos faltantes
- ✅ Finalización automática por etapa
- ✅ Reportes y diagnósticos

---

## 🚀 Próximos Pasos Recomendados

1. **Si eres usuario nuevo:**
   - Ejecuta configuración inicial
   - Importa tus datos
   - Activa procesamiento automático
   - Procesa una llamada de prueba

2. **Si actualizas desde v2.7:**
   - Crea respaldo
   - Ejecuta `ACTUALIZAR Estructura de Tabla`
   - Verifica que todo funcione
   - Reactiva procesamiento automático

3. **Para todos:**
   - Lee las instrucciones: `📖 Ver Instrucciones`
   - Revisa configuración: `⚙️ Ver Configuración Actual`
   - Genera reporte: `📈 Generar Reporte Completo`

---

**Versión:** 2.8
**Fecha:** Noviembre 2025
**Autor:** Sistema de Seguimiento de Prácticas
**Licencia:** Uso interno

---

¡Sistema listo para uso! 🎉
