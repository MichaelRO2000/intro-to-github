# Análisis de Expansión Regional - Corporación XYZ

## Estudio de Decisión para Establecimiento de Tienda en Costa Caribe

### 📋 Descripción del Proyecto

Este repositorio contiene un análisis integral para la expansión de una tienda regional de la **Corporación XYZ** en la costa Caribe de Colombia. El estudio aplica métodos de decisión bajo incertidumbre y riesgo para seleccionar la mejor ubicación entre tres ciudades candidatas.

### 🎯 Objetivo

Determinar la ciudad óptima para establecer la primera tienda regional de Corporación XYZ entre las opciones:
- **Santa Marta**
- **Barranquilla** 
- **Cartagena**

### 💰 Parámetros del Proyecto

- **Inversión Inicial**: USD $800,000
- **Importancia Estratégica**: Establecimiento de presencia de marca y estrategia logística futura
- **Horizonte de Análisis**: Proyecciones mensuales

### 📊 Datos Analizados

#### Costos Operativos Mensuales (COP)

| Ciudad | Arriendo | Servicios Públicos | Mantenimiento | Total |
|--------|----------|-------------------|---------------|--------|
| Santa Marta | $2,000,000 | $3,500,000 | $4,000,000 | $9,500,000 |
| Barranquilla | $2,200,000 | $3,000,000 | $3,800,000 | $9,000,000 |
| Cartagena | $2,800,000 | $4,000,000 | $3,500,000 | $10,300,000 |

#### Proyecciones de Ingresos Mensuales (COP)

| Ciudad | Demanda Alta | Demanda Media | Demanda Baja |
|--------|--------------|---------------|--------------|
| Santa Marta | $15,000,000 | $12,000,000 | $5,000,000 |
| Barranquilla | $14,000,000 | $10,000,000 | $5,000,000 |
| Cartagena | $16,000,000 | $8,000,000 | $6,000,000 |

#### Escenarios de Probabilidad

**Probabilidades Iniciales:**
- Demanda Alta: 20%
- Demanda Media: 30%
- Demanda Baja: 50%

**Probabilidades con Estudio Técnico (Costo: $1,000,000):**
- Demanda Alta: 30%
- Demanda Media: 40%
- Demanda Baja: 30%

### 🔍 Metodología de Análisis

#### 1. Criterios de Decisión bajo Incertidumbre

- **Criterio de Laplace** (Equiprobabilidad)
- **Criterio Maximax** (Optimista)
- **Criterio Maximin** (Pesimista)
- **Criterio de Hurwicz** (α = 0.6)
- **Criterio de Savage** (Minimax Arrepentimiento)
- **Criterio del Valor Esperado**

#### 2. Análisis mediante Árbol de Decisiones

- Evaluación de alternativas con probabilidades iniciales
- Cálculo de valores esperados
- Análisis de sensibilidad

#### 3. Evaluación del Estudio Técnico

- Comparación con y sin información adicional
- Análisis costo-beneficio del estudio
- Valor de la información perfecta

### 📈 Resultados Principales

#### Matriz de Pagos (Utilidad Neta Mensual)

| Ciudad | Demanda Alta | Demanda Media | Demanda Baja |
|--------|--------------|---------------|--------------|
| Santa Marta | $5,500,000 | $2,500,000 | -$4,500,000 |
| Barranquilla | $5,000,000 | $1,000,000 | -$4,000,000 |
| Cartagena | $5,700,000 | -$2,300,000 | -$4,300,000 |

#### Resumen de Criterios de Decisión

| Criterio | Mejor Alternativa |
|----------|------------------|
| Laplace | Santa Marta |
| Maximax | Cartagena |
| Maximin | Barranquilla |
| Hurwicz | Santa Marta |
| Savage | Santa Marta |
| Valor Esperado | Santa Marta |

### 🏆 Recomendación Final

**SE RECOMIENDA SELECCIONAR SANTA MARTA** como ubicación para la primera tienda regional.

#### Justificación:
- ✅ **Mayor consistencia**: Santa Marta es seleccionada por 4 de 6 criterios de decisión
- ✅ **Mejor valor esperado**: $300,000 mensuales con probabilidades iniciales
- ✅ **Balance óptimo**: Combina buen potencial de ingresos con costos controlados
- ✅ **Menor riesgo relativo**: Mejor desempeño en escenarios adversos

### 📁 Archivos del Proyecto

#### `Analisis_Expansion_XYZ_Costa_Caribe.xlsx`
Archivo Excel completo con 5 hojas de análisis:

1. **Datos Base**: Información general y datos del problema
2. **Criterios Decisión**: Aplicación de 6 métodos de decisión
3. **Árbol Decisiones**: Análisis mediante árbol de decisiones
4. **Con Estudio Técnico**: Evaluación con nuevas probabilidades
5. **Recomendaciones**: Conclusiones y recomendaciones finales

#### `analisis_expansion_xyz.py`
Script Python que genera automáticamente el archivo Excel con todos los cálculos y análisis.

### 🔧 Cómo Ejecutar el Análisis

1. **Prerrequisitos**:
   ```bash
   pip install openpyxl
   ```

2. **Generar el análisis**:
   ```bash
   python3 analisis_expansion_xyz.py
   ```

3. **Abrir el archivo Excel generado**: `Analisis_Expansion_XYZ_Costa_Caribe.xlsx`

### ✅ Fortalezas del Método de Árboles de Decisión

- **Estructura Clara**: Visualización sistemática del proceso de decisión
- **Análisis Probabilístico**: Incorpora incertidumbre mediante probabilidades
- **Flexibilidad**: Evalúa múltiples alternativas simultáneamente
- **Trazabilidad**: Cada rama muestra consecuencias claras
- **Valor de la Información**: Cuantifica el beneficio de información adicional
- **Comunicación Efectiva**: Facilita explicación a stakeholders

### ⚠️ Limitaciones del Método

- **Dependencia de Probabilidades**: Resultados sensibles a probabilidades asignadas
- **Simplicidad**: Puede no capturar toda la complejidad empresarial
- **Naturaleza Estática**: No considera cambios dinámicos en el tiempo
- **Subjetividad**: Probabilidades pueden ser imprecisas
- **Complejidad Creciente**: Difícil manejo con muchas variables
- **Supuestos de Neutralidad**: Asume neutralidad al riesgo

### 📋 Consideraciones para la Implementación

- **Análisis de Sensibilidad**: Variar probabilidades ±10%
- **Factores Cualitativos**: Infraestructura, competencia, regulaciones
- **Revisión Periódica**: Actualizar proyecciones regularmente
- **Plan de Contingencia**: Preparar alternativas para baja demanda
- **Monitoreo Continuo**: Implementar KPIs post-implementación

### 🤝 Contribuciones

Este análisis fue desarrollado aplicando principios de teoría de decisiones y puede ser extendido con:
- Análisis de Monte Carlo
- Modelado dinámico
- Factores de riesgo adicionales
- Análisis de escenarios extremos

---

**Corporación XYZ - Análisis de Expansión Regional 2024**

*Desarrollado con metodologías de investigación operativa y teoría de decisiones*