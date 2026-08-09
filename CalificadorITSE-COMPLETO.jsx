import React, { useState } from 'react';
import * as XLSX from 'xlsx';

const CalificadorITSE = () => {
  // ============================================================
  // DATOS COMPLETOS - TODOS LOS TÉCNICOS
  // ============================================================
  
  const examsData = {
    'AC': {
      name: 'Artes Culinarias',
      opciones: {
        1: {
          name: 'Opción A',
          maxPuntoBruto: 33,
          questions: [
            {id: 1, enunciado: 'La sostenibilidad social es menos importante que la ambiental y la económica en el contexto del desarrollo sostenible integral.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'menos importante que', palabraCorrectaEsperada: 'tan importante como', points: 3},
            {id: 2, enunciado: 'El ACV (Análisis de Ciclo de Vida) solo mide impactos en la etapa de operación sin considerar otras fases.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo mide... sin considerar', palabraCorrectaEsperada: 'mide en todas las etapas', points: 3},
            {id: 3, enunciado: 'La Matriz de Aspectos e Impactos Ambientales (ISO 14001) es una herramienta para identificar, evaluar y priorizar impactos significativos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 4, enunciado: 'Las Fichas de Datos de Seguridad (SDS) son documentos opcionales en gestión ambiental de cocinas.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'son opcionales', palabraCorrectaEsperada: 'son obligatorias', points: 3},
            {id: 5, enunciado: 'La Economía Circular busca maximizar valor en múltiples ciclos de vida minimizando residuos irreversibles.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'maximizar... generar', palabraCorrectaEsperada: 'minimizar residuos', points: 3},
            {id: 6, enunciado: 'Un KPI de consumo hídrico mide progreso hacia metas de eficiencia operacional.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 7, enunciado: 'El análisis costo-beneficio triple no incluye dimensiones ambientales.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'no incluye', palabraCorrectaEsperada: 'incluye', points: 3},
            {id: 8, enunciado: 'Los costos invisibles en operaciones culinarias incluyen pérdidas ambientales y sociales.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 9, enunciado: 'La huella de carbono en cocinas profesionales únicamente incluye consumo de gas.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'únicamente', palabraCorrectaEsperada: 'incluye', points: 3},
            {id: 10, enunciado: 'La Ley 41/1998 de Panamá ignora regulaciones para operaciones de servicios de alimentos.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'ignora', palabraCorrectaEsperada: 'promueve', points: 3},
            {id: 11, enunciado: 'Un KRI anticipa riesgos en desempeño operacional mediante umbrales de control.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 12, enunciado: 'La comunicación de sostenibilidad corporativa debe ocultar impactos negativos.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'ocultar', palabraCorrectaEsperada: 'comunicar', points: 3},
            {id: 13, enunciado: 'Los ODS 12 y 13 funcionan desconectados en gestión ambiental culinaria.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'desconectados', palabraCorrectaEsperada: 'vinculados', points: 3},
            {id: 14, enunciado: 'La Economía Circular busca mantener recursos en ciclos de uso el mayor tiempo posible.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 15, enunciado: 'La ecoeficiencia busca hacer más con menos recursos y residuos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1}
          ]
        },
        2: {
          name: 'Opción B',
          maxPuntoBruto: 25,
          questions: [
            {id: 1, enunciado: 'Un resultado de análisis costo-beneficio positivo es más preferible que considerar aspectos ambientales.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'más preferible', palabraCorrectaEsperada: 'menos preferible', points: 3},
            {id: 2, enunciado: 'La Matriz AIA (ISO 14001) identifica aspectos y evalúa impactos significativos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 3, enunciado: 'La responsabilidad de gestión ambiental en cocinas es compartida entre todos los niveles.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 4, enunciado: 'Los impactos directos e indirectos en cocinas generan el mismo nivel de riesgo.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'el mismo impacto', palabraCorrectaEsperada: 'diferentes impactos', points: 3},
            {id: 5, enunciado: 'Un KPI debe ser comprensible para todos los niveles de la organización.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 6, enunciado: 'El análisis costo-beneficio ambiental solo considera dinero, no costos ambientales y sociales.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'únicamente al dinero', palabraCorrectaEsperada: 'costos ambientales y sociales', points: 3},
            {id: 7, enunciado: 'La Ley 41/1998 es la base legislativa ambiental de Panamá.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 8, enunciado: 'La rendición de cuentas ambiental considera solo aspectos económicos.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo económicos', palabraCorrectaEsperada: 'ambientales y sociales', points: 3},
            {id: 9, enunciado: 'Las certificaciones de sostenibilidad culinaria son solo marketing sin verificación rigurosa.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo marketing', palabraCorrectaEsperada: 'verificables y rigurosos', points: 3},
            {id: 10, enunciado: 'Un menú sostenible integra origen de ingredientes, gestión de residuos y transparencia.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 11, enunciado: 'La ecoeficiencia en cocinas busca hacer más con menos recursos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 12, enunciado: 'La huella de carbono corporativa integra consumo energético, viajes, materiales y transporte.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 13, enunciado: 'La Ley 54/2022 es legislación ambiental secundaria en Panamá.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 14, enunciado: 'La educación ambiental corporativa integra conocimiento con cambios de comportamiento.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 15, enunciado: 'Los tres ejes (ambiental, social, económico) son independientes en sostenibilidad.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1}
          ]
        },
        3: {
          name: 'Opción C',
          maxPuntoBruto: 29,
          questions: [
            {id: 1, enunciado: 'La Matriz de Aspectos e Impactos Ambientales solo identifica problemas sin orientar soluciones.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo identifica', palabraCorrectaEsperada: 'identifica y orienta', points: 3},
            {id: 2, enunciado: 'El análisis costo-beneficio triple en cocinas es metodología estándar.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 3, enunciado: 'Los residuos orgánicos de cocinas tienen cero valor económico.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'cero valor', palabraCorrectaEsperada: 'valor económico', points: 3},
            {id: 4, enunciado: 'Un KRI es indicador de riesgo que advierte sobre vulnerabilidades operacionales.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 5, enunciado: 'La huella ecológica en operaciones culinarias es infinita y nunca puede reducirse.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'infinita y nunca', palabraCorrectaEsperada: 'medible y reducible', points: 3},
            {id: 6, enunciado: 'El cambio climático en Panamá tiene impacto real en operaciones culinarias.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'teórico sin impacto', palabraCorrectaEsperada: 'impacto real', points: 3},
            {id: 7, enunciado: 'La Economía Circular aplica a negocios de cualquier tamaño.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 8, enunciado: 'La Ley 41/1998 regula operaciones de servicios de alimentos en Panamá.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 9, enunciado: 'La responsabilidad de gestión ambiental recae únicamente en el productor de alimentos.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'únicamente productor', palabraCorrectaEsperada: 'responsabilidad compartida', points: 3},
            {id: 10, enunciado: 'El ACV (Análisis de Ciclo de Vida) incluye todas las etapas del producto.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 11, enunciado: 'La rendición de cuentas ambiental es compromiso corporativo fundamental.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 12, enunciado: 'Un indicador efectivo mide solo desempeño sin anticipar riesgos.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo desempeño', palabraCorrectaEsperada: 'desempeño y riesgos', points: 3},
            {id: 13, enunciado: 'La PML (Producción Más Limpia) busca optimizar procesos existentes.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 14, enunciado: 'La implementación de sostenibilidad en cocinas no genera costos iniciales.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'no genera costos', palabraCorrectaEsperada: 'genera costos', points: 3},
            {id: 15, enunciado: 'La sensibilización ambiental es paso inicial para desarrollar cultura sostenible.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1}
          ]
        }
      }
    },
    'TM': {
      name: 'Tecnologías Metalmecánicas',
      opciones: {
        1: {
          name: 'Opción 1',
          maxPuntoBruto: 33,
          questions: [
            {id: 1, enunciado: 'La sostenibilidad debe considerar solo aspectos ambientales, ignorando lo social y económico.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'más preferible', palabraCorrectaEsperada: 'menos preferible que reducir y reutilizar', points: 3},
            {id: 2, enunciado: 'El sistema de gestión ambiental requiere documentación formal.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 3, enunciado: 'Un KPI de energía renovable establece alarma si la disponibilidad del sistema cae por debajo del 90%.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'KPI', palabraCorrectaEsperada: 'KRI', points: 3},
            {id: 4, enunciado: 'La Matriz de Leopold y la Matriz AIA (ISO 14001) son herramientas equivalentes.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'menos', palabraCorrectaEsperada: 'tan', points: 3},
            {id: 5, enunciado: 'El impacto ambiental indirecto en metalurgia es irrelevante para la evaluación de sostenibilidad.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'irrelevante', palabraCorrectaEsperada: 'crítica', points: 3},
            {id: 6, enunciado: 'Los Alcances 1, 2 y 3 del GHG Protocol aplican a operaciones MRO.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 7, enunciado: 'El análisis costo-beneficio ambiental no incluye externalidades sociales.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'no incluye', palabraCorrectaEsperada: 'incluye', points: 3},
            {id: 8, enunciado: 'Un aspecto ambiental es la consecuencia del impacto en la operación.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'es la consecuencia', palabraCorrectaEsperada: 'es lo que hace la empresa; el impacto es la consecuencia', points: 3},
            {id: 9, enunciado: 'La responsabilidad de sostenibilidad en talleres es competencia de todos los niveles.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 10, enunciado: 'Un KRI de seguridad es más fácil de medir que un KPI de producción.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'es más fácil de medir', palabraCorrectaEsperada: 'requiere igual rigor en su diseño y medición', points: 3},
            {id: 11, enunciado: 'La Economía Circular en metalurgia busca que residuos terminen en rellenos.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'terminen en rellenos', palabraCorrectaEsperada: 'sean minimizados, reutilizados o recuperados', points: 3},
            {id: 12, enunciado: 'La Ley 41/1998 aplica a talleres de mantenimiento industrial en Panamá.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 13, enunciado: 'La implementación de ecoeficiencia en metalurgia no genera costos de inversión inicial.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'no genera', palabraCorrectaEsperada: 'genera', points: 3},
            {id: 14, enunciado: 'El análisis costo-beneficio triple integra aspectos económicos, ambientales y sociales.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 15, enunciado: 'Los impactos ambientales en metalurgia se limitan únicamente al ruido de operaciones.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'únicamente el ruido', palabraCorrectaEsperada: 'la contaminación de agua y suelo por químicos', points: 3}
          ]
        },
        2: {
          name: 'Opción 2',
          maxPuntoBruto: 33,
          questions: [
            {id: 1, enunciado: 'La Economía Circular en metalurgia busca que residuos se minimicen, reutilicen o recuperen.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'terminen en rellenos', palabraCorrectaEsperada: 'se minimicen, reutilicen o recuperen', points: 3},
            {id: 2, enunciado: 'Un sistema de gestión ambiental requiere compromiso de todos los niveles.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 3, enunciado: 'El aspecto ambiental es la consecuencia de lo que hace la empresa.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'es la consecuencia', palabraCorrectaEsperada: 'es lo que hace la empresa; el impacto es la consecuencia', points: 3},
            {id: 4, enunciado: 'La jerarquía de residuos (reducir, reutilizar, reciclar) es más preferible que desechar.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'más preferible', palabraCorrectaEsperada: 'menos preferible; reducir y reutilizar tienen prioridad', points: 3},
            {id: 5, enunciado: 'Los costos invisibles en metalurgia no deben incluirse en evaluaciones económicas.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'no deben incluirse', palabraCorrectaEsperada: 'deben incluirse', points: 3},
            {id: 6, enunciado: 'La Ley 41/1998 de Panamá es marco legal para operaciones industriales.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 7, enunciado: 'Un KPI mide desempeño operacional hacia objetivos.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'mide el desempeño', palabraCorrectaEsperada: 'anticipa riesgos', points: 3},
            {id: 8, enunciado: 'La responsabilidad ambiental es compartida en la organización.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 9, enunciado: 'El impacto indirecto en metalurgia es menos importante que el impacto directo.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'directo', palabraCorrectaEsperada: 'indirecto', points: 3},
            {id: 10, enunciado: 'La educación ambiental en talleres es componente importante del sistema de gestión.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 11, enunciado: 'La gestión de residuos peligrosos no requiere procedimientos especiales en metalurgia.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'no requiere', palabraCorrectaEsperada: 'requiere', points: 3},
            {id: 12, enunciado: 'Los residuos tóxicos en metalurgia incluyen solo residuos sólidos.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'únicamente residuos sólidos', palabraCorrectaEsperada: 'aguas residuales contaminadas con zinc y químicos', points: 3},
            {id: 13, enunciado: 'La ecoeficiencia busca hacer más con menos recursos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 14, enunciado: 'El análisis de ciclo de vida integra todas las fases del producto.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 15, enunciado: 'La sostenibilidad en metalurgia es responsabilidad solo de gerencia.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1}
          ]
        },
        3: {
          name: 'Opción 3',
          maxPuntoBruto: 33,
          questions: [
            {id: 1, enunciado: 'La educación ambiental en metalurgia no es importante para el desempeño.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'no es', palabraCorrectaEsperada: 'es', points: 3},
            {id: 2, enunciado: 'Un KPI debe ser específico, medible y realista.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 3, enunciado: 'La ecoeficiencia busca reducir consumo de recursos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 4, enunciado: 'Los sistemas de gestión ambiental no contienen procedimientos de verificación.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'no contienen', palabraCorrectaEsperada: 'contienen', points: 3},
            {id: 5, enunciado: 'Los aspectos ambientales en metalurgia tienen importancia más que igual en el desempeño global.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'más importante', palabraCorrectaEsperada: 'igualmente importante; ambos son complementarios', points: 3},
            {id: 6, enunciado: 'La sostenibilidad considera solo aspectos ambientales en metalurgia.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'únicamente aspectos ambientales', palabraCorrectaEsperada: 'ambiental, económico y social', points: 3},
            {id: 7, enunciado: 'Los efluentes de metalurgia no generan contaminación de agua y suelo.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'no genera', palabraCorrectaEsperada: 'genera contaminación de agua y suelo', points: 3},
            {id: 8, enunciado: 'La Ley 54/2022 de Panamá complementa la Ley 41/1998.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 9, enunciado: 'El análisis costo-beneficio ambiental integra aspectos económicos y sociales.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 10, enunciado: 'Los impactos indirectos no deben incluirse en evaluaciones ambientales.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'no deben incluirse', palabraCorrectaEsperada: 'deben incluirse', points: 3},
            {id: 11, enunciado: 'La responsabilidad ambiental en talleres es compartida entre técnicos y gerencia.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 12, enunciado: 'Los residuos especiales en metalurgia requieren manejo especial.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 13, enunciado: 'La Matriz AIA es herramienta para identificar impactos significativos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 14, enunciado: 'El sistema de gestión ambiental es compromiso de largo plazo.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 15, enunciado: 'La comunicación de sostenibilidad debe ser transparente y verificable.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1}
          ]
        }
      }
    },
    'SE+CO': {
      name: 'Servicios Empresariales + Construcción',
      opciones: {
        1: {
          name: 'Opción Única',
          maxPuntoBruto: 36,
          questions: [
            {id: 1, enunciado: 'La Matriz de Aspectos e Impactos Ambientales solo identifica problemas.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo', palabraCorrectaEsperada: 'también identifica soluciones', points: 3},
            {id: 2, enunciado: 'La Ley 41/1998 es la base legislativa ambiental de Panamá.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 3, enunciado: 'El desperdicio afecta únicamente al ambiente.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'únicamente', palabraCorrectaEsperada: 'también', points: 3},
            {id: 4, enunciado: 'Un KRI de agua alerta sobre vulnerabilidad de recursos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 5, enunciado: 'La huella ecológica nunca puede reducirse en operaciones constructivas.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'nunca', palabraCorrectaEsperada: 'siempre', points: 3},
            {id: 6, enunciado: 'La Economía Circular mantiene valor máximo de recursos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 7, enunciado: 'La responsabilidad de sostenibilidad recae solo en proveedores.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo', palabraCorrectaEsperada: 'compartida', points: 3},
            {id: 8, enunciado: 'La rendición de cuentas ambiental implica transparencia.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 9, enunciado: 'La Ley 54/2022 reconoce el cambio climático en Panamá.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 10, enunciado: 'Un menú sostenible integra origen, residuos y transparencia.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 11, enunciado: 'Un análisis de viabilidad ambiental debe considerar solo aspectos técnicos.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo', palabraCorrectaEsperada: 'también económicos y sociales', points: 3},
            {id: 12, enunciado: 'El liderazgo ético en gestión ambiental considera impactos ambientales, sociales y económicos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 13, enunciado: 'La evaluación del impacto ambiental según ODS debe considerar solo el ODS 13.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo', palabraCorrectaEsperada: 'múltiples ODS', points: 3},
            {id: 14, enunciado: 'El cambio climático en Panamá tiene consecuencias reales para operaciones empresariales.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 15, enunciado: 'La medición continua mediante KPI y KRI permite detectar tendencias antes de problemas.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1}
          ]
        }
      }
    },
    'OH': {
      name: 'Operaciones Hoteleras',
      opciones: {
        1: {
          name: 'Opción A',
          maxPuntoBruto: 45,
          questions: [
            {id: 1, enunciado: 'La Educación Ambiental en hoteles solo se enfoca en cumplir leyes.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo', palabraCorrectaEsperada: 'también genera cambios de comportamiento', points: 3},
            {id: 2, enunciado: 'La sostenibilidad social es menos importante que la ambiental y económica.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'menos importante que', palabraCorrectaEsperada: 'tan importante como', points: 3},
            {id: 3, enunciado: 'Los impactos ambientales directos incluyen el consumo de agua en piscinas.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 4, enunciado: 'La Matriz AIA identifica y prioriza impactos ambientales por significancia.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 5, enunciado: 'Un KPI efectivo debe ser complejo para demostrar profesionalismo.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'complejo', palabraCorrectaEsperada: 'simple y comprensible', points: 3},
            {id: 6, enunciado: 'La Economía Circular busca mantener recursos en uso el mayor tiempo posible.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 7, enunciado: 'El ACV solo mide impactos en la fase de operación del producto.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo', palabraCorrectaEsperada: 'todas las fases', points: 3},
            {id: 8, enunciado: 'Los costos invisibles son montos de dinero ocultados en secreto en la contabilidad.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'ocultados en secreto', palabraCorrectaEsperada: 'costos ambientales y sociales no contabilizados', points: 3},
            {id: 9, enunciado: 'Los KRI anticipan riesgos futuros en desempeño operacional.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 10, enunciado: 'La Ley 41/1998 es la base legislativa ambiental de Panamá.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 11, enunciado: 'La responsabilidad de sostenibilidad recae únicamente en proveedores.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'únicamente', palabraCorrectaEsperada: 'compartida', points: 3},
            {id: 12, enunciado: 'La huella de carbono de un hotel incluye emisiones de energía, agua y transporte.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 13, enunciado: 'La ecoeficiencia busca hacer más con menos recursos y residuos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 14, enunciado: 'El análisis costo-beneficio ambiental incluye solo números financieros.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo', palabraCorrectaEsperada: 'también ambientales y sociales', points: 3},
            {id: 15, enunciado: 'Los residuos sólidos de hoteles deben ser clasificados según su naturaleza.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1}
          ]
        },
        2: {
          name: 'Opción B',
          maxPuntoBruto: 45,
          questions: [
            {id: 1, enunciado: 'El diagnóstico ambiental solo identifica problemas, sin buscar soluciones.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo', palabraCorrectaEsperada: 'también identifica soluciones', points: 3},
            {id: 2, enunciado: 'Los impactos indirectos en hoteles no pueden medirse ni evaluarse.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'no pueden', palabraCorrectaEsperada: 'pueden', points: 3},
            {id: 3, enunciado: 'La Ley 54/2022 es la legislación ambiental primaria de Panamá.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'primaria', palabraCorrectaEsperada: 'complementaria/secundaria', points: 3},
            {id: 4, enunciado: 'La rendición de cuentas ambiental genera transparencia y confianza con stakeholders.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 5, enunciado: 'Un hotel puede reducir su huella ecológica mediante optimización de recursos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 6, enunciado: 'La Economía Circular busca maximizar valor mientras se minimiza residuos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 7, enunciado: 'Los Datos de Seguridad (FDS) solo son requeridos por hospitales, no hoteles.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo', palabraCorrectaEsperada: 'también en hoteles', points: 3},
            {id: 8, enunciado: 'El ACV (Análisis de Ciclo de Vida) incluye todas las fases desde extracción hasta fin de vida.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 9, enunciado: 'Los ODS son compromisos voluntarios sin aplicación obligatoria en hoteles.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'voluntarios', palabraCorrectaEsperada: 'marco global con orientación vinculante', points: 3},
            {id: 10, enunciado: 'La ecoeficiencia busca aumentar productividad con menos recursos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1},
            {id: 11, enunciado: 'Los costos ambientales son irrelevantes en decisiones empresariales de hoteles.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'irrelevantes', palabraCorrectaEsperada: 'relevantes/críticos', points: 3},
            {id: 12, enunciado: 'La Matriz AIA bajo ISO 14001 identifica aspectos pero no prioriza impactos.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'no prioriza', palabraCorrectaEsperada: 'prioriza', points: 3},
            {id: 13, enunciado: 'El desempeño ambiental de un hotel es responsabilidad solo de gerencia.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo', palabraCorrectaEsperada: 'compartida/todos', points: 3},
            {id: 14, enunciado: 'La PML implica cambiar completos sistemas, no solo optimizar procesos.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'cambiar completos', palabraCorrectaEsperada: 'optimizar', points: 3},
            {id: 15, enunciado: 'La sensibilización ambiental es el paso inicial para desarrollar cultura sostenible.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 1}
          ]
        }
      }
    },
    'OL': {
      name: 'Operaciones Logísticas',
      opciones: {
        1: {
          name: 'Opción A',
          maxPuntoBruto: 45,
          questions: [
            {id: 1, enunciado: 'La Matriz de Aspectos e Impactos Ambientales solo identifica problemas.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo', palabraCorrectaEsperada: 'identifica y prioriza', points: 3},
            {id: 2, enunciado: 'La Ley 41/1998 es base legislativa ambiental en Panamá.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 2},
            {id: 3, enunciado: 'El desperdicio afecta únicamente al ambiente.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'únicamente', palabraCorrectaEsperada: 'también', points: 3},
            {id: 4, enunciado: 'Un KRI de agua alerta sobre vulnerabilidad de recursos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 2},
            {id: 5, enunciado: 'La huella ecológica nunca puede reducirse en operaciones logísticas.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'nunca', palabraCorrectaEsperada: 'siempre', points: 3},
            {id: 6, enunciado: 'La Economía Circular mantiene valor máximo de recursos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 2},
            {id: 7, enunciado: 'La responsabilidad de sostenibilidad recae solo en proveedores.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo', palabraCorrectaEsperada: 'compartida', points: 3},
            {id: 8, enunciado: 'La rendición de cuentas ambiental implica transparencia.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 2},
            {id: 9, enunciado: 'La Ley 54/2022 reconoce cambio climático en Panamá.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 2},
            {id: 10, enunciado: 'Un menú sostenible integra origen, residuos y transparencia.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 2},
            {id: 11, enunciado: 'El Análisis de Ciclo de Vida solo incluye fase de distribución.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo', palabraCorrectaEsperada: 'incluye desde cuna hasta tumba', points: 3},
            {id: 12, enunciado: 'Los impactos indirectos en logística no pueden medirse.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'no pueden', palabraCorrectaEsperada: 'pueden', points: 3},
            {id: 13, enunciado: 'La ecoeficiencia busca hacer más con menos recursos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 2},
            {id: 14, enunciado: 'El análisis de costo-beneficio son solo números financieros.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo', palabraCorrectaEsperada: 'incluye', points: 3},
            {id: 15, enunciado: 'La sostenibilidad social se refiere solo a salarios.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo', palabraCorrectaEsperada: 'también incluye', points: 3}
          ]
        }
      }
    },
    'MA': {
      name: 'Mantenimiento de Aeronaves',
      opciones: {
        1: {
          name: 'Opción A',
          maxPuntoBruto: 45,
          questions: [
            {id: 1, enunciado: 'La educación ambiental en talleres aeronáuticos debe integrar conciencia y prácticas sostenibles en todas las operaciones.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 2, enunciado: 'Los impactos ambientales indirectos son menos importantes que los impactos directos en un análisis ambiental.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'menos importante', palabraCorrectaEsperada: 'tan importante (o igual de importante)', points: 3},
            {id: 3, enunciado: 'La Matriz de Aspectos e Impactos Ambientales (ISO 14001) es una herramienta para identificar, evaluar y priorizar impactos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 4, enunciado: 'Un KPI debe ser tan complejo que solo sea comprensible para la gerencia de ingeniería.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'tan complejo', palabraCorrectaEsperada: 'claro, simple y comprensible para todos', points: 3},
            {id: 5, enunciado: 'Los alcances 1, 2 y 3 del GHG Protocol permiten medir emisiones directas e indirectas de carbono en operaciones MRO.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 6, enunciado: 'La Economía Circular busca mantener recursos en ciclos de uso el mayor tiempo posible, reduciendo residuos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 7, enunciado: 'El ACV (Análisis de Ciclo de Vida) solo mide impactos en la fase de operación o uso del producto.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo', palabraCorrectaEsperada: 'mide desde extracción de materias primas hasta disposición final', points: 3},
            {id: 8, enunciado: 'Los KRI son indicadores que anticipan riesgos futuros en desempeño operacional y seguridad.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 9, enunciado: 'La Ley 41/1998 de Panamá es la base legislativa ambiental del país y aplica a operaciones industriales y aeronáuticas.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 10, enunciado: 'Los costos invisibles en MRO incluyen solo dinero perdido por robo de repuestos.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo', palabraCorrectaEsperada: 'son costos ambientales y sociales no contabilizados', points: 3},
            {id: 11, enunciado: 'Las Fichas de Datos de Seguridad (SDS) contienen información sobre composición química, peligros y medidas de emergencia.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 12, enunciado: 'La clasificación GHS (Globally Harmonized System) es obligatoria para el etiquetado de sustancias peligrosas en talleres aeronáuticos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 13, enunciado: 'El análisis costo-beneficio triple en MRO integra únicamente aspectos económicos, ignorando lo ambiental y social.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'únicamente aspectos económicos', palabraCorrectaEsperada: 'integra aspectos ambientales, económicos y sociales', points: 3},
            {id: 14, enunciado: 'La ecoeficiencia en mantenimiento aeronáutico significa producir igual o mejor servicio con menor consumo de recursos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 15, enunciado: 'La responsabilidad de gestión ambiental recae únicamente en proveedores externos, no en el personal técnico interno.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'únicamente en proveedores externos', palabraCorrectaEsperada: 'es compartida entre todos los niveles: técnicos, supervisores y gerencia', points: 3}
          ]
        },
        2: {
          name: 'Opción B',
          maxPuntoBruto: 45,
          questions: [
            {id: 1, enunciado: 'La Matriz AIA (ISO 14001) permite identificar y priorizar impactos ambientales significativos en operaciones de mantenimiento.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 2, enunciado: 'Un KPI debe ser tan simple que cualquier técnico del taller pueda entender qué se está midiendo y por qué importa.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 3, enunciado: 'Los residuos peligrosos aeronáuticos (RTP) como solventes y aceites pueden ser disposición en incineración controlada según legislación ambiental panameña.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 4, enunciado: 'El GHG Protocol Alcance 1 incluye solo emisiones del transporte de personal hacia el taller MRO.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo emisiones del transporte', palabraCorrectaEsperada: 'todas las emisiones directas de fuentes propias del taller', points: 3},
            {id: 5, enunciado: 'Los análisis costo-beneficio triple deben ser ignorados en talleres pequeños para reducir complejidad administrativa.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'deben ser ignorados', palabraCorrectaEsperada: 'deben considerarse en toda escala de operación', points: 3},
            {id: 6, enunciado: 'Un KRI de derrames químicos podría ser: (número de incidentes / 1000 horas trabajadas) × 100.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 7, enunciado: 'La sostenibilidad tridimensional implica únicamente maximizar ganancias económicas del taller.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'únicamente maximizar ganancias', palabraCorrectaEsperada: 'integrar aspectos ambientales, sociales y económicos', points: 3},
            {id: 8, enunciado: 'Las Fichas SDS en formato digital son menos válidas que las impresas para cumplimiento ambiental.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'menos válidas', palabraCorrectaEsperada: 'tienen igual validez legal si están debidamente registradas', points: 3},
            {id: 9, enunciado: 'El ACV (Life Cycle Assessment) evalúa impactos ambientales desde la extracción de materias primas hasta disposición final del producto.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 10, enunciado: 'La Economía Circular solo aplica a grandes corporaciones, no a talleres medianos de mantenimiento.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo aplica a grandes corporaciones', palabraCorrectaEsperada: 'aplica a operaciones de cualquier tamaño para optimizar recursos', points: 3},
            {id: 11, enunciado: 'En la Fase 2 del Proyecto Final, los estudiantes utilizan la Matriz de Aspectos e Impactos para evaluar significancia de impactos identificados.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 12, enunciado: 'Un técnico debe ignorar indicios de contaminación ambiental si no afectan directamente su seguridad personal.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'debe ignorar', palabraCorrectaEsperada: 'debe reportar y tomar acciones', points: 3},
            {id: 13, enunciado: 'La responsabilidad de gestión ambiental es compartida entre todos los niveles de la organización, desde técnicos hasta gerencia.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 14, enunciado: 'El Sistema GHS de clasificación de sustancias peligrosas es obligatorio solo en países europeos, no en Panamá.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo en países europeos', palabraCorrectaEsperada: 'es obligatorio en Panamá y aplicable globalmente', points: 3},
            {id: 15, enunciado: 'Un indicador efectivo debe responder a cambios reales en operaciones y generar acciones de mejora, no solo recolectar datos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3}
          ]
        }
      }
    },
    'IA': {
      name: 'Inteligencia Artificial',
      opciones: {
        1: {
          name: 'Opción 1',
          maxPuntoBruto: 45,
          questions: [
            {id: 1, enunciado: 'La sostenibilidad social es menos importante que la ambiental y la económica en el contexto del desarrollo sostenible integral.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'menos', palabraCorrectaEsperada: 'tan', points: 3},
            {id: 2, enunciado: 'El GHG Protocol (Protocolo de Gases de Efecto Invernadero) propone tres alcances: Alcance 1 (emisiones directas), Alcance 2 (energía comprada) y Alcance 3 (emisiones indirectas de la cadena de valor).', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 3, enunciado: 'Un KPI de residuos tóxicos y peligrosos (RTP) tiene como fórmula: (kg de RTP generados ÷ kg de residuos totales) × 100, con meta de eliminar el 100% de residuos peligrosos.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'kg de RTP generados', palabraCorrectaEsperada: 'kg entregados a gestor autorizado', points: 3},
            {id: 4, enunciado: 'Los datasets en inteligencia artificial NO requieren protocolo de privacidad porque todos los datos son públicos por naturaleza.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'NO requieren', palabraCorrectaEsperada: 'requieren', points: 3},
            {id: 5, enunciado: 'La capacidad total de almacenamiento en servidores puede exceder el 90% de utilización sin generar alertas de riesgo en un datacenter de IA.', respuestaEsperada: 'F', palabraIncorrectaEsperada: '90%', palabraCorrectaEsperada: '80%', points: 3},
            {id: 6, enunciado: 'Una matriz de aspectos e impactos ambientales AIA evalúa la significancia de impactos multiplicando Magnitud × Extensión × Reversibilidad (M×E×R).', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 7, enunciado: 'El entrenamiento de modelos de machine learning consume energía insignificante comparado con otros procesos industriales, por lo que no requiere optimización.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'insignificante', palabraCorrectaEsperada: 'significativo', points: 3},
            {id: 8, enunciado: 'La huella hídrica de un datacenter se relaciona únicamente con el agua potable consumida por los empleados, no con sistemas de refrigeración.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'únicamente con el agua potable', palabraCorrectaEsperada: 'incluye principalmente sistemas de refrigeración y climatización', points: 3},
            {id: 9, enunciado: 'La economía circular en IA busca reutilizar componentes electrónicos, recuperar minerales críticos de residuos, y extender la vida útil de servidores mediante mantenimiento preventivo.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 10, enunciado: 'Un KRI es una herramienta de alerta temprana que anticipa riesgos o condiciones insostenibles antes de que se conviertan en problemas críticos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 11, enunciado: 'La Ley 41/1998 (Ley General de Ambiente de Panamá) es el único marco regulatorio que aplica a empresas de tecnología e inteligencia artificial en el país.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'único', palabraCorrectaEsperada: 'principal', points: 3},
            {id: 12, enunciado: 'La huella de carbono de un sistema de inteligencia artificial SOLO se calcula en Alcance 1 (emisiones directas) del GHG Protocol.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'SOLO en Alcance 1', palabraCorrectaEsperada: 'en Alcances 1, 2 y 3', points: 3},
            {id: 13, enunciado: 'Un protocolo de privacidad en datasets garantiza que solo usuarios autorizados accedan a datos personales y que se cumple con regulaciones como GDPR y Ley 51/2012.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 14, enunciado: 'La acumulación de versiones antiguas de modelos sin depuración en servidores es un factor menor y no compromete la eficiencia energética de un datacenter.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'menor', palabraCorrectaEsperada: 'significativo', points: 3},
            {id: 15, enunciado: 'La ética en inteligencia artificial incluye el derecho de las máquinas a tomar decisiones autónomas sin supervisión humana en asuntos que afecten a comunidades vulnerables.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'sin supervisión humana', palabraCorrectaEsperada: 'con supervisión y auditoria humana permanente', points: 3}
          ]
        },
        2: {
          name: 'Opción 2',
          maxPuntoBruto: 45,
          questions: [
            {id: 1, enunciado: 'El consumo de energía en centros de datos es insignificante en comparación con emisiones totales de carbono global.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'insignificante', palabraCorrectaEsperada: 'significativo', points: 3},
            {id: 2, enunciado: 'La Ley 51/2012 de Panamá regula protección de datos personales y privacidad en la recopilación de información.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 3, enunciado: 'Un aspecto ambiental crítico en IA es la acumulación de versiones antiguas de modelos sin limpieza de datos en servidores, que consume espacio y energía.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'consumex', palabraCorrectaEsperada: 'consume', points: 3},
            {id: 4, enunciado: 'La sostenibilidad en IA solo considera impactos ambientales, sin tomar en cuenta dimensiones sociales o económicas.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo considera impactos ambientales', palabraCorrectaEsperada: 'integra ambiental, social y económico', points: 3},
            {id: 5, enunciado: 'Una matriz de aspectos e impactos AIA identificada y priorizada permite enfoque de gestión estratégico en los impactos más significativos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 6, enunciado: 'Los minerales críticos como cobalto y litio se encuentran en cantidad ilimitada en la naturaleza y su extracción no produce daño ambiental.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'ilimitada', palabraCorrectaEsperada: 'limitada', points: 3},
            {id: 7, enunciado: 'La huella ecológica mide la capacidad de regeneración de recursos naturales necesarios para sustentar consumo de energía y materiales en sistemas de IA.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 8, enunciado: 'El GDPR y leyes similares son marcos regulatorios que SOLO aplican en Europa y no tienen relevancia en Panamá.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'SOLO aplican en Europa', palabraCorrectaEsperada: 'aplican globalmente incluyendo Panamá', points: 3},
            {id: 9, enunciado: 'Un análisis costo-beneficio triple considera impactos económicos, ambientales y sociales de una inversión en IA sostenible.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 10, enunciado: 'La ética en inteligencia artificial excluye decisiones algorítmicas que afecten a grupos vulnerables sin protecciones.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'excluye decisiones', palabraCorrectaEsperada: 'incluye protecciones para decisiones', points: 3},
            {id: 11, enunciado: 'Green Computing busca minimizar consumo de energía, usar fuentes renovables, y reducir residuos electrónicos en operaciones de IA.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 12, enunciado: 'Un KRI de \'nivel de almacenamiento en servidores\' con alerta al 80% es innecesario porque la capacidad máxima es expansible.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'innecesario', palabraCorrectaEsperada: 'necesario', points: 3},
            {id: 13, enunciado: 'La fase de despliegue de sistemas de IA tiene impactos ambientales menores comparada con entrenamiento.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'menores', palabraCorrectaEsperada: 'significativos y continuos', points: 3},
            {id: 14, enunciado: 'La economía circular en tecnología promueve diseño, uso, reutilización y recuperación de valor sin generación irreversible de residuos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 15, enunciado: 'La reproducibilidad ética de algoritmos es un concepto irrelevante porque siempre que un modelo produce buenos resultados es éticamente válido.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'irrelevante', palabraCorrectaEsperada: 'esencial', points: 3}
          ]
        }
      }
    },
    'EB': {
      name: 'Gestión Ejecutiva Bilingüe',
      opciones: {
        1: {
          name: 'Opción Única',
          maxPuntoBruto: 45,
          questions: [
            {id: 1, enunciado: 'La Educación Ambiental corporativa moderna desarrolla conocimiento integrado con cambios de comportamiento, mentalidad y prácticas operacionales.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'únicamente', palabraCorrectaEsperada: 'también genera cambios de comportamiento', points: 3},
            {id: 2, enunciado: 'Un KRI de energía renovable establece alarma si la disponibilidad del sistema cae por debajo del 90%.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'mide', palabraCorrectaEsperada: 'advierte', points: 3},
            {id: 3, enunciado: 'La Matriz ISO 14001 AIA identifica aspecto, impacto, severidad y probabilidad del impacto potencial.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 4, enunciado: 'Un resultado de análisis costo-beneficio positivo es necesario mas no suficiente para autorizar una propuesta ambiental.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'suficiente pero no necesaria', palabraCorrectaEsperada: 'necesaria pero no suficiente', points: 3},
            {id: 5, enunciado: 'La Economía Circular busca maximizar valor en múltiples ciclos de vida del producto; PML busca optimizar procesos existentes.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'sinónimos exactamente iguales', palabraCorrectaEsperada: 'conceptos complementarios con diferencias metodológicas', points: 3},
            {id: 6, enunciado: 'Una ficha SDS es documento vivo que se actualiza cuando hay cambios en formulación o nueva información de riesgos.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 7, enunciado: 'El precio del descuido ambiental corporativo incluye sanciones, pérdida de reputación, costos de remediación y retención de talento.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'solo multas regulatorias', palabraCorrectaEsperada: 'multas, reputación, remediación, talento, clientes', points: 3},
            {id: 8, enunciado: 'Un KRI que no genera acciones trimestrales se considera funcional y debe mantenerse igual.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 9, enunciado: 'Los tres ejes (ambiental, social, económico) son interdependientes en sostenibilidad corporativa moderna.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'independiente de', palabraCorrectaEsperada: 'interdependientes con', points: 3},
            {id: 10, enunciado: 'Un KPI de consumo hídrico "(m³ reducidos/m³ línea base)×100" mide progreso hacia meta de eficiencia.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'todos los residuos sean reciclables', palabraCorrectaEsperada: 'meta reconoce % realista alcanzable', points: 3},
            {id: 11, enunciado: 'La huella de carbono corporativa integra consumo energético, viajes de trabajo, materiales y transporte de suministros.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'únicamente', palabraCorrectaEsperada: 'también incluye viajes y materiales', points: 3},
            {id: 12, enunciado: 'Los ODS 12 y 13 cumplen objetivos separados sin conexión en gestión ambiental empresarial.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 13, enunciado: 'La Ley 41/1998 Panamá junto con Ley 54/2022 forman el marco normativo ambiental aplicable a operaciones empresariales.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 14, enunciado: 'Un KPI mide avance hacia objetivos; un KRI advierte sobre riesgos mediante umbrales de control.', respuestaEsperada: 'V', palabraIncorrectaEsperada: 'N/A', palabraCorrectaEsperada: 'N/A', points: 3},
            {id: 15, enunciado: 'La comunicación de sostenibilidad ejecutiva es bilateral involucrando empleados, gerencia y stakeholders externos.', respuestaEsperada: 'F', palabraIncorrectaEsperada: 'unilateral, únicamente hacia stakeholders externos', palabraCorrectaEsperada: 'bilateral involucrando empleados y gerencia', points: 3}
          ]
        }
      }
    }
  };

  // ============================================================
  // ESTADOS
  // ============================================================
  const [stage, setStage] = useState('select');
  const [selectedTech, setSelectedTech] = useState(null);
  const [selectedOp, setSelectedOp] = useState(null);
  const [studentName, setStudentName] = useState('');
  const [currentQuestionIndex, setCurrentQuestionIndex] = useState(0);
  const [responses, setResponses] = useState([]);
  const [loading, setLoading] = useState(false);
  const [feedback, setFeedback] = useState('');

  // ============================================================
  // FUNCIONES
  // ============================================================

  const getCurrentExam = () => {
    if (!selectedTech || !selectedOp) return null;
    return examsData[selectedTech]?.opciones[selectedOp] || null;
  };

  const currentExam = getCurrentExam();
  const currentQuestion = currentExam?.questions[currentQuestionIndex];

  const handleStartGrading = () => {
    if (!selectedTech || !selectedOp || !studentName.trim()) {
      alert('Por favor completa todos los campos');
      return;
    }
    setResponses(new Array(currentExam.questions.length).fill(null));
    setStage('grading');
    setCurrentQuestionIndex(0);
  };

  const analyzeResponse = async (studentResponse) => {
    if (!currentQuestion) return null;

    setLoading(true);
    setFeedback('Analizando respuesta con Claude...');

    try {
      const prompt = `Eres un evaluador académico estricto de exámenes de Educación Ambiental Panamá.

CONTEXTO DEL EXAMEN:
- Pregunta: "${currentQuestion.enunciado}"
- Respuesta esperada: ${currentQuestion.respuestaEsperada}
- Palabra/frase incorrecta esperada: "${currentQuestion.palabraIncorrectaEsperada}"
- Palabra/frase correcta esperada: "${currentQuestion.palabraCorrectaEsperada}"

RESPUESTA DEL ESTUDIANTE:
- Marcó: ${studentResponse.answer}
- Palabra incorrecta ingresada: "${studentResponse.incorrectWord || ''}"
- Palabra correcta ingresada: "${studentResponse.correctWord || ''}"

CRITERIOS DE EVALUACIÓN:
1. ¿Marcó V o F correctamente?
2. Si marcó F:
   - ¿La palabra incorrecta identifica el error conceptual del enunciado?
   - ¿La palabra correcta es un sinónimo válido, contextualmente apropiado?
   - ¿Es sustancialmente diferente a la palabra incorrecta?
3. Valida en estos niveles:
   - Sencillos: palabras directas (ej: "cielo"→"azul")
   - Complejos: frases relacionadas (ej: "reduce costos"→"minimize expenses")
   - Muy complejos: conceptos que reformulan manteniendo la esencia

RESPONDE SOLO CON JSON (sin markdown ni explicaciones):
{
  "isCorrect": boolean,
  "score": número (0, 1, 2 o 3),
  "reasoning": "explicación breve",
  "synonymLevel": "sencillo|complejo|muy complejo|no aceptado",
  "contextualFit": "aceptado|parcialmente|rechazado"
}`;

      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 250,
          messages: [{ role: "user", content: prompt }]
        })
      });

      const data = await response.json();
      const textContent = data.content?.[0]?.text || '{}';
      const jsonMatch = textContent.match(/\{[\s\S]*\}/);
      const result = jsonMatch ? JSON.parse(jsonMatch[0]) : { isCorrect: false, score: 0, reasoning: 'Error al analizar' };

      setFeedback(`${result.reasoning} (Sinónimo: ${result.synonymLevel}, Contexto: ${result.contextualFit})`);
      return result;
    } catch (error) {
      console.error('Error API:', error);
      setFeedback('Error al conectar con Claude API. Revisa manualmente.');
      return { isCorrect: false, score: 0, reasoning: 'Error en análisis' };
    } finally {
      setLoading(false);
    }
  };

  const handleSubmitQuestion = async () => {
    if (!currentQuestion) return;

    const answerSelect = document.getElementById('answer-select');
    const incorrectWordInput = document.getElementById('incorrect-word');
    const correctWordInput = document.getElementById('correct-word');

    const studentResponse = {
      answer: answerSelect?.value || '',
      incorrectWord: incorrectWordInput?.value || '',
      correctWord: correctWordInput?.value || ''
    };

    if (!studentResponse.answer) {
      alert('Marca V o F');
      return;
    }

    setLoading(true);

    let evaluation;

    if (studentResponse.answer === 'V') {
      evaluation = currentQuestion.respuestaEsperada === 'V'
        ? { score: currentQuestion.points, isCorrect: true, reasoning: '✓ Verdadero correcto' }
        : { score: 0, isCorrect: false, reasoning: '✗ Verdadero incorrecto' };
    } else {
      if (!studentResponse.incorrectWord || !studentResponse.correctWord) {
        evaluation = {
          score: 0,
          isCorrect: false,
          reasoning: '✗ Falso incompleto (falta palabra/frase)'
        };
      } else if (currentQuestion.respuestaEsperada === 'F') {
        evaluation = await analyzeResponse(studentResponse) || {
          score: 0,
          isCorrect: false,
          reasoning: 'Error en análisis'
        };
      } else {
        evaluation = {
          score: 0,
          isCorrect: false,
          reasoning: '✗ Falso incorrecto (respuesta correcta es V)'
        };
      }
    }

    const newResponses = [...responses];
    newResponses[currentQuestionIndex] = {
      questionId: currentQuestion.id,
      enunciado: currentQuestion.enunciado,
      studentResponse,
      evaluation,
      maxScore: currentQuestion.points
    };
    setResponses(newResponses);

    if (currentQuestionIndex < currentExam.questions.length - 1) {
      setCurrentQuestionIndex(currentQuestionIndex + 1);
      setFeedback('');
      if (answerSelect) answerSelect.value = '';
      if (incorrectWordInput) incorrectWordInput.value = '';
      if (correctWordInput) correctWordInput.value = '';
    } else {
      setStage('results');
    }

    setLoading(false);
  };

  const generateExcel = () => {
    const exam = getCurrentExam();
    const totalBruto = responses.reduce((sum, r) => sum + (r?.evaluation?.score || 0), 0);
    const maxBruto = exam.maxPuntoBruto;
    const calificacionNormalizada = (totalBruto / maxBruto) * 70;
    const calificacion100 = (calificacionNormalizada / 70) * 100;

    const desglose = responses.map((r, idx) => ({
      Pregunta: idx + 1,
      Respuesta: r.studentResponse.answer === 'V' ? 'V' : `F: ${r.studentResponse.correctWord}`,
      Puntos: r.evaluation.score,
      Máximo: r.maxScore,
      Observación: r.evaluation.reasoning.substring(0, 50)
    }));

    const wsDesglose = XLSX.utils.json_to_sheet(desglose);
    wsDesglose['!cols'] = [{ wch: 10 }, { wch: 20 }, { wch: 8 }, { wch: 8 }, { wch: 40 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsDesglose, 'Desglose');

    const resumenData = [{
      Estudiante: studentName,
      Técnico: examsData[selectedTech].name,
      Opción: selectedOp,
      'Puntos Brutos': totalBruto,
      'Máximo Bruto': maxBruto,
      'Calificación /70': calificacionNormalizada.toFixed(2),
      'Calificación /100': calificacion100.toFixed(2)
    }];

    const wsResumen = XLSX.utils.json_to_sheet(resumenData);
    wsResumen['!cols'] = [{ wch: 20 }, { wch: 30 }, { wch: 10 }, { wch: 15 }, { wch: 15 }, { wch: 18 }, { wch: 18 }];
    XLSX.utils.book_append_sheet(wb, wsResumen, 'Resumen', 0);

    XLSX.writeFile(wb, `${studentName}_${examsData[selectedTech].name}_Op${selectedOp}.xlsx`);
  };

  // ============================================================
  // RENDER: SELECCIÓN
  // ============================================================
  if (stage === 'select') {
    return (
      <div style={{ padding: '30px', fontFamily: 'Arial, sans-serif', maxWidth: '650px', margin: '0 auto', backgroundColor: '#f5f5f5', minHeight: '100vh' }}>
        <div style={{ backgroundColor: 'white', padding: '30px', borderRadius: '8px', boxShadow: '0 2px 8px rgba(0,0,0,0.1)' }}>
          <h1 style={{ color: '#1F5C3A', marginBottom: '30px', textAlign: 'center' }}>📋 Calificador Sección B - ITSE</h1>

          <div style={{ marginBottom: '20px' }}>
            <label style={{ display: 'block', fontWeight: 'bold', marginBottom: '8px', color: '#333' }}>Técnico:</label>
            <select
              value={selectedTech || ''}
              onChange={(e) => { setSelectedTech(e.target.value); setSelectedOp(null); }}
              style={{ width: '100%', padding: '12px', fontSize: '14px', border: '1px solid #ddd', borderRadius: '4px' }}
            >
              <option value="">-- Selecciona un técnico --</option>
              {Object.entries(examsData).map(([key, val]) => (
                <option key={key} value={key}>{val.name}</option>
              ))}
            </select>
          </div>

          {selectedTech && (
            <div style={{ marginBottom: '20px' }}>
              <label style={{ display: 'block', fontWeight: 'bold', marginBottom: '8px', color: '#333' }}>Opción:</label>
              <select
                value={selectedOp || ''}
                onChange={(e) => setSelectedOp(Number(e.target.value))}
                style={{ width: '100%', padding: '12px', fontSize: '14px', border: '1px solid #ddd', borderRadius: '4px' }}
              >
                <option value="">-- Selecciona una opción --</option>
                {Object.entries(examsData[selectedTech].opciones).map(([key, val]) => (
                  <option key={key} value={key}>{val.name}</option>
                ))}
              </select>
            </div>
          )}

          <div style={{ marginBottom: '30px' }}>
            <label style={{ display: 'block', fontWeight: 'bold', marginBottom: '8px', color: '#333' }}>Nombre del estudiante:</label>
            <input
              type="text"
              value={studentName}
              onChange={(e) => setStudentName(e.target.value)}
              placeholder="Ej: Juan Pérez García"
              style={{ width: '100%', padding: '12px', fontSize: '14px', border: '1px solid #ddd', borderRadius: '4px', boxSizing: 'border-box' }}
            />
          </div>

          <button
            onClick={handleStartGrading}
            disabled={!selectedTech || !selectedOp || !studentName.trim()}
            style={{
              width: '100%',
              padding: '14px',
              backgroundColor: !selectedTech || !selectedOp || !studentName.trim() ? '#ccc' : '#1F5C3A',
              color: 'white',
              border: 'none',
              borderRadius: '4px',
              cursor: !selectedTech || !selectedOp || !studentName.trim() ? 'not-allowed' : 'pointer',
              fontSize: '16px',
              fontWeight: 'bold',
              transition: 'background-color 0.3s'
            }}
            onMouseOver={(e) => { if (selectedTech && selectedOp && studentName.trim()) e.target.style.backgroundColor = '#153a28'; }}
            onMouseOut={(e) => { if (selectedTech && selectedOp && studentName.trim()) e.target.style.backgroundColor = '#1F5C3A'; }}
          >
            🚀 Comenzar Calificación
          </button>
        </div>
      </div>
    );
  }

  // ============================================================
  // RENDER: CALIFICACIÓN
  // ============================================================
  if (stage === 'grading' && currentQuestion) {
    const progress = ((currentQuestionIndex + 1) / currentExam.questions.length) * 100;

    return (
      <div style={{ padding: '30px', fontFamily: 'Arial, sans-serif', maxWidth: '800px', margin: '0 auto', backgroundColor: '#f5f5f5', minHeight: '100vh' }}>
        <div style={{ backgroundColor: 'white', padding: '30px', borderRadius: '8px', boxShadow: '0 2px 8px rgba(0,0,0,0.1)' }}>
          <div style={{ marginBottom: '15px' }}>
            <h2 style={{ color: '#1F5C3A', margin: '0 0 5px 0' }}>Calificador ITSE</h2>
            <p style={{ margin: '0', color: '#666', fontSize: '14px' }}>
              <strong>{examsData[selectedTech].name}</strong> | Opción {selectedOp} | Estudiante: <strong>{studentName}</strong>
            </p>
          </div>

          <div style={{ marginBottom: '25px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '8px', fontSize: '12px', color: '#666' }}>
              <span>Progreso</span>
              <span>{currentQuestionIndex + 1} de {currentExam.questions.length}</span>
            </div>
            <div style={{ width: '100%', backgroundColor: '#e0e0e0', borderRadius: '4px', height: '8px', overflow: 'hidden' }}>
              <div style={{ width: progress + '%', backgroundColor: '#1F5C3A', height: '100%', transition: 'width 0.3s' }}></div>
            </div>
          </div>

          <div style={{ backgroundColor: '#f0f8f4', padding: '18px', borderRadius: '4px', marginBottom: '25px', borderLeft: '4px solid #1F5C3A' }}>
            <p style={{ margin: '0 0 10px 0', fontWeight: 'bold', color: '#1F5C3A' }}>Pregunta {currentQuestion.id}:</p>
            <p style={{ margin: '0', fontSize: '16px', lineHeight: '1.5', color: '#333' }}>{currentQuestion.enunciado}</p>
          </div>

          <div style={{ marginBottom: '20px' }}>
            <label style={{ display: 'block', fontWeight: 'bold', marginBottom: '8px', color: '#333' }}>Respuesta del estudiante:</label>
            <select id="answer-select" style={{ width: '100%', padding: '12px', fontSize: '14px', border: '1px solid #ddd', borderRadius: '4px' }}>
              <option value="">-- Verdadero o Falso --</option>
              <option value="V">Verdadero (V)</option>
              <option value="F">Falso (F)</option>
            </select>
          </div>

          {document.getElementById('answer-select')?.value === 'F' && (
            <>
              <div style={{ marginBottom: '15px' }}>
                <label style={{ display: 'block', fontWeight: 'bold', marginBottom: '8px', color: '#333' }}>Palabra/frase incorrecta:</label>
                <input
                  id="incorrect-word"
                  type="text"
                  placeholder="Lo que el estudiante identificó como incorrecto"
                  style={{ width: '100%', padding: '12px', fontSize: '14px', border: '1px solid #ddd', borderRadius: '4px', boxSizing: 'border-box' }}
                />
              </div>
              <div style={{ marginBottom: '20px' }}>
                <label style={{ display: 'block', fontWeight: 'bold', marginBottom: '8px', color: '#333' }}>Palabra/frase correcta:</label>
                <input
                  id="correct-word"
                  type="text"
                  placeholder="Lo que el estudiante propone como correcto"
                  style={{ width: '100%', padding: '12px', fontSize: '14px', border: '1px solid #ddd', borderRadius: '4px', boxSizing: 'border-box' }}
                />
              </div>
            </>
          )}

          {feedback && (
            <div style={{ backgroundColor: '#e3f2fd', padding: '14px', borderRadius: '4px', marginBottom: '20px', borderLeft: '4px solid #0288d1' }}>
              <p style={{ margin: '0', fontSize: '13px', lineHeight: '1.5', color: '#01579b' }}>
                <strong>Análisis Claude:</strong> {feedback}
              </p>
            </div>
          )}

          <button
            onClick={handleSubmitQuestion}
            disabled={loading}
            style={{
              width: '100%',
              padding: '14px',
              backgroundColor: loading ? '#ccc' : '#1F5C3A',
              color: 'white',
              border: 'none',
              borderRadius: '4px',
              cursor: loading ? 'wait' : 'pointer',
              fontSize: '16px',
              fontWeight: 'bold',
              transition: 'background-color 0.3s'
            }}
            onMouseOver={(e) => { if (!loading) e.target.style.backgroundColor = '#153a28'; }}
            onMouseOut={(e) => { if (!loading) e.target.style.backgroundColor = '#1F5C3A'; }}
          >
            {loading ? '⏳ Analizando...' : currentQuestionIndex < currentExam.questions.length - 1 ? '→ Siguiente' : '✓ Finalizar'}
          </button>
        </div>
      </div>
    );
  }

  // ============================================================
  // RENDER: RESULTADOS
  // ============================================================
  if (stage === 'results') {
    const exam = getCurrentExam();
    const totalBruto = responses.reduce((sum, r) => sum + (r?.evaluation?.score || 0), 0);
    const maxBruto = exam.maxPuntoBruto;
    const calificacionNormalizada = (totalBruto / maxBruto) * 70;
    const calificacion100 = (calificacionNormalizada / 70) * 100;

    const correctas = responses.filter(r => r?.evaluation?.isCorrect).length;
    const incorrectas = responses.length - correctas;

    return (
      <div style={{ padding: '30px', fontFamily: 'Arial, sans-serif', maxWidth: '800px', margin: '0 auto', backgroundColor: '#f5f5f5', minHeight: '100vh' }}>
        <div style={{ backgroundColor: 'white', padding: '30px', borderRadius: '8px', boxShadow: '0 2px 8px rgba(0,0,0,0.1)' }}>
          <h2 style={{ color: '#1F5C3A', marginBottom: '25px', textAlign: 'center' }}>✅ Calificación Completada</h2>

          <div style={{ backgroundColor: '#f0f8f4', padding: '25px', borderRadius: '8px', marginBottom: '30px', border: '2px solid #1F5C3A', textAlign: 'center' }}>
            <p style={{ margin: '0 0 10px 0', color: '#666', fontSize: '14px' }}>
              {examsData[selectedTech].name} | Opción {selectedOp}
            </p>
            <p style={{ margin: '0 0 15px 0', fontSize: '18px', color: '#333' }}>
              <strong>{studentName}</strong>
            </p>
            <p style={{ margin: '0 0 10px 0', fontSize: '32px', fontWeight: 'bold', color: '#1F5C3A' }}>
              {calificacion100.toFixed(2)} / 100
            </p>
            <p style={{ margin: '0', fontSize: '12px', color: '#666' }}>
              ({totalBruto.toFixed(1)} ÷ {maxBruto} pts brutos → {calificacionNormalizada.toFixed(2)} / 70 normalizados)
            </p>
            <p style={{ margin: '15px 0 0 0', fontSize: '14px', color: '#333' }}>
              ✓ {correctas} correctas | ✗ {incorrectas} incorrectas
            </p>
          </div>

          <h3 style={{ color: '#333', marginBottom: '15px', fontSize: '16px' }}>Desglose por pregunta:</h3>
          <div style={{ maxHeight: '300px', overflowY: 'auto', marginBottom: '25px' }}>
            {responses.map((r, idx) => (
              <div
                key={idx}
                style={{
                  backgroundColor: r?.evaluation?.isCorrect ? '#e8f5e9' : '#ffebee',
                  padding: '12px',
                  marginBottom: '8px',
                  borderRadius: '4px',
                  borderLeft: `4px solid ${r?.evaluation?.isCorrect ? '#4caf50' : '#f44336'}`,
                  fontSize: '13px'
                }}
              >
                <strong style={{ color: r?.evaluation?.isCorrect ? '#2e7d32' : '#c62828' }}>
                  P{idx + 1}: {r?.evaluation?.reasoning}
                </strong>
                <span style={{ float: 'right', fontWeight: 'bold' }}>
                  {r?.evaluation?.score?.toFixed(1) || 0} / {r?.maxScore || 3} pts
                </span>
              </div>
            ))}
          </div>

          <button
            onClick={generateExcel}
            style={{
              width: '100%',
              padding: '14px',
              backgroundColor: '#F9A825',
              color: '#1F5C3A',
              border: 'none',
              borderRadius: '4px',
              cursor: 'pointer',
              fontSize: '16px',
              fontWeight: 'bold',
              marginBottom: '12px',
              transition: 'background-color 0.3s'
            }}
            onMouseOver={(e) => { e.target.style.backgroundColor = '#e89b1f'; }}
            onMouseOut={(e) => { e.target.style.backgroundColor = '#F9A825'; }}
          >
            📥 Descargar Excel
          </button>

          <button
            onClick={() => {
              setStage('select');
              setSelectedTech(null);
              setSelectedOp(null);
              setStudentName('');
              setResponses([]);
              setCurrentQuestionIndex(0);
            }}
            style={{
              width: '100%',
              padding: '14px',
              backgroundColor: '#1F5C3A',
              color: 'white',
              border: 'none',
              borderRadius: '4px',
              cursor: 'pointer',
              fontSize: '16px',
              fontWeight: 'bold',
              transition: 'background-color 0.3s'
            }}
            onMouseOver={(e) => { e.target.style.backgroundColor = '#153a28'; }}
            onMouseOut={(e) => { e.target.style.backgroundColor = '#1F5C3A'; }}
          >
            + Calificar Otro Estudiante
          </button>
        </div>
      </div>
    );
  }

  return null;
};

export default CalificadorITSE;
