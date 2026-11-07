# Neodata MCP

Servidor [MCP](https://modelcontextprotocol.io) para usar **Neodata Precios Unitarios (PU Win+)** desde Claude
(app de escritorio o Claude Code). Con él puedes pedirle a Claude que lea tus presupuestos, matrices e insumos,
que exporte a Excel, que importe conceptos desde Excel o que edite la obra. También puede consultar **Construbase**.

> Proyecto independiente, sin relación con Neodata S.A. de C.V. Úsalo bajo tu responsabilidad y
> trabaja con respaldos. Las escrituras ya hacen uno automático.

## Cómo funciona
Neodata guarda cada obra como una base SQL Server (`.mdf`) en **SQL Server LocalDB** `(localdb)\MSSQLLocalDB`.
El servidor se conecta con tu usuario de Windows, así que no usa la contraseña `sa` de Neodata. Para leer no hace falta tener Neodata abierto.

## Requisitos
- Windows con **Neodata PU Win+** instalado. Neodata ya trae LocalDB y el *ODBC Driver 17 for SQL Server*.
- **Python 3.10 o superior**.
- La app **Claude Desktop** o **Claude Code**.

## Instalación
1. Descarga este repositorio (botón **Code → Download ZIP**) y descomprímelo, por ejemplo en `C:\neodata_mcp`.
2. Instala las dependencias:
   ```
   pip install -r C:\neodata_mcp\requirements.txt
   ```
3. Registra el servidor:
   - **Claude Code:**
     ```
     claude mcp add neodata --scope user -- python -X utf8 C:\neodata_mcp\neodata_mcp.py
     ```
   - **Claude Desktop:** agrega esto en `claude_desktop_config.json` (en *Configuración → Desarrollador → Editar config*):
     ```json
     {
       "mcpServers": {
         "neodata": {
           "command": "python",
           "args": ["-X", "utf8", "C:\\neodata_mcp\\neodata_mcp.py"]
         }
       }
     }
     ```
4. Reinicia Claude o abre una sesión nueva. Pídele *"usa neodata_estado"* para comprobar que encuentra tus obras.

## Herramientas
| Leer | Archivos | Escribir (dry_run=True por omisión) |
|---|---|---|
| `neodata_estado`, `neodata_listar_obras` | `neodata_exportar_excel` | `neodata_crear_partida` |
| `neodata_resumen` | `neodata_plantilla_importacion` | `neodata_guardar_concepto` |
| `neodata_conceptos` | `neodata_importar_excel` | `neodata_guardar_insumo` |
| `neodata_matriz` (análisis de P.U.) | | `neodata_definir_matriz` |
| `neodata_insumos`, `neodata_explosion_insumos` | | `neodata_eliminar_concepto` |
| `neodata_consulta_sql` (solo SELECT) | | `neodata_respaldar`, `neodata_listar_respaldos`, `neodata_restaurar_respaldo` |

Ejemplos de lo que puedes pedir:
- *"Dame el resumen de mi obra por partida"*
- *"Muéstrame la matriz de CIM-001"*
- *"Exporta el presupuesto a Excel con factor de sobrecosto 1.28"*
- *"Busca en Construbase el precio de la plantilla de concreto de 5 cm"*
- *"Hazme la plantilla de importación"*, y después *"importa este Excel a mi obra"*

## Seguridad al escribir
- Toda escritura es **simulación** por omisión: se ejecuta, muestra el resultado y se deshace. Para guardar hay que pedirlo explícitamente.
- **Se niega a escribir si Neodata está abierto**, para que Neodata no sobrescriba los cambios.
- Antes de cada escritura real hace un **respaldo `.bak`** en `C:\NeodataWinPlusObras\respaldos_mcp`, y cualquiera se puede restaurar.
- Usa los **procedimientos almacenados de Neodata** (`NeoConceptos_GrabarActualizar`, `NeoMatrices_GrabarSoloMatriz`, etc.),
  todo en una transacción: si algo falla, no se guarda nada.
- Guarda el **costo directo**. El **precio unitario** (indirectos, financiamiento, utilidad y cargos) lo calcula
  Neodata con su libro de sobrecosto. Después de editar, abre la obra en Neodata y ejecuta **Cálculo**.

## Construbase
Descarga Construbase gratis desde [neodata.mx/descargas](https://neodata.mx/descargas) (versión *LDB*). Descomprime el
`.mdf` en `C:\Neodata2021Obras` y consúltalo como una obra más, por ejemplo con
`obra="Construbase_Octubre_2026"` y `presupuesto` igual a `CONSTRUCTOR`, `CONCURSOS`, `REMODELADOR` o `DESARROLLADOR`.
La primera consulta tarda **3–4 min**, porque la base pesa ~850 MB y se abre en frío. Las siguientes tardan ~2 s.

## Configuración opcional (variables de entorno)
| Variable | Por omisión |
|---|---|
| `NEODATA_CARPETAS` | `C:\NeodataWinPlusObras;C:\Neodata2021Obras;C:\Neodata2018Obras` |
| `NEODATA_RESPALDOS` | `C:\NeodataWinPlusObras\respaldos_mcp` |
| `NEODATA_SALIDA` | `Documentos\Neodata_MCP` (exportaciones a Excel) |
| `NEODATA_SERVIDOR` | `(localdb)\MSSQLLocalDB` |
| `NEODATA_DRIVER` | `ODBC Driver 17 for SQL Server` |

## Modelo de datos (por si quieres extenderlo)
`PuCatalogo` (códigos) → `PuExpIns` (código dentro de un presupuesto) → `PuExpInsCostos` (costo por moneda; 1 = M.N.).
`PuMatrices.IdCodigoMatriz` e `IdCodigoInsumo` apuntan a `PuExpIns`. El importe de cada renglón es costo × cantidad
(o costo ÷ cantidad si `Dividir`). Un insumo con `EsPorcentaje = <tipo>` vale un porcentaje del subtotal de ese tipo;
por ejemplo, herramienta menor = 3 % de la mano de obra. Los conceptos (`PuPresupuestosConceptos`) cuelgan de las
partidas (`PuPresupuestosPartidas`), que forman un árbol con `IdPartidaPadre`.

Probado con Neodata PU Win+ 2023.50 (estructura de base v21.1.0).
