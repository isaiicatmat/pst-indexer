# Testing - PST Indexer

Guía para ejecutar y escribir tests para el proyecto PST Indexer.

## Suite de tests incluida

El proyecto incluye una suite completa de tests en `test_email_searcher.py` que cubre:

- ✅ **24 tests** que validan la funcionalidad completa
- ✅ **Edge cases** como caracteres especiales, unicode, SQL injection
- ✅ **Casos de error** y situaciones límite
- ✅ **Rendimiento** con 1000+ correos

## Ejecución rápida

```bash
# Ejecutar todos los tests
python test_email_searcher.py

# Con pytest (más detallado)
pytest test_email_searcher.py -v

# Con reporte de cobertura
pytest test_email_searcher.py --cov=email_searcher
```

## Tests cubiertos

### Funcionalidad básica

- `test_database_initialization` - BD se crea correctamente
- `test_database_schema` - Esquema de tabla correcto
- `test_add_single_email` - Agregar correo único
- `test_index_stats` - Estadísticas correctas

### Búsquedas

- `test_search_by_sender` - Búsqueda por remitente
- `test_search_by_subject` - Búsqueda por asunto
- `test_search_by_content` - Búsqueda por contenido
- `test_search_by_date_range` - Búsqueda por rango de fechas
- `test_search_combined_criteria` - Búsqueda multi-criterios
- `test_search_no_results` - Búsqueda sin resultados
- `test_search_case_insensitive` - Búsqueda sin distinguir mayúsculas
- `test_date_sorting` - Resultados ordenados por fecha DESC

### Edge Cases

- `test_duplicate_emails` - No se agregan duplicados
- `test_long_email_content` - Correos muy largos (10,000 caracteres)
- `test_null_values_handling` - Manejo de valores vacíos
- `test_max_results_limit` - Límite de 1000 resultados
- `test_unicode_content` - Caracteres unicode (中文, 日本語, etc.)
- `test_special_characters` - Caracteres especiales y 'quotes'
- `test_sql_injection_attempt` - Protección SQL injection
- `test_very_old_date` - Correos de 1990+
- `test_future_date` - Correos con fechas futuras
- `test_whitespace_only_search` - Búsqueda solo con espacios
- `test_empty_database_stats` - Stats con BD vacía

## Estructura de tests

```python
class TestEmailIndexer(unittest.TestCase):
    """Tests unitarios de la clase EmailIndexer"""
    
    def setUp(self):
        # Crear DB temporal antes de cada test
        
    def tearDown(self):
        # Limpiar después de cada test
```

## Agregar nuevos tests

Para agregar un test nuevo:

```python
def test_nueva_funcionalidad(self):
    """Test: Descripción clara de qué se prueba"""
    # Arrange - preparar datos
    conn = sqlite3.connect(self.db_path)
    c = conn.cursor()
    # ... insertar datos
    
    # Act - ejecutar
    results = self.indexer.search(...)
    
    # Assert - validar
    self.assertEqual(len(results), expected_count)
```

## Cobertura de código

Generar reporte de cobertura:

```bash
pip install pytest-cov
pytest test_email_searcher.py --cov=email_searcher --cov-report=html
```

Esto genera `htmlcov/index.html` con visualización detallada.

## CI/CD

Para integrar en GitHub Actions:

```yaml
name: Tests

on: [push, pull_request]

jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v2
      - uses: actions/setup-python@v2
        with:
          python-version: 3.9
      - run: pip install -r dev-requirements.txt
      - run: python -m pytest test_email_searcher.py -v
```

## Resultados esperados

```
Ran 24 tests in 0.066s
OK
```

Todos los tests deben pasar antes de hacer merge a main.

## Notas

- Los tests NO requieren archivos PST reales
- Cada test usa una BD temporal aislada
- Sin efectos secundarios entre tests
- Completa en < 100ms

---

Para más información sobre desarrollo, ver [DEVELOPMENT.md](DEVELOPMENT.md)
