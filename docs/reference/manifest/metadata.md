# <a name="metadata-element"></a>Elemento MetaData

Define as configurações de metadados usadas por uma função personalizada no Excel.

## <a name="attributes"></a>Atributos

Nenhum

## <a name="child-elements"></a>Elementos filho

|  Elemento  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Sim  | ID do recurso do arquivo HTML usado por funções personalizadas. |

## <a name="example"></a>Exemplo

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
